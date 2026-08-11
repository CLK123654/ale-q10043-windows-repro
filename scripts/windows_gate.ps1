[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$RepositoryRoot,
  [Parameter(Mandatory = $true)]
  [string]$EvidenceRoot
)

$ErrorActionPreference = 'Stop'
if (Get-Variable -Name PSNativeCommandUseErrorActionPreference -ErrorAction SilentlyContinue) {
  $PSNativeCommandUseErrorActionPreference = $false
}

$ArtifactsRoot = Join-Path $RepositoryRoot 'artifacts'
$ScratchRoot = Join-Path $env:RUNNER_TEMP ('playwright-checkout-' + [guid]::NewGuid().ToString('N'))
$ReportNames = @('case_results.csv', 'network_contract.csv', 'evidence_manifest.csv', 'run_summary.json')
$RequiredSourceFiles = @('playwright.config.mjs', 'tests/subscription_checkout.spec.mjs')

New-Item -ItemType Directory -Force -Path $EvidenceRoot | Out-Null
New-Item -ItemType Directory -Force -Path $ScratchRoot | Out-Null

function Assert-True {
  param([bool]$Condition, [string]$Message)
  if (-not $Condition) {
    throw $Message
  }
}

function Get-Sha256 {
  param([string]$PathValue)
  return (Get-FileHash -Algorithm SHA256 -LiteralPath $PathValue).Hash.ToLowerInvariant()
}

function Get-RelativeFiles {
  param([string]$Root)
  return @(Get-ChildItem -LiteralPath $Root -File -Recurse |
    ForEach-Object { [System.IO.Path]::GetRelativePath($Root, $_.FullName).Replace('\', '/') } |
    Sort-Object)
}

function Get-SourceTreeHash {
  param([string]$Root)
  $lines = Get-ChildItem -LiteralPath $Root -File -Recurse |
    Where-Object {
      $relative = [System.IO.Path]::GetRelativePath($Root, $_.FullName).Replace('\', '/')
      $relative -notmatch '^(node_modules|output|\.playwright-artifacts|test-results)/'
    } |
    ForEach-Object {
      $relative = [System.IO.Path]::GetRelativePath($Root, $_.FullName).Replace('\', '/')
      $relative + [char]0 + (Get-Sha256 $_.FullName)
    } |
    Sort-Object
  $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]::Join("`n", $lines))
  return [Convert]::ToHexString([System.Security.Cryptography.SHA256]::HashData($bytes)).ToLowerInvariant()
}

function Get-ZipFileEntries {
  param([string]$ArchivePath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
  try {
    return @($archive.Entries |
      Where-Object { -not [string]::IsNullOrEmpty($_.Name) } |
      ForEach-Object { $_.FullName.Replace('\', '/') } |
      Sort-Object)
  } finally {
    $archive.Dispose()
  }
}

function Assert-SequenceEqual {
  param([object[]]$Actual, [object[]]$Expected, [string]$Label)
  $actualValues = @($Actual | ForEach-Object { [string]$_ })
  $expectedValues = @($Expected | ForEach-Object { [string]$_ })
  Assert-True ($actualValues.Count -eq $expectedValues.Count) "$Label count mismatch"
  for ($index = 0; $index -lt $expectedValues.Count; $index += 1) {
    Assert-True ($actualValues[$index] -ceq $expectedValues[$index]) "$Label mismatch at index $index"
  }
}

function Get-WorkbookSheetNames {
  param([string]$WorkbookPath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
  try {
    $entry = $archive.GetEntry('xl/workbook.xml')
    Assert-True ($null -ne $entry) 'workbook.xml missing'
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
      [xml]$xml = $reader.ReadToEnd()
    } finally {
      $reader.Dispose()
    }
    return @($xml.SelectNodes("//*[local-name()='sheet']") | ForEach-Object { $_.GetAttribute('name') })
  } finally {
    $archive.Dispose()
  }
}

function Get-WorkbookText {
  param([string]$WorkbookPath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
  $values = [System.Collections.Generic.List[string]]::new()
  $shared = @()
  try {
    $sharedEntry = $archive.GetEntry('xl/sharedStrings.xml')
    if ($null -ne $sharedEntry) {
      $reader = [System.IO.StreamReader]::new($sharedEntry.Open())
      try {
        [xml]$sharedXml = $reader.ReadToEnd()
      } finally {
        $reader.Dispose()
      }
      $shared = @($sharedXml.SelectNodes("//*[local-name()='si']") | ForEach-Object {
        [string]::Join('', @($_.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }))
      })
    }
    foreach ($entry in $archive.Entries) {
      if (-not $entry.FullName.StartsWith('xl/worksheets/') -or -not $entry.FullName.EndsWith('.xml')) {
        continue
      }
      $reader = [System.IO.StreamReader]::new($entry.Open())
      try {
        [xml]$xml = $reader.ReadToEnd()
      } finally {
        $reader.Dispose()
      }
      foreach ($cell in $xml.SelectNodes("//*[local-name()='c']")) {
        $type = $cell.GetAttribute('t')
        if ($type -eq 'inlineStr') {
          $parts = @($cell.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText })
          if ($parts.Count -gt 0) {
            $values.Add([string]::Join('', $parts))
          }
        } elseif ($type -eq 'str') {
          $node = $cell.SelectSingleNode("./*[local-name()='v']")
          if ($null -ne $node -and -not [string]::IsNullOrEmpty($node.InnerText)) {
            $values.Add($node.InnerText)
          }
        } elseif ($type -eq 's') {
          $node = $cell.SelectSingleNode("./*[local-name()='v']")
          if ($null -ne $node -and $node.InnerText -match '^\d+$') {
            $index = [int]$node.InnerText
            if ($index -lt $shared.Count) {
              $values.Add($shared[$index])
            }
          }
        }
      }
    }
  } finally {
    $archive.Dispose()
  }
  return @($values)
}

function Assert-SpecificationShape {
  param([string]$WorkbookPath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
  try {
    $entry = $archive.GetEntry('xl/worksheets/sheet1.xml')
    Assert-True ($null -ne $entry) 'specification worksheet missing'
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
      [xml]$xml = $reader.ReadToEnd()
    } finally {
      $reader.Dispose()
    }
    foreach ($cell in $xml.SelectNodes("//*[local-name()='c']")) {
      $reference = $cell.GetAttribute('r')
      Assert-True ($reference -match '^[AB]\d+$') "specification contains data outside two columns: $reference"
    }
  } finally {
    $archive.Dispose()
  }
}

function Get-JsonStrings {
  param([object]$Value)
  $items = [System.Collections.Generic.List[string]]::new()
  function Visit-Value {
    param([object]$Current)
    if ($null -eq $Current) {
      return
    }
    if ($Current -is [string]) {
      $items.Add($Current)
      return
    }
    if ($Current -is [System.Collections.IEnumerable] -and $Current -isnot [string] -and $Current -isnot [pscustomobject]) {
      foreach ($entry in $Current) {
        Visit-Value $entry
      }
      return
    }
    if ($Current -is [pscustomobject]) {
      foreach ($property in $Current.PSObject.Properties) {
        Visit-Value $property.Value
      }
    }
  }
  Visit-Value $Value
  return @($items)
}

function Assert-NaturalText {
  param([string[]]$Texts, [string]$Label)
  $quoteCharacters = @(
    [char]34, [char]39, [char]96, '“', '”', '‘', '’', '＂', '＇',
    '「', '」', '『', '』', '«', '»', '‹', '›', '〝', '〞', '〟', '《', '》', '〈', '〉'
  )
  $space = '[\t \u00a0\u1680\u2000-\u200a\u202f\u205f\u3000]+'
  $han = '[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]'
  $boundary = "(?:$han$space[A-Za-z0-9]|[A-Za-z0-9]$space$han|[A-Za-z]$space[0-9]|[0-9]$space[A-Za-z])"
  $riskTerms = @('此外', '至关重要', '深入探讨', '彰显', '赋能', '无缝', '不断演变的格局', '不仅', '不只是', '值得注意的是', '专家认为', '行业报告显示', '观察者指出', '未来展望', '挑战与未来', '——')
  $processTerms = @(
    '制题返修', '去AI', '修改题目', '规则调整', 'Windows复现', 'GitHub Actions',
    '双干净目录', '动态变化', '负例', '附件哈希', '飞书回读',
    ('record' + '_id'), ('file' + '_token')
  )
  foreach ($text in $Texts) {
    foreach ($character in $quoteCharacters) {
      Assert-True (-not $text.Contains([string]$character)) "$Label contains a forbidden quote"
    }
    Assert-True (-not [regex]::IsMatch($text, $boundary)) "$Label contains a mixed boundary space"
    foreach ($term in ($riskTerms + $processTerms)) {
      Assert-True (-not $text.Contains($term)) "$Label contains forbidden term $term"
    }
  }
}

function Assert-NoPublicMetadata {
  $sensitiveTerms = @(
    ('record' + '_id'), ('file' + '_token'), ('app' + '_token'), ('table' + '_id'),
    ('tmp' + '_url'), ('open.feishu.cn/' + 'open-apis/drive')
  )
  $textExtensions = @('.txt', '.md', '.json', '.mjs', '.js', '.ps1', '.yml', '.yaml', '.csv', '.html')
  foreach ($file in Get-ChildItem -LiteralPath $RepositoryRoot -File -Recurse) {
    if ($file.FullName.Contains("$([System.IO.Path]::DirectorySeparatorChar).git$([System.IO.Path]::DirectorySeparatorChar)")) {
      continue
    }
    if (-not $textExtensions.Contains($file.Extension.ToLowerInvariant())) {
      continue
    }
    $text = Get-Content -LiteralPath $file.FullName -Raw
    foreach ($term in $sensitiveTerms) {
      Assert-True (-not $text.Contains($term)) "public metadata found in $($file.Name)"
    }
  }
}

function Assert-NoLinuxArtifacts {
  param([string]$Root)
  $bannedExtensions = @('.sh', '.bash', '.zsh', '.so', '.elf', '.deb', '.rpm', '.appimage')
  foreach ($file in Get-ChildItem -LiteralPath $Root -File -Recurse) {
    Assert-True (-not $bannedExtensions.Contains($file.Extension.ToLowerInvariant())) "Linux or shell artifact found: $($file.Name)"
    $stream = [System.IO.File]::OpenRead($file.FullName)
    try {
      if ($stream.Length -ge 4) {
        $bytes = [byte[]]::new(4)
        [void]$stream.Read($bytes, 0, 4)
        $isElf = $bytes[0] -eq 0x7f -and $bytes[1] -eq 0x45 -and $bytes[2] -eq 0x4c -and $bytes[3] -eq 0x46
        Assert-True (-not $isElf) "ELF binary found: $($file.Name)"
      }
    } finally {
      $stream.Dispose()
    }
  }
  $lockPath = Join-Path $Root 'input_data/package-lock.json'
  $lock = Get-Content -LiteralPath $lockPath -Raw | ConvertFrom-Json -AsHashtable
  foreach ($property in $lock['packages'].GetEnumerator()) {
    $package = $property.Value
    if (-not $package.ContainsKey('os')) {
      continue
    }
    $systems = @($package['os'] | ForEach-Object { [string]$_ })
    $optional = $package.ContainsKey('optional') -and $package['optional'] -eq $true
    Assert-True ($optional -or $systems.Contains('win32')) "required platform-specific dependency found: $($property.Key)"
  }
}

function Assert-ZipNoLinuxArtifacts {
  param([string]$ArchivePath)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $bannedExtensions = @('.sh', '.bash', '.zsh', '.so', '.elf', '.deb', '.rpm', '.appimage')
  $archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
  try {
    foreach ($entry in $archive.Entries) {
      if ([string]::IsNullOrEmpty($entry.Name)) {
        continue
      }
      $extension = [System.IO.Path]::GetExtension($entry.Name).ToLowerInvariant()
      Assert-True (-not $bannedExtensions.Contains($extension)) "Linux artifact found in $($entry.FullName)"
      $stream = $entry.Open()
      try {
        if ($entry.Length -ge 4) {
          $bytes = [byte[]]::new(4)
          [void]$stream.Read($bytes, 0, 4)
          $isElf = $bytes[0] -eq 0x7f -and $bytes[1] -eq 0x45 -and $bytes[2] -eq 0x4c -and $bytes[3] -eq 0x46
          Assert-True (-not $isElf) "ELF binary found in $($entry.FullName)"
        }
      } finally {
        $stream.Dispose()
      }
    }
  } finally {
    $archive.Dispose()
  }
}

function Expand-TaskWorkspace {
  param([string]$Name)
  $workspace = Join-Path $ScratchRoot $Name
  New-Item -ItemType Directory -Force -Path $workspace | Out-Null
  Expand-Archive -LiteralPath (Join-Path $ArtifactsRoot '输入数据包.zip') -DestinationPath $workspace
  $inputRoot = Join-Path $workspace 'input_data'
  Expand-Archive -LiteralPath (Join-Path $ArtifactsRoot 'reference.zip') -DestinationPath $inputRoot
  $expectedRoot = Join-Path $workspace 'expected_output'
  Copy-Item -LiteralPath (Join-Path $inputRoot 'output') -Destination $expectedRoot -Recurse
  return [ordered]@{
    name = $Name
    workspace = $workspace
    input_root = $inputRoot
    output_root = Join-Path $inputRoot 'output'
    expected_root = $expectedRoot
  }
}

function Invoke-NativeCapture {
  param([string]$File, [string[]]$Arguments, [string]$WorkingDirectory, [string]$LogName)
  Push-Location $WorkingDirectory
  try {
    $lines = @(& $File @Arguments 2>&1 | ForEach-Object { $_.ToString() })
    $status = $LASTEXITCODE
  } finally {
    Pop-Location
  }
  $logPath = Join-Path $EvidenceRoot "$LogName.log"
  Set-Content -LiteralPath $logPath -Value $lines -Encoding utf8NoBOM
  return [ordered]@{
    command = "$File $([string]::Join(' ', $Arguments))"
    exit_code = $status
    log_file = [System.IO.Path]::GetRelativePath($EvidenceRoot, $logPath).Replace('\', '/')
    output_tail = [string]::Join("`n", @($lines | Select-Object -Last 30))
  }
}

function Invoke-Task {
  param([object]$Prepared, [string]$LogPrefix)
  $install = Invoke-NativeCapture 'npm.cmd' @('ci') $Prepared.input_root "$LogPrefix-npm-ci"
  Assert-True ($install.exit_code -eq 0) "$LogPrefix npm ci failed"
  $browser = Invoke-NativeCapture 'npx.cmd' @('playwright', 'install', 'chromium') $Prepared.input_root "$LogPrefix-browser-install"
  Assert-True ($browser.exit_code -eq 0) "$LogPrefix browser install failed"
  $task = Invoke-NativeCapture 'npm.cmd' @('run', 'test:e2e') $Prepared.input_root "$LogPrefix-task"
  return [ordered]@{ install = $install; browser_install = $browser; task = $task }
}

function Get-ReportSemantics {
  param([string]$OutputRoot)
  $reportRoot = Join-Path $OutputRoot 'reports'
  return [ordered]@{
    case_results = @(Import-Csv -LiteralPath (Join-Path $reportRoot 'case_results.csv') | Sort-Object case_id | ForEach-Object { $_ | ConvertTo-Json -Compress })
    network_contract = @(Import-Csv -LiteralPath (Join-Path $reportRoot 'network_contract.csv') | Sort-Object case_id | ForEach-Object { $_ | ConvertTo-Json -Compress })
    evidence_manifest = @(Import-Csv -LiteralPath (Join-Path $reportRoot 'evidence_manifest.csv') | Sort-Object case_id | ForEach-Object { $_ | ConvertTo-Json -Compress })
    run_summary = (Get-Content -LiteralPath (Join-Path $reportRoot 'run_summary.json') -Raw | ConvertFrom-Json | ConvertTo-Json -Compress -Depth 50)
  }
}

function Assert-ReportSemantics {
  param([object]$Actual, [object]$Expected, [string]$Label)
  foreach ($key in @('case_results', 'network_contract', 'evidence_manifest')) {
    Assert-SequenceEqual $Actual[$key] $Expected[$key] "$Label $key"
  }
  Assert-True ($Actual.run_summary -eq $Expected.run_summary) "$Label run_summary mismatch"
}

function Get-CsvRowsWithoutColumn {
  param([string]$CsvPath, [string]$ExcludedColumn)
  $rows = [System.Collections.Generic.List[string]]::new()
  foreach ($row in Import-Csv -LiteralPath $CsvPath | Sort-Object case_id) {
    $copy = [ordered]@{}
    foreach ($property in $row.PSObject.Properties) {
      if ($property.Name -ne $ExcludedColumn) {
        $copy[$property.Name] = $property.Value
      }
    }
    $rows.Add(($copy | ConvertTo-Json -Compress))
  }
  return @($rows)
}

function Get-SummaryWithoutTime {
  param([string]$JsonPath)
  $value = Get-Content -LiteralPath $JsonPath -Raw | ConvertFrom-Json
  $value.PSObject.Properties.Remove('fixed_browser_time')
  return ($value | ConvertTo-Json -Compress -Depth 50)
}

function Get-PngInfo {
  param([string]$PathValue)
  $bytes = [System.IO.File]::ReadAllBytes($PathValue)
  Assert-True ($bytes.Length -ge 24) "PNG is truncated: $PathValue"
  $expected = [byte[]](0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a)
  for ($index = 0; $index -lt 8; $index += 1) {
    Assert-True ($bytes[$index] -eq $expected[$index]) "PNG signature mismatch: $PathValue"
  }
  $width = [System.Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($bytes, 16))
  $height = [System.Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($bytes, 20))
  Assert-True ($width -gt 0 -and $height -gt 0) "PNG dimensions invalid: $PathValue"
  return [ordered]@{ bytes = $bytes.Length; width = $width; height = $height; sha256 = Get-Sha256 $PathValue }
}

function Get-TraceInfo {
  param([string]$PathValue)
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [System.IO.Compression.ZipFile]::OpenRead($PathValue)
  try {
    $members = @($archive.Entries | Where-Object { -not [string]::IsNullOrEmpty($_.Name) } | ForEach-Object { $_.FullName })
    Assert-True ($members.Contains('trace.trace')) "trace.trace missing: $PathValue"
    Assert-True ($members.Contains('trace.network')) "trace.network missing: $PathValue"
    $networkEntry = $archive.GetEntry('trace.network')
    $reader = [System.IO.StreamReader]::new($networkEntry.Open())
    try {
      $networkText = $reader.ReadToEnd()
    } finally {
      $reader.Dispose()
    }
    $urls = @([regex]::Matches($networkText, '"url":"(https?://[^"\s]+)"') | ForEach-Object { $_.Groups[1].Value })
    Assert-True ($urls.Count -gt 0) "trace contains no network URL: $PathValue"
    foreach ($url in $urls) {
      $uri = [Uri]$url
      Assert-True ($uri.Host -eq '127.0.0.1' -and $uri.Port -eq 4173) "external request found in trace: $url"
    }
    $browserMatch = [regex]::Match($networkText, '(?:HeadlessChrome|Chrome)/([0-9.]+)')
    return [ordered]@{
      bytes = (Get-Item -LiteralPath $PathValue).Length
      members = $members.Count
      loopback_url_count = $urls.Count
      browser_version = if ($browserMatch.Success) { $browserMatch.Groups[1].Value } else { '' }
      sha256 = Get-Sha256 $PathValue
    }
  } finally {
    $archive.Dispose()
  }
}

function Test-EvidenceFiles {
  param([object]$Prepared)
  $manifestPath = Join-Path $Prepared.output_root 'reports/evidence_manifest.csv'
  $rows = @(Import-Csv -LiteralPath $manifestPath | Sort-Object case_id)
  Assert-True ($rows.Count -eq 6) 'evidence manifest does not contain six cases'
  $traces = [ordered]@{}
  $screenshots = [ordered]@{}
  foreach ($row in $rows) {
    Assert-True ($row.trace_file.StartsWith('output/evidence/traces/')) "invalid trace path for $($row.case_id)"
    Assert-True ($row.screenshot_file.StartsWith('output/evidence/screenshots/')) "invalid screenshot path for $($row.case_id)"
    $tracePath = Join-Path $Prepared.input_root $row.trace_file
    $screenshotPath = Join-Path $Prepared.input_root $row.screenshot_file
    Assert-True (Test-Path -LiteralPath $tracePath -PathType Leaf) "trace missing for $($row.case_id)"
    Assert-True (Test-Path -LiteralPath $screenshotPath -PathType Leaf) "screenshot missing for $($row.case_id)"
    Assert-ZipNoLinuxArtifacts $tracePath
    $traces[$row.case_id] = Get-TraceInfo $tracePath
    $screenshots[$row.case_id] = Get-PngInfo $screenshotPath
  }
  return [ordered]@{ traces = $traces; screenshots = $screenshots }
}

function Invoke-CleanRun {
  param([string]$Name, [string]$LogPrefix)
  $prepared = Expand-TaskWorkspace $Name
  $before = Get-SourceTreeHash $prepared.input_root
  $expected = Get-ReportSemantics $prepared.expected_root
  $commands = Invoke-Task $prepared $LogPrefix
  Assert-True ($commands.task.exit_code -eq 0) "$Name task failed"
  $after = Get-SourceTreeHash $prepared.input_root
  Assert-True ($before -eq $after) "$Name changed source inputs"
  $actual = Get-ReportSemantics $prepared.output_root
  Assert-ReportSemantics $actual $expected $Name
  foreach ($relative in $RequiredSourceFiles) {
    Assert-True ((Get-Sha256 (Join-Path $prepared.output_root $relative)) -eq (Get-Sha256 (Join-Path $prepared.expected_root $relative))) "$Name changed $relative"
  }
  Assert-SequenceEqual (Get-RelativeFiles $prepared.output_root) (Get-RelativeFiles $prepared.expected_root) "$Name output members"
  $files = Test-EvidenceFiles $prepared
  $hashes = [ordered]@{}
  foreach ($report in $ReportNames) {
    $hashes[$report] = Get-Sha256 (Join-Path $prepared.output_root "reports/$report")
  }
  return [ordered]@{
    directory_name = $Name
    input_tree_sha256_before = $before
    input_tree_sha256_after = $after
    report_sha256 = $hashes
    semantics = $actual
    evidence_files = $files
    commands = $commands
    exit_code = $commands.task.exit_code
  }
}

function Invoke-TimeMutation {
  $prepared = Expand-TaskWorkspace '时间规则变化 中文 空格'
  $policyPath = Join-Path $prepared.input_root 'policy/test_contract.json'
  $policy = Get-Content -LiteralPath $policyPath -Raw | ConvertFrom-Json
  Assert-True ($policy.fixed_browser_time -eq '2026-07-30T09:30:00Z') 'baseline browser time mismatch'
  $policy.fixed_browser_time = '2026-07-30T10:00:00Z'
  Set-Content -LiteralPath $policyPath -Value ($policy | ConvertTo-Json -Depth 20) -Encoding utf8NoBOM
  $before = Get-SourceTreeHash $prepared.input_root
  $commands = Invoke-Task $prepared 'mutation'
  Assert-True ($commands.task.exit_code -eq 0) 'time mutation failed'
  $after = Get-SourceTreeHash $prepared.input_root
  Assert-True ($before -eq $after) 'time mutation run changed source inputs'
  $actualCasePath = Join-Path $prepared.output_root 'reports/case_results.csv'
  $expectedCasePath = Join-Path $prepared.expected_root 'reports/case_results.csv'
  $actualRows = @(Import-Csv -LiteralPath $actualCasePath | Sort-Object case_id)
  Assert-True ($actualRows.Count -eq 6) 'time mutation case count mismatch'
  foreach ($row in $actualRows) {
    Assert-True ($row.browser_time -eq '2026-07-30T10:00:00.000Z') "time mutation missing for $($row.case_id)"
  }
  Assert-SequenceEqual (Get-CsvRowsWithoutColumn $actualCasePath 'browser_time') (Get-CsvRowsWithoutColumn $expectedCasePath 'browser_time') 'time mutation case outcomes'
  $actualNetwork = @(Import-Csv -LiteralPath (Join-Path $prepared.output_root 'reports/network_contract.csv') | Sort-Object case_id | ForEach-Object { $_ | ConvertTo-Json -Compress })
  $expectedNetwork = @(Import-Csv -LiteralPath (Join-Path $prepared.expected_root 'reports/network_contract.csv') | Sort-Object case_id | ForEach-Object { $_ | ConvertTo-Json -Compress })
  Assert-SequenceEqual $actualNetwork $expectedNetwork 'time mutation network contract'
  $actualManifest = @(Import-Csv -LiteralPath (Join-Path $prepared.output_root 'reports/evidence_manifest.csv') | Sort-Object case_id | ForEach-Object { $_ | ConvertTo-Json -Compress })
  $expectedManifest = @(Import-Csv -LiteralPath (Join-Path $prepared.expected_root 'reports/evidence_manifest.csv') | Sort-Object case_id | ForEach-Object { $_ | ConvertTo-Json -Compress })
  Assert-SequenceEqual $actualManifest $expectedManifest 'time mutation evidence manifest'
  Assert-True ((Get-SummaryWithoutTime (Join-Path $prepared.output_root 'reports/run_summary.json')) -eq (Get-SummaryWithoutTime (Join-Path $prepared.expected_root 'reports/run_summary.json'))) 'time mutation changed summary controls'
  $files = Test-EvidenceFiles $prepared
  return [ordered]@{
    changed_rule = 'fixed_browser_time'
    before = '2026-07-30T09:30:00Z'
    after = '2026-07-30T10:00:00Z'
    observed_browser_times = @('2026-07-30T10:00:00.000Z')
    outcome_semantics_unchanged = $true
    evidence_files = $files
    commands = $commands
    exit_code = $commands.task.exit_code
    pass = $true
  }
}

function Invoke-NegativeRun {
  $prepared = Expand-TaskWorkspace '缺失报价输入 中文 空格'
  Remove-Item -LiteralPath (Join-Path $prepared.input_root 'fixtures/quote_responses.json')
  $commands = Invoke-Task $prepared 'negative'
  Assert-True ($commands.task.exit_code -ne 0) 'missing fixture returned success'
  $reportRoot = Join-Path $prepared.output_root 'reports'
  $evidencePath = Join-Path $prepared.output_root 'evidence'
  $residualReports = if (Test-Path -LiteralPath $reportRoot) { @(Get-RelativeFiles $reportRoot) } else { @() }
  $residualEvidence = if (Test-Path -LiteralPath $evidencePath) { @(Get-RelativeFiles $evidencePath) } else { @() }
  Assert-True ($residualReports.Count -eq 0) 'negative run left reports'
  Assert-True ($residualEvidence.Count -eq 0) 'negative run left evidence'
  return [ordered]@{
    removed_input = 'fixtures/quote_responses.json'
    exit_code = $commands.task.exit_code
    residual_reports = $residualReports
    residual_evidence = $residualEvidence
    commands = $commands
    pass = $true
  }
}

function Write-Evidence {
  param([object]$Evidence)
  $Evidence | ConvertTo-Json -Depth 50 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'evidence.json') -Encoding utf8NoBOM
}

$evidence = [ordered]@{
  task_asset_id = 'playwright_subscription_checkout_acceptance'
  runner_label = 'windows-2025'
  runner_image = $env:ImageOS
  runner_os = $env:RUNNER_OS
  commit_sha = $env:GITHUB_SHA
  workflow_run_id = $env:GITHUB_RUN_ID
  generated_at_utc = [DateTime]::UtcNow.ToString('o')
  pass = $false
}

try {
  Assert-True ($env:RUNNER_OS -eq 'Windows') 'workflow is not running on Windows'
  $os = Get-CimInstance Win32_OperatingSystem
  $evidence.operating_system = [ordered]@{ caption = $os.Caption; version = $os.Version; build = $os.BuildNumber }

  $nodeVersion = (& node --version).Trim()
  Assert-True ($LASTEXITCODE -eq 0 -and $nodeVersion -match '^v24\.') "unexpected Node.js version: $nodeVersion"
  $npmVersion = (& npm.cmd --version).Trim()
  Assert-True ($LASTEXITCODE -eq 0) 'npm version command failed'
  $evidence.software_versions = [ordered]@{ node = $nodeVersion; npm = $npmVersion; playwright = '1.62.1' }

  $manifest = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'manifest.json') -Raw | ConvertFrom-Json
  Assert-True ($manifest.task_asset_id -eq 'playwright_subscription_checkout_acceptance') 'task asset id mismatch'
  Assert-True ($manifest.task_asset_id -match '^[a-z][a-z0-9]*(?:_[a-z0-9]+)+$') 'task asset id is not lower snake case'
  Assert-True ($manifest.playwright_version -eq '1.62.1') 'Playwright version mismatch'

  $attachmentHashes = [ordered]@{}
  foreach ($name in @('输入数据包.zip', 'reference.zip', '关键标准答案.xlsx', '任务规格转化.xlsx')) {
    $path = Join-Path $ArtifactsRoot $name
    Assert-True (Test-Path -LiteralPath $path -PathType Leaf) "attachment missing: $name"
    $actualHash = Get-Sha256 $path
    Assert-True ($actualHash -eq [string]$manifest.attachments.$name) "attachment hash mismatch: $name"
    $attachmentHashes[$name] = $actualHash
  }
  $evidence.attachment_sha256 = $attachmentHashes

  $inputMembers = Get-ZipFileEntries (Join-Path $ArtifactsRoot '输入数据包.zip')
  $expectedInputMembers = @(
    'input_data/README.md',
    'input_data/app/checkout.html',
    'input_data/app/checkout.js',
    'input_data/app/styles.css',
    'input_data/cases/subscription_checkout_cases.csv',
    'input_data/fixtures/quote_responses.json',
    'input_data/package-lock.json',
    'input_data/package.json',
    'input_data/policy/test_contract.json',
    'input_data/starter/playwright.config.mjs',
    'input_data/starter/tests/subscription_checkout.spec.mjs',
    'input_data/tools/run-task.mjs',
    'input_data/tools/static-server.mjs'
  ) | Sort-Object
  Assert-SequenceEqual $inputMembers $expectedInputMembers 'input archive members'

  $referenceMembers = Get-ZipFileEntries (Join-Path $ArtifactsRoot 'reference.zip')
  $expectedReferenceMembers = [System.Collections.Generic.List[string]]::new()
  $expectedReferenceMembers.Add('output/playwright.config.mjs')
  $expectedReferenceMembers.Add('output/tests/subscription_checkout.spec.mjs')
  foreach ($name in $ReportNames) { $expectedReferenceMembers.Add("output/reports/$name") }
  foreach ($caseId in @('P01_PRO_MONTHLY_COUPON', 'P02_TEAM_EU_VAT', 'P03_STUDENT_AGE_LOCK', 'P04_STARTER_PAYPAL_REGION', 'P05_ENTERPRISE_APPROVAL', 'P06_PRO_RETRY_STALE')) {
    $expectedReferenceMembers.Add("output/evidence/traces/$caseId.zip")
    $expectedReferenceMembers.Add("output/evidence/screenshots/$caseId.png")
  }
  Assert-SequenceEqual $referenceMembers @($expectedReferenceMembers | Sort-Object) 'reference archive members'
  $evidence.archive_members = [ordered]@{ input = $inputMembers; reference = $referenceMembers }

  $answerSheets = Get-WorkbookSheetNames (Join-Path $ArtifactsRoot '关键标准答案.xlsx')
  Assert-SequenceEqual $answerSheets @($manifest.answer_sheets) 'answer workbook sheets'
  $specificationSheets = Get-WorkbookSheetNames (Join-Path $ArtifactsRoot '任务规格转化.xlsx')
  Assert-SequenceEqual $specificationSheets @($manifest.specification_sheets) 'specification workbook sheets'
  Assert-SpecificationShape (Join-Path $ArtifactsRoot '任务规格转化.xlsx')
  $specificationText = Get-WorkbookText (Join-Path $ArtifactsRoot '任务规格转化.xlsx')
  Assert-True ($specificationText.Contains('playwright_subscription_checkout_acceptance')) 'task asset id missing from specification'
  foreach ($forbidden in @('学科', '难度', '任务名称', '任务概要', '预计工时', '经济价值', 'Windows验证过程', ('线上题目' + 'ID'))) {
    Assert-True (-not $specificationText.Contains($forbidden)) "forbidden specification field found: $forbidden"
  }
  $evidence.workbook_sheets = [ordered]@{ answer = @($answerSheets); specification = @($specificationSheets); specification_columns = 2 }

  $staticReview = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'qa/static_text_gate.json') -Raw | ConvertFrom-Json
  $humanizer = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'qa/humanizer_review.json') -Raw | ConvertFrom-Json
  $coverage = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'qa/requirement_score_coverage.json') -Raw | ConvertFrom-Json
  $similarity = Get-Content -LiteralPath (Join-Path $RepositoryRoot 'qa/cross_task_similarity.json') -Raw | ConvertFrom-Json
  Assert-True ($staticReview.pass -eq $true -and $staticReview.violations.Count -eq 0 -and $staticReview.humanizer_risk_hits.Count -eq 0) 'static language review failed'
  Assert-True ($humanizer.pass -eq $true -and $humanizer.score -ge 45 -and $humanizer.online_ai_detection_used_as_conclusion -eq $false) 'humanizer review failed'
  $dimensionNames = @($humanizer.dimensions.PSObject.Properties.Name | Sort-Object)
  Assert-SequenceEqual $dimensionNames @('authenticity', 'directness', 'refinement', 'rhythm', 'trust') 'humanizer dimensions'
  $dimensionTotal = ($humanizer.dimensions.PSObject.Properties.Value | Measure-Object -Sum).Sum
  Assert-True ($dimensionTotal -eq $humanizer.score) 'humanizer score does not match five dimensions'
  Assert-True ($coverage.pass -eq $true -and $coverage.unscored_public_requirements.Count -eq 0 -and $coverage.score_items_without_public_requirement.Count -eq 0) 'requirement coverage failed'
  Assert-True ($similarity.pass -eq $true -and $similarity.comparison_count -eq 217 -and $similarity.threshold_hits.Count -eq 0) 'cross-task similarity gate failed'

  Assert-NoPublicMetadata
  $scan = Expand-TaskWorkspace '文本与兼容扫描 中文 空格'
  Assert-NoLinuxArtifacts $scan.workspace
  Assert-ZipNoLinuxArtifacts (Join-Path $ArtifactsRoot '输入数据包.zip')
  Assert-ZipNoLinuxArtifacts (Join-Path $ArtifactsRoot 'reference.zip')

  $naturalTexts = [System.Collections.Generic.List[string]]::new()
  foreach ($file in Get-ChildItem -LiteralPath (Join-Path $RepositoryRoot 'task') -File) {
    $naturalTexts.Add((Get-Content -LiteralPath $file.FullName -Raw))
  }
  foreach ($name in @('README.md', 'SOURCES.md')) {
    $naturalTexts.Add((Get-Content -LiteralPath (Join-Path $RepositoryRoot $name) -Raw))
  }
  foreach ($text in Get-WorkbookText (Join-Path $ArtifactsRoot '关键标准答案.xlsx')) { $naturalTexts.Add($text) }
  foreach ($text in $specificationText) { $naturalTexts.Add($text) }
  $naturalTexts.Add((Get-Content -LiteralPath (Join-Path $scan.input_root 'README.md') -Raw))
  foreach ($row in Import-Csv -LiteralPath (Join-Path $scan.input_root 'cases/subscription_checkout_cases.csv')) {
    foreach ($property in $row.PSObject.Properties) {
      if (-not [string]::IsNullOrEmpty([string]$property.Value)) { $naturalTexts.Add([string]$property.Value) }
    }
  }
  $fixture = Get-Content -LiteralPath (Join-Path $scan.input_root 'fixtures/quote_responses.json') -Raw | ConvertFrom-Json
  foreach ($text in Get-JsonStrings $fixture) { $naturalTexts.Add($text) }
  $html = Get-Content -LiteralPath (Join-Path $scan.input_root 'app/checkout.html') -Raw
  $visible = [regex]::Replace($html, '(?is)<script\b[^>]*>.*?</script>', ' ')
  $visible = [regex]::Replace($visible, '(?is)<style\b[^>]*>.*?</style>', ' ')
  $visible = [regex]::Replace($visible, '(?s)<[^>]+>', ' ')
  $visible = [System.Net.WebUtility]::HtmlDecode($visible)
  $naturalTexts.Add($visible)
  Assert-NaturalText @($naturalTexts) 'candidate-facing text'
  $evidence.natural_language = [ordered]@{
    humanizer_method = $humanizer.review_method
    humanizer_score = $humanizer.score
    online_detector_used_as_conclusion = $false
    scanned_sources = $naturalTexts.Count
    static_zero_hit = $true
    pass = $true
  }
  $evidence.compatibility_scan = [ordered]@{ linux_or_shell_artifacts = 0; required_platform_specific_dependencies = 0; pass = $true }

  $first = Invoke-CleanRun '第一轮 中文 空格' 'clean-one'
  $second = Invoke-CleanRun '第二轮 中文 空格' 'clean-two'
  Assert-ReportSemantics $first.semantics $second.semantics 'two clean runs'
  $evidence.clean_runs = @($first, $second)
  $evidence.mutation = Invoke-TimeMutation
  $evidence.negative = Invoke-NegativeRun
  $browserVersions = @($first.evidence_files.traces.Values | ForEach-Object { $_.browser_version } | Where-Object { $_ } | Sort-Object -Unique)
  Assert-True ($browserVersions.Count -eq 1) 'browser version was not captured consistently'
  $evidence.software_versions.chromium = $browserVersions[0]
  $evidence.pass = $true
  Write-Evidence $evidence
} catch {
  $evidence.error = $_.Exception.Message
  $evidence.pass = $false
  Write-Evidence $evidence
  throw
} finally {
  if (Test-Path -LiteralPath $ScratchRoot) {
    Remove-Item -LiteralPath $ScratchRoot -Recurse -Force
  }
}
