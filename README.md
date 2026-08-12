# 订阅结账页的Playwright交互回归与证据留存

这个仓库保存订阅结账页任务的正文、四个附件和Windows端质量检查。输入材料包含本地结账页、产品场景、报价响应和浏览器策略。测试启动真实Chromium，访问本机页面，拦截报价请求并生成场景报表、trace和截图。

四个附件位于artifacts目录，任务正文位于task目录。工作流使用windows-2025和Node.js24检查附件结构、原生Windows兼容性和Chromium业务结果。

在Windows PowerShell中可执行以下命令：

    ./scripts/windows_gate.ps1 -RepositoryRoot $PWD -EvidenceRoot $env:TEMP/ale-q10043-evidence

安装依赖和Chromium时需要联网。正式测试只访问127.0.0.1，不调用外部报价服务。
