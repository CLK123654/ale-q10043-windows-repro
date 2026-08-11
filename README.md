# 订阅结账页的Playwright交互回归与证据留存

这个仓库只保存这一道任务的正文、四个最终附件和独立运行门禁。输入材料是一套本地结账页、六条产品场景、报价响应和浏览器策略。测试会启动真实Chromium，访问本机页面，拦截报价请求并生成四份报表、六份trace和六张截图。

四个附件位于artifacts目录。任务正文位于task目录。工作流使用windows-2025和Node.js24，每次都在两个带中文和空格的新目录中运行，然后执行时间规则变化和缺失报价输入测试。运行证据由工作流保存为windows-2025-evidence。

在Windows PowerShell中可执行以下命令：

    ./scripts/windows_gate.ps1 -RepositoryRoot $PWD -EvidenceRoot $env:TEMP/ale-q10043-evidence

安装依赖和Chromium时需要联网。正式测试只访问127.0.0.1，不调用外部报价服务。
