# MEMORY.md - Long-Term Memory

## User Preferences
- **对话语言:** 中文（用户要求以后所有对话内容使用中文）
- **OpenClaw 操作：** 只有 `restart` / `stop` / `start` / `run` / `install` / `uninstall` / 版本更新（`update` / `upgrade`）等重启/开关/升级类命令需要打开独立终端窗口运行（`start cmd /k "..."`），让用户看到实时输出；其他查询类命令（`status` / `health` / `probe` 等）正常执行即可
- **操作完成清理临时文件（2026-06-17 约定）：** 任何文件 / 配置 / 调试操作结束后，主动扫描并删除这次产生的中间文件。临时文件包括但不限于：base64 解码用的临时 JSON、tee 输出的诊断报告、`GetTempFileName()` 创建的空 .tmp 残留。**保留**主配置、OpenClaw 自动生成的备份、用户数据目录；**保留**自己做的回滚备份（带时间戳的 `*.bak.pre-*-YYYYMMDD-HHMMSS`，用户没明确说删时不主动删）。删除前要列清单（要删 / 要保留各是什么），删完报告数量。

---

## 工作目录：收保.NETFramework 升级

- **目录：** `D:\work\收保.NETFramework升级`
- **截止日期：** 2026年8月之前会用到的文件
- **子目录：** `coding` / `测试式样书` / `详细设计` / `进度表` 等
- **提醒：** 2026-08 之前这些文件都需要保留，不要清理/移动

### "更新svn" 约定（2026-06-16 用户约定）

**截止 2026 年 8 月之前**，用户说「更新 svn」「更新 svn 文档」一律默认指：

- 对 `D:\work\收保.NETFramework升级\` 下的**所有 SVN 仓库**（`coding` / `测试式样书` / `详细设计` / `进度表`）执行 `svn update`。
- 不用问"是 update 还是 commit"，也不用先只查单个仓库。
- 操作流程：先 `svn status -u` 看每个仓库的差距 → 列汇总给用户 → 逐个 `svn update` → 汇报结果（已最新 / 修改数 / 新增文件）。
- 8 月之后此约定作废，届时再确认新规则。

### 打开文件的命名规则（2026-06-12 用户约定）

打开文件用「**类型前缀 + 画面 ID**」的格式，避免同名画面在不同文档里冲突：

- `测试式样书 D0206` → 测试式样书
- `详设 D0206` → 详细设计
- `基设 D0206` → 基本设计
- （其他类别待补充：单体测试结果、課題管理台帳、進捗表 等）

打开同一份文件如果有多个版本，默认**最新版本**，并告知用户已开哪个版本。

### 关掉文档的顺序（2026-06-12 用户约定）

关掉文档之前**先查询当前打开了哪些**（Excel 用 COM `GetActiveObject("Excel.Application")` + `Workbooks`；Edge / 其他应用用 `Get-Process` 看 `MainWindowTitle`），列出来让用户确认要关的是哪些，**再执行关闭**。避免误关。

## D:\claw 文件夹整理规则

### 整理原则

1. **单次对话产生的单个文件** → 放入 `杂项/YYYY-MM` 目录
2. **单次对话产生多个文件** → 创建独立目录，目录名取自对话主题
3. **同一主题的多个文件簇** → 归类到对应主题目录
4. **相关文件聚集到一个目录** → 按主题/项目分组
5. **每次整理前先读取文件夹内容**，根据实际情况判断分类，不依赖自动脚本

### 目录结构（截至 2026-05-29）

```
D:\claw\
├── 2048游戏开发\          # 2048游戏开发
├── FPS游戏辅助\          # FPS游戏辅助相关
├── Roguelike\            # Roguelike游戏开发
├── 报关价格分析\           # 报关价格分析（2026-04）
├── 中职成绩扣分处理\        # 中职成绩扣分处理
├── 中职教学平台文档\        # 教师端操作手册文档
├── 成绩处理\              # 成绩导入/处理脚本
├── 每日新闻推送\            # 新闻推送脚本
├── 日语翻译插件\            # 油猴脚本+SPEC
├── 睿牙问题合集\            # 睿牙问题说明+子目录
├── 杂项\                  # 零散单文件（YYYY-MM子目录）
├── check\                # 检查脚本
├── find\                 # 查找脚本
├── gen\                  # 生成脚本
├── test\                 # 测试脚本
├── 2026年4月25级期中考\    # 期中考试相关文件
├── gateway.log            # OpenClaw日志，暂时保留
└── organize.py            # （已删除，不再使用）
```

### 清理对象
- `__pycache__` — Python缓存，适时删除
- `_docx_temp`、`_final_extracted`、`_ref_extracted` — 文档处理临时目录，完成后删除

### 命名建议
- 主题目录：用中文简明描述，如"2048游戏开发"
- 零散文件：按对话内容创建 YYYY-MM 子目录

## 数据库操作安全规则
- 对数据库只能执行**查询（SELECT）**，如需执行更新（UPDATE/INSERT/DELETE）等操作，必须先征得用户确认。

## GitHub 上传规则
- **不要随意往 GitHub 推送代码**，只有在用户**明确要求上传到 GitHub** 时才执行 push 操作。
- 主动创建本地文件、初始化 git 仓库都可以，但 push 必须等用户明确授权。

## SVN 提交规则（2026-06-12 用户约定）

- **默认不写 commit message**——除非用户明确说了要写提交信息或给出具体内容。
- 在 `D:\work\收保.NETFramework升级\coding\` 下，提交用：
  ```powershell
  svn commit --non-interactive --force-log -m "" -- <文件路径>
  ```
  - `--` 分隔选项和路径，避免文件名被当成 -m 参数
  - `--non-interactive --force-log -m ""` 三个一起用才能跳过 message 提示提交空 msg
- 提交前先 svn status / svn diff 看改了什么，让用户确认
- 外部动作（commit）需用户明确授权
- **commit 可以代做**（2026-06-15 用户更新规则，替代之前的"自己提交"）：明确说"提交 XX"时按 SVN 标准流程执行，但仍先 status 汇报。message 默认空。

## 中日翻译助手（2026-06-12 用户约定）

- **触发方式**：用户说「翻译 xxx」/「把这段翻译成日文」/「translate xxx」等，自动 spawn 一个翻译 sub-agent 处理。
- **sub-agent 任务模板**：
  ```
  你是一个中日双向翻译助手。请把下面这段 [方向] 翻译成 [目标语言]。
  
  原文：[文本]
  
  要求：
  - 只输出翻译结果，不要解释
  - 翻译要自然，符合目标语言表达习惯
  - 保留原文的换行、代码块、特殊格式
  ```
- **方向自动检测**：如果用户没指定方向，sub-agent 自己根据文本判断（中日混合时优先识别主体语言）。
- **不要本地脚本/服务**：翻译功能只在 webchat 里通过 sub-agent 实现，不另起进程。
- **质量优先**：用 LLM 翻译，不用免费机器翻译 API。

### .xlsm 大文件无法 Start-Process 打开（2026-06-15 / 2026-06-16）

- **现象**：`D:\work\收保.NETFramework升级\进度表\【収入保険.NETFramework】武漢側開発スケジュール.xlsm`（1.4MB，含 7.3MB sheet1 + 6.6MB drawing + 23KB 宏 + ActiveX 控件）用 `Start-Process` 启动时，Excel 进程存活约 2.5 秒后立即退出，文件无进程锁（FileStream 共享读/复制均成功）。
- **原因**：猜测与 `Start-Process` 启动的 Excel 上下文（shell 集成、用户配置文件加载、宏信任）有关，跟文件本身内容关系不大。
- **解决（推荐）**：**用 PowerShell COM 方式直接打开**（绕开 Start-Process 的坑，已验证 2026-06-16 成功）：
  ```powershell
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $true
  $excel.Workbooks.Open("D:\work\收保.NETFramework升级\进度表\【収入保険.NETFramework】武漢側開発スケジュール.xlsm")
  ```
  返回 Workbook 对象 → 进程稳定存活。
- **备选**：手动在资源管理器里双击打开（也能成，就是用户得多动一下手）。
- **判断标准**：如果是大文件 + 宏 + ActiveX → 直接用 COM 方式开，别试 Start-Process。

## 浏览器翻译插件项目（2026-06-12）

- **目录**：`D:\claw\msn-ja-translator-v2\`
- **用途**：MSN 日本站专用浏览器翻译插件，调用 MiniMax LLM（OpenAI 兼容 API）
- **API 端点**：`https://api.minimax.io/v1/chat/completions`
- **默认模型**：`MiniMax-M3`
- **核心文件**：
  - `msn-ja-translator-v2.user.js` — 油猴脚本主程序
  - `SPEC.md` — 设计文档
  - `README.md` — 安装/使用说明
- **v1 区别**：v1（`D:\claw\日语翻译插件\`）用 MyMemory 免费 API + hover 触发；v2 改用 MiniMax LLM + 右键菜单触发，带语法解释
- **API key**：用户在插件设置中自填，存 `GM_setValue`（本地），不硬编码

---

*最后更新: 2026-06-12*