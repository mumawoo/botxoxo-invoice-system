# Installation and Usage Guide / 安装使用说明

This guide explains how to install, configure, and run Botxoxo Invoice System on Windows.

本说明用于在 Windows 上安装、配置和运行 Botxoxo Invoice System。

## 1. Download / 下载

English:

1. Open the GitHub repository.
2. Click **Code -> Download ZIP**, or download the v1.0 release package.
3. Unzip the folder, for example to `C:\Users\<you>\botxoxo-invoice-system`.

中文：

1. 打开 GitHub 项目页面。
2. 点击 **Code -> Download ZIP**，或者下载 v1.0 release 包。
3. 解压到本地目录，例如 `C:\Users\<you>\botxoxo-invoice-system`。

## 2. Install Python Dependencies / 安装 Python 依赖

English:

Open PowerShell in the project folder and run:

```powershell
python -m pip install -r requirements.txt
```

中文：

在项目目录打开 PowerShell，运行：

```powershell
python -m pip install -r requirements.txt
```

If this computer uses a special Python path, set `INVOICE_SYSTEM_PYTHON` before starting the Workbench.

如果电脑使用特殊 Python 路径，可以先设置 `INVOICE_SYSTEM_PYTHON`，Workbench 会优先使用它。

## 3. Create `.env` / 创建 `.env`

English:

```powershell
Copy-Item .env.example .env
notepad .env
```

Fill in your own values:

```env
TELEGRAM_BOT_TOKEN=
TELEGRAM_ALLOWED_USER_IDS=
TELEGRAM_AUTO_PROCESS=true
TELEGRAM_LANGUAGE=en
COMPANY_PROFILE=default

QWEN_API_KEY=
QWEN_MODEL=qwen-vl-max
QWEN_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions
ENABLE_QWEN_SCAN=true
```

中文：

```powershell
Copy-Item .env.example .env
notepad .env
```

填入你自己的配置：

```env
TELEGRAM_BOT_TOKEN=
TELEGRAM_ALLOWED_USER_IDS=
TELEGRAM_AUTO_PROCESS=true
TELEGRAM_LANGUAGE=zh
COMPANY_PROFILE=default

QWEN_API_KEY=
QWEN_MODEL=qwen-vl-max
QWEN_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions
ENABLE_QWEN_SCAN=true
```

Never commit `.env`; it contains private keys and Telegram tokens.

不要把 `.env` 提交到 Git；它包含私人的 key 和 Telegram token。

## 4. Start the Workbench / 启动 Workbench

English:

The Workbench startup file is:

```text
scripts\workbench.cmd
```

Double-click it in File Explorer, or run:

```powershell
.\scripts\workbench.cmd
```

It launches:

```text
scripts\workbench.ps1
```

中文：

Workbench 启动文件是：

```text
scripts\workbench.cmd
```

可以在资源管理器里双击，也可以在 PowerShell 运行：

```powershell
.\scripts\workbench.cmd
```

它实际会启动：

```text
scripts\workbench.ps1
```

## 5. Telegram Setup / Telegram 设置

English:

1. Create a bot with Telegram BotFather.
2. Put the token into `TELEGRAM_BOT_TOKEN`.
3. Start the bot in save-only mode first.
4. Send `/whoami` to the bot.
5. Put the returned numeric ID into `TELEGRAM_ALLOWED_USER_IDS`.
6. Restart the bot with auto scan enabled.

Useful check:

```powershell
python -m invoice_system telegram --check
```

中文：

1. 用 Telegram BotFather 创建 bot。
2. 把 token 填到 `TELEGRAM_BOT_TOKEN`。
3. 先用 save-only 模式启动 bot。
4. 给 bot 发送 `/whoami`。
5. 把返回的数字 ID 填到 `TELEGRAM_ALLOWED_USER_IDS`。
6. 重新启动 bot，并开启自动扫描。

检查命令：

```powershell
python -m invoice_system telegram --check
```

## 6. Daily Use / 日常使用

English:

1. Send receipt photos to the Telegram bot.
2. The bot saves the photo and scans it.
3. Review the manual Excel with `/excel`.
4. Correct rows with `/change` or mark bad rows with `/del`.
5. Use `/checked` only when you need the finance output.
6. Use `/submit`, review the preview, then send `confirm` to archive the batch.

中文：

1. 把发票照片发给 Telegram bot。
2. bot 保存照片并扫描。
3. 用 `/excel` 查看人工复核 Excel。
4. 用 `/change` 修改行，或用 `/del` 标记错误行。
5. 只有需要财务版时才用 `/checked`。
6. 用 `/submit` 查看提交预览，确认后发送 `confirm` 归档本批次。

## 7. Common Telegram Commands / 常用 Telegram 命令

```text
/status                 Queue and bot status / 队列和 bot 状态
/excel                  Send manual workbook / 发送人工复核 Excel
/report                 Current reimbursement summary / 当前报销汇总
/change 021 type Gas    Change category / 修改类别
/change 021 2026-07-01  Change date / 修改日期
/change 021 + note      Add note / 增加备注
/del 021 022            Mark rows deleted / 标记删除
/group 044 045          Group two crops / 合并两张 crop
/rollback               Remove latest photo / 回滚最近一张照片
/crops                  Send recent crops / 发送最近 crop 图片
/checked                Build checked finance output / 生成财务版
/submit                 Preview submit / 提交预览
/lang zh                Chinese replies / 中文回复
/lang en                English replies / 英文回复
```

## 8. Output Files / 输出文件

English:

- Manual review workbook: `报销明细_2026_xlsx.xlsx`
- Checked finance workbook: `报销_checked_2026.xlsx`
- Working crop folder: `crops/`
- Final finance crops: `final_crops/food/` and `final_crops/other/`

中文：

- 人工复核表：`报销明细_2026_xlsx.xlsx`
- 财务检查表：`报销_checked_2026.xlsx`
- 日常 crop 目录：`crops/`
- 财务最终图片：`final_crops/food/` 和 `final_crops/other/`

## 9. Privacy Reminder / 隐私提醒

English:

Do not upload real receipts, Excel files, `.env`, logs, or archives to GitHub. Runtime data is local by design.

中文：

不要把真实发票、Excel、`.env`、日志、压缩包上传到 GitHub。运行数据默认只保存在本地。
