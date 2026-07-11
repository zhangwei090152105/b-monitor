---
AIGC:
    Label: "1"
    ContentProducer: 001191440300708461136T1XGW3
    ProduceID: a646dcce4fee2e748b65eaefcef56d09_622dc23f7cf011f1baf4525400bff409
    ReservedCode1: K/FkA8mY/fknyT7Qe+eLNTVhK4tLPW10LheKkA0TMUxfMprbYW8q3hG+5CldK/YwvT+kTRT6IY7DHnPJwSXOFlUlmiz19HN/5wsrW5zdWa1wSnG2X+4xzg1k+Zg8VxmKPuEBZP0BBhoYQdY3YhpcuGKhlLKLHTyF7F+PAqcdYfR5db4sz285SFQLPBQ=
    ContentPropagator: 001191440300708461136T1XGW3
    PropagateID: a646dcce4fee2e748b65eaefcef56d09_622dc23f7cf011f1baf4525400bff409
    ReservedCode2: K/FkA8mY/fknyT7Qe+eLNTVhK4tLPW10LheKkA0TMUxfMprbYW8q3hG+5CldK/YwvT+kTRT6IY7DHnPJwSXOFlUlmiz19HN/5wsrW5zdWa1wSnG2X+4xzg1k+Zg8VxmKPuEBZP0BBhoYQdY3YhpcuGKhlLKLHTyF7F+PAqcdYfR5db4sz285SFQLPBQ=
---

# 吉客云 B 单未完结监控 — GitHub Actions 部署指南

## 概述

本项目将吉客云 B 单（出库/入库申请单）的未完结监控改造为 **GitHub Actions 定时任务**，每天北京时间 9:00 自动执行以下流水线：

1. **API 数据拉取** — 从吉客云 Open API 获取出库/入库申请单
2. **Excel 报告生成** — 4 Sheet 汇总+明细报告
3. **邮件发送** — HTML 正文 + Excel 附件
4. **企业微信推送** — Markdown 摘要（Top 5 品牌）

---

## 前提条件

1. **GitHub 账号** — 注册 [github.com](https://github.com)
2. **吉客云开放平台凭证** — 在吉客云后台获取 `AppKey` 和 `AppSecret`
3. **QQ邮箱 SMTP 授权码** — 用于发送邮件报告（或其它 SMTP 服务）
4. **企业微信群机器人 Webhook** — 用于企微消息推送（可选）

---

## 快速部署（5 步）

### 第 1 步：创建私有仓库

> ⚠️ **强烈建议创建私有仓库**，因为脚本涉及企业 API 凭证和业务数据。

```bash
# 在 GitHub 上创建私有仓库，例如: jky-b-monitor

# 将本文件夹内容推送到仓库
git init
git add .
git commit -m "初始化吉客云B单监控"
git remote add origin https://github.com/你的用户名/jky-b-monitor.git
git push -u origin main
```

### 第 2 步：配置 GitHub Secrets

所有敏感信息通过 GitHub Secrets 管理，不会暴露在代码中。

进入仓库页面 → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**，依次添加以下 Secrets：

| Secret 名称 | 说明 | 示例值 |
|---|---|---|
| `JKY_APPKEY` | 吉客云开放平台 AppKey | `47457790` |
| `JKY_APPSECRET` | 吉客云开放平台 AppSecret | `413416a89ad944a6b...` |
| `SMTP_HOST` | SMTP 服务器地址 | `smtp.qq.com` |
| `SMTP_PORT` | SMTP 端口 | `465` |
| `SMTP_USER` | 发件邮箱地址 | `yourname@qq.com` |
| `SMTP_PASS` | SMTP 授权码（非邮箱密码） | `qbauihjiiuschife` |
| `TO_ADDR` | 收件人邮箱 | `wei.zhang@example.com` |
| `CC_ADDR` | 抄送邮箱 | `cc@example.com` |
| `WECOM_WEBHOOK_1` | 企微群机器人 Webhook 1 | `https://qyapi.weixin.qq.com/...` |
| `WECOM_WEBHOOK_2` | 企微群机器人 Webhook 2（可选） | `https://qyapi.weixin.qq.com/...` |
| `MIAODA_URL` | 数字看板公网 URL（可选） | `https://xxx.github.io/b-monitor/` |

> **获取各凭证的指引**：
> - **吉客云**：开放平台 → 应用管理 → 查看 AppKey/AppSecret
> - **QQ邮箱 SMTP**：QQ邮箱 → 设置 → 账户 → POP3/SMTP 服务 → 开启并获取授权码
> - **企微 Webhook**：企微群 → 群设置 → 群机器人 → 添加 → 复制 webhook 地址

### 第 3 步：手动触发测试

1. 进入仓库 → **Actions** 标签
2. 左侧选择 **"吉客云B单未完结监控"**
3. 点击 **"Run workflow"** 下拉 → 可选填日期 → 点击绿色 **"Run workflow"** 按钮
4. 等待约 1-2 分钟，查看运行结果

### 第 4 步：确认定时任务生效

Workflow 已配置为每天 **UTC 1:00** 自动执行（对应北京时间 **9:00**）。

- 定时任务在首次推送后自动生效，无需额外操作
- 如果仓库超过 60 天无活动，GitHub 会自动暂停定时任务，届时需要手动触发一次来恢复

### 第 5 步：查看运行日志 & Excel 报告

- **运行日志**：Actions 标签 → 点击具体运行记录 → 展开 `执行 B 单监控` 步骤
- **Excel 报告**：运行记录底部 → **Artifacts** → 下载 `b-monitor-report`（保留 30 天）

---

## 文件结构

```
jky-b-monitor/
├── .github/
│   └── workflows/
│       └── jky-b-monitor.yml    # GitHub Actions workflow 配置
├── main.py                       # 主脚本（整合 4 阶段流水线）
├── config.example.py             # 配置文件模板（环境变量占位）
├── requirements.txt              # Python 依赖
├── README.md                     # 本文件
└── output/                       # 运行时生成的输出目录
    ├── raw_stockout_*.json       # 出库原始数据
    ├── raw_stockin_*.json        # 入库原始数据
    └── jkyun_b_monitor_*.xlsx    # Excel 报告
```

---

## Excel 报告处理方案

由于 GitHub Actions 运行环境没有持久化的本地文件系统，生成的 Excel 报告有以下处理方式：

| 方案 | 说明 | 推荐度 |
|---|---|---|
| **邮件附件**（已内置） | 脚本通过 SMTP 将 Excel 作为附件发送给收件人。每次运行后邮箱即可收到。 | ⭐⭐⭐ 推荐 |
| **GitHub Artifacts**（已内置） | Workflow 自动上传 Excel 到 GitHub，保留 30 天，可在 Actions 页面下载。 | ⭐⭐ 备份 |
| **云存储上传**（可选扩展） | 可扩展上传到阿里云 OSS / 腾讯云 COS / OneDrive 等。需额外配置。 | ⭐ 按需 |

> 当前方案：邮件附件（主）+ GitHub Artifacts 备份（辅），无需额外配置即可覆盖日常需求。

---

## 时区与时序说明

| 项目 | 时间 |
|---|---|
| GitHub Actions cron 表达式 | `0 1 * * *`（UTC 时间） |
| 对应北京时间 | 每天 **9:00** |
| 对应 UTC+8（无夏令时） | 每天 09:00 CST |

> GitHub Actions 的 `schedule` 使用 **UTC 时间**，不支持直接指定时区。
> 如需修改执行时间，请在 [crontab.guru](https://crontab.guru/) 换算 UTC 时间后修改 `.github/workflows/jky-b-monitor.yml` 中的 `cron` 字段。

---

## 常见问题

### Q1：定时任务没有按时执行？

- 确认仓库在近 60 天内有活动（commit / push / manual trigger）
- GitHub Actions 定时任务可能在高峰期延迟 5-30 分钟，属正常现象
- 检查 Actions 标签页的 workflow 是否处于 `active` 状态

### Q2：执行时间超过 6 小时怎么办？

- Workflow 已配置 `timeout-minutes: 30`，正常情况下 1-2 分钟即可完成
- GitHub Actions 免费版单次运行上限为 6 小时，本脚本远低于此限制
- 如果数据量巨大导致超时，可考虑分天拉取并缓存中间结果（advanced）

### Q3：GitHub Actions 免费额度够用吗？

| 项目 | 免费额度 | 本脚本消耗 |
|---|---|---|
| 运行时间 | 2,000 分钟/月（私有仓库） | ~2 分钟/天 ≈ 60 分钟/月 |
| 存储（Artifacts） | 500 MB | ~0.5 MB/次 ≈ 15 MB/月 |
| API 请求 | 1,000 次/小时 | 约 200-300 次/次运行 |

> 免费额度完全够用，每月仅消耗约 3% 的运行时间。

### Q4：如何修改执行时间？

编辑 `.github/workflows/jky-b-monitor.yml` 中的 `cron` 字段：

```yaml
on:
  schedule:
    # UTC 时间换算：北京时间 9:00 = UTC 1:00
    - cron: '0 1 * * *'
```

将 `0 1` 改为你需要的 UTC 时间（注意：北京时间 = UTC + 8）。

### Q5：邮件发送失败？

检查项：
- `SMTP_PASS` 是**授权码**而非邮箱登录密码（QQ邮箱需在设置中单独生成）
- `SMTP_HOST` 和 `SMTP_PORT` 与邮箱服务商匹配
- 如果使用企业邮箱，可能需要修改 `SMTP_HOST` 为对应地址

### Q6：企微推送乱码/格式错乱？

企微 Markdown 有严格渲染限制，脚本已适配以下规则：
- 使用全角空格（U+3000）代替 `&nbsp;`
- 使用 `\n\n\n` 强制空行
- Emoji 限定在 Unicode 6.0 范围内

### Q7：如何升级到包含妙搭看板的完整方案？

妙搭看板更新需要调用妙搭 API 或部署静态页面到 GitHub Pages，不在本 `main.py` 范围内。如需扩展：
1. 将看板部署到 GitHub Pages
2. 使用 `peaceiris/actions-gh-pages` Action 自动部署
3. 邮件和企微中的链接指向你的 GitHub Pages 地址

---

## 本地测试

在推送前可本地验证脚本：

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 设置环境变量（Windows PowerShell）
$env:JKY_APPKEY="你的AppKey"
$env:JKY_APPSECRET="你的AppSecret"
$env:SMTP_USER="yourname@qq.com"
$env:SMTP_PASS="你的授权码"
$env:TO_ADDR="收件人@example.com"
$env:WECOM_WEBHOOK_1="你的Webhook地址"

# 3. 运行
python main.py
```

---

## 维护说明

- **版本**：V1.0（2026-07-11）
- **依赖**：Python 3.11+, openpyxl
- **数据源**：吉客云 Open API v2
- **原始 Skill 根目录**：`skills/custom/jky-b-monitor/`
*（内容由AI生成，仅供参考）*
