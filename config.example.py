"""
吉客云 B 单未完结监控 - 配置文件模板

⚠️ 所有敏感信息通过环境变量读取，严禁在此文件中硬编码任何密钥/Token。
   部署到 GitHub 时，将这些值配置到 Settings → Secrets and variables → Actions → Secrets。
   本地测试时，复制此文件为 config.py 并填入真实值（config.py 已加入 .gitignore）。
"""

import os


# ==============================
#  吉客云 Open API 凭证
#  来源：吉客云开放平台 → 应用管理 → 查看 AppKey/AppSecret
# ==============================
JKY_APPKEY    = os.environ.get('JKY_APPKEY', '')
JKY_APPSECRET = os.environ.get('JKY_APPSECRET', '')
JKY_API_URL   = 'https://open.jackyun.com/open/openapi/do'


# ==============================
#  邮件 SMTP 配置
#  示例：QQ邮箱 → smtp.qq.com:465
#  发送邮箱需开启 SMTP 服务并获取授权码
# ==============================
SMTP_HOST = os.environ.get('SMTP_HOST', 'smtp.qq.com')
SMTP_PORT = int(os.environ.get('SMTP_PORT', '465'))
SMTP_USER = os.environ.get('SMTP_USER', '')
SMTP_PASS = os.environ.get('SMTP_PASS', '')  # SMTP 授权码，非邮箱密码
TO_ADDR   = os.environ.get('TO_ADDR', '')     # 主收件人
CC_ADDR   = os.environ.get('CC_ADDR', '')     # 抄送


# ==============================
#  企业微信 Webhook
#  获取：企微群 → 群设置 → 群机器人 → 添加 → 复制 webhook 地址
#  支持多个 webhook，留空字符串表示不推送到该群
# ==============================
WECOM_WEBHOOKS = [
    url for url in [
        os.environ.get('WECOM_WEBHOOK_1', ''),
        os.environ.get('WECOM_WEBHOOK_2', ''),
    ] if url
]

# ==============================
#  妙搭看板链接
#  数字看板的公网 URL（GitHub Pages 或其他托管地址）
# ==============================
MIAODA_URL = os.environ.get('MIAODA_URL', '')


# ==============================
#  可选：手动指定数据日期
#  格式 YYYYMMDD，留空则使用当天日期
# ==============================
B_MONITOR_DATE = os.environ.get('B_MONITOR_DATE', '')
