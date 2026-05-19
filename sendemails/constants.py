import re


EMAIL_PATTERN = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
NAME_EMAIL_PATTERN = re.compile(
    r'([^\s<（(,;，；]*?)\s*(?:<|（)?\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})'
)

EMAIL_PROVIDERS = {
    "Outlook / Microsoft 365": {
        "mode": "graph",
        "description": "使用 Microsoft Graph OAuth 2.0 登录，适合 Outlook 和 Microsoft 365。",
    },
    "阿里企业邮箱": {
        "mode": "smtp",
        "server": "smtp.qiye.aliyun.com",
        "port": "465",
        "description": "使用 SMTP 登录，通常需要三方客户端安全密码。",
    },
    "QQ邮箱": {
        "mode": "smtp",
        "server": "smtp.qq.com",
        "port": "465",
        "description": "使用 SMTP 登录，通常需要授权码。",
    },
    "网易邮箱": {
        "mode": "smtp",
        "server": "smtp.163.com",
        "port": "465",
        "description": "使用 SMTP 登录，通常需要授权码。",
    },
    "Gmail": {
        "mode": "smtp",
        "server": "smtp.gmail.com",
        "port": "587",
        "description": "使用 SMTP 登录，通常需要应用专用密码。",
    },
    "其他 SMTP": {
        "mode": "smtp",
        "server": "",
        "port": "",
        "description": "手动填写 SMTP 服务器和端口。",
    },
}

EMAIL_HEADER_KEYWORDS = (
    "email", "e-mail", "mail", "邮箱", "邮箱地址", "电子邮箱", "收件邮箱", "收件人邮箱"
)
NAME_HEADER_KEYWORDS = (
    "收件人", "姓名", "名字", "name", "收件人姓名", "联系人", "recipient"
)

GRAPH_SCOPES = "offline_access openid profile User.Read Mail.Send"


class DeviceFlowCancelled(Exception):
    pass
