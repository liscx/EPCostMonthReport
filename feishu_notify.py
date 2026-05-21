"""
飞书消息通知模块 - 用于从 Python 脚本直接发送图片/文本到飞书聊天。

用法:
    from feishu_notify import send_image, send_text
    send_image("/path/to/qr_code.png", "请扫码登录")
    send_text("操作完成")
"""

import os
import json
import urllib.request
import urllib.error

# --- 配置 ---
_ENV_PATH = os.path.join(os.path.expanduser("~"), "AppData", "Local", "hermes", ".env")
# 优先级: 环境变量 FEISHU_NOTIFY_CHAT_ID > hermes .env FEISHU_HOME_CHANNEL
_DEFAULT_CHAT_ID = os.environ.get("FEISHU_NOTIFY_CHAT_ID", "")
# 私聊目标：设了这个就发私聊，不发到当前会话
_DEFAULT_USER_OPEN_ID = os.environ.get("FEISHU_USER_OPEN_ID", "")
_DOMAIN = "https://open.feishu.cn"


def _get_default_chat_id() -> str:
    """获取默认 chat_id，优先级: 参数 > 环境变量 > hermes .env FEISHU_HOME_CHANNEL。未配置时返回空字符串。"""
    if _DEFAULT_CHAT_ID:
        return _DEFAULT_CHAT_ID
    # 尝试从 hermes .env 读取 FEISHU_HOME_CHANNEL
    if os.path.exists(_ENV_PATH):
        with open(_ENV_PATH, "r", encoding="utf-8") as f:
            for line in f:
                if line.strip().startswith("FEISHU_HOME_CHANNEL="):
                    return line.strip().split("=", 1)[1]
    return ""


def _load_env():
    """从 hermes .env 文件读取飞书凭证"""
    app_id, app_secret = "", ""
    if os.path.exists(_ENV_PATH):
        with open(_ENV_PATH, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("FEISHU_APP_ID="):
                    app_id = line.split("=", 1)[1]
                elif line.startswith("FEISHU_APP_SECRET="):
                    app_secret = line.split("=", 1)[1]
    if not app_id or not app_secret:
        raise RuntimeError("FEISHU_APP_ID / FEISHU_APP_SECRET 未配置")
    return app_id, app_secret


def _get_tenant_token(app_id: str, app_secret: str) -> str:
    """获取 tenant_access_token"""
    url = f"{_DOMAIN}/open-apis/auth/v3/tenant_access_token/internal"
    body = json.dumps({"app_id": app_id, "app_secret": app_secret}).encode()
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = json.loads(resp.read())
    if data.get("code") != 0:
        raise RuntimeError(f"获取 token 失败: {data}")
    return data["tenant_access_token"]


def _upload_image(token: str, image_path: str) -> str:
    """上传图片到飞书，返回 image_key"""
    url = f"{_DOMAIN}/open-apis/im/v1/images"
    boundary = "----HermesBoundary"
    filename = os.path.basename(image_path)

    with open(image_path, "rb") as f:
        image_data = f.read()

    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="image_type"\r\n\r\n'
        f"message\r\n"
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="image"; filename="{filename}"\r\n'
        f"Content-Type: image/png\r\n\r\n"
    ).encode() + image_data + f"\r\n--{boundary}--\r\n".encode()

    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": f"multipart/form-data; boundary={boundary}",
    })
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read())
    if data.get("code") != 0:
        raise RuntimeError(f"上传图片失败: {data}")
    return data["data"]["image_key"]


def _send_message(token: str, receive_id: str, msg_type: str, content: str, id_type: str = "chat_id"):
    """发送消息到飞书"""
    url = f"{_DOMAIN}/open-apis/im/v1/messages?receive_id_type={id_type}"
    body = json.dumps({
        "receive_id": receive_id,
        "msg_type": msg_type,
        "content": content,
    }).encode()
    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    })
    with urllib.request.urlopen(req, timeout=10) as resp:
        data = json.loads(resp.read())
    if data.get("code") != 0:
        raise RuntimeError(f"发送消息失败: {data}")
    return data


def _resolve_target(chat_id: str) -> tuple[str, str]:
    """解析发送目标，返回 (receive_id, id_type)。未配置时返回 (None, None)。"""
    if chat_id:
        return chat_id, "chat_id"
    if _DEFAULT_USER_OPEN_ID:
        return _DEFAULT_USER_OPEN_ID, "open_id"
    default_chat_id = _get_default_chat_id()
    if default_chat_id:
        return default_chat_id, "chat_id"
    return None, None


def send_image(image_path: str, text: str = "", chat_id: str = ""):
    """
    发送图片到飞书聊天。

    Args:
        image_path: 图片文件路径
        text: 可选的文字说明（作为独立文本消息发送）
        chat_id: 目标聊天 ID，默认发私聊（需设置 FEISHU_USER_OPEN_ID）
    """
    receive_id, id_type = _resolve_target(chat_id)
    if receive_id is None:
        print("[飞书] 未配置通知目标，跳过图片发送")
        return
    app_id, app_secret = _load_env()
    token = _get_tenant_token(app_id, app_secret)
    if text:
        _send_message(token, receive_id, "text", json.dumps({"text": text}), id_type)
    image_key = _upload_image(token, image_path)
    _send_message(token, receive_id, "image", json.dumps({"image_key": image_key}), id_type)
    print(f"[飞书] 图片已发送: {image_path}")


def send_text(text: str, chat_id: str = ""):
    """
    发送文本消息到飞书聊天。

    Args:
        text: 消息内容
        chat_id: 目标聊天 ID，默认发私聊（需设置 FEISHU_USER_OPEN_ID）
    """
    receive_id, id_type = _resolve_target(chat_id)
    if receive_id is None:
        print("[飞书] 未配置通知目标，跳过消息发送")
        return
    app_id, app_secret = _load_env()
    token = _get_tenant_token(app_id, app_secret)
    _send_message(token, receive_id, "text", json.dumps({"text": text}), id_type)
    print(f"[飞书] 消息已发送: {text[:50]}...")
