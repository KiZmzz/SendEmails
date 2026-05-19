import json
import time
from email.header import Header
from email.mime.text import MIMEText
from email.utils import formataddr, make_msgid, formatdate

from PyQt6.QtCore import QThread, pyqtSignal

from .constants import GRAPH_SCOPES, DeviceFlowCancelled
from .utils import SharedAPI


class MicrosoftAuthWorker(QThread):
    flow_ready = pyqtSignal(dict)
    login_success = pyqtSignal(dict, dict)
    login_failed = pyqtSignal(str)
    
    def __init__(self, tenant, client_id):
        super().__init__()
        self.tenant = tenant
        self.client_id = client_id
        self.is_cancelled = False
        self.active_device_code = None

    def cancel(self):
        self.is_cancelled = True

    def run(self):
        try:
            flow = self.request_device_code()
            self.active_device_code = flow["device_code"]
            self.flow_ready.emit(flow)
            
            token_data = self.poll_device_code_token(flow)
            profile = self.fetch_graph_profile(token_data["access_token"])
            self.login_success.emit(token_data, profile)
        except DeviceFlowCancelled:
            pass
        except Exception as exc:
            self.login_failed.emit(str(exc))

    def request_device_code(self):
        url = f"https://login.microsoftonline.com/{self.tenant}/oauth2/v2.0/devicecode"
        status, payload = SharedAPI.post_form(url, {"client_id": self.client_id, "scope": GRAPH_SCOPES})
        if status >= 400:
            raise RuntimeError(SharedAPI.format_oauth_error(payload))
        return payload

    def poll_device_code_token(self, flow):
        url = f"https://login.microsoftonline.com/{self.tenant}/oauth2/v2.0/token"
        interval = int(flow.get("interval", 5))
        expires_at = time.time() + int(flow.get("expires_in", 900))

        while time.time() < expires_at:
            if self.is_cancelled:
                raise DeviceFlowCancelled()

            status, payload = SharedAPI.post_form(
                url,
                {
                    "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                    "client_id": self.client_id,
                    "device_code": flow["device_code"],
                },
            )

            if status < 400 and payload.get("access_token"):
                payload["obtained_at"] = time.time()
                return payload

            error = payload.get("error")
            if error == "authorization_pending":
                time.sleep(interval)
                continue
            if error == "slow_down":
                time.sleep(interval + 5)
                continue

            raise RuntimeError(SharedAPI.format_oauth_error(payload))

        raise RuntimeError("设备码已过期，请重新发起微软登录。")

    def fetch_graph_profile(self, access_token):
        status, payload = SharedAPI.graph_request(
            "GET",
            "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName",
            access_token,
        )
        if status >= 400:
            raise RuntimeError(SharedAPI.format_graph_error(payload))
        return payload


class SmtpTestWorker(QThread):
    success = pyqtSignal()
    failed = pyqtSignal(str)

    def __init__(self, config):
        super().__init__()
        self.config = config

    def run(self):
        try:
            with SharedAPI.create_smtp_client(self.config["smtp_server"], self.config["smtp_port"]) as client:
                client.login(self.config["account"], self.config["password"])
            self.success.emit()
        except Exception as exc:
            self.failed.emit(str(exc))


class EmailSendWorker(QThread):
    progress = pyqtSignal(int, int, int, int)
    complete = pyqtSignal(bool, int, list)

    def __init__(self, config, subject, content, recipients, token_data=None, send_interval=0):
        super().__init__()
        self.config = config
        self.subject = subject
        self.content = content
        self.recipients = recipients
        self.token_data = token_data
        self.send_interval = send_interval
        self.mode = "graph" if token_data is not None else "smtp"

    def run(self):
        if self.mode == "graph":
            self.run_graph()
        else:
            self.run_smtp()

    def run_smtp(self):
        success_count = 0
        failed = []
        try:
            with SharedAPI.create_smtp_client(self.config["smtp_server"], self.config["smtp_port"]) as client:
                client.login(self.config["account"], self.config["password"])
                for index, recipient in enumerate(self.recipients, start=1):
                    try:
                        name, email = recipient if isinstance(recipient, tuple) else ("未知", recipient)
                        message = MIMEText(self.content, "html", "utf-8")
                        message["From"] = formataddr(
                            (str(Header(self.config["nickname"], "utf-8")), self.config["account"])
                        )
                        message["To"] = formataddr((str(Header(name, "utf-8")), email))
                        message["Subject"] = Header(self.subject, "utf-8")
                        message["Message-ID"] = make_msgid()
                        message["Date"] = formatdate()
                        client.sendmail(self.config["account"], [email], message.as_string())
                        success_count += 1
                    except Exception as exc:
                        name, email = recipient if isinstance(recipient, tuple) else ("未知", recipient)
                        failed.append((email, str(exc)))
                    self.progress.emit(index, len(self.recipients), success_count, len(failed))
                    if self.send_interval > 0 and index < len(self.recipients):
                        time.sleep(self.send_interval)
        except Exception as exc:
            self.complete.emit(False, 0, [(self.config["account"], str(exc))])
            return

        self.complete.emit(True, success_count, failed)

    def run_graph(self):
        success_count = 0
        failed = []
        try:
            self.refresh_graph_token_if_needed()
            token = self.token_data["access_token"]
        except Exception as exc:
            self.complete.emit(False, 0, [(self.config["account"], str(exc))])
            return

        for index, recipient in enumerate(self.recipients, start=1):
            try:
                name, email = recipient if isinstance(recipient, tuple) else ("未知", recipient)
                payload = self.build_graph_message_payload(name, email, self.subject, self.content)
                status, response = SharedAPI.graph_request(
                    "POST",
                    "https://graph.microsoft.com/v1.0/me/sendMail",
                    token,
                    payload,
                )
                if status >= 400:
                    raise RuntimeError(SharedAPI.format_graph_error(response))
                success_count += 1
            except Exception as exc:
                name, email = recipient if isinstance(recipient, tuple) else ("未知", recipient)
                failed.append((email, str(exc)))
            self.progress.emit(index, len(self.recipients), success_count, len(failed))

        self.complete.emit(True, success_count, failed)

    def build_graph_message_payload(self, name, email, subject, content):
        return {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "HTML",
                    "content": content,
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": email,
                            "name": name,
                        }
                    }
                ],
            },
            "saveToSentItems": True,
        }

    def refresh_graph_token_if_needed(self):
        expires_in = int(self.token_data.get("expires_in", 3600))
        obtained_at = float(self.token_data.get("obtained_at", 0))
        if time.time() < obtained_at + expires_in - 120:
            return

        refresh_token = self.token_data.get("refresh_token")
        if not refresh_token:
            raise RuntimeError("微软授权已过期，请重新登录。")

        url = f"https://login.microsoftonline.com/{self.config['tenant']}/oauth2/v2.0/token"
        status, payload = SharedAPI.post_form(
            url,
            {
                "client_id": self.config['client_id'],
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
                "scope": GRAPH_SCOPES,
            },
        )
        if status >= 400 or "access_token" not in payload:
            raise RuntimeError(SharedAPI.format_oauth_error(payload))

        payload["obtained_at"] = time.time()
        if "refresh_token" not in payload:
            payload["refresh_token"] = refresh_token
        self.token_data = payload
