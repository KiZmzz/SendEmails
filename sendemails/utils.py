import json
import smtplib
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import keyring


def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS) / relative_path
    return Path(__file__).parent.parent / relative_path


class ConfigManager:
    APP_NAME = "SendEmailsTool"
    CONFIG_FILE = get_resource_path("accounts.json")
    
    @classmethod
    def load_accounts(cls):
        if not cls.CONFIG_FILE.exists():
            return {}
        try:
            with open(cls.CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
            
    @classmethod
    def save_account(cls, config):
        account = config.get("account")
        if not account:
            return
            
        accounts = cls.load_accounts()
        secret = ""
        config_to_save = dict(config)
        
        if config["mode"] == "smtp":
            secret = config_to_save.pop("password", "")
        else:
            token_data = config_to_save.pop("token_data", None)
            if token_data:
                secret = json.dumps(token_data)
                
        accounts[account] = config_to_save
        
        with open(cls.CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(accounts, f, ensure_ascii=False, indent=4)
            
        if secret:
            try:
                keyring.set_password(cls.APP_NAME, account, secret)
            except Exception as e:
                print("Failed to save to keyring:", e)
                
    @classmethod
    def get_secret(cls, account, mode):
        try:
            secret = keyring.get_password(cls.APP_NAME, account)
            if not secret:
                return None
            if mode == "smtp":
                return secret
            else:
                return json.loads(secret)
        except Exception as e:
            print("Failed to get from keyring:", e)
            return None
            
    @classmethod
    def delete_account(cls, account):
        accounts = cls.load_accounts()
        if account in accounts:
            del accounts[account]
            with open(cls.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(accounts, f, ensure_ascii=False, indent=4)
        
        try:
            keyring.delete_password(cls.APP_NAME, account)
        except Exception:
            pass


class TemplateManager:
    TEMPLATE_FILE = get_resource_path("templates.json")
    
    @classmethod
    def load_templates(cls):
        if not cls.TEMPLATE_FILE.exists():
            return {}
        try:
            with open(cls.TEMPLATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
            
    @classmethod
    def save_template(cls, name, subject, html_content):
        templates = cls.load_templates()
        templates[name] = {
            "subject": subject,
            "content": html_content
        }
        with open(cls.TEMPLATE_FILE, "w", encoding="utf-8") as f:
            json.dump(templates, f, ensure_ascii=False, indent=4)
            
    @classmethod
    def delete_template(cls, name):
        templates = cls.load_templates()
        if name in templates:
            del templates[name]
            with open(cls.TEMPLATE_FILE, "w", encoding="utf-8") as f:
                json.dump(templates, f, ensure_ascii=False, indent=4)


class SharedAPI:
    @staticmethod
    def post_form(url, data):
        body = urllib.parse.urlencode(data).encode("utf-8")
        request = urllib.request.Request(
            url,
            data=body,
            method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded"},
        )
        return SharedAPI.perform_request(request)

    @staticmethod
    def graph_request(method, url, access_token, payload=None):
        headers = {"Authorization": f"Bearer {access_token}"}
        data = None
        if payload is not None:
            headers["Content-Type"] = "application/json"
            data = json.dumps(payload).encode("utf-8")
        request = urllib.request.Request(url, data=data, method=method, headers=headers)
        return SharedAPI.perform_request(request)

    @staticmethod
    def perform_request(request):
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                raw = response.read().decode("utf-8", errors="replace")
                return response.getcode(), json.loads(raw) if raw else {}
        except urllib.error.HTTPError as exc:
            raw = exc.read().decode("utf-8", errors="replace")
            try:
                payload = json.loads(raw) if raw else {}
            except json.JSONDecodeError:
                payload = {"error_description": raw or str(exc)}
            return exc.code, payload
        except Exception as exc:
            raise RuntimeError(str(exc)) from exc

    @staticmethod
    def format_oauth_error(payload):
        if not isinstance(payload, dict):
            return str(payload)
        error = payload.get("error", "oauth_error")
        description = payload.get("error_description") or payload.get("message") or "微软授权失败。"
        return f"{error}\n{description}"

    @staticmethod
    def format_graph_error(payload):
        if isinstance(payload, dict) and "error" in payload:
            error_block = payload["error"]
            if isinstance(error_block, dict):
                code = error_block.get("code", "graph_error")
                message = error_block.get("message", "Microsoft Graph 请求失败。")
                return f"{code}\n{message}"
        return str(payload)

    @staticmethod
    def create_smtp_client(smtp_server, smtp_port):
        if smtp_port == 465:
            client = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
        else:
            client = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
            client.ehlo()
            if smtp_port in (587, 25):
                client.starttls()
                client.ehlo()
        return client
