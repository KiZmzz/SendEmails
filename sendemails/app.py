import sys

from PyQt6.QtCore import Qt, QThread, pyqtSignal, QUrl
from PyQt6.QtGui import QIcon, QFont
from PyQt6.QtWidgets import QApplication
from qfluentwidgets import (
    FluentWindow, NavigationItemPosition, SubtitleLabel, InfoBar, MessageBox,
    setThemeColor, setTheme, Theme, FluentIcon as FIF,
)

from .constants import EMAIL_PROVIDERS
from .interfaces import AccountInterface, ContentInterface, SendInterface
from .utils import get_resource_path
from .workers import EmailSendWorker


class EmailSenderApp(FluentWindow):
    def __init__(self):
        super().__init__()
        self._disable_windows_mica_background()
        self.setWindowTitle("批量邮件发送工具")
        self.setWindowIcon(QIcon(str(get_resource_path("app.ico"))))
        self.setMinimumSize(1200, 700)
        self.resize(1200, 780)
        
        self.accountInterface = AccountInterface(self)
        self.contentInterface = ContentInterface(self)
        self.sendInterface = SendInterface(self)

        self.initNavigation()
        self.sendInterface.sendRequested.connect(self.start_sending)
        self.widgetLayout.setContentsMargins(0, 0, 0, 0)
        
        # 调整导航和标题栏
        self.navigationInterface.setExpandWidth(220)
        
        panel = self.navigationInterface.panel
        btn = panel.returnButton
        for p_layout in (panel.layout(),):
            for i in range(p_layout.count()):
                sub = p_layout.itemAt(i)
                if sub and sub.layout():
                    sub_layout = sub.layout()
                    for j in range(sub_layout.count()):
                        item = sub_layout.itemAt(j)
                        if item and item.widget() == btn:
                            sub_layout.takeAt(j)
                            break
        btn.setParent(None)
        btn.deleteLater()
        
        tb = self.titleBar
        tb.titleLabel.setParent(None)
        tb.titleLabel.deleteLater()
        if hasattr(tb, "iconLabel"):
            tb.iconLabel.hide()
        tb.setFixedHeight(32)

    def _disable_windows_mica_background(self):
        if sys.platform != "win32":
            return

        if hasattr(self, "setMicaEffectEnabled"):
            self.setMicaEffectEnabled(False)

        if hasattr(self, "setCustomBackgroundColor"):
            self.setCustomBackgroundColor("#f5f7fb", "#202020")

    def initNavigation(self):
        self.addSubInterface(self.accountInterface, FIF.SETTING, "账号设置", NavigationItemPosition.TOP)
        self.addSubInterface(self.contentInterface, FIF.EDIT, "邮件内容", NavigationItemPosition.TOP)
        self.addSubInterface(self.sendInterface, FIF.SEND, "收件人与发送", NavigationItemPosition.TOP)

    def start_sending(self):
        if not self.sendInterface.recipients:
            InfoBar.error("错误", "请先添加收件人", parent=self)
            self.switchTo(self.sendInterface)
            return
            
        subject = self.contentInterface.subjectEdit.text().strip()
        self.contentInterface.webView.page().runJavaScript(
            "getHtmlContent();", 
            lambda html: self._continue_sending(subject, html)
        )
        
    def _continue_sending(self, subject, raw_html):
        content = str(raw_html).strip() if raw_html else ""
        if not content or content == "<p><br></p>":
            InfoBar.error("错误", "邮件正文不能为空", parent=self)
            self.switchTo(self.contentInterface)
            return

        try:
            config = self.accountInterface.get_config()
        except ValueError as e:
            InfoBar.error("配置错误", str(e), parent=self)
            self.switchTo(self.accountInterface)
            return
            
        self.sendInterface.sendBtn.setEnabled(False)
        self.sendInterface.progressBar.show()
        self.sendInterface.progressBar.setMaximum(len(self.sendInterface.recipients))
        self.sendInterface.progressBar.setValue(0)
        self.sendInterface.statusLabel.setText("正在发送...")
        
        send_interval = self.sendInterface.get_send_interval()
        token_data = config.get("token_data")
        self.worker = EmailSendWorker(config, subject, content, self.sendInterface.recipients, token_data, send_interval)
        self.worker.progress.connect(self.on_send_progress)
        self.worker.complete.connect(self.on_send_complete)
        self.worker.start()

    def on_send_progress(self, index, total, success, fail):
        self.sendInterface.progressBar.setValue(index)
        self.sendInterface.statusLabel.setText(f"发送进度 {index}/{total}，成功 {success}，失败 {fail}。")

    def on_send_complete(self, connected, success, failed):
        self.sendInterface.sendBtn.setEnabled(True)
        if not connected:
            error = failed[0][1] if failed else "未知错误"
            self.sendInterface.statusLabel.setText("发送失败！")
            MessageBox("发送失败", error, self).exec()
            return
            
        fail_count = len(failed)
        if fail_count == 0:
            self.sendInterface.statusLabel.setText("发送完成，全部成功！")
            InfoBar.success("发送完成", f"全部发送成功，共 {success} 封。", parent=self, duration=5000)
        else:
            self.sendInterface.statusLabel.setText(f"发送完成，成功 {success} 封，失败 {fail_count} 封。")
            detail = "\n".join([f"{em}: {rs}" for em, rs in failed[:5]])
            if fail_count > 5:
                detail += f"\n... 还有 {fail_count - 5} 个失败"
            MessageBox("部分发送失败", f"成功 {success} 封，失败 {fail_count} 封。\n\n失败详情：\n{detail}", self).exec()
