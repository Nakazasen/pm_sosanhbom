"""Feature F25: Outlook COM Email Automation Module.

Automates generation, preview, and sending of HTML notification emails to team members
and department leaders using Windows Outlook COM Interop (win32com.client.Dispatch("Outlook.Application")).
Provides graceful mock/offline fallback for headless environments and automated unit testing.
"""

from __future__ import annotations

import datetime
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)


@dataclass
class EmailPreview:
    """Represents a generated email notification payload ready for preview or dispatch."""

    subject: str
    recipients_to: list[str]
    recipients_cc: list[str]
    html_body: str
    attachment_paths: list[Path] = field(default_factory=list)
    overall_status: str = "OK"  # "OK" or "NG"
    timestamp: str = field(default_factory=lambda: datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))

    @property
    def to_string(self) -> str:
        """Semicolon-delimited 'To' string."""
        return "; ".join(self.recipients_to)

    @property
    def cc_string(self) -> str:
        """Semicolon-delimited 'CC' string."""
        return "; ".join(self.recipients_cc)


class OutlookMailer:
    """Manages HTML email formatting and Outlook COM lifecycle."""

    def __init__(
        self,
        default_sender: str = "vn_pe03@dtvn.kyocera.com",
        default_cc: list[str] | None = None,
        com_dispatch: Any | None = None,
        test_mode: bool = True,
        test_recipient: str = "vinh.bd@dtvn.kyocera.com",
    ) -> None:
        """Initialize mailer with optional mockable COM dispatch."""
        self.default_sender = default_sender
        self.default_cc = default_cc or ["vinh.bd@dtvn.kyocera.com"]
        self._custom_dispatch = com_dispatch
        self.test_mode = test_mode
        self.test_recipient = test_recipient

    def build_email_preview(
        self,
        model_name: str,
        target_date: str,
        overall_status: str,
        recipients_to: str | list[str],
        recipients_cc: str | list[str] | None = None,
        sub_unit_statuses: dict[str, str] | None = None,
        summary_stats: dict[str, Any] | None = None,
        attachment_path: str | Path | None = None,
        custom_notes: str = "",
    ) -> EmailPreview:
        """Construct the EmailPreview payload containing rendered HTML and metadata."""
        to_list = [r.strip() for r in (recipients_to.split(";") if isinstance(recipients_to, str) else recipients_to) if r.strip()]

        if recipients_cc is not None:
            cc_list = [c.strip() for c in (recipients_cc.split(";") if isinstance(recipients_cc, str) else recipients_cc) if c.strip()]
        else:
            cc_list = list(self.default_cc)

        status_tag = "HOÀN TẤT - OK" if overall_status == "OK" else "CẦN GIẢI TRÌNH - NG"
        subject = f"[{status_tag}] Báo cáo đối soát BOM tự động - Model {model_name} ({target_date})"

        html_body = self.generate_html_body(
            model_name=model_name,
            target_date=target_date,
            overall_status=overall_status,
            sub_unit_statuses=sub_unit_statuses,
            summary_stats=summary_stats,
            attachment_path=attachment_path,
            custom_notes=custom_notes,
        )

        attachments: list[Path] = []
        if attachment_path:
            p = Path(attachment_path).resolve()
            if p.exists():
                attachments.append(p)
            else:
                logger.warning("Attachment path does not exist on disk: %s", p)

        return EmailPreview(
            subject=subject,
            recipients_to=to_list,
            recipients_cc=cc_list,
            html_body=html_body,
            attachment_paths=attachments,
            overall_status=overall_status,
        )

    def generate_html_body(
        self,
        model_name: str,
        target_date: str,
        overall_status: str,
        sub_unit_statuses: dict[str, str] | None = None,
        summary_stats: dict[str, Any] | None = None,
        attachment_path: str | Path | None = None,
        custom_notes: str = "",
    ) -> str:
        """Generate a corporate, responsive HTML email template for reconciliation notification."""
        status_color = "#28a745" if overall_status == "OK" else "#dc3545"
        status_label = "HOÀN TẤT - KHỚP 100% (OK)" if overall_status == "OK" else "PHÁT HIỆN SAI KHÁC CẦN GIẢI TRÌNH (NG)"
        now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        stats = summary_stats or {}
        total_parts = stats.get("total_parts", 0)
        ok_count = stats.get("ok_count", 0)
        ng_count = stats.get("ng_count", 0)
        missing_count = stats.get("missing_count", 0)
        warning_count = stats.get("warning_count", 0)

        # Build sub-units status table rows
        unit_rows_html = ""
        sub_units = sub_unit_statuses or {
            "LSU": "OK",
            "DLP": "OK",
            "DRUM": "OK",
            "IMAGE": "OK",
            "FUSER": "OK",
            "DP": "OK",
            "ISU": "OK",
            "HONTAI": "OK",
        }
        for idx, (u_name, u_stat) in enumerate(sub_units.items(), start=1):
            badge_color = "#28a745" if u_stat == "OK" else ("#dc3545" if u_stat == "NG" else "#6c757d")
            unit_rows_html += f"""
            <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 8px; text-align: center;">{idx}</td>
                <td style="padding: 8px; font-weight: bold;">{u_name}</td>
                <td style="padding: 8px; text-align: center;">
                    <span style="background-color: {badge_color}; color: #ffffff; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: bold;">
                        {u_stat}
                    </span>
                </td>
            </tr>
            """

        attachment_note = ""
        if attachment_path:
            p_name = Path(attachment_path).name
            attachment_note = f"""
            <p style="margin-top: 15px; font-size: 13px; color: #495057;">
                📎 <strong>Tệp đính kèm:</strong> <code>{p_name}</code> (Vui lòng xem chi tiết 7 Sheet trong tệp báo cáo).
            </p>
            """

        custom_notes_html = ""
        if custom_notes:
            custom_notes_html = f"""
            <div style="margin-top: 15px; padding: 10px; background-color: #fff3cd; border-left: 4px solid #ffeeba; color: #856404; font-size: 13px;">
                <strong>Ghi chú từ Trưởng nhóm:</strong><br/>
                {custom_notes}
            </div>
            """

        action_required_html = ""
        if overall_status == "NG":
            action_required_html = """
            <div style="margin-top: 20px; padding: 12px; background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; color: #721c24;">
                <h4 style="margin: 0 0 8px 0; font-size: 14px;">⚠️ HÀNH ĐỘNG YÊU CẦU:</h4>
                <ul style="margin: 0; padding-left: 20px; font-size: 13px;">
                    <li>Các phụ trách công đoạn có linh kiện <strong>NG</strong> vui lòng kiểm tra lại CTTT, đối chiếu với PLM TC14 và SAP R3.</li>
                    <li>Điền lý do giải thích chi tiết vào cột <strong>Giải thích</strong> và gửi lại trước 16:30 chiều nay.</li>
                </ul>
            </div>
            """

        html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333333; line-height: 1.5; }}
        .container {{ max-width: 680px; margin: 0 auto; border: 1px solid #dcdcdc; border-radius: 6px; overflow: hidden; }}
        .header {{ background-color: #1f497d; color: #ffffff; padding: 18px 24px; }}
        .content {{ padding: 20px 24px; background-color: #ffffff; }}
        .footer {{ background-color: #f4f6f9; color: #6c757d; padding: 12px 24px; font-size: 11px; text-align: center; border-top: 1px solid #e0e0e0; }}
        table {{ width: 100%; border-collapse: collapse; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2 style="margin: 0; font-size: 18px;">HỆ THỐNG ĐỐI SOÁT BOM TỰ ĐỘNG KYOCERA (SSBOM)</h2>
            <p style="margin: 4px 0 0 0; font-size: 12px; opacity: 0.85;">Thông báo kết quả đối soát dữ liệu CTTT vs PLM Active Workspace vs SAP R3</p>
        </div>

        <div class="content">
            <p style="font-size: 14px;">Kính gửi Anh/Chị phụ trách công đoạn và Ban Trưởng nhóm,</p>

            <div style="background-color: #f8f9fa; border-left: 4px solid {status_color}; padding: 12px; margin: 15px 0;">
                <table style="font-size: 13px;">
                    <tr><td style="width: 160px; font-weight: bold;">Mã Model máy:</td><td><strong style="color: #1f497d;">{model_name}</strong></td></tr>
                    <tr><td style="font-weight: bold;">Ngày hiệu lực đối soát:</td><td>{target_date}</td></tr>
                    <tr><td style="font-weight: bold;">Thời gian thực hiện:</td><td>{now_str}</td></tr>
                    <tr><td style="font-weight: bold;">Đánh giá tổng thể:</td>
                        <td><span style="color: {status_color}; font-weight: bold; font-size: 14px;">{status_label}</span></td>
                    </tr>
                </table>
            </div>

            <h3 style="font-size: 14px; color: #1f497d; margin-top: 20px; border-bottom: 2px solid #1f497d; padding-bottom: 4px;">
                📊 TỔNG HỢP CHỈ SỐ ĐỐI SOÁT
            </h3>
            <table style="font-size: 13px; margin-top: 10px;">
                <tr style="background-color: #f1f3f5;">
                    <th style="padding: 6px 10px; text-align: left;">Hạng mục</th>
                    <th style="padding: 6px 10px; text-align: right;">Số lượng</th>
                </tr>
                <tr>
                    <td style="padding: 6px 10px; border-bottom: 1px solid #e0e0e0;">Tổng số linh kiện đối soát:</td>
                    <td style="padding: 6px 10px; text-align: right; border-bottom: 1px solid #e0e0e0;"><strong>{total_parts}</strong></td>
                </tr>
                <tr>
                    <td style="padding: 6px 10px; border-bottom: 1px solid #e0e0e0; color: #28a745;">Linh kiện trùng khớp hoàn toàn (OK):</td>
                    <td style="padding: 6px 10px; text-align: right; border-bottom: 1px solid #e0e0e0; color: #28a745;"><strong>{ok_count}</strong></td>
                </tr>
                <tr>
                    <td style="padding: 6px 10px; border-bottom: 1px solid #e0e0e0; color: #dc3545;">Linh kiện sai khác số lượng/rev (NG):</td>
                    <td style="padding: 6px 10px; text-align: right; border-bottom: 1px solid #e0e0e0; color: #dc3545;"><strong>{ng_count}</strong></td>
                </tr>
                <tr>
                    <td style="padding: 6px 10px; border-bottom: 1px solid #e0e0e0;">Linh kiện thiếu trên CTTT (PLM Omission):</td>
                    <td style="padding: 6px 10px; text-align: right; border-bottom: 1px solid #e0e0e0;"><strong>{missing_count}</strong></td>
                </tr>
                <tr>
                    <td style="padding: 6px 10px; border-bottom: 1px solid #e0e0e0;">Cảnh báo mã Barcode / MSI:</td>
                    <td style="padding: 6px 10px; text-align: right; border-bottom: 1px solid #e0e0e0;"><strong>{warning_count}</strong></td>
                </tr>
            </table>

            <h3 style="font-size: 14px; color: #1f497d; margin-top: 25px; border-bottom: 2px solid #1f497d; padding-bottom: 4px;">
                🏢 TIẾN ĐỘ THEO TỪNG CÔNG ĐOẠN / SUB-UNIT
            </h3>
            <table style="font-size: 12px; margin-top: 10px;">
                <thead>
                    <tr style="background-color: #1f497d; color: #ffffff;">
                        <th style="padding: 8px; width: 40px;">STT</th>
                        <th style="padding: 8px; text-align: left;">Công đoạn</th>
                        <th style="padding: 8px; text-align: center; width: 100px;">Trạng thái</th>
                    </tr>
                </thead>
                <tbody>
                    {unit_rows_html}
                </tbody>
            </table>

            {action_required_html}
            {custom_notes_html}
            {attachment_note}

            <p style="margin-top: 25px; font-size: 13px;">
                Trân trọng,<br/>
                <strong>Bộ phận Kỹ thuật Sản xuất 3 (PE-3 / Cơ khí)</strong><br/>
                Kyocera Document Solutions Vietnam Co., Ltd.
            </p>
        </div>

        <div class="footer">
            Báo cáo được khởi tạo tự động từ hệ thống Kyocera Automated BOM Reconciliation System (SSBOM).<br/>
            Vui lòng không trả lời trực tiếp email này nếu không có thắc mắc kỹ thuật.
        </div>
    </div>
</body>
</html>
"""
        return html

    def create_mail_item(self, preview: EmailPreview) -> Any:
        """Create and populate an Outlook MailItem COM object using win32com."""
        outlook_app = self._get_outlook_application()
        if outlook_app is None:
            raise RuntimeError("Outlook COM Application is not available on this environment.")

        # 0 corresponds to OlItemType.olMailItem
        mail_item = outlook_app.CreateItem(0)
        mail_item.To = preview.to_string
        if preview.recipients_cc:
            mail_item.CC = preview.cc_string
        mail_item.Subject = preview.subject
        mail_item.HTMLBody = preview.html_body

        for att_path in preview.attachment_paths:
            mail_item.Attachments.Add(str(att_path))

        return mail_item

    def preview_in_outlook(self, preview: EmailPreview) -> bool:
        """Open the Outlook Compose dialog displaying the email for human inspection."""
        try:
            mail_item = self.create_mail_item(preview)
            mail_item.Display(False)
            logger.info("Opened Outlook email preview for: %s", preview.subject)
            return True
        except Exception as exc:
            logger.error("Failed to display Outlook email preview: %s", exc)
            return False

    def send_via_outlook(self, preview: EmailPreview) -> bool:
        """Directly send the email via Outlook COM."""
        try:
            mail_item = self.create_mail_item(preview)
            mail_item.Send()
            logger.info("Successfully sent email via Outlook for: %s", preview.subject)
            return True
        except Exception as exc:
            logger.error("Failed to send email via Outlook: %s", exc)
            return False

    def build_task_assignment_email(
        self,
        machine_type: str,
        start_date: str,
        quantity: int | str,
        phase: str,
        deadline_copy: str,
        deadline_verify: str,
        attachment_path: str | Path,
        recipients_to: str | list[str] | None = None,
        recipients_cc: str | list[str] | None = None,
    ) -> EmailPreview:
        """Build email requesting engineers to perform BOM comparison (Mail yêu cầu phụ trách so sánh BOM)."""
        subject = f'So sánh BOM mã hàng mới "{machine_type}"'

        # Process real recipients
        if recipients_to:
            real_to = [r.strip() for r in (recipients_to.split(";") if isinstance(recipients_to, str) else recipients_to) if r.strip()]
        else:
            real_to = [
                "KDTVN-ProductionEngineering_Mecha1_Local_@kdcf.onmicrosoft.com",
                "KDTVN-ProductionEngineering_Mecha2_Local_@kdcf.onmicrosoft.com",
                "KDTVN-ProductionEngineering_Mecha3_Local_@kdcf.onmicrosoft.com",
                "KDTVN-Production-Engineering_Mecha32@dtvn.kyocera.com",
            ]

        if recipients_cc:
            real_cc = [c.strip() for c in (recipients_cc.split(";") if isinstance(recipients_cc, str) else recipients_cc) if c.strip()]
        else:
            real_cc = ["KDTVN-Production-Engineering_Management@kdcf.onmicrosoft.com"]

        # Apply Test Mode if enabled
        if self.test_mode:
            final_to = [self.test_recipient]
            final_cc = []
            test_banner = (
                f'<div style="background-color: #FFF3CD; border: 1px solid #FFEEBA; color: #856404; padding: 10px; margin-bottom: 15px; border-radius: 4px;">'
                f'<strong>[CHẾ ĐỘ THỬ NGHIỆM / TEST MODE]</strong><br>'
                f'Thư này được chuyển hướng tới: <code>{self.test_recipient}</code><br>'
                f'Người nhận thực tế khi chạy chính thức:<br>'
                f'• To: {"; ".join(real_to)}<br>'
                f'• CC: {"; ".join(real_cc)}'
                f'</div>'
            )
        else:
            final_to = real_to
            final_cc = real_cc
            test_banner = ""

        att_path_str = str(attachment_path).replace("\\", "/")

        html_body = f"""<html>
<body style="font-family: 'Times New Roman', Times, serif; font-size: 14px; line-height: 1.6; color: #222;">
{test_banner}
<p>Dear all,</p>

<p>Dự định từ "{start_date}" sẽ sản xuất "{quantity}" mã hàng mới giai đoạn "{phase}" của máy "{machine_type}"</p>

<ol>
    <li><strong>Mọi người kiểm tra các hàng mục xác nhận trước sản xuất: 18 điểm (chi tiết trong file hàng mục chữ ý khi sản xuất mã hàng mới)</strong><br>
    Khi hoàn thành hàng mục này thì check OK vào hàng mục đó. Yêu cầu trước ngày sản xuất 3 ngày phải hoàn thành 18 điểm chú ý.</li>

    <li><strong>Chuẩn bị so sánh BOM</strong><br>
    Mọi người copy danh sách linh kiện và MSI mới nhất vào link dưới<br>
    Hạn hoàn thành copy danh sách linh kiện và MSI : <span style="background-color: yellow;">trong ngày "{deadline_copy}"</span><br>
    Hạn hoàn thành xác nhận sai khác : <span style="background-color: yellow;">trong ngày "{deadline_verify}"</span></li>
</ol>

<p><span style="background-color: yellow;">Ưu tiên hoàn thành so sánh BOM hiện tại để đảm bảo sản xuất tốt xác nhân hiệu quả so sánh BOM tự động.</span><br>
</p>

<p>Link so sánh BOM hiện tại: <a href="file:///{att_path_str}">"{attachment_path}"</a></p>

</body>
</html>"""

        attachments: list[Path] = []
        p = Path(attachment_path)
        if p.exists() and p.is_file():
            attachments.append(p.resolve())

        return EmailPreview(
            subject=subject,
            recipients_to=final_to,
            recipients_cc=final_cc,
            html_body=html_body,
            attachment_paths=attachments,
            overall_status="OK",
        )

    def build_management_review_email(
        self,
        machine_type: str,
        production_date: str,
        attachment_path: str | Path,
        recipients_to: str | list[str] | None = None,
        recipients_cc: str | list[str] | None = None,
    ) -> EmailPreview:
        """Build email requesting managers to review completed BOM (Mail nhờ quản lý check BOM)."""
        subject = f'So sánh BOM mã hàng mới "{machine_type}"'

        if recipients_to:
            real_to = [r.strip() for r in (recipients_to.split(";") if isinstance(recipients_to, str) else recipients_to) if r.strip()]
        else:
            real_to = ["KDTVN-Production-Engineering_Management@kdcf.onmicrosoft.com"]

        if recipients_cc:
            real_cc = [c.strip() for c in (recipients_cc.split(";") if isinstance(recipients_cc, str) else recipients_cc) if c.strip()]
        else:
            real_cc = [
                "KDTVN-ProductionEngineering_Mecha1_Local_@kdcf.onmicrosoft.com",
                "KDTVN-ProductionEngineering_Mecha2_Local_@kdcf.onmicrosoft.com",
                "KDTVN-ProductionEngineering_Mecha3_Local_@kdcf.onmicrosoft.com",
                "KDTVN-Production-Engineering_Mecha32@dtvn.kyocera.com",
            ]

        # Apply Test Mode if enabled
        if self.test_mode:
            final_to = [self.test_recipient]
            final_cc = []
            test_banner = (
                f'<div style="background-color: #FFF3CD; border: 1px solid #FFEEBA; color: #856404; padding: 10px; margin-bottom: 15px; border-radius: 4px;">'
                f'<strong>[CHẾ ĐỘ THỬ NGHIỆM / TEST MODE]</strong><br>'
                f'Thư này được chuyển hướng tới: <code>{self.test_recipient}</code><br>'
                f'Người nhận thực tế khi chạy chính thức:<br>'
                f'• To: {"; ".join(real_to)}<br>'
                f'• CC: {"; ".join(real_cc)}'
                f'</div>'
            )
        else:
            final_to = real_to
            final_cc = real_cc
            test_banner = ""

        att_path_str = str(attachment_path).replace("\\", "/")

        html_body = f"""<html>
<body style="font-family: 'Times New Roman', Times, serif; font-size: 14px; line-height: 1.6; color: #222;">
{test_banner}
<p>Dear các anh quản lý,</p>

<p>Mọi người đã hoàn thành so sánh BOM mã hàng mới "{machine_type}" bên dưới.<br>
Các anh kiểm tra lại giúp em với.</p>

<p>Ngày sản xuất dự kiến: <span style="background-color: yellow;">trong ngày "{production_date}"</span><br> </p> 

<p>Link so sánh BOM hiện tại: <a href="file:///{att_path_str}">"{attachment_path}"</a></p>
</body>
</html>"""

        attachments: list[Path] = []
        p = Path(attachment_path)
        if p.exists() and p.is_file():
            attachments.append(p.resolve())

        return EmailPreview(
            subject=subject,
            recipients_to=final_to,
            recipients_cc=final_cc,
            html_body=html_body,
            attachment_paths=attachments,
            overall_status="OK",
        )

    def _get_outlook_application(self) -> Any:
        """Obtain Outlook Application instance via win32com or custom mock."""
        if self._custom_dispatch is not None:
            return self._custom_dispatch

        try:
            import win32com.client
            return win32com.client.Dispatch("Outlook.Application")
        except Exception as exc:
            logger.warning("win32com.client.Dispatch('Outlook.Application') failed: %s", exc)
            return None


# High-level helper shortcut
def generate_reconciliation_email(
    model_name: str,
    target_date: str,
    overall_status: str,
    recipients_to: str | list[str],
    attachment_path: str | Path | None = None,
    **kwargs: Any,
) -> EmailPreview:
    """Helper to quickly build an EmailPreview notification payload."""
    mailer = OutlookMailer()
    return mailer.build_email_preview(
        model_name=model_name,
        target_date=target_date,
        overall_status=overall_status,
        recipients_to=recipients_to,
        attachment_path=attachment_path,
        **kwargs,
    )
