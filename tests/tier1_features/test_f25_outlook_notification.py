"""Feature F25: Outlook HTML Email Notification Isolation Tests.

Verifies:
1. HTML template generation summarizing reconciliation results (OK/NG count).
2. Email recipient list formatting (semicolon-separated addresses).
3. Verifying attachment file exists prior to triggering send.
4. Draft mode creation vs direct background dispatch.
5. Graceful handling of Outlook COM unavailability.
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock
import pytest

from src.reporting.outlook_mailer import EmailPreview, OutlookMailer


class TestF25OutlookNotification:
    """Test suite for Feature F25: Outlook HTML Email Notification."""

    def test_f25_html_email_template_rendering(self) -> None:
        """Test 1: Render HTML summary template with counts and status badges."""
        mailer = OutlookMailer()
        preview_ok = mailer.build_email_preview(
            model_name="Virgo",
            target_date="2026-09-17",
            overall_status="OK",
            recipients_to="leader@kyocera.com",
            summary_stats={"total_parts": 100, "ok_count": 100, "ng_count": 0},
        )
        assert "HOÀN TẤT - OK" in preview_ok.subject
        assert "Virgo" in preview_ok.html_body
        assert "100" in preview_ok.html_body

        preview_ng = mailer.build_email_preview(
            model_name="Virgo",
            target_date="2026-09-17",
            overall_status="NG",
            recipients_to="leader@kyocera.com",
            summary_stats={"total_parts": 100, "ok_count": 95, "ng_count": 5},
        )
        assert "CẦN GIẢI TRÌNH - NG" in preview_ng.subject
        assert "NG" in preview_ng.html_body

    def test_f25_recipients_list_formatting(self) -> None:
        """Test 2: Format recipient list with valid delimiter."""
        mailer = OutlookMailer()
        preview = mailer.build_email_preview(
            model_name="Virgo",
            target_date="2026-09-17",
            overall_status="OK",
            recipients_to="leader@kyocera.com; member1@kyocera.com",
            recipients_cc=["lead_cc@kyocera.com"],
        )
        assert preview.to_string == "leader@kyocera.com; member1@kyocera.com"
        assert preview.cc_string == "lead_cc@kyocera.com"

    def test_f25_attachment_file_validation(self, tmp_path: Path) -> None:
        """Test 3: Attachment must exist on disk before creating mail item."""
        valid_file = tmp_path / "report.xlsx"
        valid_file.write_text("DUMMY")

        mailer = OutlookMailer()
        preview = mailer.build_email_preview(
            model_name="Virgo",
            target_date="2026-09-17",
            overall_status="OK",
            recipients_to="leader@kyocera.com",
            attachment_path=valid_file,
        )
        assert len(preview.attachment_paths) == 1

        missing_file = tmp_path / "non_existent.xlsx"
        preview_missing = mailer.build_email_preview(
            model_name="Virgo",
            target_date="2026-09-17",
            overall_status="OK",
            recipients_to="leader@kyocera.com",
            attachment_path=missing_file,
        )
        assert len(preview_missing.attachment_paths) == 0

    def test_f25_draft_mode_vs_direct_send(self) -> None:
        """Test 4: Support both Display (draft review) and Send via Outlook COM."""
        mock_dispatch = MagicMock()
        mock_item = MagicMock()
        mock_dispatch.CreateItem.return_value = mock_item

        mailer = OutlookMailer(com_dispatch=mock_dispatch)
        preview = EmailPreview(
            subject="Test Sub",
            recipients_to=["a@kyocera.com"],
            recipients_cc=[],
            html_body="<p>Test</p>",
        )

        res_preview = mailer.preview_in_outlook(preview)
        assert res_preview is True
        mock_item.Display.assert_called_once()

        res_send = mailer.send_via_outlook(preview)
        assert res_send is True
        mock_item.Send.assert_called_once()

    def test_f25_outlook_com_error_handling(self) -> None:
        """Test 5: Handle Outlook application COM unavailable error cleanly."""
        mock_dispatch = MagicMock()
        mock_dispatch.CreateItem.side_effect = RuntimeError("COM failed")

        mailer = OutlookMailer(com_dispatch=mock_dispatch)
        preview = EmailPreview(
            subject="Test Sub",
            recipients_to=["a@kyocera.com"],
            recipients_cc=[],
            html_body="<p>Test</p>",
        )
        assert mailer.send_via_outlook(preview) is False
        assert mailer.preview_in_outlook(preview) is False
