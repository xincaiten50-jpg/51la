#!/usr/bin/env python3
"""
notifier.py — Send reports via WeCom webhook and/or Email (VI/ZH).
Each channel is independent; failure in one does not affect the other.
Per-channel success/failure is reported to the caller.
"""

import asyncio
import base64
import hashlib
import ssl
import time as time_module
from dataclasses import dataclass
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List, Optional, Tuple

import requests

from config import Config, get_email_recipients, normalize_smtp_password


# ==================== RESULT TYPES ====================

@dataclass
class NotificationResult:
    channel: str           # "wecom" | "gmail"
    success: bool
    message: str
    error: Optional[str] = None


# ==================== WECOM ====================

def send_wecom(
    webhook_url: str,
    screenshot_paths: List[str],
    uv_values: List[int],
    pv_values: List[int],
    excel_path: str,
    machine_names: List[str],
    delay_seconds: float = 1.0,
    data_date: Optional[datetime] = None,  # date this data belongs to
) -> NotificationResult:
    """
    Send 4 screenshots + Excel file + UV/PV summary to WeCom webhook.
    Applies delay_seconds between each API call to avoid rate limiting.
    Returns NotificationResult with success/failure and details.
    """
    print("\n" + "=" * 50)
    print("SENDING TO WECOM")
    print("=" * 50)

    if not webhook_url:
        return NotificationResult(channel="wecom", success=False, message="", error="Webhook URL is empty")

    def _delay():
        if delay_seconds > 0:
            time_module.sleep(delay_seconds)

    try:
        # Separator at start
        _wecom_send_text(webhook_url, "-----------------------------------------------")
        _delay()

        # Send 4 screenshots
        print("\n[INFO] Sending 4 screenshots...")
        for i, screenshot_path in enumerate(screenshot_paths):
            if not screenshot_path or not os.path.exists(screenshot_path):
                print(f"[WARN] Screenshot {i+1} not found: {screenshot_path}")
                continue
            ok = _wecom_send_image(webhook_url, screenshot_path)
            print(f"[{'OK' if ok else 'FAIL'}] Screenshot {i+1}/{len(screenshot_paths)}")
            _delay()

        # Send Excel file
        print("\n[INFO] Sending Excel file...")
        excel_ok = _wecom_send_file(webhook_url, excel_path)
        print(f"[{'OK' if excel_ok else 'FAIL'}] Excel file upload")
        _delay()

        # Send UV/PV summary text
        yesterday = data_date.strftime("%Y-%m-%d") if data_date else _get_yesterday_date_str()
        text = f"📊 Bao cao UV + PV ({yesterday})\n\n"
        for i, uv in enumerate(uv_values):
            pv = pv_values[i]
            text += f"{machine_names[i]}: UV={uv:,} | PV={pv:,}\n"
        text += f"\nTong: UV={sum(uv_values):,} | PV={sum(pv_values):,}"
        _wecom_send_text(webhook_url, text)
        _delay()

        # Separator at end
        _wecom_send_text(webhook_url, "-----------------------------------------------")

        print("\n[INFO] WeCom send complete")
        print("=" * 50)
        return NotificationResult(channel="wecom", success=True, message="All WeCom messages sent")

    except Exception as e:
        return NotificationResult(channel="wecom", success=False, message="", error=str(e)[:100])


def _wecom_send_text(url: str, content: str) -> bool:
    """Send a text message to WeCom webhook. Returns True on success."""
    try:
        resp = requests.post(url, json={"msgtype": "text", "text": {"content": content}}, timeout=15)
        return resp.status_code == 200 and resp.json().get("errcode") == 0
    except Exception:
        return False


def _wecom_send_image(url: str, image_path: str) -> bool:
    """Send an image to WeCom webhook. Returns True on success."""
    try:
        with open(image_path, "rb") as f:
            image_data = f.read()
        md5_hash = hashlib.md5(image_data).hexdigest()
        base64_data = base64.b64encode(image_data).decode("utf-8")
        resp = requests.post(url, json={
            "msgtype": "image",
            "image": {"base64": base64_data, "md5": md5_hash}
        }, timeout=30)
        return resp.json().get("errcode") == 0
    except Exception:
        return False


def _wecom_send_file(url: str, file_path: str) -> bool:
    """Upload and send Excel file to WeCom webhook. Returns True on success."""
    if not os.path.exists(file_path):
        print(f"[WARN] Excel file not found: {file_path}")
        return False
    try:
        with open(file_path, "rb") as f:
            file_data = f.read()
        key = url.split("key=")[1] if "key=" in url else ""
        upload_url = f"https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={key}&type=file"
        resp = requests.post(upload_url, files={"media": ("51la.xlsx", file_data, "application/vnd.ms-excel")}, timeout=30)
        result = resp.json()
        if result.get("errcode") != 0:
            print(f"[ERROR] WeCom file upload failed: {result.get('errmsg')}")
            return False
        media_id = result.get("media_id")
        file_msg = {"msgtype": "file", "file": {"media_id": media_id}}
        resp2 = requests.post(url, json=file_msg, timeout=15)
        return resp2.json().get("errcode") == 0
    except Exception as e:
        print(f"[ERROR] WeCom file send error: {str(e)[:80]}")
        return False


# ==================== EMAIL ====================

def send_email(
    cfg: Config,
    lang: str,
    screenshot_paths: List[str],
    uv_values: List[int],
    pv_values: List[int],
    excel_path: str,
    excel_screenshot_path: Optional[str],
    data_date: Optional[datetime] = None,  # date this data belongs to
    attach_excel_file: bool = True,  # whether to attach Excel file
) -> NotificationResult:
    """
    Send an email report in the specified language (vi or zh).
    Attaches website screenshots inline and optionally Excel file as attachment.
    Returns NotificationResult with per-channel status.
    """
    print("\n" + "=" * 50)
    print(f"SENDING EMAIL ({lang.upper()})")
    print("=" * 50)

    smtp_to_recipients, smtp_cc_recipients = get_email_recipients(cfg, lang)
    smtp_recipients = smtp_to_recipients + smtp_cc_recipients
    smtp_password = normalize_smtp_password(cfg.smtp_password)

    if not cfg.smtp_sender or not smtp_to_recipients:
        return NotificationResult(channel="gmail", success=False, message="", error="SMTP_SENDER or SMTP_TO is empty")
    if not smtp_password:
        return NotificationResult(channel="gmail", success=False, message="", error="SMTP_PASSWORD is empty")

    total_uv = sum(uv_values)
    total_pv = sum(pv_values)
    yesterday = data_date.strftime("%Y-%m-%d") if data_date else _get_yesterday_date_str()

    # Build email — language-specific content
    if lang == "vi":
        subject = f"51.la Bao cao - {yesterday}"
        header = "📊 Bao cao du lieu 51.la"
        intro = f"Dưới đây là bảng tổng hợp UV và PV của 4 máy ngày <b>{yesterday}</b>:"
        th_machine = "May"
        th_uv = "UV (Luot truy cap)"
        th_pv = "PV (Luot xem trang)"
        total_label = "Tong cong"
        screenshot_header = "📸 Screenshot du lieu thuc te:"
        footer_text = "Email nay duoc gui tu he thong tu dong 51.la."
    else:
        subject = f"51.la 数据报告 - {yesterday}"
        header = "📊 51.la 数据报告"
        intro = f"下面是 <b>{yesterday}</b> 的4台机器 UV 和 PV 数据汇总："
        th_machine = "机器"
        th_uv = "UV (访问量)"
        th_pv = "PV (页面浏览量)"
        total_label = "总计"
        screenshot_header = "📸 实际网站数据截图："
        footer_text = "此邮件由 51.la 数据采集系统自动发送。"

    # HTML body
    html_body = f"""<html>
<head>
    <style>
        body {{ font-family: Arial, sans-serif; color: #333; line-height: 1.6; }}
        h2 {{ color: #2c3e50; text-align: center; }}
        table {{ border-collapse: collapse; width: 100%; max-width: 600px; margin: 20px auto; }}
        th, td {{ border: 1px solid #ddd; padding: 12px; text-align: center; }}
        th {{ background-color: #4CAF50; color: white; font-weight: bold; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        .screenshot-section {{ margin-top: 30px; text-align: center; }}
        .screenshot-section img {{ max-width: 90%; border: 1px solid #ddd; border-radius: 8px; margin: 10px 0; }}
        .footer {{ margin-top: 40px; text-align: center; font-size: 14px; color: #666; padding-top: 20px; border-top: 1px solid #eee; }}
        .total-row {{ background-color: #e8f5e8 !important; font-weight: bold; font-size: 16px; }}
    </style>
</head>
<body>
    <h2>{header} - {yesterday}</h2>
    <p>{intro}</p>
    <table>
        <tr>
            <th>{th_machine}</th>
            <th>{th_uv}</th>
            <th>{th_pv}</th>
        </tr>"""

    for i in range(4):
        html_body += f"""<tr>
            <td>{cfg.machine_names[i]}</td>
            <td>{uv_values[i]:,}</td>
            <td>{pv_values[i]:,}</td>
        </tr>"""

    html_body += f"""<tr class="total-row">
        <td><strong>{total_label}</strong></td>
        <td><strong>{total_uv:,}</strong></td>
        <td><strong>{total_pv:,}</strong></td>
    </tr>
    </table>

    <div class="screenshot-section">
        <h3>{screenshot_header}</h3>"""

    for i in range(4):
        html_body += f"""<div>
            <p><strong>{cfg.machine_names[i]}:</strong></p>
            <img src="cid:screenshot_{i+1}" alt="Screenshot {i+1}" />
        </div>"""

    html_body += """    </div>
    <div class="footer">
        <p>""" + footer_text + f"""</p>
        <p>System time: """ + _now_str() + """</p>
    </div>
</body>
</html>"""

    # Build MIME message
    msg = MIMEMultipart("related")
    msg["Subject"] = subject
    msg["From"] = cfg.smtp_sender
    msg["To"] = ", ".join(smtp_to_recipients)
    if smtp_cc_recipients:
        msg["Cc"] = ", ".join(smtp_cc_recipients)

    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # Inline website screenshots
    for i, screenshot_path in enumerate(screenshot_paths):
        if not os.path.exists(screenshot_path):
            print(f"[WARN] Screenshot {i+1} not found: {screenshot_path}")
            continue
        try:
            with open(screenshot_path, "rb") as f:
                img_data = f.read()
            img = MIMEImage(img_data, name=f"screenshot_{i+1}.png")
            img.add_header("Content-ID", f"<screenshot_{i+1}>")
            img.add_header("Content-Disposition", "inline", filename=f"screenshot_{i+1}.png")
            msg.attach(img)
            print(f"[INFO] Attached screenshot: screenshot_{i+1}.png")
        except Exception as e:
            print(f"[ERROR] Failed to attach screenshot {i+1}: {str(e)[:80]}")

    # Excel screenshot (inline, if available)
    if excel_screenshot_path and os.path.exists(excel_screenshot_path):
        try:
            with open(excel_screenshot_path, "rb") as f:
                img_data = f.read()
            img = MIMEImage(img_data, name="excel_screenshot.png")
            img.add_header("Content-ID", "<excel_screenshot>")
            img.add_header("Content-Disposition", "inline", filename="excel_screenshot.png")
            msg.attach(img)
            print("[INFO] Attached Excel screenshot")
        except Exception as e:
            print(f"[ERROR] Failed to attach Excel screenshot: {str(e)[:80]}")

    # Attach Excel file
    if attach_excel_file and os.path.exists(excel_path):
        try:
            with open(excel_path, "rb") as f:
                excel_data = f.read()
        except Exception as e:
            return NotificationResult(channel="gmail", success=False, message="", error=f"Cannot read Excel: {e}")
        attachment = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        attachment.set_payload(excel_data)
        encoders.encode_base64(attachment)
        attachment.add_header("Content-Disposition", "attachment", filename="51la.xlsx")
        msg.attach(attachment)
        print("[INFO] Attached Excel file: 51la.xlsx")
    elif not attach_excel_file:
        print("[INFO] Excel attachment disabled — not attaching")
    else:
        print(f"[WARN] Excel file not found: {excel_path}")

    # Send via SMTP — respects smtp_use_ssl, smtp_use_tls, smtp_auth_required
    context = ssl.create_default_context()
    max_retries = 2

    def _mask_recipients_email(rlist):
        return _mask_recipients(rlist)

    for attempt in range(1, max_retries + 1):
        try:
            print(f"\n[INFO] Connecting to SMTP ({cfg.smtp_provider}: {cfg.smtp_server}:{cfg.smtp_port}), attempt {attempt}...")
            if cfg.smtp_use_ssl:
                # Port 465 with implicit SSL
                with smtplib.SMTP_SSL(cfg.smtp_server, cfg.smtp_port, context=context) as server:
                    server.set_debuglevel(1 if cfg.smtp_debug else 0)
                    if cfg.smtp_auth_required:
                        server.login(cfg.smtp_sender, smtp_password)
                    server.sendmail(cfg.smtp_sender, smtp_recipients, msg.as_string())
            elif cfg.smtp_use_tls:
                # Port 587 with STARTTLS
                with smtplib.SMTP(cfg.smtp_server, cfg.smtp_port) as server:
                    server.set_debuglevel(1 if cfg.smtp_debug else 0)
                    server.ehlo()
                    server.starttls(context=context)
                    server.ehlo()
                    if cfg.smtp_auth_required:
                        server.login(cfg.smtp_sender, smtp_password)
                    server.sendmail(cfg.smtp_sender, smtp_recipients, msg.as_string())
            else:
                # Plain text (rare — only for local testing)
                with smtplib.SMTP(cfg.smtp_server, cfg.smtp_port) as server:
                    server.set_debuglevel(1 if cfg.smtp_debug else 0)
                    if cfg.smtp_auth_required:
                        server.login(cfg.smtp_sender, smtp_password)
                    server.sendmail(cfg.smtp_sender, smtp_recipients, msg.as_string())

            masked_to = _mask_recipients_email(smtp_to_recipients)
            masked_cc = _mask_recipients_email(smtp_cc_recipients) if smtp_cc_recipients else []
            print("[SUCCESS] Email sent!")
            print(f"   To (masked): {', '.join(masked_to)}")
            if masked_cc:
                print(f"   Cc (masked): {', '.join(masked_cc)}")
            print("=" * 50)
            return NotificationResult(channel="gmail", success=True, message=f"Email sent to {smtp_to_recipients}")
        except smtplib.SMTPAuthenticationError as e:
            print(f"[ERROR] SMTP auth failed (attempt {attempt}): {e}")
            if attempt < max_retries:
                time_module.sleep(attempt * 5)
            else:
                return NotificationResult(channel="gmail", success=False, message="", error=f"SMTP auth failed: {e}")
        except Exception as e:
            print(f"[ERROR] SMTP error (attempt {attempt}): {e}")
            if attempt < max_retries:
                time_module.sleep(attempt * 5)
            else:
                return NotificationResult(channel="gmail", success=False, message="", error=str(e)[:100])

    return NotificationResult(channel="gmail", success=False, message="", error="SMTP send failed after retries")


# ==================== FAILURE ALERT ====================

def send_failure_alert(cfg: Config, date_str: str, attempts: int, last_error: Optional[str]) -> bool:
    """
    Send a plain-text alert to the sender address when the scheduled run keeps
    failing after all retries — so the operator notices the same day instead of
    discovering missing report emails days later.
    Best-effort: never raises; returns True if the alert was sent.
    """
    try:
        smtp_password = normalize_smtp_password(cfg.smtp_password)
        if not cfg.smtp_sender or not smtp_password:
            print("[WARN] Failure alert skipped: SMTP sender/password not configured")
            return False

        subject = f"[ALERT] 51.la scraper FAILED {attempts}x on {date_str}"
        body = (
            "The scheduled 51.la report could NOT be completed today.\n\n"
            f"Date: {date_str}\n"
            f"Attempts made: {attempts}\n"
            f"Last error: {last_error or 'unknown'}\n\n"
            "Check the machine, then restart if needed:\n"
            "  pm2 logs 51la-daily --lines 100\n"
            "  pm2 restart 51la-daily\n\n"
            "(This alert is sent by the scraper itself.)"
        )
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = cfg.smtp_sender
        msg["To"] = cfg.smtp_sender

        context = ssl.create_default_context()
        if cfg.smtp_use_ssl:
            server = smtplib.SMTP_SSL(cfg.smtp_server, cfg.smtp_port, context=context)
        else:
            server = smtplib.SMTP(cfg.smtp_server, cfg.smtp_port)
        with server:
            if cfg.smtp_use_tls:
                server.ehlo()
                server.starttls(context=context)
                server.ehlo()
            if cfg.smtp_auth_required:
                server.login(cfg.smtp_sender, smtp_password)
            server.sendmail(cfg.smtp_sender, [cfg.smtp_sender], msg.as_string())

        print(f"[OK] Failure alert sent to {cfg.smtp_sender}")
        return True
    except Exception as e:
        print(f"[WARN] Failure alert could not be sent: {e}")
        return False


# ==================== NOTIFY ALL ====================

def notify_all(
    cfg: Config,
    method: str,         # "gmail" | "wecom" | "both"
    lang: str,           # "vi" | "zh"
    screenshot_paths: List[str],
    uv_values: List[int],
    pv_values: List[int],
    excel_path: str,
    excel_screenshot_path: Optional[str],
    data_date: Optional[datetime] = None,  # date this data belongs to (for email subject/body)
    attach_excel_file: bool = True,  # whether to attach Excel in email
) -> List[NotificationResult]:
    """
    Send notifications via configured channels.
    Returns list of NotificationResult — one per channel attempted.
    Each channel is attempted independently; one failure does not block others.
    """
    results: List[NotificationResult] = []
    delay = getattr(cfg, "wecom_send_delay_seconds", 1.0)

    if method in {"gmail", "both"}:
        result = send_email(cfg, lang, screenshot_paths, uv_values, pv_values, excel_path, excel_screenshot_path, data_date=data_date, attach_excel_file=attach_excel_file)
        results.append(result)

    if method in {"wecom", "both"}:
        result = send_wecom(cfg.wecom_webhook_url, screenshot_paths, uv_values, pv_values, excel_path, cfg.machine_names, delay_seconds=delay, data_date=data_date)
        results.append(result)

    return results


def _mask_recipients(recipients: List[str]) -> List[str]:
    """Mask all but the first and last characters of each email for safe display."""
    masked = []
    for r in recipients:
        if "@" in r and len(r) > 4:
            local, domain = r.split("@", 1)
            masked_local = local[0] + "***" + local[-1] if len(local) > 2 else "***"
            masked.append(f"{masked_local}@{domain}")
        else:
            masked.append("***")
    return masked


def _get_email_recipients_for_dryrun(cfg: Config, lang: str) -> Tuple[List[str], List[str]]:
    """
    Return masked recipient lists for dry-run display.
    Does not call SMTP or send anything.
    """
    to_list, cc_list = get_email_recipients(cfg, lang)
    return _mask_recipients(to_list), _mask_recipients(cc_list)


# ==================== HELPERS ====================

import smtplib
import os
from datetime import datetime


def _get_yesterday_date_str() -> str:
    from datetime import datetime, timedelta
    from zoneinfo import ZoneInfo
    vn_tz = ZoneInfo("Asia/Ho_Chi_Minh")
    now_vn = datetime.now(vn_tz)
    yesterday = (now_vn - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    return yesterday.strftime("%Y-%m-%d")


def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")