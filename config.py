#!/usr/bin/env python3
"""
config.py — Load .env, validate, and export typed configuration.
All secrets come from .env only. No hardcoded credentials.
"""

import os
from dataclasses import dataclass, field
from typing import List


def _parse_list(raw: str) -> List[str]:
    """Parse comma-separated list, returning non-empty trimmed items."""
    return [item.strip() for item in raw.split(",") if item.strip()]


def _env(key: str, default: str = "") -> str:
    return os.getenv(key, default)


@dataclass
class Config:
    notification_method: str          # gmail | wecom | both
    email_lang: str                    # vi | zh

    wecom_webhook_url: str

    # SMTP provider fields
    smtp_provider: str                 # gmail | workspace_relay | exmail (legacy QQ)
    smtp_server: str
    smtp_port: int
    smtp_use_tls: bool
    smtp_use_ssl: bool
    smtp_auth_required: bool
    smtp_sender: str
    smtp_password: str
    smtp_debug: bool

    smtp_to_vi: List[str]
    smtp_cc_vi: List[str]
    smtp_to_zh: List[str]
    smtp_cc_zh: List[str]

    excel_path: str
    sheet_name: str

    schedule_hour: int
    schedule_minute: int
    timezone: str

    machine_names: List[str]
    scraper_links: List[str]
    scraper_password: str
    wecom_send_delay_seconds: float = 1.0

    # Monthly workbook (pre-created)
    use_precreated_monthly_files: bool = False
    auto_create_monthly_workbook: bool = False  # opt-in only
    reports_dir: str = "reports"
    report_file_pattern: str = "51la_{YYYY-MM}.xlsx"
    current_workbook_path: str = "51la_current.xlsx"
    copy_to_current: bool = True

    allow_real_email_send: bool = False  # hard safety gate — default OFF

    # Internal state
    _errors: List[str] = field(default_factory=list)
    _email_deprecation_warned: bool = False

    @property
    def errors(self) -> List[str]:
        return self._errors


def load() -> Config:
    """Load and validate configuration from environment variables."""
    errors: List[str] = []

    # Normalize notification method: email → gmail (deprecated)
    raw_method = _env("NOTIFICATION_METHOD", "gmail").strip().lower()
    email_deprecation_warned = False
    if raw_method == "email":
        print("[WARN] --method email is deprecated. Use --method gmail instead.")
        raw_method = "gmail"
        email_deprecation_warned = True

    cfg = Config(
        notification_method=raw_method,
        email_lang=_env("EMAIL_LANG", "zh").strip().lower(),
        wecom_webhook_url=_env("WECOM_WEBHOOK_URL", "").strip(),
        smtp_provider=_env("SMTP_PROVIDER", "gmail").strip().lower(),
        smtp_server=_env("SMTP_SERVER", "").strip(),
        smtp_port=int(_env("SMTP_PORT", "587").strip()),
        smtp_use_tls=_env("SMTP_USE_TLS", "true").strip().lower() in {"1", "true", "yes", "on"},
        smtp_use_ssl=_env("SMTP_USE_SSL", "false").strip().lower() in {"1", "true", "yes", "on"},
        smtp_auth_required=_env("SMTP_AUTH_REQUIRED", "true").strip().lower() in {"1", "true", "yes", "on"},
        smtp_sender=_env("SMTP_SENDER", "").strip(),
        smtp_password=_env("SMTP_PASSWORD", "").strip(),
        smtp_debug=_env("SMTP_DEBUG", "0").strip() in {"1", "true", "yes", "on"},
        smtp_to_vi=_parse_list(_env("SMTP_TO_VI", "")),
        smtp_cc_vi=_parse_list(_env("SMTP_CC_VI", "")),
        smtp_to_zh=_parse_list(_env("SMTP_TO_ZH", "")),
        smtp_cc_zh=_parse_list(_env("SMTP_CC_ZH", "")),
        excel_path=_env("EXCEL_PATH", "51la.xlsx").strip(),
        sheet_name=_env("SHEET_NAME", "Sheet1").strip(),
        schedule_hour=int(_env("SCHEDULE_HOUR", "8").strip()),
        schedule_minute=int(_env("SCHEDULE_MINUTE", "50").strip()),
        timezone=_env("TIMEZONE", "Asia/Ho_Chi_Minh").strip(),
        machine_names=_parse_list(_env("MACHINE_NAMES", "11-02-01,11-02-02,11-02-03,11-02-04")),
        scraper_links=_parse_list(_env("SCRAPER_LINKS", "")),
        scraper_password=_env("SCRAPER_PASSWORD", "").strip(),
        wecom_send_delay_seconds=float(_env("WECOM_SEND_DELAY_SECONDS", "1.0").strip()),
        use_precreated_monthly_files=_env("USE_PRECREATED_MONTHLY_FILES", "false").strip().lower() in {"1", "true", "yes", "on"},
        auto_create_monthly_workbook=_env("AUTO_CREATE_MONTHLY_WORKBOOK", "false").strip().lower() in {"1", "true", "yes", "on"},
        reports_dir=_env("REPORTS_DIR", "reports").strip(),
        report_file_pattern=_env("REPORT_FILE_PATTERN", "51la_{YYYY-MM}.xlsx").strip(),
        current_workbook_path=_env("CURRENT_WORKBOOK_PATH", "51la_current.xlsx").strip(),
        copy_to_current=_env("COPY_TO_CURRENT", "true").strip().lower() in {"1", "true", "yes", "on"},
        allow_real_email_send=_env("ALLOW_REAL_EMAIL_SEND", "false").strip().lower() in {"1", "true", "yes", "on"},
        _email_deprecation_warned=email_deprecation_warned,
    )

    # --- Validate notification method ---
    if cfg.notification_method not in {"gmail", "wecom", "both"}:
        errors.append("NOTIFICATION_METHOD must be gmail, wecom, or both")

    # --- Validate email language ---
    if cfg.email_lang not in {"vi", "zh"}:
        errors.append("EMAIL_LANG must be vi or zh")

    # --- Validate WeCom ---
    if cfg.notification_method in {"wecom", "both"}:
        if not cfg.wecom_webhook_url:
            errors.append("WECOM_WEBHOOK_URL is required when method is wecom or both")

    # --- Validate Gmail/email ---
    if cfg.notification_method in {"gmail", "both"}:
        if not cfg.smtp_server:
            errors.append("SMTP_SERVER is required when method is gmail or both")
        if not cfg.smtp_sender:
            errors.append("SMTP_SENDER is required when method is gmail or both")
        if not cfg.smtp_port:
            errors.append("SMTP_PORT is required when method is gmail or both")
        if cfg.smtp_provider == "gmail" and not cfg.smtp_password:
            errors.append("SMTP_PASSWORD is required when SMTP_PROVIDER=gmail")
        if cfg.smtp_provider == "workspace_relay" and cfg.smtp_auth_required and not cfg.smtp_password:
            errors.append("SMTP_PASSWORD is required when SMTP_PROVIDER=workspace_relay and SMTP_AUTH_REQUIRED=true")
        if not cfg.smtp_use_tls and not cfg.smtp_use_ssl:
            errors.append("Either SMTP_USE_TLS or SMTP_USE_SSL must be enabled")
        # At minimum one TO address must be set for the selected language
        if cfg.email_lang == "vi" and not cfg.smtp_to_vi:
            errors.append("SMTP_TO_VI is required when EMAIL_LANG=vi")
        if cfg.email_lang == "zh" and not cfg.smtp_to_zh:
            errors.append("SMTP_TO_ZH is required when EMAIL_LANG=zh")

    # --- Validate scraper settings ---
    if len(cfg.scraper_links) != 4:
        errors.append(f"SCRAPER_LINKS must have exactly 4 entries, got {len(cfg.scraper_links)}")
    if not cfg.scraper_password:
        errors.append("SCRAPER_PASSWORD is required")

    # --- Validate machine names ---
    if len(cfg.machine_names) != 4:
        errors.append(f"MACHINE_NAMES must have exactly 4 entries, got {len(cfg.machine_names)}: {cfg.machine_names}")

    cfg._errors = errors
    return cfg


def validate_for_run(cfg: Config) -> None:
    """
    Raise ValueError if config has errors.
    Call this after load() to surface all issues before running.
    """
    if cfg.errors:
        raise ValueError("Config validation failed:\n" + "\n".join(f"  - {e}" for e in cfg.errors))


def normalize_smtp_password(password: str) -> str:
    """QQ app passwords may be copied with spaces — strip all whitespace."""
    return "".join(password.split())


def get_email_recipients(cfg: Config, lang: str) -> tuple:
    """
    Return (to_list, cc_list) for the given language.
    lang: 'vi' or 'zh'
    """
    if lang == "vi":
        return cfg.smtp_to_vi, cfg.smtp_cc_vi
    return cfg.smtp_to_zh, cfg.smtp_cc_zh