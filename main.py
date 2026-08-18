#!/usr/bin/env python3
"""
main.py — Entry point for merged 51.la scraper.

Usage:
  python main.py --method email --lang vi
  python main.py --method email --lang zh
  python main.py --method wecom
  python main.py --method both --lang zh
  python main.py --schedule-mode --method both --lang zh
"""

import argparse
import asyncio
from dataclasses import dataclass
from datetime import date, datetime, timedelta
import json
import os
import sys
import time as time_module
from pathlib import Path
from zoneinfo import ZoneInfo

# Set working directory to this file's location
os.chdir(Path(__file__).parent)

from dotenv import load_dotenv

from config import load, validate_for_run, Config
import scraper
import notifier
import workbook_manager


LOCK_FILE = ".running"
RUN_STATE_FILE = ".run_state"
SCREENSHOT_DIR = "screenshots"

# Scheduled-run resilience: catch-up + retry + alert
MAX_RUN_ATTEMPTS = 5          # attempts per day before giving up and alerting
RETRY_DELAY_SECONDS = 15 * 60  # wait between failed attempts


# ==================== RUN CONTEXT ====================

@dataclass
class RunContext:
    """Per-run resolved date/workbook state. Built fresh inside every run_once() call."""
    report_date: datetime             # datetime at 00:00 for the report_date
    report_date_date: date           # pure date object (for filename/path matching)
    workbook_path: Path              # resolved workbook path
    workbook_mode: str               # "precreated" | "auto_create" | "legacy"
    current_copy_path: Path          # 51la_current.xlsx
    workbook_created: bool = False   # True only when auto-create created a new file
    date_source: str = "default-yesterday"  # "default-yesterday" | "--report-date" | "--run-date"


def _resolve_run_context(cfg, args) -> RunContext:
    """
    Build a fresh RunContext for this run.

    Computes report_date, resolves the workbook path, validates the workbook
    month matches report_date's month. Exits on hard failures.

    This MUST be called at the start of every run_once() so that report_date
    and workbook path are recomputed for every scheduled execution.
    """
    # --report-date and --run-date are mutually exclusive
    if args.report_date and args.run_date:
        print()
        print("=" * 60)
        print("ERROR: --report-date and --run-date cannot be used together.")
        print("  --report-date: direct override of report_date")
        print("  --run-date:    simulate current run date (report_date = run_date - 1)")
        print("=" * 60)
        sys.exit(1)

    run_date_parsed = None
    if args.run_date:
        try:
            run_date_parsed = datetime.strptime(args.run_date, "%Y-%m-%d").date()
        except ValueError:
            print(f"[ERROR] Invalid --run-date format: {args.run_date} (expected YYYY-MM-DD)")
            sys.exit(1)

    # Compute report_date
    report_date_date = workbook_manager.get_report_date(args.report_date, run_date_parsed)
    report_date_dt = datetime.combine(report_date_date, datetime.min.time())

    # Resolve workbook
    try:
        workbook_path, mode, was_created = workbook_manager.resolve_workbook_for_report(
            cfg, args, report_date_date
        )
    except workbook_manager.MonthlyWorkbookMissingError as e:
        print()
        print("=" * 60)
        print("BLOCKED:", e)
        print("=" * 60)
        sys.exit(1)

    # Mismatch guard: workbook month must match report_date month
    if mode in {"precreated", "auto_create"}:
        try:
            workbook_manager.validate_workbook_month_match(report_date_date, workbook_path)
        except ValueError as e:
            print()
            print("=" * 60)
            print(str(e))
            print("=" * 60)
            sys.exit(1)

    # Determine date_source label
    if args.report_date:
        date_source = "--report-date"
    elif args.run_date:
        date_source = "--run-date"
    else:
        date_source = "default-yesterday"

    return RunContext(
        report_date=report_date_dt,
        report_date_date=report_date_date,
        workbook_path=workbook_path,
        workbook_mode=mode,
        current_copy_path=Path(cfg.current_workbook_path),
        workbook_created=was_created,
        date_source=date_source,
    )


def _log_runtime_info(cfg, ctx: RunContext) -> None:
    """Log runtime date/workbook info before scraping/writing."""
    process_start = datetime.now()
    try:
        tz = ZoneInfo(cfg.timezone)
        tz_now = datetime.now(tz)
        tz_now_str = tz_now.isoformat()
    except Exception:
        tz_now_str = datetime.now().isoformat()
    print(f"[INFO] Process start time: {process_start.isoformat()}")
    print(f"[INFO] Trigger time: {process_start.isoformat()}")
    print(f"[INFO] System local time: {process_start.isoformat()}")
    print(f"[INFO] Configured timezone: {cfg.timezone}")
    print(f"[INFO] Timezone now: {tz_now_str}")
    print(f"[INFO] Computed report_date: {ctx.report_date.strftime('%Y-%m-%d')}")
    print(f"[INFO] Date source: {ctx.date_source}")
    print(f"[INFO] Workbook mode: {ctx.workbook_mode}")
    print(f"[INFO] Workbook: {ctx.workbook_path}")


# ==================== LOCK UTILITIES ====================

def is_process_alive(pid: int) -> bool:
    """Check if a process with given PID is alive. Works on Windows and macOS/Linux."""
    try:
        import platform
        if platform.system() == "Windows":
            import ctypes
            PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
            SYNCHRONIZE = 0x00100000
            STILL_ACTIVE = 259

            kernel = ctypes.WinDLL("kernel32", use_last_error=True)
            handle = kernel.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE, False, pid)
            if handle == 0:
                return False
            try:
                exit_code = ctypes.DWORD()
                if kernel.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                    return exit_code.value == STILL_ACTIVE
                return False
            finally:
                kernel.CloseHandle(handle)
        else:
            import os
            os.kill(pid, 0)
            return True
    except ProcessLookupError:
        return False  # No such process — lock is stale, safe to remove
    except PermissionError:
        return True  # Process exists but owned by another user
    except Exception:
        return False


def acquire_lock(lock_path: str = LOCK_FILE, mode: str = "once") -> bool:
    """
    Atomically acquire lock using os.O_CREAT|os.O_EXCL.
    Returns False if another instance holds a live lock.
    Removes stale lock if the holding process is dead.
    Stores PID, mode, timestamp, and command in lock file.
    """
    import time
    if os.path.exists(lock_path):
        # Try to read existing lock
        try:
            with open(lock_path, "r") as f:
                lines = f.read().splitlines()
            if lines:
                first_line = lines[0].strip()
                try:
                    lock_pid = int(first_line)
                except ValueError:
                    lock_pid = None

                lock_mode = lines[1].strip() if len(lines) > 1 else "?"
                lock_ts = lines[2].strip() if len(lines) > 2 else "?"
                lock_cmd = lines[3].strip() if len(lines) > 3 else ""

                if lock_pid is not None and is_process_alive(lock_pid):
                    print(f"[ERROR] Another instance is already running (PID {lock_pid}, mode={lock_mode}, started={lock_ts}).")
                    print(f"[INFO] Remove {lock_path} if this is incorrect.")
                    return False
                else:
                    # Stale lock — remove it
                    print(f"[INFO] Removing stale lock (PID {lock_pid} is not alive)")
                    os.remove(lock_path)
        except Exception as e:
            print(f"[WARN] Could not read lock file: {e}. Removing it.")
            try:
                os.remove(lock_path)
            except Exception:
                pass

    # Acquire atomically
    try:
        fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        lock_time = time.strftime("%Y-%m-%d %H:%M:%S")
        cmd = " ".join(sys.argv)
        content = f"{os.getpid()}\n{mode}\n{lock_time}\n{cmd}"
        os.write(fd, content.encode("utf-8"))
        os.close(fd)
        print(f"[INFO] Lock acquired: PID={os.getpid()} mode={mode}")
        return True
    except FileExistsError:
        print(f"[ERROR] Another instance acquired lock just now.")
        return False
    except Exception as e:
        print(f"[WARN] Could not create lock file: {e}. Proceeding without lock.")
        return True  # Don't block on lock failure


def release_lock(lock_path: str = LOCK_FILE):
    """Remove lock file on normal exit or interrupt."""
    try:
        if os.path.exists(lock_path):
            os.remove(lock_path)
            print(f"[INFO] Lock released.")
    except Exception:
        pass


# ==================== RUN STATE (scheduler) ====================

def _read_run_state() -> dict:
    """Read persisted scheduler state: last_success date + per-day attempt counters.

    Survives process restarts so a machine reboot / PM2 restart can never
    duplicate a successful run or forget that today's run is still missing.
    """
    try:
        with open(RUN_STATE_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return {}


def _write_run_state(state: dict) -> None:
    try:
        with open(RUN_STATE_FILE, "w") as f:
            json.dump(state, f)
    except Exception as e:
        print(f"[WARN] Could not write {RUN_STATE_FILE}: {e}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="51.la UV/PV scraper — sends WeCom and/or Gmail reports",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python main.py --method gmail --lang zh       # Full flow: scrape + Excel + Gmail
  python main.py --method gmail --dry-run       # Dry run: build payloads, no send
  python main.py --method gmail --no-notification  # Scrape + Excel only (no email)
  python main.py --method gmail --dry-run --excel-path test_51la.xlsx  # Dry run with test workbook
  python main.py --schedule-mode --method gmail --lang zh  # Scheduled daily Gmail"""

    )
    parser.add_argument(
        "--method",
        choices=["gmail", "wecom", "both", "email"],
        required=False,
        help="Notification method: gmail, wecom, or both (email is deprecated — use gmail)"
    )
    parser.add_argument(
        "--lang",
        choices=["vi", "zh"],
        required=False,
        help="Email language: vi (Vietnamese) or zh (Chinese)"
    )
    parser.add_argument(
        "--schedule-mode",
        action="store_true",
        help="Run in scheduled mode (daily at 8:50 AM Vietnam time)"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Build notification payloads but do not send them"
    )
    parser.add_argument(
        "--no-excel-write",
        action="store_true",
        help="Skip writing to Excel file"
    )
    parser.add_argument(
        "--skip-scrape",
        action="store_true",
        help="Skip scraping (use existing screenshots if available)"
    )
    parser.add_argument(
        "--allow-real-email",
        action="store_true",
        help="Allow real email sending (requires ALLOW_REAL_EMAIL_SEND=true in .env)"
    )
    parser.add_argument(
        "--excel-path",
        type=str,
        default=None,
        help="Override Excel file path (default: from .env)"
    )
    parser.add_argument(
        "--no-notification",
        action="store_true",
        help="Scrape and write Excel but do not send any notifications"
    )
    parser.add_argument(
        "--monthly-workbook",
        action="store_true",
        help="Use monthly workbook mode (reports/51la_YYYY-MM.xlsx)"
    )
    parser.add_argument(
        "--legacy-excel",
        action="store_true",
        help="Force legacy single-workbook mode (use EXCEL_PATH)"
    )
    parser.add_argument(
        "--template-path",
        type=str,
        default=None,
        help="Override template workbook path"
    )
    parser.add_argument(
        "--reports-dir",
        type=str,
        default=None,
        help="Override reports directory"
    )
    parser.add_argument(
        "--report-date",
        type=str,
        default=None,
        help="Direct override for report date (YYYY-MM-DD). Use this to test specific months."
    )
    parser.add_argument(
        "--run-date",
        type=str,
        default=None,
        help="Simulate current run date for computing report_date=yesterday (YYYY-MM-DD). Do not use with --report-date."
    )
    parser.add_argument(
        "--copy-to-current",
        action="store_true",
        default=None,
        help="Enable copy of monthly workbook to 51la_current.xlsx"
    )
    parser.add_argument(
        "--no-copy-to-current",
        action="store_true",
        default=None,
        help="Disable copy of monthly workbook to 51la_current.xlsx"
    )
    parser.add_argument(
        "--precreated-monthly",
        action="store_true",
        help="Use pre-created monthly workbooks from reports/ (default when USE_PRECREATED_MONTHLY_FILES=true in .env)"
    )
    parser.add_argument(
        "--auto-create-monthly",
        action="store_true",
        help="Allow auto-creation of monthly workbooks (not recommended — use pre-created files instead)"
    )
    parser.add_argument(
        "--report-file-pattern",
        type=str,
        default=None,
        help="Override report file pattern (default: 51la_{YYYY-MM}.xlsx)"
    )
    return parser.parse_args()


async def run_once(cfg: Config, args: argparse.Namespace) -> list:
    """Execute one scrape → Excel → notify cycle. Returns notification results (empty if dry-run).

    The RunContext is built fresh at the start of this call so that report_date
    and workbook path are recomputed for every scheduled execution. This is the
    critical fix for the PM2 daily/--schedule-mode date-staleness bug.
    """
    # Build per-run context (computes report_date, resolves workbook, validates month match)
    ctx = _resolve_run_context(cfg, args)

    # Log runtime date/workbook info BEFORE scraping/writing
    _log_runtime_info(cfg, ctx)

    dry_run = getattr(args, "dry_run", False)
    no_notification = getattr(args, "no_notification", False)
    no_excel_write = getattr(args, "no_excel_write", False)
    skip_scrape = getattr(args, "skip_scrape", False)

    if no_notification:
        mode_label = "NOTIFICATION SKIPPED"
    elif dry_run:
        mode_label = "DRY-RUN"
    else:
        mode_label = "ONCE MODE"
    print("=" * 60)
    print(f"51.LA SCRAPER — {mode_label}")
    print("=" * 60)
    print(f"Method: {args.method or cfg.notification_method}")
    lang = args.lang or cfg.email_lang
    print(f"Language: {lang}")
    print(f"Excel: {ctx.workbook_path}")
    print(f"Workbook mode: {ctx.workbook_mode}")
    if dry_run:
        print("[DRY-RUN] Notifications will NOT be sent")
    print("=" * 60)

    # Configure scraper
    scraper.configure(
        links=cfg.scraper_links,
        password=cfg.scraper_password,
        screenshot_dir=SCREENSHOT_DIR,
    )

    # Step 1: Scrape (or skip)
    if skip_scrape:
        print("\n[Step 1] SKIPPED (--skip-scrape)")
        print("[WARN] Using existing screenshot files if available")
        # Return empty results in skip mode
        from dataclasses import replace
        report = None
    else:
        # Browser self-heal: verify Playwright Chromium before scraping,
        # reinstall automatically if a cache cleaner removed it
        browser_ok, browser_msg = await scraper.ensure_browser_available()
        if not browser_ok:
            raise RuntimeError(f"Playwright browser unavailable: {browser_msg}")
        print(f"[OK] Browser check: {browser_msg}")

        print("\n[Step 1] Scraping UV/PV from 4 pages...")
        report = await scraper.scrape_all_uv(data_date=ctx.report_date)

    # Step 2: Display results
    scrape_errors = []
    if report is not None:
        print("\nSCRAPING RESULTS:")
        for i, uv in enumerate(report.uv_results):
            pv = report.pv_results[i]
            col = scraper.COLUMNS[i]
            machine = cfg.machine_names[i] if i < len(cfg.machine_names) else f"Machine {i+1}"
            status = report.statuses[i]
            if status == "PASS":
                ok_mark = "✓ PASS"
                print(f"  [{ok_mark}] {machine} (Column {col}): UV={uv:,} PV={pv:,}")
            elif status == "NO_DATA":
                ok_mark = "⚠ NO_DATA"
                print(f"  [{ok_mark}] {machine} (Column {col}): UV=0 PV=0 (51.la shows zero for this date)")
            else:  # SCRAPE_ERROR
                ok_mark = "✗ SCRAPE_ERROR"
                scrape_errors.append(i)
                err_msg = ""
                print(f"  [{ok_mark}] {machine} (Column {col}): UV={uv:,} PV={pv:,} [selector failed]")

        if scrape_errors:
            print(f"\n⚠ PARTIAL FAILURE: {len(scrape_errors)} page(s) had SCRAPE_ERROR.")
            print("  Affected pages:", [f"#{i+1} ({cfg.machine_names[i]})" for i in scrape_errors])
            print("  Excel is still written below (failed cells preserved); notification is skipped.")
    else:
        print("\n[SKIP] No scrape report available (--skip-scrape mode)")

    # Step 3: Write to Excel
    if no_excel_write:
        print("\n[Step 2] SKIPPED (--no-excel-write)")
        excel_write_ok = True
        excel_msg = "(no write)"
    elif report is not None:
        print("\n[Step 2] Writing to Excel...")
        allow_create = (ctx.workbook_mode != "precreated")
        excel_write_ok, excel_msg = scraper.write_uv_to_excel(
            report.uv_results,
            str(ctx.workbook_path),
            cfg.sheet_name,
            report.statuses,
            report_date=report.data_date or ctx.report_date,  # use data_date if available, fallback to report_date
            allow_create_date_row=allow_create,
        )
        if not excel_write_ok:
            print(f"[ERROR] Excel write failed: {excel_msg}")
        else:
            print(f"[OK] {excel_msg}")
    else:
        print("\n[Step 2] SKIPPED (no scrape report available)")
        excel_write_ok = True
        excel_msg = "(no write)"

    # Any SCRAPE_ERROR → abort before notification so the scheduler retries,
    # instead of emailing a report with zeros for the failed pages. Excel has
    # already been saved above (SCRAPE_ERROR cells preserve prior values), so
    # data scraped in this attempt is not lost.
    if scrape_errors:
        failed_pages = ", ".join(f"#{i+1} ({cfg.machine_names[i]})" for i in scrape_errors)
        raise RuntimeError(
            f"{len(scrape_errors)}/{len(report.uv_results)} pages returned SCRAPE_ERROR "
            f"({failed_pages}) — Excel saved; aborting before notification for retry"
        )

    # Step 4: Excel screenshot
    excel_screenshot_path = ""
    excel_screenshot_ok = False
    if report is not None and args.method in {"gmail", "both"}:
        screenshot_path = os.path.join(SCREENSHOT_DIR, "excel_data.png")
        excel_screenshot_ok, excel_msg = scraper.capture_excel_as_image(
            str(ctx.workbook_path), screenshot_path, cfg.sheet_name
        )
        if excel_screenshot_ok:
            excel_screenshot_path = screenshot_path
            print(f"[OK] Excel screenshot: {excel_screenshot_path}")
        else:
            print(f"[WARN] Excel screenshot skipped: {excel_msg}")

    # Step 5: Notify (or dry-run/no-notification summary)
    results = []
    if dry_run:
        if no_notification:
            print("\n[Step 3] NOTIFICATION SKIPPED — Excel was updated but no email was sent.")
            method = args.method or cfg.notification_method
            print(f"  Method: {method}")
            print(f"  Excel: {ctx.workbook_path} (updated)")
            if report:
                print(f"  Scraped: {sum(1 for s in report.screenshot_paths if s)} screenshots")
        else:
            print("\n[Step 3] DRY-RUN — Notification plan:")
            method = args.method or cfg.notification_method
            if method in {"gmail", "both"}:
                to_list, cc_list = notifier._get_email_recipients_for_dryrun(cfg, lang)
                print(f"  GMAIL [{lang}]:")
                print(f"    To (masked): {to_list}")
                print(f"    Cc (masked): {cc_list}")
                print(f"    Attachments: {ctx.workbook_path.name} + {sum(1 for s in report.screenshot_paths if s) if report else 0} website screenshots")
            if method in {"wecom", "both"}:
                print(f"  WECOM:")
                print(f"    Webhook: https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=***")
                print(f"    Attachments: {ctx.workbook_path.name} + {sum(1 for s in report.screenshot_paths if s) if report else 0} website screenshots")
                print(f"    Per-page UV/PV:")
                if report:
                    for i in range(len(report.uv_results)):
                        print(f"      {cfg.machine_names[i]}: UV={report.uv_results[i]:,} PV={report.pv_results[i]:,}")
            print("\n[DRY-RUN] No notifications were sent.")
    else:
        print("\n[Step 3] Sending notifications...")
        results = notifier.notify_all(
            cfg=cfg,
            method=args.method or cfg.notification_method,
            lang=lang,
            screenshot_paths=report.screenshot_paths if report else [],
            uv_values=report.uv_results if report else [],
            pv_values=report.pv_results if report else [],
            excel_path=str(ctx.workbook_path),
            excel_screenshot_path=excel_screenshot_path,
            data_date=report.data_date if report else None,
            attach_excel_file=cfg.attach_excel_file,
        )

    # Step 6: Copy to current (after successful Excel write, in precreated/auto_create mode)
    if (
        ctx.workbook_mode in {"precreated", "auto_create"}
        and cfg.copy_to_current
        and not no_excel_write
        and excel_write_ok
    ):
        workbook_manager.copy_to_current(ctx.workbook_path, ctx.current_copy_path)

    return results


def print_summary(results: list, elapsed: float, dry_run: bool = False, no_notification: bool = False):
    """Print final execution summary."""
    label = ""
    if no_notification:
        label = " (NOTIFICATION SKIPPED)"
    elif dry_run:
        label = " (DRY-RUN)"
    print("\n" + "=" * 60)
    print("EXECUTION SUMMARY" + label)
    print("=" * 60)
    if no_notification:
        print("  [NOTIFICATION SKIPPED] Excel was updated. No email sent.")
    elif dry_run:
        print("  [DRY-RUN] No channels attempted.")
    else:
        for r in results:
            status = "✓ SUCCESS" if r.success else "✗ FAILED"
            print(f"  {r.channel.upper()}: {status}")
            if r.error:
                print(f"    Error: {r.error}")
            else:
                print(f"    {r.message}")
    print(f"\nElapsed: {elapsed:.1f}s")
    print("=" * 60)


def run_scheduled(cfg: Config, args: argparse.Namespace):
    """Daily scheduler at SCHEDULE_HOUR:SCHEDULE_MINUTE with catch-up + retry + alert.

    Each scheduled run calls run_once(cfg, args), which builds a fresh RunContext
    so that report_date and workbook path are recomputed for every iteration —
    NOT captured at scheduler startup.

    Resilience (see .run_state file):
    - CATCH-UP: fires when now >= schedule time and today's run has not succeeded
      yet — so a machine restart / sleep / PM2 restart that misses the exact
      scheduled minute still runs the report the same day.
    - RETRY: a failed run (exception, all-pages SCRAPE_ERROR, or all notification
      channels failing) is retried up to MAX_RUN_ATTEMPTS times, spaced
      RETRY_DELAY_SECONDS apart.
    - ALERT: when attempts are exhausted, sends an email to the sender address
      so the operator finds out the same day.

    .env is reloaded on every iteration so config changes take effect without restart.
    """
    # Resolve .env path for reload on each iteration
    script_dir = Path(__file__).parent
    env_path = script_dir / ".env"
    dotenv_path = str(env_path) if env_path.exists() else None
    print("=" * 60)
    print("51.LA SCRAPER — SCHEDULED MODE")
    print("=" * 60)
    print(f"Schedule: {cfg.schedule_hour:02d}:{cfg.schedule_minute:02d} ({cfg.timezone}) — catch-up enabled")
    print(f"Retry: up to {MAX_RUN_ATTEMPTS} attempts/day, {RETRY_DELAY_SECONDS // 60} min apart, then alert email")
    print(f"Method: {args.method or cfg.notification_method}")
    print(f"Lang: {args.lang or cfg.email_lang}")
    print("[INFO] Press Ctrl+C to stop")
    print("=" * 60)

    # First deployment: seed state with today so the scheduler does not fire a
    # catch-up run for a slot that may have been handled before this change.
    state = _read_run_state()
    if "last_success" not in state:
        try:
            tz = ZoneInfo(cfg.timezone)
            today_str = datetime.now(tz).strftime("%Y-%m-%d")
        except Exception:
            today_str = datetime.now().strftime("%Y-%m-%d")
        _write_run_state({"last_success": today_str, "attempts": 0, "attempt_date": today_str})
        print(f"[INFO] Initialized {RUN_STATE_FILE} with today ({today_str}) — no catch-up on first start.")

    next_attempt_at = 0.0  # epoch seconds; 0 = no retry backoff pending

    while True:
        try:
            # Reload .env on every iteration so config changes take effect without restart
            load_dotenv(dotenv_path=dotenv_path, override=True)
            cfg = load()

            tz = ZoneInfo(cfg.timezone)
            now_tz = datetime.now(tz)
            today_str = now_tz.strftime("%Y-%m-%d")

            state = _read_run_state()
            # Reset attempt counter on a new day
            if state.get("attempt_date") != today_str:
                state = {"last_success": state.get("last_success"), "attempts": 0, "attempt_date": today_str}
                _write_run_state(state)
            attempts = int(state.get("attempts") or 0)
            ran_today = state.get("last_success") == today_str

            target = now_tz.replace(hour=cfg.schedule_hour, minute=cfg.schedule_minute, second=0, microsecond=0)
            tomorrow_target = target + timedelta(days=1)

            if (not ran_today) and now_tz >= target \
                    and attempts < MAX_RUN_ATTEMPTS \
                    and time_module.time() >= next_attempt_at:
                is_catchup = now_tz > target + timedelta(minutes=2)
                label = "CATCH-UP" if is_catchup else "scheduled"
                print(f"\n[{now_tz.strftime('%Y-%m-%d %H:%M:%S')}] Starting {label} run (attempt {attempts + 1}/{MAX_RUN_ATTEMPTS})...")
                start = time_module.time()
                failure_reason = None
                try:
                    results = asyncio.run(run_once(cfg, args))
                    elapsed = time_module.time() - start
                    dry_run = getattr(args, "dry_run", False)
                    no_notification = getattr(args, "no_notification", False)
                    print_summary(results, elapsed, dry_run=dry_run, no_notification=no_notification)
                    # Treat "all notification channels failed" as a failed run
                    if results and not any(r.success for r in results):
                        failure_reason = "; ".join((r.error or r.message) for r in results)[:300]
                        print(f"\n[ERROR] All notification channels failed: {failure_reason}")
                except Exception as e:
                    failure_reason = str(e)[:300]
                    print(f"\n[ERROR] Run failed: {failure_reason}")

                if failure_reason is None:
                    _write_run_state({"last_success": today_str, "attempts": 0, "attempt_date": today_str})
                    print(f"[{datetime.now(tz).strftime('%Y-%m-%d %H:%M:%S')}] Run succeeded — next run tomorrow at {cfg.schedule_hour:02d}:{cfg.schedule_minute:02d}.")
                else:
                    attempts += 1
                    _write_run_state({"last_success": state.get("last_success"), "attempts": attempts, "attempt_date": today_str})
                    if attempts >= MAX_RUN_ATTEMPTS:
                        print(f"[ERROR] Gave up after {attempts} attempts — sending failure alert email.")
                        notifier.send_failure_alert(cfg, today_str, attempts, failure_reason)
                    else:
                        next_attempt_at = time_module.time() + RETRY_DELAY_SECONDS
                        print(f"[INFO] Retry #{attempts + 1}/{MAX_RUN_ATTEMPTS} in {RETRY_DELAY_SECONDS // 60} minutes.")
                time_module.sleep(5)
                continue

            # Waiting — display what we're waiting for
            if now_tz < target:
                next_event, note = target, "next run"
            elif ran_today:
                next_event, note = tomorrow_target, "next run"
            elif attempts >= MAX_RUN_ATTEMPTS:
                next_event, note = tomorrow_target, "next run (gave up today)"
            else:
                next_event, note = tomorrow_target, "retry"
                if next_attempt_at > time_module.time():
                    next_event = datetime.fromtimestamp(next_attempt_at, tz)

            wait_secs = max(0.0, (next_event - now_tz).total_seconds())
            wait_h = int(wait_secs // 3600)
            wait_m = int((wait_secs % 3600) // 60)
            print(f"\r[{now_tz.strftime('%H:%M')}] Waiting for {note}: {wait_h}h {wait_m}m...", end="", flush=True)
            time_module.sleep(30)

        except KeyboardInterrupt:
            print("\n\n[INFO] Scheduler stopped by user.")
            break
        except Exception as e:
            print(f"\n[ERROR] Scheduler error: {e}")
            time_module.sleep(60)


def main():
    # Find .env file (next to this script)
    script_dir = Path(__file__).parent
    env_path = script_dir / ".env"
    dotenv_path = str(env_path) if env_path.exists() else None
    load_dotenv(dotenv_path=dotenv_path)

    args = parse_args()

    # Load config
    cfg = load()

    # Normalize deprecated --method email → gmail
    if args.method == "email":
        print("[WARN] --method email is deprecated. Use --method gmail instead.")
        args.method = "gmail"

    # Mutual exclusion: --precreated-monthly and --legacy-excel
    if getattr(args, "precreated_monthly", False) and getattr(args, "legacy_excel", False):
        print()
        print("=" * 60)
        print("ERROR: --precreated-monthly and --legacy-excel cannot be used together.")
        print("  --precreated-monthly: use pre-created monthly workbook from reports/")
        print("  --legacy-excel:        use legacy single workbook (EXCEL_PATH)")
        print("=" * 60)
        sys.exit(1)

    # Apply CLI args as overrides to config
    if args.method:
        cfg.notification_method = args.method
    if args.lang:
        cfg.email_lang = args.lang
    if args.excel_path:
        cfg.excel_path = args.excel_path

    # Override config with CLI flags
    if args.template_path:
        cfg.template_path = args.template_path
    if args.reports_dir:
        cfg.reports_dir = args.reports_dir
    if args.report_file_pattern:
        cfg.report_file_pattern = args.report_file_pattern
    if args.copy_to_current:
        cfg.copy_to_current = True
    if args.no_copy_to_current:
        cfg.copy_to_current = False

    # NOTE: report_date and workbook path are NO LONGER computed here.
    # They are recomputed inside _resolve_run_context() at the start of every
    # run_once() call — including every scheduled iteration in --schedule-mode.
    # This is the fix for the PM2 daily date-staleness bug.

    # Determine workbook mode (stable across runs — does not depend on report_date).
    # Used only by the safety gates below; the actual workbook path is resolved
    # per-run by run_once() → _resolve_run_context().
    workbook_mode = workbook_manager.get_workbook_mode(cfg, args)

    # --no-notification means scrape + write Excel but no notifications
    if getattr(args, "no_notification", False):
        args.dry_run = True

    # Safety gate: check before running if real email would be attempted
    dry_run = getattr(args, "dry_run", False)
    allow_real_email = getattr(args, "allow_real_email", False)
    method = cfg.notification_method
    no_excel_write = getattr(args, "no_excel_write", False)

    # BLOCK: real email + --no-excel-write (stale attachment risk)
    if not dry_run and method in {"gmail", "both"} and no_excel_write:
        if workbook_mode in {"precreated", "auto_create"}:
            print()
            print("=" * 60)
            print("BLOCKED: Cannot send real Gmail because the monthly workbook was not updated in this run.")
            print("  Monthly workbook is stale (--no-excel-write was used).")
            print("  Recipients would see an outdated Excel attachment.")
            print()
            print("  To update monthly workbook without sending email, use:")
            print("    --no-notification --precreated-monthly")
            print()
            print("  To run full flow (scrape + Excel + Gmail):")
            print("    --method gmail --precreated-monthly --allow-real-email")
            print("=" * 60)
        else:
            print()
            print("=" * 60)
            print("BLOCKED: Real Gmail send cannot use --no-excel-write.")
            print("  The Excel attachment would be stale (not updated).")
            print("  Recipients would see old data in the attached Excel.")
            print()
            print("  To update Excel without sending email, use:")
            print("    --no-notification --excel-path 51la.xlsx")
            print()
            print("  To run full flow (scrape + Excel + Gmail):")
            print("    --method gmail --allow-real-email")
            print("=" * 60)
        sys.exit(1)

    if not dry_run and method in {"gmail", "both"}:
        if not cfg.allow_real_email_send:
            print()
            print("=" * 60)
            print("BLOCKED: Real email sending is disabled.")
            print("  ALLOW_REAL_EMAIL_SEND is not set to true in .env.")
            print("  Use --dry-run to test without sending real email.")
            print("  Or set ALLOW_REAL_EMAIL_SEND=true in .env and pass --allow-real-email.")
            print("=" * 60)
            sys.exit(1)
        if not allow_real_email:
            print()
            print("=" * 60)
            print("BLOCKED: Real email sending requires --allow-real-email flag.")
            print("  ALLOW_REAL_EMAIL_SEND=true is set in .env, but --allow-real-email was not passed.")
            print("  Use --dry-run to test without sending real email.")
            print("=" * 60)
            sys.exit(1)

    # Validate full config
    try:
        validate_for_run(cfg)
    except ValueError as e:
        print(f"[CONFIG ERROR] {e}")
        sys.exit(1)

    # Acquire lock
    lock_mode = "scheduled" if args.schedule_mode else "once"
    if not acquire_lock(mode=lock_mode):
        sys.exit(1)

    try:
        if args.schedule_mode:
            run_scheduled(cfg, args)
        else:
            start = time_module.time()
            try:
                results = asyncio.run(run_once(cfg, args))
            except workbook_manager.DateRowNotFoundError as e:
                print()
                print("=" * 60)
                print("BLOCKED:", e)
                print("=" * 60)
                sys.exit(1)
            elapsed = time_module.time() - start

            print_summary(results, elapsed, dry_run=dry_run, no_notification=getattr(args, "no_notification", False))
    finally:
        release_lock()


if __name__ == "__main__":
    # Set UTF-8 output encoding for Windows
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)

    main()