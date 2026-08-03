#!/usr/bin/env python3
from __future__ import annotations

import argparse
import signal
import subprocess
import sys
import time
from pathlib import Path

from gallery_sync_pipeline import CATEGORY_DIRS, ROOT, is_gallery_asset

DEFAULT_LOG_FILE = ROOT / "scripts" / "logs" / "gallery-watch.log"
PIPELINE_SCRIPT = ROOT / "scripts" / "gallery_sync_pipeline.py"

STOP_REQUESTED = False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Watch Website 2 gallery folders and automatically trigger the sync pipeline when new photos arrive."
    )
    parser.add_argument(
        "--publish",
        action="store_true",
        help="After syncing, also commit and push the gallery changes so GitHub Actions deploys automatically.",
    )
    parser.add_argument(
        "--commit-message",
        default="Auto-sync gallery photo folders",
        help="Commit message to use together with --publish.",
    )
    parser.add_argument(
        "--skip-optimize",
        action="store_true",
        help="Pass through to the sync pipeline to skip optimized homepage image generation.",
    )
    parser.add_argument(
        "--skip-check",
        action="store_true",
        help="Pass through to the sync pipeline to skip final verification.",
    )
    parser.add_argument(
        "--full-site-check",
        action="store_true",
        help="Pass through to the sync pipeline to also validate the broader site links.",
    )
    parser.add_argument(
        "--poll-interval",
        type=float,
        default=2.0,
        help="How often to re-scan the folders in seconds. Default: 2.0",
    )
    parser.add_argument(
        "--settle-seconds",
        type=float,
        default=6.0,
        help="How long files must stay unchanged before the sync runs. Default: 6.0",
    )
    parser.add_argument(
        "--retry-seconds",
        type=float,
        default=20.0,
        help="How long to wait before retrying if the pipeline fails. Default: 20.0",
    )
    parser.add_argument(
        "--run-on-start",
        action="store_true",
        help="Run the sync pipeline immediately once when the watcher starts.",
    )
    parser.add_argument(
        "--max-runtime",
        type=float,
        help="Optional test mode. Stop automatically after this many seconds.",
    )
    parser.add_argument(
        "--log-file",
        default=str(DEFAULT_LOG_FILE),
        help=f"Append watcher logs to this file. Default: {DEFAULT_LOG_FILE}",
    )
    return parser.parse_args()


def request_stop(_signum: int, _frame: object) -> None:
    global STOP_REQUESTED
    STOP_REQUESTED = True


def install_signal_handlers() -> None:
    signal.signal(signal.SIGINT, request_stop)
    signal.signal(signal.SIGTERM, request_stop)


def append_log(log_path: Path, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(message + "\n")


def log(message: str, log_path: Path) -> None:
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {message}"
    print(line, flush=True)
    append_log(log_path, line)


def watched_directories() -> list[Path]:
    return [ROOT / rel for rel in CATEGORY_DIRS.values()]


def build_snapshot() -> dict[str, tuple[int, int]]:
    snapshot: dict[str, tuple[int, int]] = {}

    for directory in watched_directories():
        if not directory.exists():
            continue
        for file in sorted(directory.iterdir(), key=lambda item: item.name.lower()):
            if not is_gallery_asset(file):
                continue
            try:
                stat = file.stat()
            except FileNotFoundError:
                continue
            relative = file.relative_to(ROOT).as_posix()
            snapshot[relative] = (stat.st_size, stat.st_mtime_ns)

    return snapshot


def describe_changes(previous: dict[str, tuple[int, int]], current: dict[str, tuple[int, int]]) -> str:
    previous_keys = set(previous)
    current_keys = set(current)

    added = current_keys - previous_keys
    removed = previous_keys - current_keys
    changed = {key for key in previous_keys & current_keys if previous[key] != current[key]}

    return f"+{len(added)} new, ~{len(changed)} updated, -{len(removed)} removed"


def build_pipeline_command(args: argparse.Namespace) -> list[str]:
    command = [sys.executable, str(PIPELINE_SCRIPT)]
    if args.publish:
        command.extend(["--publish", "--commit-message", args.commit_message])
    if args.skip_optimize:
        command.append("--skip-optimize")
    if args.skip_check:
        command.append("--skip-check")
    if args.full_site_check:
        command.append("--full-site-check")
    return command


def run_pipeline(args: argparse.Namespace, log_path: Path) -> bool:
    command = build_pipeline_command(args)
    log("Starting automatic gallery sync pipeline.", log_path)
    result = subprocess.run(
        command,
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )

    if result.stdout.strip():
        for line in result.stdout.strip().splitlines():
            log(f"[pipeline] {line}", log_path)

    if result.stderr.strip():
        for line in result.stderr.strip().splitlines():
            log(f"[pipeline:stderr] {line}", log_path)

    if result.returncode == 0:
        log("Automatic gallery sync finished successfully.", log_path)
        return True

    log(f"Automatic gallery sync failed with exit code {result.returncode}.", log_path)
    return False


def main() -> int:
    args = parse_args()
    log_path = Path(args.log_file).expanduser().resolve()

    install_signal_handlers()

    if args.poll_interval <= 0:
        raise SystemExit("--poll-interval must be greater than zero")
    if args.settle_seconds < 0:
        raise SystemExit("--settle-seconds cannot be negative")
    if args.retry_seconds < 0:
        raise SystemExit("--retry-seconds cannot be negative")
    if args.max_runtime is not None and args.max_runtime <= 0:
        raise SystemExit("--max-runtime must be greater than zero when provided")

    directories = watched_directories()
    log("Gallery watcher started.", log_path)
    for directory in directories:
        log(f"Watching: {directory}", log_path)

    last_snapshot = build_snapshot()
    started_at = time.monotonic()
    next_run_at = time.monotonic() if args.run_on_start else None

    if args.run_on_start:
        log("Initial sync scheduled immediately on startup.", log_path)

    while not STOP_REQUESTED:
        now = time.monotonic()
        current_snapshot = build_snapshot()

        if current_snapshot != last_snapshot:
            summary = describe_changes(last_snapshot, current_snapshot)
            log(
                f"Detected gallery folder change ({summary}). Waiting {args.settle_seconds:.1f}s for files to settle.",
                log_path,
            )
            last_snapshot = current_snapshot
            next_run_at = now + args.settle_seconds
        elif next_run_at is not None and now >= next_run_at:
            success = run_pipeline(args, log_path)
            last_snapshot = build_snapshot()
            if success:
                next_run_at = None
            else:
                next_run_at = time.monotonic() + args.retry_seconds
                log(f"Retry scheduled in {args.retry_seconds:.1f}s.", log_path)

        if args.max_runtime is not None and now - started_at >= args.max_runtime:
            log("Max runtime reached. Stopping watcher.", log_path)
            break

        time.sleep(args.poll_interval)

    log("Gallery watcher stopped.", log_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
