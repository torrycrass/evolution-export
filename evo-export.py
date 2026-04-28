#!/usr/bin/env python3
"""
evo-export: Evolution Mail Folder Export Tool
=============================================
Exports GNOME Evolution local mail folders to portable, standards-compliant
formats suitable for long-term archival or migration to another mail client.

Run with no arguments to launch the interactive menu.
Run with subcommands for non-interactive / scripted use.

Supported Evolution storage formats (auto-detected):
  Maildir++  Modern Evolution 3.x on Debian/Ubuntu/Fedora
  mbox       Older Evolution installations

Output formats:
  maildir    One file per message in Maildir structure (default; lossless
             when source is already Maildir++)
  mbox       One mbox file per folder (Thunderbird, Apple Mail, Mutt)
  eml        Individual RFC 2822 .eml files (universally importable)

Post-export compression:
  --compress / -z    Create a .tar.gz archive of the output directory

Requirements:
  Python 3.10+  No third-party packages required.

Project: https://github.com/your-username/evo-export
License: MIT
"""

import email
import email.policy
import email.utils
import logging
import mailbox
import re
import shutil
import sys
import tarfile
from argparse import ArgumentParser, RawDescriptionHelpFormatter
from datetime import datetime
from pathlib import Path

# ---------------------------------------------------------------------------
# Defaults
# ---------------------------------------------------------------------------

DEFAULT_EVOLUTION_BASE = Path.home() / ".local/share/evolution/mail/local"

_SKIP_EXTENSIONS = frozenset({
    ".cmeta", ".ibex", ".ibex.index", ".ibex.index.data",
    ".ev-summary", ".ev-summary-meta", ".stamps", ".uidvalidity", ".lock",
})
_SKIP_NAMES = frozenset({"folders.db", "..maildir++", "..cmeta"})


# ---------------------------------------------------------------------------
# Storage format detection
# ---------------------------------------------------------------------------

class StoreFormat:
    MAILDIR_PP = "maildir++"
    MBOX       = "mbox"


def detect_store_format(base_dir: Path) -> str:
    """
    Auto-detect whether an Evolution local mail store uses Maildir++ or mbox.

    Maildir++ indicators (any one is sufficient):
      - Presence of the ``..maildir++`` sentinel file
      - Root directory contains cur/ new/ tmp/ subdirectories
    """
    if (base_dir / "..maildir++").exists():
        return StoreFormat.MAILDIR_PP
    if all((base_dir / d).is_dir() for d in ("cur", "new", "tmp")):
        return StoreFormat.MAILDIR_PP
    return StoreFormat.MBOX


# ---------------------------------------------------------------------------
# Camel _XX name encoding / decoding
# ---------------------------------------------------------------------------

def camel_encode(name: str) -> str:
    """
    Encode a logical folder name using Camel _XX hex escaping.
    Characters outside [A-Za-z0-9.-] are replaced with _{hex}.
    Example: ``my_folder`` -> ``my_5Ffolder``
    """
    safe = frozenset(
        "abcdefghijklmnopqrstuvwxyz"
        "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        "0123456789.-"
    )
    return "".join(c if c in safe else f"_{ord(c):02X}" for c in name)


def camel_decode(encoded: str) -> str:
    """Decode a Camel _XX hex-encoded folder name. Strips any leading dot."""
    name = encoded.lstrip(".")

    def _sub(m: re.Match) -> str:
        try:
            return chr(int(m.group(1), 16))
        except ValueError:
            return m.group(0)

    return re.sub(r"_([0-9A-Fa-f]{2})", _sub, name)


# ---------------------------------------------------------------------------
# Shared utilities
# ---------------------------------------------------------------------------

def human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if abs(n) < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def sanitize(name: str) -> str:
    """Strip characters unsafe for filesystem path components."""
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip()


def is_artifact(path: Path) -> bool:
    """Return True if the path is an Evolution metadata artifact to skip."""
    name = path.name
    if name in _SKIP_NAMES:
        return True
    return any(name.endswith(ext) for ext in _SKIP_EXTENSIONS)


class ExportStats:
    def __init__(self):
        self.folders = self.messages = self.errors = self.bytes_read = 0

    def summary(self, dry_run: bool, output: Path | None = None,
                archive: Path | None = None,
                originals_removed: bool = False) -> str:
        lines = [
            "",
            "=" * 55,
            f"Export {'(dry run) ' if dry_run else ''}complete",
            f"  Folders processed  : {self.folders}",
            f"  Messages exported  : {self.messages}",
            f"  Data processed     : {human_size(self.bytes_read)}",
            f"  Errors             : {self.errors}",
        ]
        if not dry_run:
            if output and not originals_removed:
                lines.append(f"  Output directory   : {output.resolve()}")
            if archive:
                lines.append(f"  Archive created    : {archive.resolve()}")
            if originals_removed:
                lines.append(f"  Originals removed  : yes (archive only)")
        lines.append("=" * 55)
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# Compression
# ---------------------------------------------------------------------------

def compress_output(output_dir: Path, log: logging.Logger,
                    remove_originals: bool = False) -> Path | None:
    """
    Create a .tar.gz archive of output_dir alongside it.
    Returns the archive Path on success, None on failure.
    Archive is named <output_dir_name>.tar.gz and placed in output_dir.parent.
    If remove_originals is True, the source directory is deleted after a
    successful archive is verified — the archive is checked for size > 0
    before any removal to guard against a silent write failure.
    """
    archive_path = output_dir.parent / (output_dir.name + ".tar.gz")
    log.info(f"Compressing {output_dir} -> {archive_path} ...")
    try:
        with tarfile.open(archive_path, "w:gz") as tar:
            tar.add(output_dir, arcname=output_dir.name)
        size = archive_path.stat().st_size
        log.info(f"Archive created: {archive_path} ({human_size(size)})")
    except (OSError, tarfile.TarError) as e:
        log.error(f"Compression failed: {e}")
        return None

    if remove_originals:
        # Safety check: only remove if the archive is non-empty
        if size == 0:
            log.error("Archive is 0 bytes — refusing to remove originals.")
            return archive_path
        try:
            shutil.rmtree(output_dir)
            log.info(f"Removed original export directory: {output_dir}")
        except OSError as e:
            log.error(f"Could not remove originals: {e}")

    return archive_path


# ===========================================================================
# Maildir++ backend
# ===========================================================================

def _is_maildir_folder(path: Path) -> bool:
    return (path.is_dir()
            and (path / "cur").is_dir()
            and (path / "new").is_dir()
            and (path / "tmp").is_dir())


class MaildirFolder:
    """
    A single folder within an Evolution Maildir++ store.

    Naming convention:
      ``.root_5Fname.Subfolder Name``
       │  Camel-encoded root        └─ literal subfolder name(s)
       └─ Maildir++ leading-dot marker
    """

    def __init__(self, path: Path, raw_name: str):
        self.path     = path
        self.raw_name = raw_name

        parts               = raw_name.split(".", 1)
        self.root_encoded   = parts[0]
        self.root_logical   = camel_decode(self.root_encoded)
        self.subfolder_path = parts[1] if len(parts) > 1 else ""
        self.logical_name   = (
            f"{self.root_logical}/{self.subfolder_path}"
            if self.subfolder_path else self.root_logical
        )
        self.depth = self.subfolder_path.count(".") + 1 if self.subfolder_path else 0

    @property
    def message_paths(self) -> list[Path]:
        msgs = []
        for sub in ("cur", "new"):
            d = self.path / sub
            if d.is_dir():
                msgs.extend(sorted(p for p in d.iterdir() if p.is_file()))
        return msgs

    @property
    def message_count(self) -> int:
        try:
            return len(self.message_paths)
        except OSError:
            return -1

    @property
    def size_bytes(self) -> int:
        try:
            return sum(p.stat().st_size for p in self.message_paths)
        except OSError:
            return 0

    def output_path(self) -> Path:
        root = sanitize(self.root_logical)
        if not self.subfolder_path:
            return Path(root)
        parts = [sanitize(p) for p in self.subfolder_path.split(".")]
        return Path(root).joinpath(*parts)


def maildir_discover_all(base_dir: Path) -> list[MaildirFolder]:
    folders = []
    for entry in sorted(base_dir.iterdir()):
        if not entry.name.startswith(".") or is_artifact(entry):
            continue
        if not _is_maildir_folder(entry):
            continue
        raw = entry.name[1:]
        if raw:
            folders.append(MaildirFolder(entry, raw))
    return folders


def maildir_discover_target(base_dir: Path, root_logical: str) -> list[MaildirFolder]:
    encoded = camel_encode(root_logical)
    matches = [
        f for f in maildir_discover_all(base_dir)
        if (f.root_encoded.lower() == encoded.lower()
            or f.root_logical.lower() == root_logical.lower())
    ]
    matches.sort(key=lambda f: (f.depth, f.subfolder_path.lower()))
    return matches


def maildir_resolve_root(base_dir: Path, target: str) -> str | None:
    enc = camel_encode(target)
    for f in maildir_discover_all(base_dir):
        if (f.root_logical.lower() == target.lower()
                or f.root_encoded.lower() in (target.lower(), enc.lower())):
            return f.root_logical
    return None


# ---------------------------------------------------------------------------
# Maildir++ exporters
# ---------------------------------------------------------------------------

def _maildir_export_maildir(folder: MaildirFolder, out_root: Path,
                             stats: ExportStats, dry_run: bool,
                             log: logging.Logger):
    dest = out_root / folder.output_path()
    log.info(f"  {'[DRY] ' if dry_run else ''}Maildir: {folder.logical_name} -> {dest}")
    if not dry_run:
        for sub in ("cur", "new", "tmp"):
            (dest / sub).mkdir(parents=True, exist_ok=True)
    copied = 0
    for mp in folder.message_paths:
        if not dry_run:
            try:
                shutil.copy2(mp, dest / mp.parent.name / mp.name)
                copied += 1
            except (OSError, shutil.Error) as e:
                log.error(f"    FAILED {mp.name}: {e}")
                stats.errors += 1
        else:
            copied += 1
    stats.folders += 1
    stats.messages += copied
    stats.bytes_read += folder.size_bytes
    log.info(f"    {copied} messages ({human_size(folder.size_bytes)})")


def _maildir_export_mbox(folder: MaildirFolder, out_root: Path,
                          stats: ExportStats, dry_run: bool,
                          log: logging.Logger):
    dest = out_root / (str(folder.output_path()) + ".mbox")
    log.info(f"  {'[DRY] ' if dry_run else ''}mbox: {folder.logical_name} -> {dest}")
    msg_paths = folder.message_paths
    if not msg_paths:
        log.info("    (empty, skipping)")
        stats.folders += 1
        return
    if not dry_run:
        dest.parent.mkdir(parents=True, exist_ok=True)
        try:
            out_box = mailbox.mbox(str(dest), create=True)
            out_box.lock()
        except Exception as e:
            log.error(f"    Cannot create mbox: {e}")
            stats.errors += 1
            return
    converted = 0
    for mp in msg_paths:
        if dry_run:
            converted += 1
            continue
        try:
            raw = mp.read_bytes()
            msg = mailbox.mboxMessage(email.message_from_bytes(raw))
            sender = (msg.get("Return-Path") or msg.get("From") or "unknown@unknown")
            sender = sender.strip().strip("<>").split()[-1]
            try:
                dt      = email.utils.parsedate_to_datetime(msg.get("Date", ""))
                date_fmt = dt.strftime("%a %b %d %H:%M:%S %Y")
            except Exception:
                date_fmt = datetime.now().strftime("%a %b %d %H:%M:%S %Y")
            msg.set_from(sender, date_fmt)
            out_box.add(msg)
            converted += 1
        except Exception as e:
            log.error(f"    FAILED {mp.name}: {e}")
            stats.errors += 1
    if not dry_run:
        out_box.flush()
        out_box.unlock()
        out_box.close()
    stats.folders += 1
    stats.messages += converted
    stats.bytes_read += folder.size_bytes
    log.info(f"    {converted} messages ({human_size(folder.size_bytes)})")


def _maildir_export_eml(folder: MaildirFolder, out_root: Path,
                         stats: ExportStats, dry_run: bool,
                         log: logging.Logger):
    dest = out_root / folder.output_path()
    log.info(f"  {'[DRY] ' if dry_run else ''}EML: {folder.logical_name} -> {dest}")
    msg_paths = folder.message_paths
    if not msg_paths:
        log.info("    (empty, skipping)")
        stats.folders += 1
        return
    if not dry_run:
        dest.mkdir(parents=True, exist_ok=True)
    exported = 0
    for seq, mp in enumerate(msg_paths):
        try:
            raw = mp.read_bytes()
            msg = email.message_from_bytes(raw)
            try:
                dt     = email.utils.parsedate_to_datetime(msg.get("Date", ""))
                prefix = dt.strftime("%Y%m%d_%H%M%S")
            except Exception:
                prefix = f"undated_{seq:06d}"
            slug     = sanitize(msg.get("Subject", "no_subject")[:50]).replace(" ", "_")
            filename = f"{prefix}_{slug}_{seq:05d}.eml"
            if not dry_run:
                (dest / filename).write_bytes(raw)
            exported += 1
        except Exception as e:
            log.error(f"    FAILED {mp.name}: {e}")
            stats.errors += 1
    stats.folders += 1
    stats.messages += exported
    stats.bytes_read += folder.size_bytes
    log.info(f"    {exported} messages ({human_size(folder.size_bytes)})")


# ===========================================================================
# mbox backend
# ===========================================================================

class MboxFolder:
    """A single folder within an Evolution mbox local mail store."""

    def __init__(self, mbox_path: Path, logical_name: str, depth: int = 0):
        self.mbox_path    = mbox_path
        self.logical_name = logical_name
        self.depth        = depth
        self.sbd_dir      = mbox_path.parent / (mbox_path.name + ".sbd")

    @property
    def message_count(self) -> int:
        if not self.mbox_path.is_file():
            return 0
        try:
            with self.mbox_path.open("rb") as f:
                return sum(1 for line in f if line.startswith(b"From "))
        except OSError:
            return -1

    @property
    def size_bytes(self) -> int:
        try:
            return self.mbox_path.stat().st_size
        except OSError:
            return 0

    def output_path(self) -> Path:
        parts = [sanitize(p) for p in self.logical_name.split("/")]
        return Path(*parts)


def mbox_discover_all(base_dir: Path, depth: int = 0) -> list[MboxFolder]:
    folders: list[MboxFolder] = []
    if not base_dir.exists():
        return folders
    for entry in sorted(base_dir.iterdir()):
        if entry.name.startswith(".") or is_artifact(entry) or entry.suffix == ".sbd":
            continue
        if not entry.is_file():
            continue
        logical = camel_decode(entry.name)
        folder  = MboxFolder(entry, logical_name=logical, depth=depth)
        folders.append(folder)
        if folder.sbd_dir.exists():
            for sub in mbox_discover_all(folder.sbd_dir, depth + 1):
                sub.logical_name = f"{logical}/{sub.logical_name}"
                folders.append(sub)
    return folders


def mbox_find_folder(base_dir: Path, target: str) -> MboxFolder | None:
    encoded = camel_encode(target)
    for candidate in (base_dir / encoded, base_dir / target):
        if candidate.is_file():
            return MboxFolder(candidate, logical_name=target)
    for f in mbox_discover_all(base_dir):
        if f.logical_name.lower() == target.lower():
            return f
    return None


# ---------------------------------------------------------------------------
# mbox exporters
# ---------------------------------------------------------------------------

def _mbox_export_mbox(folder: MboxFolder, out_root: Path,
                       stats: ExportStats, dry_run: bool, log: logging.Logger):
    dest = out_root / (str(folder.output_path()) + ".mbox")
    log.info(f"  {'[DRY] ' if dry_run else ''}mbox copy: {folder.logical_name} -> {dest}")
    if not dry_run:
        dest.parent.mkdir(parents=True, exist_ok=True)
        try:
            shutil.copy2(folder.mbox_path, dest)
        except (OSError, shutil.Error) as e:
            log.error(f"    FAILED: {e}")
            stats.errors += 1
            return
    stats.folders += 1
    stats.messages += folder.message_count
    stats.bytes_read += folder.size_bytes
    log.info(f"    {folder.message_count} messages ({human_size(folder.size_bytes)})")


def _mbox_export_maildir(folder: MboxFolder, out_root: Path,
                          stats: ExportStats, dry_run: bool, log: logging.Logger):
    dest = out_root / folder.output_path()
    log.info(f"  {'[DRY] ' if dry_run else ''}mbox->Maildir: {folder.logical_name} -> {dest}")
    if not folder.mbox_path.is_file():
        log.warning("    No mbox file, skipping.")
        return
    if not dry_run:
        for sub in ("cur", "new", "tmp"):
            (dest / sub).mkdir(parents=True, exist_ok=True)
    try:
        src = mailbox.mbox(str(folder.mbox_path), create=False)
    except Exception as e:
        log.error(f"    Cannot open mbox: {e}")
        stats.errors += 1
        return
    count = 0
    for idx, (_, msg) in enumerate(src.items()):
        try:
            ts = int(email.utils.parsedate_to_datetime(
                msg.get("Date", "")).timestamp())
        except Exception:
            ts = int(datetime.now().timestamp())
        filename = f"{ts}.{idx:06d}.evo-export"
        if not dry_run:
            try:
                (dest / "cur" / filename).write_bytes(
                    msg.as_bytes(policy=email.policy.compat32))
                count += 1
            except OSError as e:
                log.error(f"    FAILED {filename}: {e}")
                stats.errors += 1
        else:
            count += 1
    src.close()
    stats.folders += 1
    stats.messages += count
    stats.bytes_read += folder.size_bytes
    log.info(f"    {count} messages ({human_size(folder.size_bytes)})")


def _mbox_export_eml(folder: MboxFolder, out_root: Path,
                      stats: ExportStats, dry_run: bool, log: logging.Logger):
    dest = out_root / folder.output_path()
    log.info(f"  {'[DRY] ' if dry_run else ''}mbox->EML: {folder.logical_name} -> {dest}")
    if not folder.mbox_path.is_file():
        log.warning("    No mbox file, skipping.")
        return
    if not dry_run:
        dest.mkdir(parents=True, exist_ok=True)
    try:
        src = mailbox.mbox(str(folder.mbox_path), create=False)
    except Exception as e:
        log.error(f"    Cannot open mbox: {e}")
        stats.errors += 1
        return
    count = 0
    for seq, (_, msg) in enumerate(src.items()):
        try:
            dt     = email.utils.parsedate_to_datetime(msg.get("Date", ""))
            prefix = dt.strftime("%Y%m%d_%H%M%S")
        except Exception:
            prefix = f"undated_{seq:06d}"
        slug     = sanitize(msg.get("Subject", "no_subject")[:50]).replace(" ", "_")
        filename = f"{prefix}_{slug}_{seq:05d}.eml"
        if not dry_run:
            try:
                (dest / filename).write_bytes(
                    msg.as_bytes(policy=email.policy.compat32))
                count += 1
            except OSError as e:
                log.error(f"    FAILED {filename}: {e}")
                stats.errors += 1
        else:
            count += 1
    src.close()
    stats.folders += 1
    stats.messages += count
    stats.bytes_read += folder.size_bytes
    log.info(f"    {count} messages ({human_size(folder.size_bytes)})")


# ===========================================================================
# Unified dispatch
# ===========================================================================

_EXPORTERS = {
    (StoreFormat.MAILDIR_PP, "maildir"): _maildir_export_maildir,
    (StoreFormat.MAILDIR_PP, "mbox"):    _maildir_export_mbox,
    (StoreFormat.MAILDIR_PP, "eml"):     _maildir_export_eml,
    (StoreFormat.MBOX,       "mbox"):    _mbox_export_mbox,
    (StoreFormat.MBOX,       "maildir"): _mbox_export_maildir,
    (StoreFormat.MBOX,       "eml"):     _mbox_export_eml,
}


def get_folders(base_dir: Path, store_fmt: str,
                target: str | None = None,
                recursive: bool = True) -> list:
    if store_fmt == StoreFormat.MAILDIR_PP:
        if target:
            root    = maildir_resolve_root(base_dir, target)
            if root is None:
                return []
            folders = maildir_discover_target(base_dir, root)
            if not recursive:
                folders = [f for f in folders if f.depth == 0]
        else:
            folders = maildir_discover_all(base_dir)
        return folders
    else:
        if target:
            root = mbox_find_folder(base_dir, target)
            return [root] if root else []
        return mbox_discover_all(base_dir)


def run_export(base_dir: Path, store_fmt: str, folders: list,
               output: Path, fmt: str, dry_run: bool, compress: bool,
               remove_originals: bool, log: logging.Logger) -> ExportStats:
    """Core export runner used by both CLI and menu paths."""
    stats    = ExportStats()
    exporter = _EXPORTERS[(store_fmt, fmt)]

    for folder in folders:
        exporter(folder, output, stats, dry_run, log)

    archive = None
    if compress and not dry_run and stats.errors == 0:
        # Compress the actual export subdirectory (e.g. output/uwm_archive),
        # not the output root itself — avoids permission errors when output is
        # a top-level path like /home/user and the archive would land in /home.
        export_root = output / Path(folders[0].output_path().parts[0])
        archive = compress_output(export_root, log, remove_originals)
        if archive is None:
            stats.errors += 1

    print(stats.summary(dry_run,
                        output if not dry_run else None,
                        archive if not dry_run else None,
                        originals_removed=remove_originals and archive is not None))
    return stats


# ===========================================================================
# Interactive menu
# ===========================================================================

_MENU_DIVIDER  = "─" * 54
_MENU_HEADER   = "═" * 54


def _clear():
    print("\033[2J\033[H", end="")


def _banner(base_dir: Path, store_fmt: str):
    print(_MENU_HEADER)
    print("  evo-export  ·  Evolution Mail Export Tool")
    print(_MENU_HEADER)
    print(f"  Store : {base_dir}")
    print(f"  Format: {store_fmt}")
    print(_MENU_DIVIDER)


def _prompt_choice(prompt: str, options: list[str],
                   allow_back: bool = True) -> int | None:
    """
    Display a numbered menu and return the 0-based index of the selection.
    Returns None if the user selects 'Back' or enters an invalid choice
    after being given the option.
    """
    for i, opt in enumerate(options, 1):
        print(f"  {i}. {opt}")
    if allow_back:
        print(f"  0. Back / Cancel")
    print()
    while True:
        raw = input(f"  {prompt} [{'0-' if allow_back else ''}{len(options)}]: ").strip()
        if raw == "0" and allow_back:
            return None
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return int(raw) - 1
        print(f"  Please enter a number between {'0' if allow_back else '1'} "
              f"and {len(options)}.")


def _prompt_yn(prompt: str, default: bool = False) -> bool:
    default_str = "Y/n" if default else "y/N"
    while True:
        raw = input(f"  {prompt} [{default_str}]: ").strip().lower()
        if raw == "":
            return default
        if raw in ("y", "yes"):
            return True
        if raw in ("n", "no"):
            return False
        print("  Please enter y or n.")


def _prompt_path(prompt: str, default: str = "") -> Path:
    while True:
        display = f" (default: {default})" if default else ""
        raw = input(f"  {prompt}{display}: ").strip()
        if not raw and default:
            raw = default
        if raw:
            p = Path(raw).expanduser()
            return p
        print("  Path cannot be empty.")


def menu_list(base_dir: Path, store_fmt: str, log: logging.Logger):
    """Menu action: display all folders."""
    _clear()
    _banner(base_dir, store_fmt)
    print("  All local mail folders\n")

    folders = get_folders(base_dir, store_fmt)
    if not folders:
        print("  No folders found.")
        input("\n  Press Enter to return...")
        return

    print(f"  {'FOLDER':<42} {'MSGS':>6}  {'SIZE':>10}")
    print(f"  {_MENU_DIVIDER}")

    prev_root = None
    for f in folders:
        if store_fmt == StoreFormat.MAILDIR_PP:
            root    = f.root_logical
            depth   = f.depth
            display = f.subfolder_path if f.subfolder_path else f"[{f.root_logical}]"
        else:
            root    = f.logical_name.split("/")[0]
            depth   = f.depth
            display = f.logical_name

        if root != prev_root:
            if prev_root is not None:
                print()
            prev_root = root

        indent  = "  " * depth
        mc      = f.message_count
        mc_str  = str(mc) if mc >= 0 else "err"
        print(f"  {indent + display:<42} {mc_str:>6}  {human_size(f.size_bytes):>10}")

    total_msgs = sum(f.message_count for f in folders if f.message_count >= 0)
    total_size = sum(f.size_bytes    for f in folders)
    print(f"\n  Total: {len(folders)} folders | {total_msgs} messages | "
          f"{human_size(total_size)}")
    input("\n  Press Enter to return to main menu...")


def menu_export(base_dir: Path, store_fmt: str, log: logging.Logger):
    """Menu action: guided export workflow."""

    # ── Step 1: Discover and select folder ──────────────────────────────────
    _clear()
    _banner(base_dir, store_fmt)
    print("  STEP 1 OF 5 — Select folder to export\n")

    all_folders = get_folders(base_dir, store_fmt)
    if not all_folders:
        print("  No folders found.")
        input("\n  Press Enter to return...")
        return

    # Build unique root names for the selection list
    if store_fmt == StoreFormat.MAILDIR_PP:
        roots = list(dict.fromkeys(f.root_logical for f in all_folders))
    else:
        roots = list(dict.fromkeys(
            f.logical_name.split("/")[0] for f in all_folders))

    idx = _prompt_choice("Select a folder group", roots)
    if idx is None:
        return
    target_root = roots[idx]

    # ── Step 2: Subfolders ───────────────────────────────────────────────────
    _clear()
    _banner(base_dir, store_fmt)
    print(f"  STEP 2 OF 5 — Scope  ({target_root})\n")

    target_folders = get_folders(base_dir, store_fmt, target=target_root)
    subfolder_count = sum(1 for f in target_folders if
                          (f.depth > 0 if store_fmt == StoreFormat.MAILDIR_PP
                           else "/" in f.logical_name))

    print(f"  Found {len(target_folders)} folder(s) under '{target_root}'")
    if subfolder_count:
        print(f"  ({subfolder_count} subfolder(s) available)\n")
        recursive = _prompt_yn("Include subfolders?", default=True)
    else:
        print("  (no subfolders found)\n")
        recursive = False

    if not recursive:
        target_folders = [f for f in target_folders if
                          (f.depth == 0 if store_fmt == StoreFormat.MAILDIR_PP
                           else "/" not in f.logical_name)]

    total_msgs = sum(f.message_count for f in target_folders)
    total_size = sum(f.size_bytes    for f in target_folders)
    print(f"\n  Scope: {len(target_folders)} folder(s), "
          f"{total_msgs} messages, {human_size(total_size)}")

    # ── Step 3: Output format ────────────────────────────────────────────────
    _clear()
    _banner(base_dir, store_fmt)
    print(f"  STEP 3 OF 5 — Output format  ({target_root})\n")

    fmt_options = [
        "maildir  — One file per message (lossless; best for long-term archival)",
        "mbox     — One file per folder  (Thunderbird, Apple Mail, Mutt)",
        "eml      — Individual .eml files (maximum portability)",
    ]
    fmt_keys = ["maildir", "mbox", "eml"]
    fmt_idx  = _prompt_choice("Choose output format", fmt_options)
    if fmt_idx is None:
        return
    chosen_fmt = fmt_keys[fmt_idx]

    # ── Step 4: Output path + options ────────────────────────────────────────
    _clear()
    _banner(base_dir, store_fmt)
    print(f"  STEP 4 OF 5 — Options  ({target_root} → {chosen_fmt})\n")

    default_out = str(Path.home() / "mail-backup" / sanitize(target_root))
    output      = _prompt_path("Output directory", default=default_out)
    compress    = _prompt_yn("Compress output to .tar.gz when done?", default=False)
    if compress:
        remove_originals = _prompt_yn(
            "Remove original export files after compression?", default=False)
    else:
        remove_originals = False
    dry_run     = _prompt_yn("Dry run only (no files written)?", default=False)

    # ── Step 5: Confirm ──────────────────────────────────────────────────────
    _clear()
    _banner(base_dir, store_fmt)
    print(f"  STEP 5 OF 5 — Confirm export\n")
    print(f"  Folder       : {target_root}")
    print(f"  Subfolders   : {'yes' if recursive else 'no'}")
    print(f"  Folders      : {len(target_folders)}")
    print(f"  Messages     : {total_msgs}")
    print(f"  Output format: {chosen_fmt}")
    print(f"  Output path  : {output}")
    if compress:
        print(f"  Compress     : yes (.tar.gz)")
        print(f"  Keep originals: {'no — archive only' if remove_originals else 'yes'}")
    else:
        print(f"  Compress     : no")
    print(f"  Dry run      : {'yes' if dry_run else 'no'}")
    print()

    if not _prompt_yn("Proceed with export?", default=True):
        print("  Export cancelled.")
        input("\n  Press Enter to return...")
        return

    print()
    stats = run_export(base_dir, store_fmt, target_folders,
                       output, chosen_fmt, dry_run, compress,
                       remove_originals, log)

    input("\n  Press Enter to return to main menu...")


def cmd_menu(base_dir: Path, log: logging.Logger):
    """Interactive menu entry point."""
    store_fmt = detect_store_format(base_dir)

    while True:
        _clear()
        _banner(base_dir, store_fmt)
        print()
        options = [
            "List all mail folders",
            "Export a folder",
            "Quit",
        ]
        for i, opt in enumerate(options, 1):
            print(f"  {i}. {opt}")
        print()

        raw = input("  Choose an option [1-3]: ").strip()
        if not raw.isdigit() or int(raw) not in range(1, 4):
            continue

        choice = int(raw)
        if choice == 1:
            menu_list(base_dir, store_fmt, log)
        elif choice == 2:
            menu_export(base_dir, store_fmt, log)
        elif choice == 3:
            _clear()
            print("  Goodbye.")
            sys.exit(0)


# ===========================================================================
# CLI commands
# ===========================================================================

def cmd_list(args, log: logging.Logger):
    base      = Path(args.evolution_dir)
    if not base.exists():
        log.error(f"Evolution mail directory not found: {base}")
        sys.exit(1)
    store_fmt = detect_store_format(base)
    folders   = get_folders(base, store_fmt)

    if not folders:
        print(f"\nNo folders found in: {base}")
        return

    print(f"\nEvolution local mail store: {base}")
    print(f"Storage format: {store_fmt}\n")
    print(f"  {'LOGICAL NAME':<45} {'FILESYSTEM NAME':<45} {'MSGS':>6} {'SIZE':>10}")
    print(f"  {'-' * 110}")

    prev_root = None
    for f in folders:
        if store_fmt == StoreFormat.MAILDIR_PP:
            root    = f.root_logical
            depth   = f.depth
            display = f.subfolder_path if f.subfolder_path else f"[{f.root_logical}]"
            fs_name = "." + f.raw_name
        else:
            root    = f.logical_name.split("/")[0]
            depth   = f.depth
            display = f.logical_name
            fs_name = f.mbox_path.name

        if root != prev_root:
            if prev_root is not None:
                print()
            prev_root = root

        indent  = "  " * depth
        mc      = f.message_count
        mc_str  = str(mc) if mc >= 0 else "err"
        print(f"  {indent + display:<45} {fs_name:<45} {mc_str:>6} "
              f"{human_size(f.size_bytes):>10}")

    total_msgs = sum(f.message_count for f in folders if f.message_count >= 0)
    total_size = sum(f.size_bytes    for f in folders)
    print(f"\n  Total: {len(folders)} folders | {total_msgs} messages | "
          f"{human_size(total_size)}\n")


def cmd_info(args, log: logging.Logger):
    base      = Path(args.evolution_dir)
    if not base.exists():
        log.error(f"Evolution mail directory not found: {base}")
        sys.exit(1)
    store_fmt = detect_store_format(base)
    folders   = get_folders(base, store_fmt, target=args.folder)

    if not folders:
        log.error(f"No folder matching '{args.folder}' found.")
        log.info("Run 'list' to see all available folders.")
        sys.exit(1)

    total_msgs = sum(f.message_count for f in folders)
    total_size = sum(f.size_bytes    for f in folders)

    print(f"\nFolder       : {args.folder}")
    print(f"Store format : {store_fmt}")
    print(f"Store path   : {base}")
    print(f"{'=' * 65}")

    for f in folders:
        depth    = f.depth
        label    = (f.subfolder_path or "[root]") if store_fmt == StoreFormat.MAILDIR_PP \
                   else f.logical_name
        path_str = str(f.path if store_fmt == StoreFormat.MAILDIR_PP else f.mbox_path)
        indent   = "  " * depth
        print(f"  {indent}├─ {label:<38} {f.message_count:>5} msgs  "
              f"{human_size(f.size_bytes):>10}")
        print(f"  {indent}│  {path_str}")

    print(f"\n  Total: {total_msgs} messages, {human_size(total_size)}")

    for f in folders:
        if store_fmt == StoreFormat.MAILDIR_PP:
            samples = [email.message_from_bytes(mp.read_bytes())
                       for mp in f.message_paths[:5]]
        else:
            try:
                box     = mailbox.mbox(str(f.mbox_path), create=False)
                samples = [msg for _, msg in list(box.items())[:5]]
            except Exception:
                continue

        if samples:
            print(f"\nSample messages from '{f.logical_name}':")
            print(f"  {'DATE':<28} {'FROM':<35} SUBJECT")
            print(f"  {'-' * 95}")
            for msg in samples:
                try:
                    date = (msg.get("Date", "")[:27] or "unknown").strip()
                    frm  = (msg.get("From",  "unknown")[:33]).strip()
                    subj = (msg.get("Subject", "(no subject)")[:50]).strip()
                    print(f"  {date:<28} {frm:<35} {subj}")
                except Exception:
                    pass
            break
    print()


def cmd_export(args, log: logging.Logger):
    base      = Path(args.evolution_dir)
    output    = Path(args.output)
    if not base.exists():
        log.error(f"Evolution mail directory not found: {base}")
        sys.exit(1)

    store_fmt = detect_store_format(base)
    log.info(f"Detected storage format: {store_fmt}")

    folders = get_folders(base, store_fmt, target=args.folder,
                          recursive=args.recursive)
    if not folders:
        log.error(f"No folder matching '{args.folder}' found.")
        log.info("Run 'list' to see available folders.")
        sys.exit(1)

    log.info(f"Exporting {len(folders)} folder(s) -> {args.format}")
    if args.dry_run:
        log.info("*** DRY RUN — no files will be written ***")

    stats = run_export(base, store_fmt, folders,
                       output, args.format, args.dry_run,
                       args.compress, args.remove_originals, log)
    if stats.errors:
        sys.exit(1)


# ===========================================================================
# Argument parser
# ===========================================================================

def build_parser() -> ArgumentParser:
    parser = ArgumentParser(
        prog="evo-export",
        description=(
            "Export GNOME Evolution local mail folders to portable formats.\n"
            "Run with no arguments to launch the interactive menu."
        ),
        formatter_class=RawDescriptionHelpFormatter,
        epilog="""
Storage formats (auto-detected):
  maildir++  Modern Evolution 3.x (Debian/Ubuntu/Fedora)
  mbox       Older Evolution installations

Output formats:
  maildir    One file per message in Maildir structure [default]
  mbox       One mbox file per folder (Thunderbird, Apple Mail, Mutt)
  eml        Individual RFC 2822 .eml files

Examples:
  %(prog)s                                     # interactive menu
  %(prog)s list
  %(prog)s info Archive
  %(prog)s export Archive -o ~/backup --dry-run -r
  %(prog)s export Archive -o ~/backup -f mbox -r -z
  %(prog)s export Archive -o ~/backup -f maildir -r
  %(prog)s -e /custom/path list
        """,
    )
    parser.add_argument(
        "--evolution-dir", "-e",
        default=str(DEFAULT_EVOLUTION_BASE),
        metavar="DIR",
        dest="evolution_dir",
        help=f"Path to Evolution local mail directory "
             f"(default: {DEFAULT_EVOLUTION_BASE})",
    )
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Enable debug-level logging")

    sub = parser.add_subparsers(dest="command")   # not required — menu if absent

    sub.add_parser("list", help="List all local mail folders")

    p_info = sub.add_parser("info", help="Show details for a folder")
    p_info.add_argument("folder", help="Folder name (e.g., Archive)")

    p_exp = sub.add_parser("export", help="Export a folder")
    p_exp.add_argument("folder", help="Folder name (e.g., Archive)")
    p_exp.add_argument("-o", "--output", required=True, metavar="DIR",
                       help="Output directory (created if needed)")
    p_exp.add_argument("-f", "--format",
                       choices=["maildir", "mbox", "eml"],
                       default="maildir",
                       help="Output format (default: maildir)")
    p_exp.add_argument("-r", "--recursive", action="store_true",
                       help="Include subfolders")
    p_exp.add_argument("-z", "--compress", action="store_true",
                       help="Compress output to .tar.gz after export")
    p_exp.add_argument("-R", "--remove-originals", action="store_true",
                       dest="remove_originals",
                       help="Remove the export directory after successful "
                            "compression (only meaningful with -z)")
    p_exp.add_argument("--dry-run", action="store_true",
                       help="Simulate without writing files")

    return parser


# ===========================================================================
# Entry point
# ===========================================================================

def main():
    parser = build_parser()
    args   = parser.parse_args()

    logging.basicConfig(
        format="%(levelname)-8s %(message)s",
        level=logging.DEBUG if args.verbose else logging.INFO,
        stream=sys.stderr,
    )
    log = logging.getLogger("evo-export")

    # No subcommand → interactive menu
    if not args.command:
        base = Path(args.evolution_dir)
        if not base.exists():
            log.error(f"Evolution mail directory not found: {base}")
            log.info(f"Override with: --evolution-dir <path>")
            sys.exit(1)
        cmd_menu(base, log)
        return

    {"list": cmd_list, "info": cmd_info, "export": cmd_export}[args.command](args, log)


if __name__ == "__main__":
    main()
