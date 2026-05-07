import getpass
import importlib
import os
import re
import tempfile
import uuid
import zipfile
from dataclasses import dataclass
from datetime import date, datetime, time
from pathlib import Path

import streamlit as st

# ---------------------------------------------------------------------------
# Write .streamlit/config.toml at startup so no dot-files need to exist in
# the repository.  The folder is created next to the script at runtime.
# ---------------------------------------------------------------------------
_CONFIG_DIR = Path(__file__).parent / ".streamlit"
_CONFIG_FILE = _CONFIG_DIR / "config.toml"

if not _CONFIG_FILE.exists():
    _CONFIG_DIR.mkdir(exist_ok=True)
    _CONFIG_FILE.write_text(
        "[theme]\n"
        'primaryColor = "#0B5D3B"\n'
        'backgroundColor = "#FFFFFF"\n'
        'secondaryBackgroundColor = "#F5F7FA"\n'
        'textColor = "#111827"\n'
        'font = "sans serif"\n'
        "\n"
        "[browser]\n"
        "gatherUsageStats = false\n"
        "\n"
        "[client]\n"
        "showErrorDetails = false\n"
        "\n"
        "[logger]\n"
        'level = "error"\n',
        encoding="utf-8",
    )


SUPPORT_TO = "btcs_support.dnb@broadridge.com"
SUPPORT_CC = "em.markets@dnb.no"


@dataclass
class UserContext:
	ad_id: str
	email: str
	launcher_root: Path
	log_dir: Path | None
	candidates: list[Path]


def detect_ad_id() -> str:
	username = (
		os.environ.get("USERNAME")
		or os.environ.get("USER")
		or getpass.getuser()
		or ""
	)
	return username.strip().upper()


def parse_email_from_folder(folder_name: str) -> str:
	match = re.match(r"(?P<email>[^_]+@[^_]+)", folder_name)
	return match.group("email") if match else ""


def discover_log_candidates(launcher_root: Path) -> list[Path]:
	if not launcher_root.exists():
		return []

	candidates: list[Path] = []
	for child in launcher_root.iterdir():
		if not child.is_dir():
			continue
		logs_dir = child / "logs"
		if logs_dir.is_dir():
			candidates.append(logs_dir)
	candidates.sort(key=lambda item: item.stat().st_mtime, reverse=True)
	return candidates


def build_user_context() -> UserContext:
	ad_id = detect_ad_id()
	launcher_root = Path.home() / "AppData" / "Roaming" / "BROADRIDGE" / "ullauncher"
	candidates = discover_log_candidates(launcher_root)
	log_dir = candidates[0] if candidates else None
	email = parse_email_from_folder(log_dir.parent.name) if log_dir else ""
	return UserContext(
		ad_id=ad_id,
		email=email,
		launcher_root=launcher_root,
		log_dir=log_dir,
		candidates=candidates,
	)


def normalize_date_bounds(start_date: date, end_date: date) -> tuple[datetime, datetime]:
	start_dt = datetime.combine(start_date, time.min)
	end_dt = datetime.combine(end_date, time.max)
	return start_dt, end_dt


def list_matching_files(log_dir: Path, start_date: date, end_date: date) -> list[dict]:
	start_dt, end_dt = normalize_date_bounds(start_date, end_date)
	matches: list[dict] = []

	for child in sorted(log_dir.iterdir()):
		if not child.is_file():
			continue

		modified = datetime.fromtimestamp(child.stat().st_mtime)
		if start_dt <= modified <= end_dt:
			matches.append(
				{
					"name": child.name,
					"modified": modified.strftime("%Y-%m-%d %H:%M:%S"),
					"size_kb": round(child.stat().st_size / 1024, 1),
					"path": str(child),
				}
			)

	return matches


def create_zip_archive(source_dir: Path, files: list[dict], ad_id: str) -> Path:
	timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
	safe_ad_id = re.sub(r"[^A-Za-z0-9_-]", "_", ad_id.upper()) or "USER"
	zip_name = f"{safe_ad_id}_logs_{timestamp}_{uuid.uuid4().hex[:8]}.zip"
	zip_path = Path(tempfile.gettempdir()) / zip_name

	with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
		for item in files:
			file_path = Path(item["path"])
			archive.write(file_path, arcname=file_path.name)

	return zip_path


def create_outlook_draft(ad_id: str, user_email: str, zip_path: Path, file_count: int) -> None:
	try:
		pythoncom = importlib.import_module("pythoncom")
		dispatch = getattr(importlib.import_module("win32com.client"), "Dispatch")
	except ImportError as exc:
		raise RuntimeError(
			"Outlook integration requires pywin32. Install it with 'pip install pywin32'. "
			f"Import error: {exc}"
		)

	pythoncom.CoInitialize()
	outlook = dispatch("Outlook.Application")
	mail_item = outlook.CreateItem(0)
	mail_item.To = SUPPORT_TO
	mail_item.CC = SUPPORT_CC
	mail_item.Subject = f"{ad_id.upper()} LOGS"

	details = [
		f"User ID: {ad_id.upper()}",
		f"Detected email: {user_email or 'Not detected'}",
		f"Attached archive: {zip_path.name}",
		f"Included files: {file_count}",
		"",
		"Please review the attachment and send the email when ready.",
	]
	mail_item.Body = "\n".join(details)
	mail_item.Attachments.Add(str(zip_path))
	mail_item.Display()


def init_session_state() -> None:
	st.session_state.setdefault("zip_path", "")
	st.session_state.setdefault("last_file_count", 0)
	st.session_state.setdefault("last_subject", "")


st.set_page_config(
    page_title="Broadridge Log Collector",
    page_icon="📥",
    layout="wide",
)

init_session_state()
context = build_user_context()

st.title("Broadridge Log Collector")
st.caption(
	"Collect launcher log files from the current Windows profile, zip them, and open a draft email in Outlook for the user to send."
)

with st.sidebar:
	st.subheader("Support Routing")
	st.write(f"To: {SUPPORT_TO}")
	st.write(f"CC: {SUPPORT_CC}")
	st.write("Subject format: AD/AB ID + LOGS")

	st.subheader("Auto-detection")
	st.write(f"Launcher root: {context.launcher_root}")
	st.write(f"Detected AD/AB: {context.ad_id or 'Not detected'}")
	st.write(f"Detected email: {context.email or 'Not detected'}")

col_left, col_right = st.columns([1.2, 1])

with col_left:
	st.subheader("User Details")
	ad_id = st.text_input("AD/AB number", value=context.ad_id, max_chars=20).strip().upper()
	detected_email = st.text_input(
		"Email address",
		value=context.email,
		help="This is inferred from the Broadridge launcher folder name when available.",
	).strip()

	candidate_paths = [str(path) for path in context.candidates]
	default_index = 0 if candidate_paths else None

	if candidate_paths:
		selected_log_dir = st.selectbox(
			"Detected log folder",
			options=candidate_paths,
			index=default_index,
			help="Choose the detected Broadridge launcher log folder if more than one exists.",
		)
		log_dir_value = st.text_input("Log folder path", value=selected_log_dir).strip()
	else:
		log_dir_value = st.text_input(
			"Log folder path",
			value=str(context.log_dir) if context.log_dir else "",
			help="If detection fails, paste the full logs folder path here.",
		).strip()

with col_right:
	st.subheader("Date Filter")
	filter_mode = st.radio("Filter type", options=["Single day", "Date range"], horizontal=True)
	today = date.today()

	if filter_mode == "Single day":
		selected_day = st.date_input("Modified on", value=today)
		start_date = selected_day
		end_date = selected_day
	else:
		start_date = st.date_input("Modified from", value=today)
		end_date = st.date_input("Modified to", value=today)

	subject_preview = f"{ad_id or 'ADXXXXXX'} LOGS"
	st.metric("Email subject", subject_preview)

log_dir = Path(log_dir_value) if log_dir_value else None
validation_errors: list[str] = []

if not ad_id:
	validation_errors.append("AD/AB number is required.")
if not log_dir_value:
	validation_errors.append("A log folder path is required.")
elif not log_dir or not log_dir.exists():
	validation_errors.append("The selected log folder does not exist.")
elif not log_dir.is_dir():
	validation_errors.append("The selected log folder path is not a directory.")
if start_date > end_date:
	validation_errors.append("The start date must be on or before the end date.")

if validation_errors:
	for message in validation_errors:
		st.error(message)
	matching_files: list[dict] = []
else:
	matching_files = list_matching_files(log_dir, start_date, end_date)

st.subheader("Matching Files")
if matching_files:
	st.dataframe(matching_files, use_container_width=True, hide_index=True)
	st.write(f"{len(matching_files)} files matched the selected date filter.")
else:
	st.info("No files matched the selected date filter.")

action_col_1, action_col_2 = st.columns([1, 1])

with action_col_1:
	can_prepare = bool(matching_files) and not validation_errors
	if st.button("Prepare zip and Outlook draft", type="primary", disabled=not can_prepare):
		try:
			zip_path = create_zip_archive(log_dir, matching_files, ad_id)
			create_outlook_draft(ad_id, detected_email, zip_path, len(matching_files))

			st.session_state["zip_path"] = str(zip_path)
			st.session_state["last_file_count"] = len(matching_files)
			st.session_state["last_subject"] = f"{ad_id.upper()} LOGS"

			st.success("The zip archive was created and a draft email was opened in Outlook. Review it and send it manually.")
		except Exception as exc:
			st.error(f"Failed to prepare the Outlook draft: {exc}")

with action_col_2:
	if st.session_state["zip_path"]:
		zip_path = Path(st.session_state["zip_path"])
		if zip_path.exists():
			with zip_path.open("rb") as handle:
				st.download_button(
					"Download latest zip",
					data=handle.read(),
					file_name=zip_path.name,
					mime="application/zip",
				)

st.subheader("If Detection Fails")
st.write(
	"This app can usually infer the AD/AB number from the Windows username and the email address from the Broadridge launcher folder name. "
	"If either value is blank on another machine, capture it manually in the fields above."
)
st.write(
	"A public hosted URL is not a good fit for this workflow because browser-hosted apps cannot directly read a user's Windows profile or automate desktop Outlook. "
	"This should be shared as a local Streamlit app or packaged as a Windows executable."
)


# Run with: python -m streamlit run .\streamlit_log_file_scraper.py
