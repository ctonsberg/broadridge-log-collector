import io
import re
import uuid
import zipfile
from datetime import datetime

import streamlit as st


SUPPORT_TO = "btcs_support.dnb@broadridge.com"
SUPPORT_CC = "em.markets@dnb.no"


def create_zip_bytes_from_uploads(uploaded_files: list, ad_id: str) -> tuple[str, bytes]:
	timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
	safe_ad_id = re.sub(r"[^A-Za-z0-9_-]", "_", ad_id.upper()) or "USER"
	zip_name = f"{safe_ad_id}_logs_{timestamp}_{uuid.uuid4().hex[:8]}.zip"
	buffer = io.BytesIO()

	with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
		for uploaded_file in uploaded_files:
			archive.writestr(uploaded_file.name, uploaded_file.getvalue())

	return zip_name, buffer.getvalue()


def init_session_state() -> None:
	st.session_state.setdefault("zip_name", "")
	st.session_state.setdefault("zip_bytes", b"")
	st.session_state.setdefault("last_subject", "")


st.set_page_config(
	page_title="Broadridge Log Collector Cloud",
	page_icon="📥",
	layout="wide",
)

init_session_state()

st.title("Broadridge Log Collector Cloud")
st.caption("Hosted version: upload the relevant log files from your PC, then download a zip for manual email submission.")
st.warning(
	"This hosted app cannot read folders on your PC or open Outlook. Upload the log files you want included, then download the zip and attach it to an email manually."
)

with st.sidebar:
	st.subheader("Support Routing")
	st.write(f"To: {SUPPORT_TO}")
	st.write(f"CC: {SUPPORT_CC}")
	st.write("Subject format: AD/AB ID + LOGS")
	st.subheader("Cloud Workflow")
	st.write("Hosted mode cannot inspect your local Windows profile.")
	st.write("Enter your AD/AB number manually and upload the matching files.")

col_left, col_right = st.columns([1.2, 1])

with col_left:
	st.subheader("User Details")
	ad_id = st.text_input("AD/AB number", max_chars=20).strip().upper()
	detected_email = st.text_input(
		"Email address",
		help="Optional. Add your email if support should see which user uploaded the log package.",
	).strip()
	uploaded_files = st.file_uploader(
		"Upload log files",
		accept_multiple_files=True,
		help="Open your local logs folder on your PC, select the files for the required date or date range, and upload them here.",
	)

with col_right:
	st.subheader("Email Preview")
	subject_preview = f"{ad_id or 'ADXXXXXX'} LOGS"
	st.metric("Email subject", subject_preview)
	st.caption("Choose the required files on your PC first. The server cannot filter your local folder by modified date.")

validation_errors: list[str] = []
if not ad_id:
	validation_errors.append("AD/AB number is required.")
if not uploaded_files:
	validation_errors.append("Upload one or more log files from your local logs folder.")

matching_files = []
if uploaded_files:
	matching_files = [
		{
			"name": uploaded_file.name,
			"modified": "Uploaded manually",
			"size_kb": round(len(uploaded_file.getvalue()) / 1024, 1),
			"path": "Browser upload",
		}
		for uploaded_file in uploaded_files
	]

if validation_errors:
	for message in validation_errors:
		st.error(message)

st.subheader("Selected Files")
if matching_files:
	st.dataframe(matching_files, use_container_width=True, hide_index=True)
	st.write(f"{len(matching_files)} uploaded files will be added to the zip archive.")
else:
	st.info("No files are ready yet.")

action_col_1, action_col_2 = st.columns([1, 1])

with action_col_1:
	can_prepare = bool(matching_files) and not validation_errors
	if st.button("Create zip for download", type="primary", disabled=not can_prepare):
		try:
			zip_name, zip_bytes = create_zip_bytes_from_uploads(uploaded_files, ad_id)
			st.session_state["zip_name"] = zip_name
			st.session_state["zip_bytes"] = zip_bytes
			st.session_state["last_subject"] = f"{ad_id.upper()} LOGS"
			st.success("The zip archive is ready. Download it and attach it to an email addressed to support.")
		except Exception as exc:
			st.error(f"Failed to prepare the log package: {exc}")

with action_col_2:
	if st.session_state["zip_bytes"]:
		st.download_button(
			"Download latest zip",
			data=st.session_state["zip_bytes"],
			file_name=st.session_state["zip_name"] or "logs.zip",
			mime="application/zip",
		)
	st.write(f"Send to: {SUPPORT_TO}")
	st.write(f"CC: {SUPPORT_CC}")
	st.write(f"Subject: {st.session_state['last_subject'] or subject_preview}")
	if detected_email:
		st.write(f"User email: {detected_email}")

st.subheader("How To Use The Hosted Version")
st.write(
	"Open your local logs folder on your PC, select the relevant files for the requested date or date range, upload them here, then download the zip and attach it manually in Outlook or your mail client."
)


# Run with: python -m streamlit run .\streamlit_log_file_scraper_cloud.py