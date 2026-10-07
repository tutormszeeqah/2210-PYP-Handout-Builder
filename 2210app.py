import datetime
import io
import os
import fitz  # PyMuPDF
import streamlit as st

# Word Document Libraries
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Google API Libraries
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# ==========================================
# 0. STREAMLIT PAGE CONFIG & CUSTOM STYLING
# ==========================================
st.set_page_config(
    page_title="2210 CS PYP Portal", 
    page_icon="💻",
    layout="wide"
)

# Custom Color Palette & Interface Styling
st.markdown("""
    <style>
    /* Main Page Background */
    .stApp, [data-testid="stAppViewContainer"] {
        background-color: #63D0F8 !important;
    }
    
    /* Sidebar Background */
    [data-testid="stSidebar"], [data-testid="stSidebar"] > div:first-child {
        background-color: #F0FCBB !important;
    }

    /* Global Text Color */
    html, body, [class*="css"], h1, h2, h3, h4, h5, h6, p, span, label, div, .stMarkdown {
        color: #384403 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    /* Input Fields, Selectboxes & Text Areas */
    div[data-baseweb="input"], 
    div[data-baseweb="select"] > div, 
    .stTextInput input, 
    .stSelectbox select,
    textarea {
        background-color: #CEC2F5 !important;
        color: #141801 !important;
        border-radius: 10px !important;
        border: 4px solid #A5C809 !important;
    }

    /* Buttons Styling */
    .stButton button, 
    .stDownloadButton button, 
    [data-testid="baseButton-secondary"], 
    [data-testid="baseButton-primary"] {
        background-color: #C9F40B !important;
        color: #384403 !important;
        border: 2px solid #A5C809 !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        transition: all 0.2s ease-in-out;
    }
    
    /* Hover State for Buttons */
    .stButton button:hover, .stDownloadButton button:hover {
        background-color: #A5C809 !important;
        color: #384403 !important;
        border: 3px solid #384403 !important;
    }

    /* Navigation Tab Labels */
    button[data-baseweb="tab"] p {
        font-weight: bold !important;
        font-size: 1.1rem !important;
        color: #384403 !important;
    }
    
    /* Active Tab Highlight Indicator */
    div[data-baseweb="tab-highlight"] {
        background-color: #F863E1 !important;
    }

    /* Expanders & Containers */
    [data-testid="stExpander"] {
        border: 1.5px solid #A5C809 !important;
        border-radius: 8px !important;
        background-color: #F0FCBB !important;
    }
    </style>
""", unsafe_allow_html=True)


# ==========================================
# 1. DIRECTORY MAPPING & CONFIGURATION
# ==========================================
SYLLABUS_CODE = "2210"
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# 4 Core Folders for 2210 Computer Science
LOCAL_FOLDERS = {
    "p1": "2210_Paper1",
    "p2": "2210_Paper2",
    "ms_p1": "2210_ms_P1",
    "ms_p2": "2210_ms_P2"
}

for folder_path in LOCAL_FOLDERS.values():
    os.makedirs(folder_path, exist_ok=True)


# ==========================================
# 2. GOOGLE DRIVE AUTHENTICATION & SYNC
# ==========================================
def build_drive_service(write_access=False):
    """Authenticates using Google Service Account credentials from Streamlit Secrets."""
    try:
        if "gcp_service_account" in st.secrets:
            service_account_info = dict(st.secrets["gcp_service_account"])
            scopes = ['https://www.googleapis.com/auth/drive.file'] if write_access else SCOPES
            creds = service_account.Credentials.from_service_account_info(
                service_account_info, 
                scopes=scopes
            )
            return build('drive', 'v3', credentials=creds)
        else:
            st.error("Missing [gcp_service_account] configuration in secrets.")
            return None
    except Exception as e:
        st.error(f"Authentication Error: {e}")
        return None

def sync_drive_folder_to_local(folder_key: str) -> tuple[int, str]:
    """Downloads missing files from Google Drive folder into local directory."""
    service = build_drive_service(write_access=False)
    if not service:
        return 0, "Failed to authenticate Service Account."
    
    folder_ids = st.secrets.get("drive_folders", {})
    drive_folder_id = folder_ids.get(folder_key)
    
    if not drive_folder_id:
        return 0, f"Missing drive_folder_id for `{folder_key}` in secrets."

    local_path = LOCAL_FOLDERS[folder_key]
    
    try:
        query = f"'{drive_folder_id}' in parents and trashed = false"
        drive_files = []
        page_token = None

        while True:
            response = service.files().list(
                q=query,
                fields="nextPageToken, files(id, name, mimeType)",
                pageToken=page_token,
                pageSize=100
            ).execute()
            
            drive_files.extend(response.get('files', []))
            page_token = response.get('nextPageToken', None)
            
            if not page_token:
                break

        downloaded_count = 0

        for file_info in drive_files:
            file_name = file_info['name']
            file_id = file_info['id']
            local_file_path = os.path.join(local_path, file_name)

            if not os.path.exists(local_file_path):
                request = service.files().get_media(fileId=file_id)
                with open(local_file_path, "wb") as f:
                    downloader = MediaIoBaseDownload(f, request)
                    done = False
                    while not done:
                        _, done = downloader.next_chunk()
                downloaded_count += 1

        total_local_files = len([f for f in os.listdir(local_path) if os.path.isfile(os.path.join(local_path, f))])
        return downloaded_count, f"Synced {downloaded_count} new file(s) for `{folder_key}` (Total: {total_local_files})."
        
    except Exception as e:
        return 0, f"Sync error for `{folder_key}`: {e}"

def perform_bulk_sync():
    """Syncs all 4 configured Google Drive folders."""
    total_synced = 0
    messages = []
    for f_key in LOCAL_FOLDERS.keys():
        count, msg = sync_drive_folder_to_local(f_key)
        total_synced += count
        messages.append(msg)
    return total_synced, messages


# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================
def add_page_number_to_run(run):
    """Adds a dynamic Word page number field."""
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    r = run._r
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)

def create_worksheet_docx(basket_items: list) -> io.BytesIO:
    """Generates a Word document containing selected PDF pages."""
    doc = Document()
    section = doc.sections[0]

    section.page_width = Inches(8.5)
    section.page_height = Inches(11.5)
    section.top_margin = Inches(0.3)
    section.bottom_margin = Inches(0.3)
    section.left_margin = Inches(0.3)
    section.right_margin = Inches(0.3)

    header = section.header
    header_p = header.paragraphs[0]
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_run = header_p.add_run("Page ")
    add_page_number_to_run(header_run)

    doc.add_heading(f'O Level {SYLLABUS_CODE} Computer Science Worksheet', level=1)

    for idx, item in enumerate(basket_items):
        doc.add_heading(f"Source: {item['file']} (Page {item['page'] + 1})", level=2)
        pdf_doc = fitz.open(item['path'])
        page = pdf_doc.load_page(item['page'])
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_data = io.BytesIO(pix.tobytes("png"))

        doc.add_picture(img_data, width=Inches(7.5))

        if idx < len(basket_items) - 1:
            doc.add_page_break()
        pdf_doc.close()

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def render_pdf_page_preview(filepath: str, page_num: int = 0):
    """Renders a PDF page to PNG image bytes for preview."""
    try:
        doc = fitz.open(filepath)
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        img_bytes = pix.tobytes("png")
        doc.close()
        return img_bytes
    except Exception as e:
        st.error(f"Unable to render page preview: {e}")
        return None

def execute_pdf_search(folder_key: str, keyword_string: str, variant_filter: str = "All Variants") -> list[dict]:
    """Searches PDF files in a specific folder for matching keywords and variant filter (variants 2 and 3)."""
    results = []
    keywords = [k.strip().lower() for k in keyword_string.split(",") if k.strip()]
    folder_path = LOCAL_FOLDERS[folder_key]
    
    if os.path.exists(folder_path):
        for file in os.listdir(folder_path):
            if file.endswith(".pdf"):
                file_lower = file.lower()
                
                # Apply variant filtering for Variants 2 and 3
                if variant_filter == "Variant 2" and not ("_qp_12" in file_lower or "_qp_22" in file_lower or "_12." in file_lower or "_22." in file_lower):
                    continue
                elif variant_filter == "Variant 3" and not ("_qp_13" in file_lower or "_qp_23" in file_lower or "_13." in file_lower or "_23." in file_lower):
                    continue

                filepath = os.path.join(folder_path, file)
                try:
                    doc = fitz.open(filepath)
                    for page_num in range(len(doc)):
                        text = doc[page_num].get_text().lower()
                        if all(kw in text for kw in keywords):
                            results.append({
                                "file": file, 
                                "page": page_num, 
                                "path": filepath
                            })
                    doc.close()
                except Exception:
                    continue
    return results


# ==========================================
# 4. SESSION STATE INITIALIZATION
# ==========================================
if 'handout_basket' not in st.session_state:
    st.session_state.handout_basket = []

if 'p1_results' not in st.session_state:
    st.session_state.p1_results = []
if 'p2_results' not in st.session_state:
    st.session_state.p2_results = []

if 'has_auto_synced' not in st.session_state:
    st.session_state.has_auto_synced = True
    with st.spinner("🚀 Initializing portal & auto-syncing 2210 CS files..."):
        perform_bulk_sync()


# ==========================================
# 5. STREAMLIT UI LAYOUT
# ==========================================
st.title("2210 COMPUTER SCIENCE BRUNEI ZONE 5")
st.subheader(f"💻 O Level {SYLLABUS_CODE} Computer Science PYP Portal")

# --- SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("🔄 Google Drive Sync")
    if st.button("🔄 Sync Google Drive", use_container_width=True):
        with st.spinner("Syncing Google Drive folders..."):
            synced_count, sync_msgs = perform_bulk_sync()
            st.success(f"Sync Complete! {synced_count} new file(s) downloaded.")
            for m in sync_msgs:
                st.caption(m)

    st.markdown("---")
    st.metric(label="Saved Pages in Cart", value=len(st.session_state.handout_basket))

    if st.button("🗑️ Clear Cart", use_container_width=True):
        st.session_state.handout_basket = []
        st.rerun()

# --- NAVIGATION TABS (6 TABS) ---
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "🖥️ P1 Search", 
    "🐍 P2 Search", 
    "🛒 Handout Cart", 
    "🔑 Answer P1", 
    "🔑 Answer P2", 
    "🔒 Admin PYP Upload"
])


# --- TAB 1: PAPER 1 SEARCH (COMPUTER SYSTEMS) ---
with tab1:
    st.subheader("🖥️ Paper 1 Search (Computer Systems - Variants 2 & 3)")
    
    col_v, col_kw = st.columns([1, 2])
    with col_v:
        p1_variant = st.selectbox(
            "Select Variant Filter:", 
            ["All Variants (2 & 3)", "Variant 2", "Variant 3"],
            key="p1_variant_select"
        )
    with col_kw:
        p1_kw = st.text_input(
            "Enter Search Keywords", 
            placeholder="e.g., Hexadecimal, CPU, Operating System, Encryption, Von Neumann", 
            key="p1_kw"
        )

    if st.button("Search Paper 1", key="btn_search_p1"):
        if p1_kw.strip():
            filter_var = "Variant 2" if "Variant 2" in p1_variant else ("Variant 3" if "Variant 3" in p1_variant else "All Variants")
            with st.spinner("Scanning Paper 1 PDFs..."):
                st.session_state.p1_results = execute_pdf_search("p1", p1_kw, filter_var)
        else:
            st.warning("Please enter at least one keyword.")

    if st.session_state.p1_results:
        st.write(f"Found **{len(st.session_state.p1_results)}** matching page(s):")
        for idx, item in enumerate(st.session_state.p1_results):
            with st.expander(f"📄 {item['file']} | Page {item['page'] + 1}"):
                c1, c2 = st.columns([3, 1])
                with c1:
                    preview_img = render_pdf_page_preview(item["path"], item["page"])
                    if preview_img:
                        st.image(preview_img, use_container_width=True)
                with c2:
                    if st.button("➕ Add to Cart", key=f"add_p1_{idx}"):
                        st.session_state.handout_basket.append(item)
                        st.toast(f"Added Page {item['page'] + 1} to cart!")
                        st.rerun()
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    with open(item["path"], "rb") as pdf_f:
                        st.download_button(
                            label="📥 Download Full PDF",
                            data=pdf_f,
                            file_name=item["file"],
                            mime="application/pdf",
                            key=f"dl_p1_{idx}"
                        )


# --- TAB 2: PAPER 2 SEARCH (ALGORITHMS, PROGRAMMING & LOGIC) ---
with tab2:
    st.subheader("🐍 Paper 2 Search (Algorithms & Programming)")
    
    col_v, col_kw = st.columns([1, 2])
    with col_v:
        p2_variant = st.selectbox(
            "Select Variant Filter:", 
            ["All Variants (2 & 3)", "Variant 2", "Variant 3"],
            key="p2_variant_select"
        )
    with col_kw:
        p2_kw = st.text_input(
            "Enter Search Keywords", 
            placeholder="e.g., Pseudocode, Flowchart, Trace Table, SQL, Array, Logic Gate", 
            key="p2_kw"
        )

    if st.button("Search Paper 2", key="btn_search_p2"):
        if p2_kw.strip():
            filter_var = "Variant 2" if "Variant 2" in p2_variant else ("Variant 3" if "Variant 3" in p2_variant else "All Variants")
            with st.spinner("Scanning Paper 2 PDFs..."):
                st.session_state.p2_results = execute_pdf_search("p2", p2_kw, filter_var)
        else:
            st.warning("Please enter at least one keyword.")

    if st.session_state.p2_results:
        st.write(f"Found **{len(st.session_state.p2_results)}** matching page(s):")
        for idx, item in enumerate(st.session_state.p2_results):
            with st.expander(f"📄 {item['file']} | Page {item['page'] + 1}"):
                c1, c2 = st.columns([3, 1])
                with c1:
                    preview_img = render_pdf_page_preview(item["path"], item["page"])
                    if preview_img:
                        st.image(preview_img, use_container_width=True)
                with c2:
                    if st.button("➕ Add to Cart", key=f"add_p2_{idx}"):
                        st.session_state.handout_basket.append(item)
                        st.toast(f"Added Page {item['page'] + 1} to cart!")
                        st.rerun()
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    with open(item["path"], "rb") as pdf_f:
                        st.download_button(
                            label="📥 Download Full PDF",
                            data=pdf_f,
                            file_name=item["file"],
                            mime="application/pdf",
                            key=f"dl_p2_{idx}"
                        )


# --- TAB 3: HANDOUT CART & WORKSHEET GENERATOR ---
with tab3:
    st.subheader("🛒 Handout Cart & Worksheet Generator")
    
    if len(st.session_state.handout_basket) > 0:
        st.info(f"### 📋 Selected Question Pages ({len(st.session_state.handout_basket)} items)")
        st.caption("Review your selected pages below. Remove any individual page before compiling into Word format.")
        st.markdown("---")
        
        items_to_display = list(st.session_state.handout_basket)
        
        for idx, item in enumerate(items_to_display):
            filename = item.get("file", "Unknown File")
            page_num = item.get("page", 0) + 1
            file_path = item.get("path", "")
            
            with st.expander(f"📄 Item #{idx + 1}: {filename} (Page {page_num})", expanded=False):
                col_preview, col_action = st.columns([3, 1])
                
                with col_preview:
                    if os.path.exists(file_path):
                        img_bytes = render_pdf_page_preview(file_path, item.get("page", 0))
                        if img_bytes:
                            st.image(img_bytes, caption=f"Preview: Page {page_num}", use_container_width=True)
                    else:
                        st.caption(f"Source file path: `{file_path}`")
                
                with col_action:
                    st.markdown("#### Actions")
                    remove_key = f"remove_btn_cart_item_{idx}_{filename}_{page_num}"
                    if st.button("🗑️ Remove Item", key=remove_key, use_container_width=True):
                        st.session_state.handout_basket.pop(idx)
                        st.toast(f"Removed item #{idx + 1} from cart!")
                        st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("---")
        st.markdown("### 📝 Export Handout / Worksheet")
        
        with st.spinner("Merging selected pages into Word Worksheet..."):
            doc_buffer = create_worksheet_docx(st.session_state.handout_basket)
            target_filename = f"{SYLLABUS_CODE}_CS_Worksheet.docx"

        st.download_button(
            label="🪄 Download Merged Word Document Worksheet",
            data=doc_buffer,
            file_name=target_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    else:
        st.info("🛒 Your Cart is empty! Search Paper 1 or Paper 2 and click '➕ Add to Cart' to build your worksheet.")


# --- TAB 4: ANSWER SCHEME (PAPER 1) ---
with tab4:
    st.subheader("🔑 Preview & Download Marking Schemes (Paper 1)")
    
    col_y, col_m, col_v = st.columns([1, 1.5, 1.5])
    with col_y:
        as_year = st.text_input("Year (YYYY)", value="2023", key="ms1_year")
    with col_m:
        as_month = st.selectbox("Session", [" June (s) ", " November (w) "], key="ms1_mth")
        month_code = "s" if "June" in as_month else "w"
    with col_v:
        as_variant = st.selectbox("Variant", ["12", "13"], key="ms1_var")

    cleaned_year = as_year.strip()
    short_year = cleaned_year[-2:] if len(cleaned_year) >= 2 else cleaned_year
    search_session_tag = f"{month_code}{short_year}"
    expected_ms_filename = f"{SYLLABUS_CODE}_{search_session_tag}_ms_{as_variant}.pdf"

    st.markdown("---")
    
    found_ms_files = []
    folder_path = LOCAL_FOLDERS["ms_p1"]

    if os.path.exists(folder_path):
        for file in os.listdir(folder_path):
            if file.endswith(".pdf"):
                file_lower = file.lower()
                if search_session_tag in file_lower and f"ms_{as_variant}" in file_lower:
                    found_ms_files.append(os.path.join(folder_path, file))

    if found_ms_files:
        st.success(f"Found {len(found_ms_files)} matching Answer Scheme file(s):")
        for ms_path in found_ms_files:
            ms_filename = os.path.basename(ms_path)
            with st.expander(f"🔑 Mark Scheme: {ms_filename}", expanded=True):
                with open(ms_path, "rb") as f:
                    st.download_button(
                        label=f"📥 Download {ms_filename}",
                        data=f,
                        file_name=ms_filename,
                        mime="application/pdf",
                        key=f"dl_ms1_btn_{ms_filename}"
                    )
                
                doc = fitz.open(ms_path)
                total_pages = len(doc)
                st.caption(f"📜 Showing all **{total_pages}** pages below:")
                
                with st.container(height=600):
                    for page_num in range(total_pages):
                        page_img = render_pdf_page_preview(ms_path, page_num)
                        if page_img:
                            st.image(page_img, caption=f"Page {page_num + 1} of {total_pages}", use_container_width=True)
                            if page_num < total_pages - 1:
                                st.markdown("---")
                doc.close()
    else:
        st.warning(f"No Mark Scheme found matching session `{search_session_tag}` and variant `{as_variant}` (Expected pattern: `{expected_ms_filename}`).")


# --- TAB 5: ANSWER SCHEME (PAPER 2) ---
with tab5:
    st.subheader("🔑 Preview & Download Marking Schemes (Paper 2)")
    
    col_y, col_m, col_v = st.columns([1, 1.5, 1.5])
    with col_y:
        as_year2 = st.text_input("Year (YYYY)", value="2023", key="ms2_year")
    with col_m:
        as_month2 = st.selectbox("Session", [" June (s) ", " November (w) "], key="ms2_mth")
        month_code2 = "s" if "June" in as_month2 else "w"
    with col_v:
        as_variant2 = st.selectbox("Variant", ["22", "23"], key="ms2_var")

    cleaned_year2 = as_year2.strip()
    short_year2 = cleaned_year2[-2:] if len(cleaned_year2) >= 2 else cleaned_year2
    search_session_tag2 = f"{month_code2}{short_year2}"
    expected_ms_filename2 = f"{SYLLABUS_CODE}_{search_session_tag2}_ms_{as_variant2}.pdf"

    st.markdown("---")
    
    found_ms_files2 = []
    folder_path2 = LOCAL_FOLDERS["ms_p2"]

    if os.path.exists(folder_path2):
        for file in os.listdir(folder_path2):
            if file.endswith(".pdf"):
                file_lower = file.lower()
                if search_session_tag2 in file_lower and f"ms_{as_variant2}" in file_lower:
                    found_ms_files2.append(os.path.join(folder_path2, file))

    if found_ms_files2:
        st.success(f"Found {len(found_ms_files2)} matching Answer Scheme file(s):")
        for ms_path in found_ms_files2:
            ms_filename = os.path.basename(ms_path)
            with st.expander(f"🔑 Mark Scheme: {ms_filename}", expanded=True):
                with open(ms_path, "rb") as f:
                    st.download_button(
                        label=f"📥 Download {ms_filename}",
                        data=f,
                        file_name=ms_filename,
                        mime="application/pdf",
                        key=f"dl_ms2_btn_{ms_filename}"
                    )
                
                doc = fitz.open(ms_path)
                total_pages = len(doc)
                st.caption(f"📜 Showing all **{total_pages}** pages below:")
                
                with st.container(height=600):
                    for page_num in range(total_pages):
                        page_img = render_pdf_page_preview(ms_path, page_num)
                        if page_img:
                            st.image(page_img, caption=f"Page {page_num + 1} of {total_pages}", use_container_width=True)
                            if page_num < total_pages - 1:
                                st.markdown("---")
                doc.close()
    else:
        st.warning(f"No Mark Scheme found matching session `{search_session_tag2}` and variant `{as_variant2}` (Expected pattern: `{expected_ms_filename2}`).")


# --- TAB 6: ADMIN PANEL (4 GOOGLE DRIVE REPOSITORIES) ---
with tab6:
    st.subheader("🔒 Administrator Control Panel")
    st.caption("Manage Google Drive across all 4 PYP folders")

    admin_pwd = st.secrets.get("ADMIN_PASSWORD", "")
    pwd_input = st.text_input("Enter Admin Password", type="password", key="admin_pwd_input")

    if pwd_input and pwd_input == admin_pwd:
        st.success("Authenticated as Administrator")
        st.markdown("---")
        
        st.markdown("### 🌐 Upload & Manage PYP Collections")
        st.info("💡 Click any of the 4 buttons below to open its Google Drive folder in a new tab where you can upload new PDF past papers or mark schemes.")
        
        drive_links = st.secrets.get("drive_web_links", {})

        c1, c2 = st.columns(2)
        with c1:
            st.link_button("🖥️ 1. Upload Paper 1 Question Papers", drive_links.get("p1", "https://drive.google.com"), use_container_width=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.link_button("🔑 3. Upload Answer Schemes (Paper 1)", drive_links.get("ms_p1", "https://drive.google.com"), use_container_width=True)
            
        with c2:
            st.link_button("🐍 2. Upload Paper 2 Question Papers", drive_links.get("p2", "https://drive.google.com"), use_container_width=True)
            st.markdown("<br>", unsafe_allow_html=True)
            st.link_button("🔑 4. Upload Answer Schemes (Paper 2)", drive_links.get("ms_p2", "https://drive.google.com"), use_container_width=True)

    elif pwd_input:
        st.error("Incorrect Admin Password.")


# ==========================================
# 6. PORTAL FOOTER
# ==========================================
st.markdown("---")
SCHOOL_NAME = "SBC 2210 COMPUTER SCIENCE 2026"
SCHOOL_VISION = "Quality Education, Progressive Nation"

footer_html = f"""
<div style="text-align: center; padding: 15px 0px; font-family: sans-serif;">
    <p style="margin: 0; font-size: 1.0em; font-weight: bold; color: #384403;">🏫 {SCHOOL_NAME}</p>
    <p style="margin: 5px 0; font-size: 0.9em; font-style: italic; color: #384403;">"{SCHOOL_VISION}"</p>
    <p style="margin: 5px 0 0 0; font-size: 0.85em; font-weight: 600; color: #384403;">💻 Developed by Cikgu Hajah Nurul Haziqah HN (PTES) ({SYLLABUS_CODE})</p>
</div>
"""
st.markdown(footer_html, unsafe_allow_html=True)
