import streamlit as st
import pandas as pd
import os
import zipfile
import tempfile
import logging
from io import BytesIO
from pathlib import Path
import sys
import shutil

sys.path.append(os.path.dirname(__file__))

from exam_scheduler import (
    DataLoader, ExamScheduler, SystemLogger,
    MIN_ALLOCATION_SIZE, OUTPUT_DIRECTORY, PHOTOS_DIRECTORY
)

st.set_page_config(
    page_title="Exam Seating Arrangement System",
    page_icon="🎓",
    layout="centered"
)


class UIStyleManager:
    """Manages UI styling and theming."""
    
    @staticmethod
    def apply_styles():
        st.markdown("""
            <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
            
            * {
                font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            }
            
            .main-header {
                font-size: 2.8rem;
                font-weight: 700;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                text-align: center;
                padding: 1.5rem 0 0.5rem 0;
                letter-spacing: -0.5px;
            }
            
            .sub-header {
                font-size: 1.1rem;
                color: #6b7280;
                text-align: center;
                padding-bottom: 2rem;
                font-weight: 400;
            }
            
            .stButton>button {
                width: 100%;
                border-radius: 12px;
                font-weight: 500;
                transition: all 0.3s ease;
                font-family: 'Inter', sans-serif;
                padding: 0.75rem 1rem;
                font-size: 1rem;
            }
            
            .stButton>button:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            }
            
            .stDataFrame, .dataframe, table {
                border-radius: 12px !important;
                overflow: hidden;
            }
            
            .stDataFrame > div {
                border-radius: 12px !important;
            }
            
            .stAlert {
                border-radius: 12px !important;
                font-family: 'Inter', sans-serif;
            }
            
            .stMetric {
                background-color: #f9fafb;
                padding: 1.2rem;
                border-radius: 12px;
                border: 1px solid #e5e7eb;
            }
            
            .stMetric label {
                font-weight: 500;
                color: #6b7280;
                font-size: 0.9rem;
            }
            
            .stMetric div[data-testid="stMetricValue"] {
                font-weight: 600;
                color: #111827;
                font-size: 1.8rem;
            }
            
            .stFileUploader {
                border-radius: 12px;
            }
            
            .stTextArea textarea {
                border-radius: 12px;
                font-family: 'Inter', monospace;
            }
            
            .stDownloadButton>button {
                border-radius: 12px;
                font-weight: 500;
            }
            
            .stNumberInput input, .stSelectbox select {
                border-radius: 8px;
                font-family: 'Inter', sans-serif;
            }
            
            .stProgress > div > div {
                border-radius: 8px;
            }
            
            .section-container {
                background: #ffffff;
                padding: 1.5rem;
                border-radius: 16px;
                border: 1px solid #e5e7eb;
                margin-bottom: 1.5rem;
            }
            
            .section-title {
                font-size: 1.3rem;
                font-weight: 600;
                color: #111827;
                margin-bottom: 1rem;
            }
            </style>
        """, unsafe_allow_html=True)


class SessionStateManager:
    """Manages Streamlit session state."""
    
    @staticmethod
    def initialize():
        if 'processed' not in st.session_state:
            st.session_state.processed = False
        if 'output_dir' not in st.session_state:
            st.session_state.output_dir = None
        if 'seating_df' not in st.session_state:
            st.session_state.seating_df = None
        if 'seats_df' not in st.session_state:
            st.session_state.seats_df = None


class FileExtractor:
    """Handles file extraction and preparation."""
    
    @staticmethod
    def extract_photos(zip_file, target_dir):
        extracted_count = 0
        with zipfile.ZipFile(zip_file, 'r') as archive:
            contents = archive.namelist()
            
            for item in contents:
                if item.endswith('/') or item.startswith('__MACOSX'):
                    continue
                
                basename = os.path.basename(item)
                if not basename:
                    continue
                
                source = archive.open(item)
                dest_path = os.path.join(target_dir, basename)
                
                with open(dest_path, 'wb') as dest:
                    shutil.copyfileobj(source, dest)
                extracted_count += 1
        
        return extracted_count
    
    @staticmethod
    def save_excel(upload_file, target_path):
        with open(target_path, 'wb') as f:
            f.write(upload_file.read())


class WorkspaceManager:
    """Manages temporary workspace setup."""
    
    @staticmethod
    def setup_workspace():
        workspace = tempfile.mkdtemp()
        data_dir = os.path.join(workspace, 'input')
        photos_dir = os.path.join(workspace, 'photos')
        output_dir = os.path.join(workspace, 'output')
        
        os.makedirs(data_dir, exist_ok=True)
        os.makedirs(photos_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)
        
        return workspace, data_dir, photos_dir, output_dir


class LoggingConfigurator:
    """Configures logging for the application."""
    
    @staticmethod
    def configure(output_dir):
        error_log = os.path.join(output_dir, 'error.log')
        
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(error_log, mode='w'),
                logging.StreamHandler()
            ],
            force=True
        )


class OutputCollector:
    """Collects and counts output files."""
    
    @staticmethod
    def collect_files(output_dir):
        excel_list = []
        pdf_list = []
        report_list = []
        
        if os.path.exists(output_dir):
            for root, dirs, files in os.walk(output_dir):
                for filename in files:
                    full_path = os.path.join(root, filename)
                    if filename.endswith('.xlsx'):
                        if 'op_' in filename:
                            report_list.append(full_path)
                        else:
                            excel_list.append(full_path)
                    elif filename.endswith('.pdf'):
                        pdf_list.append(full_path)
        
        return excel_list, pdf_list, report_list


class ZipArchiveBuilder:
    """Builds ZIP archives from directories."""
    
    @staticmethod
    def create_archive(source_dir):
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as archive:
            for root, dirs, files in os.walk(source_dir):
                for filename in files:
                    full_path = os.path.join(root, filename)
                    rel_path = os.path.relpath(full_path, source_dir)
                    archive.write(full_path, rel_path)
        
        buffer.seek(0)
        return buffer


class DataFrameConverter:
    """Converts data for display."""
    
    @staticmethod
    def convert_to_string_df(df):
        result = df.copy()
        for col in result.columns:
            result[col] = result[col].astype(str)
        return result
    
    @staticmethod
    def add_index(df, start=1):
        result = df.copy()
        result.index = range(start, len(result) + start)
        return result


class ProcessingOrchestrator:
    """Orchestrates the main processing workflow."""
    
    def __init__(self, excel_file, photos_zip, strategy, buffer):
        self.excel_file = excel_file
        self.photos_zip = photos_zip
        self.strategy = strategy
        self.buffer = buffer
    
    def process(self):
        workspace, data_dir, photos_dir, output_dir = WorkspaceManager.setup_workspace()
        
        excel_path = os.path.join(data_dir, 'input_data_tt.xlsx')
        FileExtractor.save_excel(self.excel_file, excel_path)
        
        status = st.empty()
        status.info("📸 Extracting photos...")
        
        photo_count = FileExtractor.extract_photos(self.photos_zip, photos_dir)
        
        status.success(f"✅ Extracted {photo_count} photo files")
        
        LoggingConfigurator.configure(output_dir)
        
        sys.path.insert(0, os.getcwd())
        
        status.info("📖 Loading data...")
        loader = DataLoader(excel_path)
        timetable, enrollments, students, rooms = loader.load_all_sheets()
        
        progress = st.progress(0)
        
        status.info("🔄 Processing seating arrangements...")
        progress.progress(0.3)
        
        import exam_scheduler
        original_output = exam_scheduler.OUTPUT_DIRECTORY
        original_photos = exam_scheduler.PHOTOS_DIRECTORY
        
        exam_scheduler.OUTPUT_DIRECTORY = output_dir
        exam_scheduler.PHOTOS_DIRECTORY = photos_dir
        
        import document_creator
        document_creator.PHOTOS_DIRECTORY = photos_dir
        
        scheduler = ExamScheduler(
            timetable, enrollments, students, rooms,
            self.strategy, self.buffer, True
        )
        
        generated = scheduler.process_all_dates()
        
        exam_scheduler.OUTPUT_DIRECTORY = original_output
        exam_scheduler.PHOTOS_DIRECTORY = original_photos
        
        progress.progress(1.0)
        status.success("✅ Processing complete!")
        
        seating_path = os.path.join(output_dir, 'op_overall_seating_arrangement.xlsx')
        seats_path = os.path.join(output_dir, 'op_seats_left.xlsx')
        
        seating_df = None
        seats_df = None
        
        if os.path.exists(seating_path):
            temp_df = pd.read_excel(seating_path)
            seating_df = DataFrameConverter.convert_to_string_df(temp_df)
            
        if os.path.exists(seats_path):
            temp_df = pd.read_excel(seats_path)
            if 'Room No.' in temp_df.columns:
                temp_df['Room No.'] = temp_df['Room No.'].astype(str)
            seats_df = temp_df
        
        return output_dir, seating_df, seats_df, len(generated)


class UIRenderer:
    """Renders UI components."""
    
    @staticmethod
    def render_header():
        st.markdown('<div class="main-header">🎓 Exam Seating Arrangement System</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Automated seating arrangement with attendance sheet generation</div>', unsafe_allow_html=True)
    
    @staticmethod
    def render_file_uploaders():
        st.markdown("### 📤 Upload Input Files")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**📁 Input Excel File**")
            excel = st.file_uploader(
                "Upload Excel File",
                type=['xlsx', 'xls'],
                help="Excel file with sheets: in_timetable, in_course_roll_mapping, in_roll_name_mapping, in_room_capacity",
                label_visibility="collapsed"
            )
        
        with col2:
            st.markdown("**📸 Student Photos (ZIP)**")
            photos = st.file_uploader(
                "Upload Photos ZIP",
                type=['zip'],
                help="ZIP file containing photos named as ROLLNUMBER.jpg (include nopic.png)",
                label_visibility="collapsed"
            )
        
        st.markdown("---")
        return excel, photos
    
    @staticmethod
    def render_config():
        st.markdown("### ⚙️ Configuration")
        c1, c2, c3 = st.columns(3)
        
        with c1:
            strategy = st.selectbox(
                "Seating Mode",
                ["dense", "sparse"],
                help="Dense: Fill rooms completely | Sparse: Max 50% per course"
            )
        
        with c2:
            buffer = st.number_input(
                "Buffer Seats",
                min_value=0,
                max_value=20,
                value=5,
                help="Number of seats to keep as buffer in each room"
            )
        
        with c3:
            st.markdown("<br>", unsafe_allow_html=True)
            btn = st.button("🚀 Process Arrangement", type="primary", use_container_width=True)
        
        st.markdown("---")
        return strategy, buffer, btn
    
    @staticmethod
    def render_metrics(excel_count, pdf_count, report_count):
        m1, m2, m3 = st.columns(3)
        with m1:
            st.metric("📄 Excel Files", excel_count)
        with m2:
            st.metric("📋 PDF Sheets", pdf_count)
        with m3:
            st.metric("📊 Reports", report_count)
        
        st.markdown("<br>", unsafe_allow_html=True)
    
    @staticmethod
    def render_download_button(output_dir):
        if st.button("📥 Download Complete Output (ZIP)", use_container_width=True, type="primary"):
            with st.spinner("Creating ZIP file..."):
                try:
                    archive = ZipArchiveBuilder.create_archive(output_dir)
                    
                    st.download_button(
                        label="⬇️ Download exam_seating_output.zip",
                        data=archive.getvalue(),
                        file_name="exam_seating_output.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                    
                except Exception as err:
                    st.error(f"❌ Error creating ZIP file: {str(err)}")
    
    @staticmethod
    def render_dataframes(seating_df, seats_df):
        if seating_df is not None:
            st.markdown("### 📋 Overall Seating Arrangement")
            display_df = DataFrameConverter.add_index(seating_df)
            st.dataframe(display_df, height=400)
            st.markdown("<br>", unsafe_allow_html=True)
        
        if seats_df is not None:
            st.markdown("### 💺 Seats Left Report")
            display_df = DataFrameConverter.add_index(seats_df)
            st.dataframe(display_df, height=400)
    
    @staticmethod
    def render_footer():
        st.markdown("---")
        st.markdown("""
        <div style="text-align: center; color: #888; padding: 1rem;">
            <small>Submitted By: <strong>Aryan (2511AI29)</strong>, <strong>Tarun Vijay (2511CS20)</strong></small>
        </div>
        """, unsafe_allow_html=True)


class Application:
    """Main application controller."""
    
    def __init__(self):
        UIStyleManager.apply_styles()
        SessionStateManager.initialize()
    
    def run(self):
        UIRenderer.render_header()
        
        excel_upload, photos_upload = UIRenderer.render_file_uploaders()
        
        strategy, buffer, process_btn = UIRenderer.render_config()
        
        if process_btn:
            self._handle_processing(excel_upload, photos_upload, strategy, buffer)
        
        if st.session_state.processed:
            self._display_results()
        
        UIRenderer.render_footer()
    
    def _handle_processing(self, excel, photos, strategy, buffer):
        if not excel:
            st.error("❌ Please upload the input Excel file!")
            return
        
        if not photos:
            st.error("❌ Please upload the photos ZIP file!")
            return
        
        with st.spinner("⏳ Processing... This may take a few minutes"):
            try:
                orchestrator = ProcessingOrchestrator(excel, photos, strategy, buffer)
                output_dir, seating_df, seats_df, pdf_count = orchestrator.process()
                
                st.session_state.processed = True
                st.session_state.output_dir = output_dir
                st.session_state.seating_df = seating_df
                st.session_state.seats_df = seats_df
                
                st.success(f"✅ Successfully generated {pdf_count} attendance sheets!")
                
            except Exception as err:
                st.error(f"❌ Error: {str(err)}")
                st.exception(err)
                logging.error(f"Processing error: {err}", exc_info=True)
    
    def _display_results(self):
        st.markdown("---")
        st.markdown("### 📊 Processing Results")
        
        output_dir = st.session_state.output_dir
        
        excel_files, pdf_files, reports = OutputCollector.collect_files(output_dir)
        
        UIRenderer.render_metrics(len(excel_files), len(pdf_files), len(reports))
        
        UIRenderer.render_download_button(output_dir)
        
        st.markdown("---")
        
        UIRenderer.render_dataframes(
            st.session_state.seating_df,
            st.session_state.seats_df
        )


if __name__ == "__main__":
    app = Application()
    app.run()