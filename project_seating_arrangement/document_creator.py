import os
import logging
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus.flowables import Image
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics import renderPDF
from PIL import Image as PILImage

PHOTOS_DIRECTORY = 'photos'
DEFAULT_PHOTO = os.path.join(PHOTOS_DIRECTORY, 'nopic.png')
COLUMNS_PER_ROW = 3
PHOTO_WIDTH = 23 * mm
PHOTO_HEIGHT = 23 * mm
CORNER_RADIUS = 3 * mm


class BorderDrawing:
    """Custom flowable for drawing rounded borders."""
    
    def __init__(self, w, h, radius, col=colors.black, fill=False):
        self.w = w
        self.h = h
        self.radius = radius
        self.col = col
        self.fill = fill
    
    def wrap(self, avail_w, avail_h):
        return (self.w, self.h)
    
    def draw(self):
        c = self.canv
        c.saveState()
        c.setStrokeColor(self.col)
        c.setLineWidth(1.5)
        if self.fill:
            c.setFillColor(colors.white)
        c.roundRect(0, 0, self.w, self.h, self.radius, 
                    stroke=1, fill=1 if self.fill else 0)
        c.restoreState()


class PhotoManager:
    """Manages student photo retrieval and validation."""
    
    def __init__(self, photos_dir=PHOTOS_DIRECTORY, default_img=DEFAULT_PHOTO):
        self.photos_dir = photos_dir
        self.default_img = default_img
        self.photo_cache = {}
        self._initialize_cache()
    
    def _initialize_cache(self):
        if not os.path.exists(self.photos_dir):
            logging.warning(f"Photos directory not found: {self.photos_dir}")
            return
        
        for filename in os.listdir(self.photos_dir):
            if filename.lower().endswith(('.jpg', '.jpeg', '.png')):
                roll_id = filename.split('.')[0].upper()
                self.photo_cache[roll_id] = os.path.join(self.photos_dir, filename)
        
        logging.info(f"Loaded {len(self.photo_cache)} student photos")
    
    def get_photo_path(self, roll_number):
        roll_upper = str(roll_number).upper()
        
        if roll_upper in self.photo_cache:
            path = self.photo_cache[roll_upper]
            logging.info(f"Found photo for roll {roll_number} at {path}")
            if os.path.exists(path):
                return path
        
        for ext in ['.jpg', '.jpeg', '.JPG', '.JPEG', '.png', '.PNG']:
            path = os.path.join(self.photos_dir, f"{roll_upper}{ext}")
            if os.path.exists(path):
                return path
        
        logging.warning(f"No photo found for roll {roll_number}")
        
        if os.path.exists(self.default_img):
            return self.default_img
        else:
            logging.error(f"nopic image not found at {self.default_img}")
            return None
    
    def validate_photo(self, path, target_w, target_h):
        try:
            img = PILImage.open(path)
            if img.mode not in ('RGB', 'RGBA'):
                img = img.convert('RGB')
            return path
        except Exception as err:
            logging.error(f"Error processing image {path}: {err}")
            return self.default_img if os.path.exists(self.default_img) else None


class StyleManager:
    """Manages paragraph and text styles."""
    
    @staticmethod
    def get_title_style():
        return ParagraphStyle(
            'Title',
            fontName='Helvetica-Bold',
            fontSize=20,
            alignment=TA_CENTER,
            spaceAfter=2*mm,
            textColor=colors.black
        )
    
    @staticmethod
    def get_header_style():
        return ParagraphStyle(
            'Header',
            fontName='Helvetica',
            fontSize=11,
            alignment=TA_LEFT,
            leading=14,
            textColor=colors.black
        )
    
    @staticmethod
    def get_name_style():
        return ParagraphStyle(
            'Name',
            fontName='Helvetica-Bold',
            fontSize=10,
            alignment=TA_LEFT,
            leading=12
        )
    
    @staticmethod
    def get_roll_style():
        return ParagraphStyle(
            'Roll',
            fontName='Helvetica',
            fontSize=9,
            alignment=TA_LEFT,
            leading=11
        )
    
    @staticmethod
    def get_signature_style():
        return ParagraphStyle(
            'Sign',
            fontName='Helvetica',
            fontSize=9,
            alignment=TA_LEFT,
            leading=11
        )
    
    @staticmethod
    def get_supervisor_header_style():
        return ParagraphStyle(
            'InvigHeader',
            fontName='Helvetica-Bold',
            fontSize=11,
            alignment=TA_CENTER,
            leading=13,
            textColor=colors.black
        )
    
    @staticmethod
    def get_placeholder_style():
        return ParagraphStyle(
            'Placeholder',
            fontName='Helvetica',
            fontSize=8,
            alignment=TA_CENTER,
            leading=10
        )


class StudentCellBuilder:
    """Builds individual student cells with photo and info."""
    
    def __init__(self, photo_mgr):
        self.photo_mgr = photo_mgr
    
    def build_cell(self, roll, name):
        photo_path = self.photo_mgr.get_photo_path(roll)
        
        photo_elem = None
        if photo_path and os.path.exists(photo_path):
            try:
                validated = self.photo_mgr.validate_photo(photo_path, PHOTO_WIDTH, PHOTO_HEIGHT)
                if validated:
                    photo_elem = Image(validated, width=PHOTO_WIDTH, height=PHOTO_HEIGHT)
                else:
                    photo_elem = self._build_placeholder()
            except Exception as err:
                logging.error(f"Error loading image for {roll}: {err}")
                photo_elem = self._build_placeholder()
        else:
            photo_elem = self._build_placeholder()
        
        info_parts = []
        info_parts.append(Paragraph(name, StyleManager.get_name_style()))
        info_parts.append(Spacer(1, 1*mm))
        info_parts.append(Paragraph(f"<b>Roll:</b> {roll}", StyleManager.get_roll_style()))
        info_parts.append(Spacer(1, 1*mm))
        info_parts.append(Paragraph("Sign:______________", StyleManager.get_signature_style()))
        
        combined = Table(
            [[photo_elem, info_parts]],
            colWidths=[PHOTO_WIDTH, None]
        )
        
        combined.setStyle(TableStyle([
            ('LINEBEFORE', (1, 0), (1, -1), 1.5, colors.black),
            ('VALIGN', (0, 0), (0, 0), 'MIDDLE'),
            ('VALIGN', (1, 0), (1, 0), 'MIDDLE'),
            ('ALIGN', (0, 0), (0, 0), 'CENTER'),
            ('LEFTPADDING', (0, 0), (0, -1), 2),
            ('RIGHTPADDING', (0, 0), (0, -1), 2),
            ('LEFTPADDING', (1, 0), (1, -1), 2*mm),
            ('RIGHTPADDING', (1, 0), (1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        
        return combined
    
    def _build_placeholder(self):
        placeholder = Table(
            [[Paragraph("📷<br/><br/>No Image<br/>Available", StyleManager.get_placeholder_style())]],
            colWidths=[PHOTO_WIDTH],
            rowHeights=[PHOTO_HEIGHT]
        )
        
        placeholder.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOX', (0, 0), (-1, -1), 1, colors.grey),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        
        return placeholder


class HeaderBuilder:
    """Builds document headers."""
    
    @staticmethod
    def build_first_row(date, day, session, room, count):
        content = f"""<b>Date:</b> {date} ({day}) &nbsp;&nbsp;|&nbsp;&nbsp; <b>Shift:</b> {session} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Room No:</b> {room} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Student count:</b> {count}"""
        
        tbl = Table(
            [[Paragraph(content, StyleManager.get_header_style())]],
            colWidths=[A4[0] - 20*mm]
        )
        tbl.setStyle(TableStyle([
            ('ROUNDEDCORNERS', [5, 5, 5, 5]),
            ('BOX', (0, 0), (-1, -1), 1.5, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        return tbl
    
    @staticmethod
    def build_second_row(course):
        content = f"""<b>Subject:</b> {course} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Stud Present:</b> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;| <b>Stud Absent:</b>"""
        
        tbl = Table(
            [[Paragraph(content, StyleManager.get_header_style())]],
            colWidths=[A4[0] - 20*mm]
        )
        tbl.setStyle(TableStyle([
            ('ROUNDEDCORNERS', [5, 5, 5, 5]),
            ('BOX', (0, 0), (-1, -1), 1.5, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        return tbl


class StudentGridBuilder:
    """Builds the main grid of students."""
    
    def __init__(self, cell_builder):
        self.cell_builder = cell_builder
    
    def build_grid(self, students):
        rows = []
        current_row = []
        
        for pos, student in enumerate(students):
            roll = student['Roll Number']
            name = student['Student Name']
            
            cell = self.cell_builder.build_cell(roll, name)
            current_row.append(cell)
            
            if len(current_row) == COLUMNS_PER_ROW or pos == len(students) - 1:
                while len(current_row) < COLUMNS_PER_ROW:
                    current_row.append('')
                
                rows.append(current_row)
                current_row = []
        
        if not rows:
            return None
        
        col_width = (A4[0] - 20*mm) / COLUMNS_PER_ROW
        
        grid = Table(
            rows,
            colWidths=[col_width] * COLUMNS_PER_ROW,
            rowHeights=None
        )
        
        grid.setStyle(TableStyle([
            ('ROUNDEDCORNERS', [5, 5, 5, 5]),
            ('GRID', (0, 0), (-1, -1), 1.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        
        return grid


class SupervisorSectionBuilder:
    """Builds the supervisor signature section."""
    
    @staticmethod
    def build_section():
        header_tbl = Table(
            [[Paragraph("Invigilator Name & Signature", StyleManager.get_supervisor_header_style())]],
            colWidths=[A4[0] - 20*mm]
        )
        header_tbl.setStyle(TableStyle([
            ('ROUNDEDCORNERS', [5, 5, 0, 0]),
            ('BOX', (0, 0), (-1, -1), 1.5, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        
        data_rows = [['Sl No.', 'Name', 'Signature']]
        for i in range(1, 11):
            data_rows.append([str(i), '', ''])
        
        data_tbl = Table(
            data_rows,
            colWidths=[20*mm, 95*mm, 75*mm],
            rowHeights=[8*mm] * 11
        )
        
        data_tbl.setStyle(TableStyle([
            ('ROUNDEDCORNERS', [0, 0, 5, 5]),
            ('GRID', (0, 0), (-1, -1), 1.5, colors.black),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('ALIGN', (0, 0), (0, -1), 'CENTER'),
            ('ALIGN', (1, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        
        return [header_tbl, Spacer(1, 0), data_tbl]


class AttendanceSheetGenerator:
    """Main class for generating attendance sheets with photographs."""
    
    def __init__(self, photos_dir=PHOTOS_DIRECTORY, default_img=DEFAULT_PHOTO):
        self.photo_mgr = PhotoManager(photos_dir, default_img)
        self.cell_builder = StudentCellBuilder(self.photo_mgr)
        self.grid_builder = StudentGridBuilder(self.cell_builder)
    
    def generate_document(self, output_path, date, day, session, room, course, students, count):
        try:
            doc = SimpleDocTemplate(
                output_path,
                pagesize=A4,
                topMargin=6*mm,
                bottomMargin=6*mm,
                leftMargin=10*mm,
                rightMargin=10*mm
            )
            
            elements = []
            
            elements.append(Paragraph("IITP Attendance System", StyleManager.get_title_style()))
            elements.append(Spacer(1, 4*mm))
            
            elements.append(HeaderBuilder.build_first_row(date, day, session, room, count))
            elements.append(Spacer(1, 2*mm))
            
            elements.append(HeaderBuilder.build_second_row(course))
            elements.append(Spacer(1, 2*mm))
            
            grid = self.grid_builder.build_grid(students)
            if grid:
                elements.append(grid)
            
            elements.append(Spacer(1, 3*mm))
            
            supervisor_parts = SupervisorSectionBuilder.build_section()
            elements.append(KeepTogether(supervisor_parts))
            
            doc.build(elements)
            logging.info(f"Generated PDF: {output_path}")
            return True
            
        except Exception as err:
            logging.error(f"Failed to generate PDF {output_path}: {err}", exc_info=True)
            return False


def generate_attendance_sheets(assignments, date_obj, date_str, day, session, 
                               allocator, names, output_dir, sheet_gen):
    """Legacy function for generating attendance sheets."""
    created = []
    failed = []
    
    os.makedirs(output_dir, exist_ok=True)
    
    docs_dir = os.path.join(output_dir, 'attendance')
    os.makedirs(docs_dir, exist_ok=True)
    logging.info(f"Attendance PDFs will be saved to: {docs_dir}")
    
    date_code = date_obj.strftime('%Y%m%d')
    session_lower = session.lower()
    
    for room_id, course_map in assignments.items():
        for course, student_ids in course_map.items():
            try:
                records = []
                for sid in sorted(student_ids):
                    name = names.get(sid, "(name not found)")
                    if name == "(name not found)":
                        logging.warning(f"Name not found for roll {sid}")
                    
                    records.append({
                        'Roll Number': sid,
                        'Student Name': name,
                        'Signature': ''
                    })
                
                filename = f"{date_code}_{session_lower}_{room_id}_{course}.pdf"
                filepath = os.path.join(docs_dir, filename)
                
                success = sheet_gen.generate_document(
                    output_path=filepath,
                    date=date_str,
                    day=day,
                    session=session,
                    room=room_id,
                    course=course,
                    students=records,
                    count=len(records)
                )
                
                if success:
                    created.append(filepath)
                else:
                    failed.append(filename)
                    logging.error(f"Failed to generate PDF: {filename}")
                
            except Exception as err:
                failed.append(f"{date_code}_{session_lower}_{room_id}_{course}.pdf")
                logging.error(f"Error generating PDF for {course} in {room_id}: {err}", exc_info=True)
    
    logging.info(f"PDF Generation Summary for {date_str} {session}:")
    logging.info(f"  ✓ Successfully generated: {len(created)} PDFs")
    if failed:
        logging.error(f"  ✗ Failed to generate: {len(failed)} PDFs")
        for item in failed:
            logging.error(f"    - {item}")
    
    return created, failed