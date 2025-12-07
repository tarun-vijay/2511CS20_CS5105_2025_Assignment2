import pandas as pd
import os
import sys
import logging
from collections import defaultdict
from datetime import datetime
from document_creator import AttendanceSheetGenerator, generate_attendance_sheets

INPUT_FILE_PATH = 'input/input_data_tt.xlsx'
OUTPUT_DIRECTORY = 'output'
PHOTOS_DIRECTORY = 'photos'
LOG_FILE_PATH = os.path.join(OUTPUT_DIRECTORY, 'error.log')
MIN_ALLOCATION_SIZE = 3


class SystemLogger:
    """Handles logging configuration for the application."""
    
    @staticmethod
    def setup_logging():
        if not os.path.exists(OUTPUT_DIRECTORY):
            os.makedirs(OUTPUT_DIRECTORY)
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(LOG_FILE_PATH, mode='w'),
                logging.StreamHandler()
            ]
        )


class DataLoader:
    """Manages loading and validation of input data."""
    
    def __init__(self, file_path):
        self.file_path = file_path
        
    def load_all_sheets(self):
        try:
            logging.info(f"Loading data from {self.file_path}")
            wb = pd.ExcelFile(self.file_path)
            
            timetable = pd.read_excel(wb, 'in_timetable')
            course_mapping = pd.read_excel(wb, 'in_course_roll_mapping')
            student_info = pd.read_excel(wb, 'in_roll_name_mapping')
            room_data = pd.read_excel(wb, 'in_room_capacity', usecols=['Room No.', 'Exam Capacity', 'Block'])
            
            logging.info("All data sheets loaded successfully.")
            logging.info(f"Loaded {len(timetable)} timetable entries")
            logging.info(f"Loaded {len(course_mapping)} course enrollments")
            logging.info(f"Loaded {len(student_info)} student name mappings")
            logging.info(f"Loaded {len(room_data)} rooms")
            
            return timetable, course_mapping, student_info, room_data
        except FileNotFoundError:
            logging.error(f"Input file not found at {self.file_path}. Please place it in the 'input' directory.")
            sys.exit(1)
        except Exception as err:
            logging.error(f"Error loading data: {err}")
            sys.exit(1)


class CourseParser:
    """Utility class for parsing course-related data."""
    
    @staticmethod
    def extract_courses(course_str):
        if pd.isna(course_str) or not isinstance(course_str, str) or course_str.strip().upper() == 'NO EXAM':
            return []
        
        courses = []
        for item in course_str.split(';'):
            clean = item.strip()
            if clean:
                if item != clean:
                    logging.info(f"Stripped extra whitespace from subject code '{item}'")
                courses.append(clean)
        return courses


class FloorCalculator:
    """Calculates floor levels from room identifiers."""
    
    @staticmethod
    def get_floor(room_id):
        room_str = str(room_id).strip()
        
        if len(room_str) >= 5 and room_str.isdigit() and room_str[:2] == '10':
            return 10
        
        if len(room_str) >= 4 and room_str[0].isdigit():
            return int(room_str[0])
        
        return 0
    
    @staticmethod
    def calculate_distance(floor1, floor2):
        return abs(floor1 - floor2)


class RoomAllocator:
    """Manages room allocation with building and floor proximity optimization."""
    
    def __init__(self, rooms, strategy, buffer):
        self.rooms = rooms
        self.strategy = strategy
        self.buffer = buffer
        
        self.usage = {r['Room No.']: 0 for r in rooms}
        self.course_tracker = {}
        
        for room in self.rooms:
            room['Floor'] = FloorCalculator.get_floor(room['Room No.'])
            room['effective_capacity'] = room['Exam Capacity'] - buffer
    
    def get_available_capacity(self, room_id, course, max_cap):
        current = self.usage[room_id]
        
        if self.strategy == 'dense':
            available = max_cap - current
        else:
            course_limit = int(max_cap * 0.5)
            key = (room_id, course)
            course_usage = self.course_tracker.get(key, 0)
            available = course_limit - course_usage
            
            total_available = max_cap - current
            available = min(available, total_available)
        
        return max(0, available)
    
    def register_allocation(self, room_id, course, count):
        self.usage[room_id] += count
        if self.strategy == 'sparse':
            key = (room_id, course)
            self.course_tracker[key] = self.course_tracker.get(key, 0) + count
    
    def allocate_course(self, course, students):
        total = len(students)
        logging.info(f"Allocating course {course} with {total} students...")
        
        buildings = defaultdict(list)
        for room in self.rooms:
            buildings[room['Block']].append(room)
        
        allocations = []
        remaining = list(students)
        
        best_building = None
        max_capacity = 0
        
        for building, building_rooms in buildings.items():
            available = sum(
                self.get_available_capacity(r['Room No.'], course, r['effective_capacity'])
                for r in building_rooms
            )
            if available >= total:
                best_building = building
                logging.info(f"  Allocating course {course} entirely in building {building}")
                break
            elif available > max_capacity:
                max_capacity = available
                best_building = building
        
        if best_building:
            building_rooms = buildings[best_building]
            sorted_rooms = sorted(
                building_rooms,
                key=lambda r: (-r['effective_capacity'], r['Floor'])
            )
            
            if sorted_rooms:
                ref_floor = sorted_rooms[0]['Floor']
                sorted_rooms = sorted(
                    sorted_rooms,
                    key=lambda r: (FloorCalculator.calculate_distance(r['Floor'], ref_floor), -r['effective_capacity'])
                )
            
            remaining = self._assign_to_rooms(
                course, remaining, sorted_rooms, allocations, best_building
            )
        
        if remaining:
            logging.warning(f"  Course {course} requires multiple buildings")
            other_rooms = [r for r in self.rooms if r['Block'] != best_building]
            sorted_other = sorted(
                other_rooms,
                key=lambda r: (-r['effective_capacity'], r['Block'], r['Floor'])
            )
            remaining = self._assign_to_rooms(
                course, remaining, sorted_other, allocations, "multiple"
            )
        
        if remaining:
            logging.error(f"  Failed to allocate {len(remaining)} students for course {course}")
        
        return allocations
    
    def _assign_to_rooms(self, course, students, rooms, allocations, building):
        remaining = list(students)
        
        for room in rooms:
            if not remaining:
                break
            
            room_id = room['Room No.']
            max_cap = room['effective_capacity']
            floor = room['Floor']
            
            available = self.get_available_capacity(room_id, course, max_cap)
            
            if available > 0:
                assign_count = min(len(remaining), available)
                
                if assign_count < MIN_ALLOCATION_SIZE and len(remaining) > assign_count:
                    continue
                
                assigned = remaining[:assign_count]
                remaining = remaining[assign_count:]
                
                allocations.append({
                    'room': room,
                    'students': assigned
                })
                
                self.register_allocation(room_id, course, assign_count)
                
                logging.info(f"    Allocated {assign_count} students to {room_id} "
                           f"(Floor {floor}, Building {room['Block']}, "
                           f"Used: {self.usage[room_id]}/{max_cap})")
        
        if remaining:
            for room in rooms:
                if not remaining:
                    break
                
                room_id = room['Room No.']
                max_cap = room['effective_capacity']
                available = self.get_available_capacity(room_id, course, max_cap)
                
                if available > 0:
                    assign_count = min(len(remaining), available)
                    assigned = remaining[:assign_count]
                    remaining = remaining[assign_count:]
                    
                    allocations.append({
                        'room': room,
                        'students': assigned
                    })
                    
                    self.register_allocation(room_id, course, assign_count)
                    
                    logging.info(f"    Allocated {assign_count} students to {room_id} (forced allocation)")
        
        return remaining
    
    def check_capacity_violations(self):
        violations = []
        for room_id, used in self.usage.items():
            room = next((r for r in self.rooms if r['Room No.'] == room_id), None)
            if room:
                max_eff = room['effective_capacity']
                if used > max_eff:
                    msg = (f"CAPACITY EXCEEDED: Room {room_id} has {used} students "
                           f"but effective capacity is {max_eff}")
                    logging.error(msg)
                    violations.append(msg)
        return violations


class ConflictDetector:
    """Detects scheduling conflicts for students."""
    
    @staticmethod
    def check_conflicts(course_enrollments, date, session):
        student_courses = defaultdict(list)
        
        for course, student_ids in course_enrollments.items():
            for sid in student_ids:
                student_courses[sid].append(course)
        
        has_conflict = False
        for sid, courses in student_courses.items():
            if len(courses) > 1:
                if not has_conflict:
                    logging.error(f"\n⚠️  CLASH DETECTED ⚠️")
                    logging.error(f"Date: {date}")
                    logging.error(f"Session: {session}")
                logging.error(f"Student {sid} enrolled in: {', '.join(courses)}")
                has_conflict = True
        
        if not has_conflict:
            logging.info("✓ No clashes detected")
        
        return has_conflict


class DocumentGenerator:
    """Generates Excel and PDF documents for seating arrangements."""
    
    @staticmethod
    def create_excel_sheet(path, course, room, date, session, students):
        df = pd.DataFrame(students)
        
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            header = f'Course: {course} | Room: {room} | Date: {date} | Session: {session}'
            header_df = pd.DataFrame([header])
            header_df.to_excel(writer, index=False, header=False, startrow=0)
            
            ws = writer.sheets['Sheet1']
            ws.merge_cells('A1:C1')
            
            df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=2)
            
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 30
            ws.column_dimensions['C'].width = 20
            
            last_row = len(df) + 4
            
            last_row += 2
            for i in range(5):
                ws[f'A{last_row + i}'] = f"TA {i+1}:"
            
            last_row += 6
            for i in range(5):
                ws[f'A{last_row + i}'] = f"Invigilator {i+1}:"


class ReportGenerator:
    """Generates summary reports."""
    
    @staticmethod
    def generate_reports(seating_data, capacity_data, room_specs):
        if seating_data:
            df = pd.DataFrame(seating_data)
            path = os.path.join(OUTPUT_DIRECTORY, 'op_overall_seating_arrangement.xlsx')
            df.to_excel(path, index=False)
        else:
            logging.warning("No seating arrangements to write")
        
        if capacity_data:
            df = pd.DataFrame(capacity_data)
            path = os.path.join(OUTPUT_DIRECTORY, 'op_seats_left.xlsx')
            df.to_excel(path, index=False)
            logging.info(f"✓ Generated seats left report: {path}")
        else:
            logging.warning("No seats left data to write")


class ExamScheduler:
    """Main orchestrator for exam scheduling."""
    
    def __init__(self, timetable, enrollments, students, rooms, strategy, buffer, gen_docs=True):
        self.timetable = timetable
        self.enrollments = enrollments
        self.students = students
        self.rooms = rooms
        self.strategy = strategy
        self.buffer = buffer
        self.gen_docs = gen_docs
        
        self.name_lookup = {}
        for _, row in students.iterrows():
            self.name_lookup[row['Roll']] = row['Name']
        
        self.doc_builder = None
        if gen_docs:
            self.doc_builder = AttendanceSheetGenerator(PHOTOS_DIRECTORY)
        
        self.all_seating = []
        self.all_capacity = []
        self.generated_pdfs = []
    
    def process_all_dates(self):
        for idx, row in self.timetable.iterrows():
            try:
                exam_date = row['Date']
                if pd.isna(exam_date):
                    continue
                    
                date_str = exam_date.strftime('%d-%m-%Y')
                day = row['Day']
                
                logging.info(f"\n{'='*80}")
                logging.info(f"Processing Date: {date_str} ({day})")
                logging.info(f"{'='*80}")
                
                for session in ['Morning', 'Evening']:
                    self._process_session(exam_date, date_str, day, session, row)
                    
            except Exception as err:
                logging.error(f"Error processing row {idx}: {err}", exc_info=True)
                continue
        
        ReportGenerator.generate_reports(self.all_seating, self.all_capacity, self.rooms)
        
        if self.gen_docs:
            logging.info(f"Total PDFs generated: {len(self.generated_pdfs)}")
        
        return self.generated_pdfs
    
    def _process_session(self, exam_date, date_str, day, session, schedule_row):
        logging.info(f"\n--- {session} Session ---")
        
        courses = CourseParser.extract_courses(schedule_row[session])
        if not courses:
            logging.info(f"No exams scheduled for {session} session. Skipping.")
            return
        
        logging.info(f"Subjects scheduled: {', '.join(courses)}")
        
        course_students = {}
        for course in courses:
            enrolled = self.enrollments[self.enrollments['course_code'] == course]['rollno'].tolist()
            course_students[course] = sorted(enrolled)
            logging.info(f"  {course}: {len(enrolled)} students enrolled")
        
        if ConflictDetector.check_conflicts(course_students, date_str, session):
            logging.warning(f"Skipping allocation for {date_str} {session} due to clashes.")
            return
        
        total_students = sum(len(s) for s in course_students.values())
        logging.info(f"Total students to allocate: {total_students}")
        
        room_records = self.rooms.to_dict('records')
        allocator = RoomAllocator(room_records, self.strategy, self.buffer)
        
        total_capacity = sum(r['effective_capacity'] for r in allocator.rooms)
        logging.info(f"Total effective capacity available: {total_capacity}")
        
        if total_students > total_capacity:
            logging.error(f"INSUFFICIENT CAPACITY: Need {total_students} seats, "
                        f"but only {total_capacity} available")
            return
        
        sorted_courses = sorted(
            course_students.items(),
            key=lambda x: len(x[1]),
            reverse=True
        )
        
        room_assignments = defaultdict(lambda: defaultdict(list))
        
        for course, enrolled in sorted_courses:
            result = allocator.allocate_course(course, enrolled)
            
            for assignment in result:
                room = assignment['room']
                students = assignment['students']
                room_id = room['Room No.']
                
                for sid in students:
                    room_assignments[room_id][course].append(sid)
        
        violations = allocator.check_capacity_violations()
        if violations:
            logging.error("Capacity violations detected! Skipping output generation.")
            return
        
        self._generate_outputs(
            exam_date, date_str, day, session, room_assignments,
            allocator
        )
    
    def _generate_outputs(self, exam_date, date_str, day, session, assignments, allocator):
        folder_date = exam_date.strftime('%d_%m_%Y')
        readable_date = exam_date.strftime('%d-%m-%Y')
        output_path = os.path.join(OUTPUT_DIRECTORY, readable_date, session)
        os.makedirs(output_path, exist_ok=True)
        
        logging.info(f"\nGenerating output files in: {output_path}")
        
        used_rooms = set()
        
        for room_id, course_map in assignments.items():
            used_rooms.add(room_id)
            
            room_info = next((r for r in allocator.rooms if r['Room No.'] == room_id), None)
            
            for course, student_ids in course_map.items():
                student_records = []
                for sid in sorted(student_ids):
                    name = self.name_lookup.get(sid, "(name not found)")
                    if name == "(name not found)":
                        logging.warning(f"Roll number {sid} not found in name mapping")
                    student_records.append({
                        'Roll Number': sid,
                        'Student Name': name,
                        'Signature': ''
                    })
                
                excel_name = f"{folder_date}_{course}_{room_id}.xlsx"
                excel_path = os.path.join(output_path, excel_name)
                DocumentGenerator.create_excel_sheet(excel_path, course, room_id, date_str, session, student_records)
                
                if self.doc_builder:
                    try:
                        pdf_name = f"{exam_date.strftime('%Y_%m_%d')}_{session.lower()}_{room_id}_{course}.pdf"
                        docs_folder = os.path.join(OUTPUT_DIRECTORY, 'attendance')
                        os.makedirs(docs_folder, exist_ok=True)
                        pdf_path = os.path.join(docs_folder, pdf_name)
                        success = self.doc_builder.generate_document(
                            output_path=pdf_path,
                            date=date_str,
                            day=day,
                            session=session,
                            room=room_id,
                            course=course,
                            students=student_records,
                            count=len(student_records)
                        )
                        
                        if success and self.generated_pdfs is not None:
                            self.generated_pdfs.append(pdf_path)
                            
                    except Exception as err:
                        logging.error(f"Failed to generate PDF for {course} in {room_id}: {err}")
                
                self.all_seating.append({
                    'Date': folder_date,
                    'Day': day,
                    'Session': session,
                    'Course Code': course,
                    'Room': room_id,
                    'Building': room_info['Block'] if room_info else '',
                    'Room Capacity': room_info['Exam Capacity'] if room_info else 0,
                    'Allocated Student Count': len(student_ids),
                    'Roll Number List': ';'.join(map(str, sorted(student_ids)))
                })
        
        for room_id in used_rooms:
            room_info = next((r for r in allocator.rooms if r['Room No.'] == room_id), None)
            if room_info:
                used = allocator.usage[room_id]
                self.all_capacity.append({
                    'Date': folder_date,
                    'Day': day,
                    'Session': session,
                    'Room No.': room_id,
                    'Exam Capacity': room_info['Exam Capacity'],
                    'Block': room_info['Block'],
                    'Allotted': used,
                    'Vacant': room_info['Exam Capacity'] - used
                })
        
        logging.info(f"✓ Generated {len(assignments)} room files (Excel + PDF)")


def import_excel_data(filepath):
    """Wrapper function for backward compatibility."""
    loader = DataLoader(filepath)
    return loader.load_all_sheets()


def execute_arrangement_process(schedule, enrollment, student_reg, venue, strategy, buffer, create_docs=True):
    """Wrapper function for backward compatibility."""
    scheduler = ExamScheduler(schedule, enrollment, student_reg, venue, strategy, buffer, create_docs)
    return scheduler.process_all_dates()


def configure_logger():
    """Wrapper function for backward compatibility."""
    SystemLogger.setup_logging()


if __name__ == "__main__":
    configure_logger()
    
    if len(sys.argv) < 3 or len(sys.argv) > 4:
        logging.error("Usage: python exam_scheduler.py <mode> <buffer> [--no-pdf]")
        logging.error("Example: python exam_scheduler.py dense 5")
        logging.error("         python exam_scheduler.py dense 5 --no-pdf")
        sys.exit(1)
    
    mode = sys.argv[1].lower()
    if mode not in ['sparse', 'dense']:
        logging.error("Invalid mode. Choose 'sparse' or 'dense'.")
        sys.exit(1)
    
    try:
        buff = int(sys.argv[2])
        if buff < 0:
            logging.error("Buffer must be a non-negative integer.")
            sys.exit(1)
    except ValueError:
        logging.error("Buffer must be an integer.")
        sys.exit(1)
    
    create_docs = True
    if len(sys.argv) == 4 and sys.argv[3] == '--no-pdf':
        create_docs = False
        logging.info("PDF generation disabled (--no-pdf flag)")
    
    logging.info(f"\n{'='*80}")
    logging.info(f"EXAM SEATING ARRANGEMENT SYSTEM")
    logging.info(f"{'='*80}")
    logging.info(f"Mode: {mode.upper()}")
    logging.info(f"Buffer: {buff} seats per room")
    logging.info(f"{'='*80}\n")
    
    try:
        timetable, course_roll, roll_name, room_cap = import_excel_data(INPUT_FILE_PATH)
        
        execute_arrangement_process(
            timetable, course_roll, roll_name, room_cap, 
            mode, buff, create_docs
        )
        
        logging.info("✓✓✓ SCRIPT COMPLETED SUCCESSFULLY ✓✓✓")
        
    except Exception as err:
        logging.critical(f"\n{'='*80}")
        logging.critical(f"CRITICAL ERROR: {err}")
        logging.critical(f"{'='*80}\n", exc_info=True)
        sys.exit(1)