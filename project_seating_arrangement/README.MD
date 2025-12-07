
<!-- 2016_04_30_morning_10502_CB204, 2016_04_30_morning_10502_CB308 : File consist different  photos of students for verification of project -->


# Exam Seating Arrangement System

An automated system for generating exam seating arrangements with PDF attendance sheets.

---

## 📋 Prerequisites

- **Python 3.12** (for virtual environment setup)
- **Docker & Docker Compose** (for containerized deployment)
- **Input Files**:
  - `input_data_tt.xlsx` - Excel file with required sheets
  - `photos.zip` - ZIP containing student photos (ROLLNUMBER.jpg format, include nopic.png)

---

## 🚀 Implementation Guide

### **Case 1: With Docker, Without Streamlit (CLI Mode)**

#### Setup & Run
```bash

# 1. Build and start the container with inline configuration
LAYOUT_TYPE=dense SEAT_BUFFER=5 docker-compose up --build


# 2. Stop the container
docker-compose down
```

#### Run Examples
```bash
# Sparse mode with 10 buffer seats
LAYOUT_TYPE=sparse SEAT_BUFFER=5 docker-compose up --build

# Dense mode
LAYOUT_TYPE=dense SEAT_BUFFER=5  docker-compose up --build
```

#### Output
Results will be in `./output/` folder (mounted volume)

---

### **Case 2: With Docker, With Streamlit (Web UI)**

#### Setup & Run
```bash

# 1. Run streamlit in docker
APP_MODE_UI=true docker-compose up --build

# OR use .env file of system with APP_MODE_UI=true
docker-compose up --build

# 2. Access web interface at: http://localhost:8501

# 3. Stop the container
docker-compose down
```

#### Usage
1. Open browser at `http://localhost:8501`
2. **Upload files via web interface**:
   - Upload `input_data_tt.xlsx`
   - Upload `photos.zip` (files will be extracted automatically)
3. **Configure**:
   - Select mode (dense/sparse)
   - Set buffer seats
4. Click **"Process Arrangement"**
5. View results and download ZIP

#### Stop
```bash
docker-compose down
```
---
### **Case 3: Without Docker, Without Streamlit (CLI Mode)**

#### Setup & Run
```bash
# 1. Create virtual environment with Python 3.12
python3.12 -m venv venv

# 2. Activate virtual environment
source venv/bin/activate       # Linux/Mac
venv\Scripts\activate          # Windows

# 3. Install dependencies
pip install -r requirements.txt


# 4. Run the code
python exam_scheduler.py dense 5

# Dense mode with 5 buffer seats
python exam_scheduler.py dense 5

# Sparse mode with 3 buffer seats
python exam_scheduler.py sparse 3

```

#### Output
Results will be in `output/` folder:
- `output/DD-MM-YYYY/Morning|Evening/*.xlsx` - Room allocation files
- `output/attendance/*.pdf` - Attendance sheets
- `output/op_overall_seating_arrangement.xlsx` - Complete arrangement
- `output/op_seats_left.xlsx` - Capacity report

---

### **Case 4: Without Docker, With Streamlit (Web UI)**

#### Setup & Run
```bash
# 1. Create virtual environment with Python 3.12
python3.12 -m venv venv

# 2. Activate virtual environment
source venv/bin/activate       # Linux/Mac
venv\Scripts\activate          # Windows

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run streamlit app
streamlit run web_interface.py
```

#### Usage
1. Open browser at `http://localhost:8501`
2. **Upload files**:
   - Upload `input_data_tt.xlsx`
   - Upload `photos.zip`
3. **Configure**:
   - Select mode (dense/sparse)
   - Set buffer seats 
4. Click **"Process Arrangement"**
5. View results inline and download ZIP

---

## 📁 Project Structure

```
.
├── exam_scheduler.py           # Main processing logic
├── web_interface.py            # Streamlit web interface
├── document_creator.py         # PDF generator
├── requirements.txt            # Python dependencies
├── Dockerfile                  # Docker configuration
├── docker-compose.yml          # Docker Compose setup
├── input/                      # Input files folder
│   └── input_data_tt.xlsx      # Required Excel file
│   └── photos.zip              # Required when using streamlit
├── photos/                     # Student photos folder
│   ├── ROLL001.jpg
│   ├── ROLL002.jpg
│   └── nopic.png               # Fallback image
└── output/                     # Generated results
    ├── attendance/             # PDF attendance sheets
    ├── DD-MM-YYYY/             # Date-wise Excel files (/morning or /evening inside it)
    ├── op_overall_seating_arrangement.xlsx
    └── op_seats_left.xlsx
```

---

## 📝 Input File Requirements

### Photos (`photos.zip` or `photos/`)
- Format: `ROLLNUMBER.jpg` (e.g., `2021CS001.jpg`)
- Include `nopic.png` for missing photos
- Supported formats: JPG, JPEG, PNG

---

## ⚙️ Configuration Options

### Seating Modes
- **dense**: Fill rooms completely with one course
- **sparse**: Maximum 50% capacity per course per room

### Buffer Seats
- Range:  seats per room
- Reduces effective capacity for safety margin

---

## 🔧 Troubleshooting

### Python Version Issues
```bash
# Check Python version
python3.12 --version

# If not installed (Ubuntu/Debian)
sudo apt install python3.12 python3.12-venv

# macOS (using Homebrew)
brew install python@3.12
```

### Docker Port Conflicts
Edit `docker-compose.yml`:
```yaml
ports:
  - "8502:8501"  # Change 8502 to any available port
```

### Permission Issues (if occur any)
```bash
sudo chmod -R 777 output/
sudo chmod -R 777 input/
sudo chmod -R 777 photos/
```

### Photo Extraction Issues (Streamlit)
- Ensure `photos.zip` doesn't have nested folders
- File names must match roll numbers exactly

---

## 📊 Output Files

### Excel Files
- **Room allocation files**: Individual files per date/session/room/course
- **Overall seating arrangement**: Complete allocation summary
- **Seats left report**: Room capacity utilization

### PDF Files
- Attendance sheets with student photos
- Format: `YYYYMMDD_session_room_course.pdf`
- Includes signature fields and invigilator section

---
