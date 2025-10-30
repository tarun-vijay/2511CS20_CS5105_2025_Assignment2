# 🚀 Streamlit Application

A simple guide to set up and run the Streamlit app locally or using Docker Compose.

---

## 💻 Run Locally (Without Docker)

Follow these steps to get started:

```bash
# 1️⃣ Create a virtual environment
python -m venv venv

# 2️⃣ Activate the environment
# Linux / macOS
source venv/bin/activate

# Windows
venv\Scripts\activate

# 3️⃣ Install dependencies
pip install -r requirements.txt

# 4️⃣ Run the Streamlit app
streamlit run app.py

```

Once started, open your browser and go to:

👉 http://localhost:8501

## 🐳 Run with Docker Compose
```bash
# 🔨 Build and start the container
docker-compose up --build

# 🛑 Stop and remove the container
docker-compose down


```

Then visit:

👉 http://localhost:8501