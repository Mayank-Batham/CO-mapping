# Course Outcome Mapping Automation

This repository contains a Streamlit application designed to automate the extraction and mapping of Course Outcomes (CO) from CIA question papers and student marks into Excel academic scheme files.

## Prerequisites

- Python 3.8 or higher installed on your system
- Git (optional, for cloning the repository)

## How to Run the Application Locally

Follow these steps to download and run the application on your own machine:

### 1. Download the Repository
You can either download the repository as a ZIP file from GitHub and extract it, or clone it using your terminal/command prompt:
```bash
git clone https://github.com/Mayank-Batham/CO-mapping.git
cd CO-mapping
```

### 2. Create a Virtual Environment (Recommended)
It's best practice to create a virtual environment to keep the application's dependencies separate from your main system.

**For Windows:**
```bash
python -m venv venv
.\venv\Scripts\activate
```

**For macOS/Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### 3. Install the Dependencies
With your virtual environment activated, install all the required Python packages using the provided `requirements.txt` file:
```bash
pip install -r requirements.txt
```

### 4. Run the Application
Start the Streamlit dashboard by running the following command:
```bash
streamlit run app.py
```

This will start the local server, and your default web browser should automatically open the application at `http://localhost:8501`. If it doesn't open automatically, simply copy and paste that link into your browser.
