# Study group optimisation

A small web app that assigns students to study groups from an Excel spreadsheet. It uses constraint optimisation (Google OR-Tools CP-SAT) to form groups of a given size, optionally same-sex or mixed, and can avoid putting students together who were in the same group before.

## Installation

1. **Clone or copy the project** into a folder on your machine.

2. **Create a virtual environment** (recommended):

   ```bash
   python3 -m venv venv
   source venv/bin/activate   # Linux/macOS
   # or: venv\Scripts\activate   # Windows
   ```

3. **Install dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

   This installs Flask, openpyxl, and ortools.

## Running the app

1. Start the server:

   ```bash
   python app.py
   ```

2. Open a browser and go to:

   **http://127.0.0.1:5000/**

3. Use the web interface to create study groups (see [Usage](#usage) below).

## Usage

### Input spreadsheet

- Upload an **Excel file (.xlsx)**.
- The app reads from the **first sheet** and expects:
  - **Column B** – sex (e.g. M/F, K/M, male/female, mand/kvinde).
  - **Column C** – first name.
  - **Column D** – last name.
- Rows without a valid sex (F or M) are skipped. The first row can be a header.

### Settings

- **Group size** – Number of students per group (e.g. 10). The solver forms as many groups as needed; some groups may be slightly smaller if the total does not divide evenly.
- **Group type**:
  - **Same-sex** – Each group contains only male or only female students.
  - **Mixed** – Groups can contain both sexes.

### Creating groups

1. Click **Run optimisation**. The app assigns students to groups and shows the result on screen.
2. You can:
   - **Download .csv** – Get a CSV of the groups (Group, Sex, Name).
   - **Use groups** – Download your original spreadsheet with a new worksheet added (e.g. `study_group_1`) containing the assignments (Group, Sex, Name). You can keep this file for later and re-upload it; the app will try to avoid putting the same pairs together again.

### Notes

- Download links are stored in memory and expire when you close or restart the app.
- The solver prefers to minimise how often students who were previously in the same group are placed together again (when prior groups are read from the file).

## Requirements

- Python 3.8+
- See `requirements.txt`: Flask, openpyxl, ortools
