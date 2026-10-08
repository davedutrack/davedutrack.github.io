# EduTrack Academic

A modern school academic portal for **attendance, examinations, student records, academic resources, analytics and communication**.

## Main modules

- Principal / Teacher / Student login
- Principal-controlled teacher accounts
- Admission-number student accounts
- Student profile editing
- Excel bulk student import with flexible column recognition
- Unit Tests, Half Yearly, Annual and custom examinations
- Subject-wise marks and result analytics
- Daily class roll call: Present / Absent
- Monthly attendance: enter working days once, then present days per student
- Student attendance dashboard with absence dates and percentage
- Class/section-specific dashboard filters remembered on the device
- Academic Hub with:
  - Circular
  - Home Work
  - Time Table
  - Tutorials
  - Question Bank
  - Syllabus
  - Co-Curricular
  - Date Sheet
  - Infirmary information
  - Communication
  - Examination documents
- Teachers and principal can upload resources and target all students or a class/section
- Students only see resources intended for their class/section
- Dark mode
- Responsive/mobile-app style navigation
- Back button and animated page transitions
- Chart-based academic analytics

## Run locally on Windows

Open Command Prompt in the folder containing `app.py`:

```bat
py -m pip install -r requirements.txt
py app.py
```

Open:

```text
http://127.0.0.1:5000
```

The SQLite database `schoolpulse.db` is created automatically on first launch. It is intentionally ignored by Git.

## GitHub

GitHub stores the source code; it does **not** run this Flask application by itself.

```bat
git init
git add .
git commit -m "Initial EduTrack Academic portal"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/edutrack-academic.git
git push -u origin main
```

Never commit real student data, passwords, `schoolpulse.db`, or uploaded school files.

## Host it online

This project is configured for a Python-capable host such as Render:

- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn app:app`
- Runtime: Python
- Secret environment variable: `SCHOOLPULSE_SECRET`

`render.yaml` is included for deployment configuration.

### Important production note

SQLite and the local `uploads/` directory are suitable for an exhibition or small prototype. For a real school's multi-user deployment, use a managed PostgreSQL database and persistent/object storage for uploaded resources. Do not expose a development server or real student records without HTTPS, backups, access controls, audit logging and appropriate school privacy/security procedures.

## Academic Hub uploads

Supported upload types:

PDF, DOC, DOCX, XLS, XLSX, PPT, PPTX, JPG, JPEG, PNG, GIF, TXT, CSV and ZIP.

Maximum upload size: 25 MB.

A teacher can select a category, title, description and optional class/section target. Students see only resources available to their class/section or resources published to everyone.

## Excel student import

Use `Student_Import_Template.xlsx` as a starting point. The importer recognizes common variations such as:

- Admission No / Admission Number / Adm No
- Student Name / Name
- Class / Standard / Grade
- Section / Sec / Division
- Roll / Roll No / Roll Number
- Gender
- DOB / Date of Birth
- Father Name
- Mother Name
- Phone / Mobile
- Email

Column order does not matter. The application previews detected columns before importing.
