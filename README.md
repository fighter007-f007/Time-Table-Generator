# Student Timetable Generator

A web app for students to get their personal timetable by entering **only
their roll number**.

## What students do

1. Open the website.
2. Enter roll number.
3. Press Enter / click **Generate My Timetable**.
4. View the personal timetable in the same six-period structure as the
   master timetable.
5. Download Excel or a copyable TXT version.

Students do **not** upload files.

## Admin data

Put the full subject student-list files and the master timetable in `data/`.
The included sample data has been sanitized to remove administrative email
sheets before being placed in the repository.

Example:

```text
data/
├── CRM.xlsx
├── MDT.xlsx
├── SDT.xlsx
├── SO.xlsx
├── BRM.xlsx
└── PGDM (2025-27) Term-V Time Table.pdf
```

For a new term, replace the subject files and timetable in `data/` and push
the changes to GitHub.

## Deploy to Streamlit Community Cloud

Use `streamlit_app.py` as the entrypoint.

1. Create a GitHub repository.
2. Upload the contents of this folder.
3. Connect GitHub to Streamlit Community Cloud.
4. Create a new app.
5. Select the repository, branch, and `streamlit_app.py`.
6. Deploy.
7. Share the generated `*.streamlit.app` URL with students.

## Important privacy note

Do not put student email addresses or other sensitive information in a public
GitHub repository. The sample files bundled here have already had the
administrative email sheets removed.

The app itself only needs roll number, name, and any section information
required to construct the timetable.

## Supported timetable formats

The current web build is designed primarily for the master timetable PDF.
The shared parser also retains support for other timetable formats used by
the desktop version where applicable.
