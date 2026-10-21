# Dergesa Reports

Windows Forms application for the existing Access database `dergesa.mdb`.

## Reports
- Hyrje dhe dalje sipas artikullit, filtered by date range.
- Gjendja aktuale using `Artikujt.Gjendjemalli`.
- Kartela e artikullit for a selected article and date range.
- CSV export for opening in Excel.

## Database connection
The application asks for an `.mdb` or `.accdb` file. Default path is `dergesa.mdb` next to the EXE.

Connection string:
`Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<database path>;Persist Security Info=False;`

The application is compiled as **x86** to work with the 32-bit Access Database Engine. Install a compatible Microsoft Access Database Engine / ACE OLEDB provider if the connection fails. Do not include the production database in this repository.

## Build using GitHub Actions
1. Create an empty GitHub repository.
2. Upload all files in this project, preserving folders.
3. Commit to `main` (or `master`).
4. Open **Actions** and wait for **Build Dergesa Reports EXE** to complete.
5. Open the completed workflow run and download artifact `DergesaReports-Windows-x86`.
6. Extract the ZIP and run `DergesaReports.exe` on a Windows PC with the compatible ACE provider installed.

## Local build
Open `DergesaReports.sln` in Visual Studio on Windows and build `Release | x86`, or use MSBuild:
`msbuild DergesaReports.sln /p:Configuration=Release /p:Platform=x86`

## Notes
- The app reads existing tables and does not create or alter database tables.
- The date filters are inclusive: the selected end date includes the whole day.
- The app excludes invoices marked `Anulluar=True`.
- Confirm stock field meaning against your workflow before relying on it for inventory accounting.
