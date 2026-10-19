This tool will generate FOD User Guides for New/Deleted/Revised Forms.

Guide:
1. Enter 5-digit Pub Number.
2. Enter Release Number.
3. Enter/Select Revenue Month and Year.
4. Browse and select last release zip file. (Must be in the I:Drive/pubnum/FOD folder)
5. Browse and select userguide reference. (Must be in the I:Drive/pubnum/FOD folder)
6. Enter path for [FT_OUTPUT_PATH] from job summary email.
7. Click Generate to proceed.
8. Click ZipFile to create Zip file of the current release.
9.  Click Exit to close the application."

Version: 1.0
For issues and concerns, please contact: brian.labrador@reedelsevier.com

## Build an EXE with Auto Py to Exe

Use Python 3.13. The project requires Python 3.13 because `aspose-words` does
not currently support Python 3.14.

1. Open PowerShell in this project folder.

   ```powershell
   py -3.13 -m venv .venv
   .\.venv\Scripts\python.exe -m pip install --upgrade pip
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
   ```

2. Start Auto Py to Exe from the same virtual environment.

   ```powershell
   .\.venv\Scripts\auto-py-to-exe.exe
   ```

3. Configure Auto Py to Exe as follows:

   - **Script Location:** `main.py`
   - **Onefile / Onedir:** choose **One Directory**.
   - **Console Window:** choose **Window Based** for the final release. Use
     **Console Based** first if you need to troubleshoot an error.
   - **Icon:** select `logo.ico`.
   - **Additional Files:** add `logo.ico` and `FOD.png` with destination `.`.

4. Click **Convert .py to .exe**.

5. Before sharing the result, make sure the folder containing `FODGen.exe`
   also contains `logo.ico` and `FOD.png`. PyInstaller may place added files
   in `_internal`, but the application loads these two files from beside the
   executable.

6. Share the entire generated application folder, including `_internal`;
   do not distribute only `FODGen.exe`.

### Recipient requirements

- Microsoft Word must be installed because the application uses Word
  automation to convert RTF and DOCX files.
- The recipient must have access to
  `\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary`.
- Runtime errors are written to `FODGen.log` beside the executable.
