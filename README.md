# QuickID — Windows and macOS source

QuickID matches MALDI-TOF ZooMS peak lists against a collagen marker reference spreadsheet. QuickID was conceived by Naihui Wang.

## Download and run

Windows: Download QuickID.exe from this repository's Releases page and double-click it. The EXE is provided separately from the source code; no Python or .NET installation is required to run it.
macOS: Download the source code and build QuickID.app using the macOS instructions below. There is currently no prebuilt Mac app in the release. The source also includes a GitHub Actions workflow that can build Apple Silicon and Intel versions.

## Using the source

Install Python 3 with Tkinter, then run `python quickid.py`. Select a reference `.xlsx` file, a folder of `.txt` or `.tab` peak lists, and a CSV output path; then click **Run**. The twelve built-in marker names default to `P1`, `A`, `A'`, `B`, `C`, `P2`, `D`, `E`, `F`, `F'`, `G`, `G'`. You can edit all twelve names or leave all twelve blank and use only the optional markers. **Clear all** clears the built-in marker names and tolerances.

The default lower tolerance is −0.3 *m/z* for every built-in marker position. The default upper tolerance is +1.3 *m/z* for the positions initially named **P1, A, A′, E, G, G′**, and +0.3 *m/z* for the other six positions. Renaming a marker does not change its tolerances. Values are entered to one decimal place.

Below the twelve built-in markers are eight optional markers in two rows. Each requires a name and both tolerance values; its name must appear as a unique column heading in the reference workbook (case-insensitive), and must not duplicate an active built-in name. Missing or incomplete optional markers are skipped. 

The output CSV puts **Sample name in the first column** and **QuickID marker count in the second**.

The application uses Python's standard library at runtime. **Users of packaged builds do not need Python or .NET installed.** To package, install Python 3 and PyInstaller on the build computer:

### Windows

```powershell
py -m pip install pyinstaller
py -m PyInstaller --noconfirm --clean --onefile --windowed --name QuickID --add-data "app.ini.example:." quickid.py
```

Distribute `dist/QuickID.exe` (built on Windows). You can build from Visual Studio's terminal; Visual Studio alone does not include Python or PyInstaller. If your computer recognizes `python` rather than `py`, substitute `python` in both commands.

### macOS

```sh
python3 -m pip install pyinstaller
python3 -m PyInstaller --noconfirm --clean --onedir --windowed --name QuickID --add-data 'app.ini.example:.' quickid.py
```

Distribute `dist/QuickID.app` as a ZIP created on the Mac. Build on the target CPU architecture (Apple Silicon or Intel), or test a universal build separately. For public Mac downloads, code signing and notarization avoid system security prompts; neither is provided here.

PyInstaller cannot create the Mac app from Windows or the Windows exe from a Mac. The source is shared; build on each operating system. A `.exe` cannot run directly on macOS.

## Limitations and review

Older `.xls` workbooks must be saved as `.xlsx`, and formula cells should have calculated values saved in the workbook. Candidate rows at or above the threshold are aids to interpretation; verify them manually against spectra and the reference. 

Run `python -m unittest discover -s tests -v` for the sample regression checks (`openpyxl` is required only by the tests, to create fixture workbooks).
