rmdir /s /q dist\randomopgaver
rmdir /s /q dist\randomopgavergui
pyinstaller --windowed --icon=icon.ico --clean randomopgavergui.pyw
pyinstaller --console --clean randomopgaver.py
copy dist\randomopgaver\randomopgaver.exe dist\randomopgavergui\randomopgaver.exe
mkdir dist\randomopgavergui\eksempel
copy eksempel\eksempel.docx  dist\randomopgavergui\eksempel\eksempel.docx
copy eksempel\eksempel.xlsx  dist\randomopgavergui\eksempel\eksempel.xlsx
copy icon.ico  dist\randomopgavergui\icon.ico
