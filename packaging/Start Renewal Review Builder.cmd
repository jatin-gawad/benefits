@echo off
setlocal
set "APP_DIR=%~dp0"
set "APP_EXE=%APP_DIR%RenewalReviewBuilder.App.exe"

if not exist "%APP_EXE%" (
  echo Renewal Review Builder is not fully extracted.
  echo.
  echo If you opened this from inside a .zip file, close this window and extract the whole zip first:
  echo 1. Right-click RenewalReviewBuilderPortable.zip
  echo 2. Choose Extract All
  echo 3. Open the extracted RenewalReviewBuilderPortable folder
  echo 4. Double-click this launcher again
  echo.
  echo Renewal Review Builder could not be found in this folder.
  echo.
  echo Expected:
  echo "%APP_EXE%"
  echo.
  pause
  exit /b 1
)

start "" "%APP_EXE%"
if errorlevel 1 (
  echo Windows blocked or could not start Renewal Review Builder.
  echo.
  echo Try these steps:
  echo 1. Move the folder to your Desktop or Documents.
  echo 2. Right-click RenewalReviewBuilder.App.exe, choose Properties, then Unblock if that button appears.
  echo 3. Ask IT to allow this internal Westland portable app if company policy blocks unsigned apps.
  echo.
  pause
  exit /b 1
)

endlocal
