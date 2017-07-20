@goto CHECK_PERMISSIONS

:BUILD_AND_PACKAGE
cd /d %~dp0

set ProgramFilesX86=%ProgramFiles%
if exist "%ProgramFiles(x86)%" set ProgramFilesX86=%ProgramFiles(x86)%
"%ProgramFilesX86%\MSBuild\14.0\Bin\MSBuild.exe" /t:Clean,Rebuild /p:Configuration=Debug ..\PerfectXL.VbaCodeAnalyzer.sln
if %errorlevel% NEQ 0 goto BUILDERROR

del setup-PerfectXL.VbaCodeAnalyzer.Host*.exe
"..\packages\Tools.InnoSetup.5.5.9\tools\ISCC.exe" make-installer.iss
if %errorlevel% NEQ 0 goto BUILDERROR

del PerfectXL.VbaCodeAnalyzer.Host*.nupkg
@powershell .\create-package.ps1 "setup-PerfectXL.VbaCodeAnalyzer.Host*.exe"
if %errorlevel% NEQ 0 goto BUILDERROR

@echo.
@echo To push your package:
@echo     .\nuget.exe push *.nupkg -Source https://nuget.perfectxl.com/nuget/choco -ApiKey {APIKEY}
@echo For installation instructions, see:
@echo     https://chocolatey.org/install
@echo         and
@echo     https://nuget.perfectxl.com/feeds/choco/PerfectXL.VbaCodeAnalyzer.Host
@echo (Don't forget that "choco install" can be invoked with a --notsilent option that
@echo enables you to select a custom install location.)
@goto :eof

:BUILDERROR
@echo ERROR: Build failed!
pause
goto :eof

:CHECK_PERMISSIONS
@echo Elevated permissions required. Detecting permissions...

net session >nul 2>&1
if %errorLevel% == 0 (
    echo Success: Elevated permissions confirmed.
    goto BUILD_AND_PACKAGE

) else (
    @echo Failure: Elevated permissions required.
    @echo (Right-click and choose: "Run as administrator")
    pause
)
