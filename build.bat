@echo off
setlocal

echo === Building DMSynth (x64) ===
if not exist build mkdir build
cd build
cmake .. -G "Visual Studio 17 2022" -A x64
if %ERRORLEVEL% neq 0 (
    echo [ERROR] CMake configuration failed ^(x64^).
    pause
    exit /b %ERRORLEVEL%
)
cmake --build . --config Release --target dm_synth
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Build failed ^(x64^).
    pause
    exit /b %ERRORLEVEL%
)
cd ..

echo.
echo === Building DMSynth (x86) ===
if not exist build32 mkdir build32
cd build32
cmake .. -G "Visual Studio 17 2022" -A Win32
if %ERRORLEVEL% neq 0 (
    echo [ERROR] CMake configuration failed ^(x86^).
    pause
    exit /b %ERRORLEVEL%
)
cmake --build . --config Release --target dm_synth
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Build failed ^(x86^).
    pause
    exit /b %ERRORLEVEL%
)
cd ..

echo.
echo ==========================================
echo   Build Successful!
echo   x64: build\Release\dm_synth.exe
echo   x86: build32\Release\dm_synth.exe
echo ==========================================
pause
