@echo off
echo ========================================
echo   PowerPoint Generator - App Launcher
echo ========================================
echo.
echo Transform your text into professional PowerPoint presentations
echo using AI and template styling automatically.
echo.
echo Available options:
echo 1. PowerPoint Generator App (app.html)
echo 2. LLM Agent Demo (index.html)
echo 3. Quick Demo (demo.html)  
echo 4. Examples Gallery (examples.html)
echo 5. Open all applications
echo.
set /p choice="Enter your choice (1-5): "

if "%choice%"=="1" (
    echo Opening PowerPoint Generator...
    start app.html
) else if "%choice%"=="2" (
    echo Opening LLM Agent Demo...
    start index.html
) else if "%choice%"=="3" (
    echo Opening Quick Demo...
    start demo.html
) else if "%choice%"=="4" (
    echo Opening Examples Gallery...
    start examples.html
) else if "%choice%"=="5" (
    echo Opening all applications...
    start app.html
    timeout /t 2 /nobreak >nul
    start index.html
    timeout /t 2 /nobreak >nul
    start demo.html
    timeout /t 2 /nobreak >nul
    start examples.html
) else (
    echo Invalid choice. Opening PowerPoint Generator...
    start app.html
)

echo.
echo ========================================
echo Application launched! Check your browser.
echo ========================================
echo.
echo PowerPoint Generator Tips:
echo - Paste any text content to convert to slides
echo - Use guidance like "investor pitch" or "training"
echo - Upload your own template or use built-in ones
echo - Get API keys from OpenAI, Anthropic, or Google
echo.
echo LLM Agent Demo Tips:
echo - Try: "Search for AI news and analyze trends"
echo - Try: "Calculate fibonacci numbers and create chart"  
echo - Try: "Get weather data and perform analysis"
echo.
pause
