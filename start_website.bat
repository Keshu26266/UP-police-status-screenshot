@echo off
echo Starting PCC Screenshot Tool website...
echo A browser window will open automatically in a moment.
echo DO NOT close this black window while you are using the site - it is the engine running it.
echo.
python -m pip install -r requirements.txt --quiet
python -m playwright install chromium
python pcc_web_app.py
pause
