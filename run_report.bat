@echo off
cd /d "%~dp0"
echo ===== %date% %time% ===== >> run_log.txt
python mis_report.py >> run_log.txt 2>&1 && python send_mail.py >> run_log.txt 2>&1