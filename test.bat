@echo off
cd /d "%~dp0"
echo line1 > test_log.txt
echo line2 >> test_log.txt
where python >> test_log.txt 2>&1
where pip >> test_log.txt 2>&1
echo line3 >> test_log.txt
