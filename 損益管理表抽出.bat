@echo off
chcp 65001 > nul
echo %~n1から抽出中...
set fullPath=%~f1
node C:\Users\jfe106005\projects\PLExtraction\PLExtraction.js "%fullPath%"
pause
explorer C:\Users\jfe106005\projects\PLExtraction\csv_output