@echo off
chcp 65001 > NUL
title GY-61 阻尼器加速度分析網頁啟動器
echo ---------------------------------------------------
echo   GY-61 阻尼器加速度分析網頁 啟動中...
echo ---------------------------------------------------
echo.
echo 正在開啟瀏覽器預覽網頁...
start http://localhost:8080/index.html
echo.
echo 若服務未啟動，正在嘗試背景啟動 Python HTTP 伺服器...
"C:\Users\927632_st.tc\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "C:\Users\927632_st.tc\Documents\減震球\阻尼器加速度分析網頁\server.py"
pause
