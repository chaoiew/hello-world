@echo off
:: 批量转换当前目录及子目录下的所有.csv文件为UTF-8无BOM
for /r %%f in (*.csv) do (
    powershell -Command "$content = Get-Content '%%f'; Set-Content -Path '%%f' -Value $content -Encoding UTF8NoBOM"
)
echo 转换完成！
pause
