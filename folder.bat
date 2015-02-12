@echo off&echo 查找空文件夹
set dd=%~1
set fn=0
if "%~1"=="" set/p dd=./：
cd/d "%dd%"
for /f "delims=" %%a in ('dir/b/s/ad')do (
  dir/a/s/b "%%a"|find/v "">nul||(set/a "fn+=1"&rmdir %%a&echo.路径 %%a 文件夹名 %%~nxa 删除成功))
if defined fn (echo 共找到 %fn% 个空文件夹)else echo 没找到空文件夹
pause