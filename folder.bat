@echo off&echo ���ҿ��ļ���
set dd=%~1
set fn=0
if "%~1"=="" set/p dd=./��
cd/d "%dd%"
for /f "delims=" %%a in ('dir/b/s/ad')do (
  dir/a/s/b "%%a"|find/v "">nul||(set/a "fn+=1"&rmdir %%a&echo.·�� %%a �ļ����� %%~nxa ɾ���ɹ�))
if defined fn (echo ���ҵ� %fn% �����ļ���)else echo û�ҵ����ļ���
pause