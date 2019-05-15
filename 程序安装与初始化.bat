@echo off
rem 下载python安装包
rem 安装python环境,暂停,手动确认安装完成,注意提示
rem pip3 install pakgs
rem 创建文件加路径(已实现,暂无法校验字符串是否路径)
echo 下载python安装包,下载好后自动开始安装
for /f "delims== tokens=2" %%i in (sfconfig.ini) do 
echo %%i|findstr /\ >nul
pause
if ERRORLEVEL equ 0 (
    if exist %%i (
    echo %%i 路径已经存在
    ) else (
        md %%i
        ))
set num=1&&call echo %%num%%
pause
