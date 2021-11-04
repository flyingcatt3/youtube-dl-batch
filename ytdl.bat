rem Author:lazzicat
rem 20211104

@ECHO OFF
chcp 65001
set error="& {[void][reflection.assembly]::loadwithpartialname('System.Windows.Forms');[void][reflection.assembly]::loadwithpartialname('System.Drawing');$notify = new-object system.windows.forms.notifyicon;$notify.icon = [System.Drawing.SystemIcons]::Information;$notify.visible = $true;$notify.showballoontip(10,'發生錯誤!','Error!',[system.windows.forms.tooltipicon]::None);(New-Object Media.SoundPlayer "C:\Windows\Media\Alarm08.wav").PlaySync();}"
set completed="& {[void][reflection.assembly]::loadwithpartialname('System.Windows.Forms');[void][reflection.assembly]::loadwithpartialname('System.Drawing');$notify = new-object system.windows.forms.notifyicon;$notify.icon = [System.Drawing.SystemIcons]::Information;$notify.visible = $true;$notify.showballoontip(10,'已完成作業','The task has completed.',[system.windows.forms.tooltipicon]::None);(New-Object Media.SoundPlayer "C:\Windows\Media\Ring03.wav").PlaySync()}"
color 0E

echo #### 正在下載...
echo #### Downloading...
echo/

MOVE "%cd%\webm\*.webm" "%cd%\webm_old\" 2>nul >nul
FOR /F "tokens=* USEBACKQ" %%F IN (`powershell -sta "add-type -as System.Windows.Forms; [windows.forms.clipboard]::GetText()"`) DO (
    echo #### 複製的內容為
    echo #### %%F
    echo/
    youtube-dl %%F -f 251 -o "%cd%\webm\%%(title)s.%%(ext)s" --console-title -ciw
    IF ERRORLEVEL 1 (
        echo/
        color 0c
        echo #### 發生錯誤
        echo #### An error occured.
        title Error!
        powershell -c %error%;
        PAUSE
        exit
        )
)

echo/
echo/
echo/

echo #### 正在轉換成 opus 格式...
echo #### Converting to opus format...
echo/
set count=0
setlocal
set output=%time::=_%_outputLog.txt
FOR /R "%cd%\webm\" %%F IN (*.webm) DO (
    rename "%%F" %count%.webm
    mkvextract "%cd%\webm\%count%.webm" tracks 0:"%cd%\opus\%%~nF.opus"
    rename "%cd%\webm\%count%.webm" "%%~nxF"
    echo %time% - %%~nF.opus>> %output%
    set /a count+=1
)

color 0a
echo/
echo/
echo/
echo #### 作業已完成
echo #### The task has completed.
title Completed.

echo/
echo #### 已完成的檔案:
echo/
type %output%
explorer %cd%\opus
powershell -c %completed%
PAUSE