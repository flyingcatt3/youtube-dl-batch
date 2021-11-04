@ECHO OFF
chcp 65001
set completed="& {[void][reflection.assembly]::loadwithpartialname('System.Windows.Forms');[void][reflection.assembly]::loadwithpartialname('System.Drawing');$notify = new-object system.windows.forms.notifyicon;$notify.icon = [System.Drawing.SystemIcons]::Information;$notify.visible = $true;$notify.showballoontip(10,'已完成作業','The task has completed.',[system.windows.forms.tooltipicon]::None);(New-Object Media.SoundPlayer "C:\Windows\Media\Ring03.wav").PlaySync()}"
set count=0
setlocal
set output=%time::=_%_outputLog.txt
FOR /R "%cd%\webm\" %%F IN (*.webm) DO (
    rename "%%F" %count%.webm
    D:\data\影片\mkvtoolnix\mkvextract.exe "%cd%\webm\%count%.webm" tracks 0:"%cd%\opus\%%~nF.opus"
    rename "%cd%\webm\%count%.webm" "%%~nxF"
    echo %time% - %%~nF.opus>> %output%
    )
    set /a count+=1


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