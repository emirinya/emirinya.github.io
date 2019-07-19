chcp 65001
wmic baseboard get serialnumber > C:\pcInfo.txt
wmic csproduct >> C:\pcInfo.txt
wmic cpu get name,caption,NumberOfCores,SocketDesignation,deviceid >> C:\pcInfo.txt
wmic memorychip >> C:\pcInfo.txt
wmic nicconfig get description,ipaddress,macaddress >> C:\pcInfo.txt
wmic diskdrive >> C:\pcInfo.txt
wmic os get caption,installdate,osarchitecture,version,csname,serialnumber >> C:\pcInfo.txt
wmic desktopmonitor get deviceid,pnpdeviceid >> C:\pcInfo.txt
type C:\pcInfo.txt > C:\pcInfor.txt
del C:\pcInfo.txt
dxdiag /whql:off /t C:\display.txt
:search
if exist "C:\display.txt" (findstr /c:"Card name" C:\display.txt >> C:\pcInfor.txt) else (goto search)
del C:\display.txt