@ECHO OFF

IF EXIST Results.TXT DEL Results.TXT

ECHO ------------------------------------------------------------------->>Results.TXT
ECHO Getting Status of Automatic Updates Service on Windows 2003 Servers>>Results.TXT 
ECHO ------------------------------------------------------------------->>Results.TXT
ECHO.>>Results.TXT 
FOR /F %%I IN (W2K3SERVERS.TXT) DO (
NETSVC "Automatic Updates" \\%%I /query>>Results.TXT
ECHO.>>Results.TXT
)

ECHO ------------------------------------------------------------------->>Results.TXT
ECHO Getting Status of Automatic Updates Service on Windows 2000 Servers>>Results.TXT 
ECHO ------------------------------------------------------------------->>Results.TXT
ECHO.>>Results.TXT 
FOR /F %%I IN (W2KSERVERS.TXT) DO (
NETSVC "Automatic Updates" \\%%I /query>>Results.TXT
ECHO.>>Results.TXT
)

NOTEPAD Results.TXT