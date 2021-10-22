@ECHO OFF

SET "ServerList=C:\Users\Automation\Documents\Ping Reports\Computers.txt"
SET "LogFile=C:\Users\Automation\Documents\Ping Reports\PingResult_%DATE:~-4%_%DATE:~3,2%_%DATE:~0,2%.txt"

If Not Exist "%ServerList%" Exit /B
>"%LogFile%" (For /F UseBackQ %%A In ("%ServerList%") Do Ping -n 1 %%A|Find "TTL=">Nul&&(Echo Yes[%%A)||Echo No [%%A])
