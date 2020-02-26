@ECHO OFF

for /f "skip=1" %%d in ('wmic os get localdatetime') do if not defined mydate set mydate=%%d
::cd C:\Users\asayeed\Desktop\Python Projects\Pipeline Cognos Report Cleanup\%mydate:~0,8%



copy "C:\Users\asayeed\Desktop\Python Projects\Pipeline Cognos Report Cleanup\%mydate:~0,8%\*.xlsx" "C:\Users\asayeed\Desktop\Python Projects\Pipeline Cognos Report Cleanup\Run\*.xlsx" >ListofCopiedFiles.txt

"C:\Users\asayeed\AppData\Local\Continuum\anaconda3\envs\CognosCleanup\python.exe" "C:\Users\asayeed\Desktop\Python Projects\Pipeline Cognos Report Cleanup\Run\CognosCleanup.py" >log.txt


