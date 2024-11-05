@echo off

::Clear Screen
cls

::Get the cdkey
For /f "tokens=2 delims=," %%a in ('wmic path SoftwareLicensingService get OA3xOriginalProductKey^,VLRenewalInterval /value /format:csv') do set key=%%a

::Show cd key
echo %key%

::Activate Windows with the cd key
::slmgr //B /ipk %key%
::slmgr //B /ato

Pause
