@echo off
cls

pushd "%~dp0"

del /A:S /F /Q "aeskey.txt"
del /A:S /F /Q "credpassword.txt"
del /A:S /F /Q "Setup.ps1"

::del /F /Q "%temp%\aeskey.txt"
::del /F /Q "%temp%\credpassword.txt"

::Self Delete
DEL /A:S "%~f0"

popd