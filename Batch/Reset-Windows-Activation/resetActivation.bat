reg load HKLM\MY_SYSTEM "D:\Windows\System32\config\system"
reg delete HKLM\MY_SYSTEM\WPA /f
reg unload HKLM\MY_SYSTEM

PAUSE
