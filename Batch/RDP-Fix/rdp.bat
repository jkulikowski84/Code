@echo off
CLS

netsh.exe advfirewall firewall add rule name="Remote Desktop - User Mode (TCP-In)" dir=in action=allow program="%%SystemRoot%%\system32\svchost.exe" service="TermService" description="Inbound rule for the Remote Desktop service to allow RDP traffic. [TCP 3389] added by LogicDaemon's script" enable=yes profile=private,domain localport=3389 protocol=tcp

netsh.exe advfirewall firewall add rule name="Remote Desktop - User Mode (UDP-In)" dir=in action=allow program="%%SystemRoot%%\system32\svchost.exe" service="TermService" description="Inbound rule for the Remote Desktop service to allow RDP traffic. [UDP 3389] added by LogicDaemon's script" enable=yes profile=private,domain localport=3389 protocol=udp

netsh advfirewall firewall add rule name="Open Remote Desktop" protocol=TCP dir=in localport=3389 action=allow

REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server" /v fDenyTSConnections /t REG_DWORD /d 0 /f

netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes