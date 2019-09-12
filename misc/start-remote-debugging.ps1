<# Start-Remote-Debugging.ps1
Version: 1.0
Description:
Use this powershell script after you have deployed a dummy C# UWP Project to Windows IoT Core.
This script will open TCP port 4020 and UDP port 3702 on the remote device
It will also move the remote debugger to a more accessible folder and launch it.
NOTE: After launching, this powershell session will become busy listening for debug requests.
#>
enter-pssession minwinpc -credential minwinpc\administrator
cd c:\data\users\defaultaccount\appdata\local\developmentfiles\vsremotetools
netsh advfirewall firewall add rule name="Remote Debugging TCP Inbound" dir=in action=allow protocol=TCP localport=4020
netsh advfirewall firewall add rule name="Remote Debugging UDP Inbound" dir=in action=allow protocol=UDP localport=3702
(Start-Process -FilePath "xcopy" -ArgumentList "*.* c:\temp\rdbg /y /s /i /d /h" -Wait -Passthru).ExitCode
(Start-Process -FilePath "c:\temp\rdbg\arm\msvsmon.exe" -ArgumentList "/silent /nostatus /nosecuritywarn /nofirewallwarn /noclrwarn" -Wait -Passthru).ExitCode