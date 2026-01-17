# Qbit_port_auto_update
Auto updates Qbitorrent listening port after enabling/re-connecting with port forwarding in Proton VPN

Schedule the .ps1 file using task scheduler. Trigger based on your needs, I use every 5 minutes. Use the below settings:

Actions: Start a program  
Program/script: powershell.exe  
Add arguments: -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\update_qbit_port.ps1"  

Remember to enable WebUI in qBittorrent and modify the username and password file path in this script.
