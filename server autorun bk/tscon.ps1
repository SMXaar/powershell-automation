$SessionId = (Get-Process -Id $PID).SessionId
Start-Process -FilePath "tscon.exe" -ArgumentList "$SessionId /dest:console" -Verb RunAs