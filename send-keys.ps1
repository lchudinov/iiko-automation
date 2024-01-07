# Start Notepad
Start-Process notepad -PassThru | Wait-Process

# Wait for Notepad to open (you may need to adjust the delay based on your system)
Start-Sleep -Seconds 2

# Send keys to Notepad
$shell = New-Object -ComObject WScript.Shell
$shell.AppActivate("Безымянный - Блокнот")  # Activate the Notepad window
Start-Sleep -Seconds 2
$shell.SendKeys("Hello, World!")