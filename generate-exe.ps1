Set-Variable -Name "name" -Value "AI Voice Assistant"

pyinstaller -i icon.ico -n $name --hidden-import=pyttsx3.drivers.sapi5 --onefile assistant.py
Copy-Item ./dist/$name.exe ./