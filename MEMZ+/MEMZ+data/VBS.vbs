' Memz+ Simulator V1.0 (Safe Edition)

Option Explicit
Dim objShell, response, i

Set objShell = CreateObject("WScript.Shell")

' ------------------------
' Фаза 0 — Перше попередження
' ------------------------
response = MsgBox("Are you sure you want to run it?", vbYesNo + vbExclamation, "MEMZ+")
If response = vbNo Then WScript.Quit

' ------------------------
' Фаза 1 — Остаточне попередження
' ------------------------
response = MsgBox("This is the last warning! Do you want to run this?!", vbYesNo + vbCritical, "MEMZ+")
If response = vbNo Then WScript.Quit

' ------------------------
' Фаза 2 — Фейковий блокнот
' ------------------------
objShell.Run "notepad"
WScript.Sleep 500
objShell.AppActivate "Notepad"
objShell.SendKeys "Your computer will be destroyed by the MEMZ+ virus :) then use your computer while you can"

' ------------------------
' Фаза 3 — Хаос
' ------------------------
For i = 1 To 15
    MsgBox "ERROR 0xMEMZ+", vbCritical, "MEMZ+"
Next

' ------------------------
' Фаза 4 — Фінал
' ------------------------
MsgBox "Ваш комп’ютер був знищений вірусом MEMZ+ ??" & vbCrLf & "Зустрічайте Nyancat!", vbInformation, "MEMZ+"

' ------------------------
' Відкриття Nyancat на екрані
' ------------------------
' Вкажи свій локальний шлях до відео / gif Nyancat
Dim nyanPath
nyanPath = "C:\Users\Public\Videos\nyancat.mp4"
objShell.Run "wmplayer """ & nyanPath & """"