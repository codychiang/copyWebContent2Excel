Dim base
base = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

Dim python, script
python = base & "\.venv\Scripts\pythonw.exe"
script = base & "\main.py"

CreateObject("WScript.Shell").Run """" & python & """ """ & script & """", 0, False
