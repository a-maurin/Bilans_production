Set shell = CreateObject("WScript.Shell")

' Dossier Bilans_production
scriptDir = WScript.ScriptFullName
scriptDir = Left(scriptDir, InStrRev(scriptDir, "\") - 1)

' Commande à exécuter : pythonw sur le script tools\config_profils.py
cmd = "cmd /c cd /d """ & scriptDir & """ && pythonw ""tools\config_profils.py"""

' 0 = fenêtre cachée, False = ne pas attendre la fin
shell.Run cmd, 0, False