Set shell = CreateObject("Shell.Application")
If Err.Number <> 0 Then 
  MsgBox "Failed to initialize. Run as administrator!", 16, "ERROR"
  WScript.Quit
End If

Set fso = CreateObject("Scripting.FileSystemObject")

workDir = "C:\Users\HP\OneDrive\Documents"
If Not fso.FolderExists(workDir) Then
  fso.CreateFolder(workDir)
End If

logFile = workDir & "\user_manager_log.txt"

Sub LogAction(action)
  Dim log
  Set log = fso.OpenTextFile(logFile, 8, True)
  log.WriteLine Now() & " - " & action
  log.Close
End Sub

Function RunCmd(cmd)
  shell.ShellExecute "cmd.exe", "/c " & cmd & " > " & workDir & "\temp.txt 2>& 1", "", "runas", 0
  WScript.Sleep 2000

If fso.FileExists(workDir & "\temp.txt") Then
  Set file = fso.OpenTextFile(workDir & "\temp.txt")
  result = file.ReadAll
  file.Close
  RunCmd = result
 Else
  RunCmd = "Command executed"
 End If
End Function

  username = InputBox("Username to remove from admin:", "Remove from Admin")
  confirm = MsgBox(" Remove" & " Harshi " & "from admin group", "1", "Confirm")
  If confirm = 6 Then 
     shell.shellExecute = RunCmd("net localgroup administrators " & " Harshi " & "/delete") 
     LogAction("Removed " & " Harshi " & "from admin") 
     MsgBox username & "removed from admin group" 
   End If            