' Start and supervise the local service at logon without a console window.
Option Explicit

Dim shell, fso, root, node, server, dist, pfx, password, manifest, registration, channelRegistration
Dim readyFile, pidFile, tokenFile, token, lockPath, lockOwnerPath, lockCancelPath, stopFile, logFile, command
Dim wmi, startup, processClass, processId, launchResult, processes, child, attempts, nodeCreatedAt
Dim ownedChildren
Dim managedLaunch, developerChannel
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

root = fso.GetParentFolderName(WScript.ScriptFullName)
node = fso.BuildPath(root, "runtime\node.exe")
server = fso.BuildPath(root, "app\local-server.mjs")
dist = fso.BuildPath(root, "app\dist")
pfx = fso.BuildPath(root, "certificate\localhost.pfx")
password = fso.BuildPath(root, "certificate\pfx-password.txt")
manifest = fso.BuildPath(root, "app\manifest.xml")
readyFile = fso.BuildPath(root, "service.ready")
pidFile = fso.BuildPath(root, "service.pid")
logFile = fso.BuildPath(root, "service.log")
tokenFile = fso.BuildPath(root, "service.instance")
Randomize
token = fso.GetTempName & "-" & CStr(CLng(Timer * 1000)) & "-" & CStr(Int(Rnd * 1000000000))
lockPath = fso.BuildPath(root, "service.launching")
lockOwnerPath = fso.BuildPath(lockPath, "owner")
lockCancelPath = fso.BuildPath(lockPath, "cancel")
stopFile = fso.BuildPath(root, "service.stop")
registration = "HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\6374B2A1-D997-4BB0-B23B-17F28561827B"
channelRegistration = "HKCU\Software\DdotExcel\Channel"

Sub AppendLog(message)
  Dim stream
  On Error Resume Next
  Set stream = fso.OpenTextFile(logFile, 8, True, 0)
  If Err.Number = 0 Then
    stream.WriteLine CStr(Now) & " launcher " & message
    stream.Close
  End If
  Err.Clear
  On Error GoTo 0
End Sub

Function ServiceHealthy()
  Dim request
  ServiceHealthy = False
  On Error Resume Next
  Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
  request.SetTimeouts 1000, 1000, 1000, 1000
  request.Open "GET", "https://localhost:3927/health", False
  request.Send
  If Err.Number = 0 Then
    ServiceHealthy = (request.Status = 200 And _
      InStr(request.ResponseText, """service"":""ddot-excel""") > 0 And _
      InStr(request.ResponseText, """status"":""running""") > 0)
  End If
  Err.Clear
  On Error GoTo 0
End Function

Function UseDeveloperChannel()
  Dim value, errorNumber
  UseDeveloperChannel = True
  On Error Resume Next
  value = shell.RegRead(channelRegistration)
  errorNumber = Err.Number
  Err.Clear
  On Error GoTo 0
  If errorNumber = 0 Then
    UseDeveloperChannel = (LCase(CStr(value)) <> "trusted-catalog")
  ElseIf errorNumber <> -2147024894 Then
    AppendLog "channel read failed: " & CStr(errorNumber)
    WScript.Quit 2
  End If
End Function

Function ReadRegistration(ByRef exists)
  Dim value, errorNumber
  exists = False
  value = ""
  On Error Resume Next
  value = shell.RegRead(registration)
  errorNumber = Err.Number
  Err.Clear
  On Error GoTo 0
  If errorNumber = 0 Then
    exists = True
  ElseIf errorNumber <> -2147024894 Then
    AppendLog "registry read failed: " & CStr(errorNumber)
    WScript.Quit 2
  End If
  ReadRegistration = CStr(value)
End Function

Sub RemoveOwnedRegistration()
  Dim exists, value, remainingExists, remaining
  value = ReadRegistration(exists)
  If Not exists Then Exit Sub
  If LCase(value) <> LCase(manifest) Then Exit Sub
  On Error Resume Next
  shell.RegDelete registration
  If Err.Number <> 0 Then
    AppendLog "registry delete failed: " & CStr(Err.Number)
    WScript.Quit 3
  End If
  Err.Clear
  On Error GoTo 0
  remaining = ReadRegistration(remainingExists)
  If remainingExists And LCase(remaining) = LCase(manifest) Then
    AppendLog "registry delete verification failed"
    WScript.Quit 3
  End If
End Sub

Sub DeleteMarker(path)
  If Not fso.FileExists(path) Then Exit Sub
  On Error Resume Next
  fso.DeleteFile path, True
  If Err.Number <> 0 Then
    AppendLog "marker delete failed: " & path & " (" & CStr(Err.Number) & ")"
    WScript.Quit 4
  End If
  Err.Clear
  On Error GoTo 0
End Sub

Function OwnsToken()
  Dim stream, value
  OwnsToken = False
  If Not fso.FileExists(tokenFile) Then Exit Function
  On Error Resume Next
  Set stream = fso.OpenTextFile(tokenFile, 1, False)
  value = stream.ReadAll
  stream.Close
  If Err.Number = 0 Then OwnsToken = (Trim(value) = token)
  If Err.Number <> 0 Then AppendLog "instance token read failed: " & CStr(Err.Number)
  Err.Clear
  On Error GoTo 0
End Function

Function ReadFileValue(path)
  Dim stream, value
  ReadFileValue = ""
  If Not fso.FileExists(path) Then Exit Function
  On Error Resume Next
  Set stream = fso.OpenTextFile(path, 1, False)
  value = stream.ReadAll
  stream.Close
  If Err.Number <> 0 Then
    AppendLog "file read failed: " & path & " (" & CStr(Err.Number) & ")"
    WScript.Quit 6
  End If
  Err.Clear
  On Error GoTo 0
  ReadFileValue = Trim(value)
End Function

Sub ReleaseLaunchLock()
  If Not fso.FolderExists(lockPath) Then Exit Sub
  If ReadFileValue(lockOwnerPath) <> token Then Exit Sub
  On Error Resume Next
  fso.DeleteFolder lockPath, True
  If Err.Number <> 0 Then AppendLog "launch lock release failed: " & CStr(Err.Number)
  Err.Clear
  On Error GoTo 0
End Sub

Sub AcquireLaunchLock()
  Dim observedOwner, index, folder, stream
  If fso.FolderExists(lockPath) Then
    observedOwner = ReadFileValue(lockOwnerPath)
    For index = 1 To 15
      If ServiceHealthy() Then WScript.Quit 0
      WScript.Sleep 1000
    Next
    If fso.FolderExists(lockPath) Then
      If ReadFileValue(lockOwnerPath) <> observedOwner Then WScript.Quit 0
      On Error Resume Next
      fso.DeleteFolder lockPath, True
      If Err.Number <> 0 Then
        AppendLog "stale launch lock cleanup failed: " & CStr(Err.Number)
        WScript.Quit 6
      End If
      Err.Clear
      On Error GoTo 0
    End If
  End If
  On Error Resume Next
  Set folder = fso.CreateFolder(lockPath)
  If Err.Number <> 0 Then
    AppendLog "launch lock acquisition failed: " & CStr(Err.Number)
    WScript.Quit 6
  End If
  Set stream = fso.CreateTextFile(lockOwnerPath, True, False)
  stream.Write token
  stream.Close
  If Err.Number <> 0 Then
    AppendLog "launch lock owner write failed: " & CStr(Err.Number)
    WScript.Quit 6
  End If
  Err.Clear
  On Error GoTo 0
  If ReadFileValue(lockOwnerPath) <> token Then
    AppendLog "launch lock verification failed"
    WScript.Quit 6
  End If
End Sub

Function LaunchCancelled()
  LaunchCancelled = (fso.FileExists(lockCancelPath) Or fso.FileExists(stopFile))
End Function

If Not fso.FileExists(node) Then
  AppendLog "node.exe is missing"
  WScript.Quit 1
End If
If Not fso.FileExists(server) Then
  AppendLog "local-server.mjs is missing"
  WScript.Quit 1
End If
managedLaunch = False
If WScript.Arguments.Count > 0 Then
  managedLaunch = (LCase(WScript.Arguments(0)) = "managed")
End If
If Not managedLaunch Then DeleteMarker stopFile
If fso.FileExists(stopFile) Then WScript.Quit 0

' A duplicate launcher must not tear down the lease of an already healthy service.
If ServiceHealthy() Then WScript.Quit 0
AcquireLaunchLock
If ServiceHealthy() Then
  ReleaseLaunchLock
  WScript.Quit 0
End If
If LaunchCancelled() Then
  ReleaseLaunchLock
  WScript.Quit 0
End If
developerChannel = UseDeveloperChannel()
RemoveOwnedRegistration
DeleteMarker readyFile
DeleteMarker pidFile
DeleteMarker tokenFile

command = """" & node & """ """ & server & """" & _
  " --root """ & dist & """" & _
  " --pfx """ & pfx & """" & _
  " --passphrase-file """ & password & """" & _
  " --port 3927" & _
  " --ready-file """ & readyFile & """" & _
  " --pid-file """ & pidFile & """" & _
  " --instance-token """ & token & """" & _
  " --token-file """ & tokenFile & """"
If developerChannel Then
  command = command & _
    " --wef-guid ""6374B2A1-D997-4BB0-B23B-17F28561827B""" & _
    " --wef-manifest """ & manifest & """"
End If
command = command & " --log-file """ & logFile & """"

' WMI exposes the Node PID, unlike WScript.Shell.Run. That lets this supervisor terminate a
' synchronous registry PowerShell child if Node is killed while the child is still running.
If LaunchCancelled() Then
  ReleaseLaunchLock
  WScript.Quit 0
End If
On Error Resume Next
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
Set startup = wmi.Get("Win32_ProcessStartup").SpawnInstance_
startup.ShowWindow = 0
Set processClass = wmi.Get("Win32_Process")
launchResult = processClass.Create(command, root, startup, processId)
If Err.Number <> 0 Or launchResult <> 0 Then
  AppendLog "node launch failed: " & CStr(Err.Number) & "/" & CStr(launchResult)
  ReleaseLaunchLock
  WScript.Quit 5
End If
Err.Clear
On Error GoTo 0

nodeCreatedAt = ""
Do
  Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & processId)
  If processes.Count = 0 Then Exit Do
  For Each child In processes
    If nodeCreatedAt = "" Then nodeCreatedAt = child.CreationDate
    If LaunchCancelled() Then
      On Error Resume Next
      child.Terminate
      Err.Clear
      On Error GoTo 0
    End If
  Next
  WScript.Sleep 500
Loop

For attempts = 1 To 20
  ownedChildren = 0
  Set processes = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE ParentProcessId = " & _
    processId & " AND Name = 'powershell.exe'")
  If processes.Count = 0 Then Exit For
  For Each child In processes
    If nodeCreatedAt <> "" And Not IsNull(child.CommandLine) And _
      InStr(LCase(CStr(child.CommandLine)), "-encodedcommand") > 0 And _
      CStr(child.CreationDate) >= nodeCreatedAt Then
      ownedChildren = ownedChildren + 1
      On Error Resume Next
      child.Terminate
      If Err.Number <> 0 Then AppendLog "child termination failed: " & CStr(Err.Number)
      Err.Clear
      On Error GoTo 0
    End If
  Next
  If ownedChildren = 0 Then Exit For
  WScript.Sleep 500
Next
If ownedChildren <> 0 Then
  AppendLog "node child tree did not terminate"
  ReleaseLaunchLock
  WScript.Quit 7
End If

If OwnsToken() Then
  RemoveOwnedRegistration
  DeleteMarker readyFile
  DeleteMarker pidFile
  DeleteMarker tokenFile
End If
ReleaseLaunchLock
AppendLog "node exited; child tree drained"
WScript.Quit 0
