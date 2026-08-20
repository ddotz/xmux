' Start the local service at logon without a console window and without waiting.
'
' The Run key used to launch powershell.exe, which spends a second or two starting the
' engine before it even reaches the server. Excel routinely won that race: it asked for
' https://localhost:3927 while nothing was listening, failed to load the manifest, and
' dropped the add-in registration — the add-in was gone after every restart, and adding it
' again worked only because by then the service had finally come up.
'
' wscript starts in milliseconds and hands straight to node, which closes that window.
' Nothing here validates the installation; manage.ps1 remains the way to start, stop, and
' health-check the service by hand.

Option Explicit

Dim shell, fso, root, node, server, dist, pfx, password, manifest, command
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

root = fso.GetParentFolderName(WScript.ScriptFullName)
node = fso.BuildPath(root, "runtime\node.exe")
server = fso.BuildPath(root, "app\local-server.mjs")
dist = fso.BuildPath(root, "app\dist")
pfx = fso.BuildPath(root, "certificate\localhost.pfx")
password = fso.BuildPath(root, "certificate\pfx-password.txt")
manifest = fso.BuildPath(root, "app\manifest.xml")

If Not fso.FileExists(node) Then WScript.Quit 1
If Not fso.FileExists(server) Then WScript.Quit 1

command = """" & node & """ """ & server & """" & _
  " --root """ & dist & """" & _
  " --pfx """ & pfx & """" & _
  " --passphrase-file """ & password & """" & _
  " --port 3927" & _
  " --ready-file """ & fso.BuildPath(root, "service.ready") & """" & _
  " --pid-file """ & fso.BuildPath(root, "service.pid") & """" & _
  " --wef-guid ""6374B2A1-D997-4BB0-B23B-17F28561827B""" & _
  " --wef-manifest """ & manifest & """"

' 0 = hidden window, False = do not wait for it to exit.
shell.Run command, 0, False
