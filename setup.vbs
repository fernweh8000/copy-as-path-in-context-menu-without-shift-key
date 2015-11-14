Set ws = CreateObject("WScript.shell"):
Set f = CreateObject("Scripting.FileSystemObject"):

::COMMENT="Copy As File を コンテキストメニューに追加"::
Sub RegisterCopyAsFileMenu(ByVal ccexepath):
	::COMMENT="環境変数にc-c.exeのPATHを追加(すでに登録済みなら追加しない)"::
	Set userEnv = ws.Environment("USER"):
	envPath = userEnv.Item("PATH"):
	If InStr(envPath & ";", ccexepath) = 0 And Right(envPath, 1) <> ";" Then envpath = envPath & ";":
	If InStr(envPath & ";", ccexepath) = 0 Then envPath = envPath & ccexepath:
	userEnv.Item("PATH") = envPath:

	::COMMENT="レジストリでコンテキストメニューに追加"::
	Call ws.RegWrite("HKEY_CLASSES_ROOT\AllFilesystemObjects\shell\fernweh.jp.CopyAsPath\", "Copy as Path", "REG_SZ"):
	Call ws.RegWrite("HKEY_CLASSES_ROOT\AllFilesystemObjects\shell\fernweh.jp.CopyAsPath\command\", """"&ccexepath&""" ""%1""", "REG_SZ"):
End Sub:

::COMMENT="フォルダを再帰的に作成(http://scripting.cocolog-nifty.com/blog/2008/11/createfolder-5c.html)"::
Sub CreateFolder(ByVal Folder):
	Dim ParentFolder:
	ParentFolder=f.GetParentFolderName(Folder):
	If ParentFolder<>"" Then If Not f.FolderExists(ParentFolder) Then CreateFolder ParentFolder:
	If Not f.FolderExists(Folder) Then f.CreateFolder Folder:
End Sub:

::COMMENT="UAC Elevationの参考: http://masahiror.hatenadiary.jp/entry/20111201/vbs_admin_run"::
Function runasCheck():
  Set args = WScript.Arguments:

  flgRunasMode = False:
  
  If args.Count > 0 Then arg1 = args.item(0):
  If UCase(arg1) = "/RUNAS" Then flgRunasMode = True:

  Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2"):
  Set osInfo = objWMI.ExecQuery("SELECT * FROM Win32_OperatingSystem"):
  flag = false:
  For Each os in osInfo:
    If Left(os.Version, 3) >= 6.0 Then flag = True:
  Next:

  Set objShell = CreateObject("Shell.Application"):
  If flgRunasMode = False And flag = True Then objShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """" & " /RUNAS", "", "runas", 1:
  If flgRunasMode = False And flag = True Then Wscript.Quit:
  
  MsgBox "Setup completed."
End Function:

runasCheck():

::COMMENT="c-c.exeをexePathからexeDestPathへコピー(上書き)"::
exePath = f.GetParentFolderName(f.GetFile(WScript.ScriptFullName)) & "\CopyToClipboard\Release\c-c.exe":
exeDestDir = ws.ExpandEnvironmentStrings("%LocalAppData%") & "\fernweh.jp\c-c":
exeDestPath = exeDestDir & "\c-c.exe":

CreateFolder(exeDestDir):
Call f.CopyFile(exePath, exeDestPath):

::COMMENT="レジストリに登録"::
If f.FileExists(exeDestPath) Then RegisterCopyAsFileMenu(exeDestPath):
