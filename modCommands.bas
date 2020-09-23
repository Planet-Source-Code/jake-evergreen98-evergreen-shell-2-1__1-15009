Attribute VB_Name = "modCommands"
Option Explicit
DefVar A-Z
Dim oldtxt
Dim iconkey
Dim cmdtoshell
Dim errcode
Dim N
Dim X
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Type POINTAPI
  X As Long
  Y As Long
End Type
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Boolean

Public Const MAX_PATH = 260
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_LARGEICON = &H0
Public Const SHGFI_SMALLICON = &H1
Public Const ILD_TRANSPARENT = &H1
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal Y&, ByVal flags&) As Long

Public shinfo As SHFILEINFO

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type
Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9
Public Const GW_OWNER = 4
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_APPWINDOW = &H40000
Public Const LB_ADDSTRING = &H180
Public Const LB_SETITEMDATA = &H19A
Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String)
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type


Public Sub pSetForegroundWindow(ByVal hwnd As Long)
Dim lForeThreadID As Long
Dim lThisThreadID As Long
Dim lReturn       As Long

If hwnd <> GetForegroundWindow() Then
    
    lForeThreadID = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
    lThisThreadID = GetWindowThreadProcessId(hwnd, ByVal 0&)
    
    If lForeThreadID <> lThisThreadID Then
        Call AttachThreadInput(lForeThreadID, lThisThreadID, True)
        lReturn = SetForegroundWindow(hwnd)
        Call AttachThreadInput(lForeThreadID, lThisThreadID, False)
    Else
       lReturn = SetForegroundWindow(hwnd)
    End If
       If IsIconic(hwnd) Then
       Call ShowWindow(hwnd, SW_RESTORE)
    Else
       Call ShowWindow(hwnd, SW_SHOW)
    End If
End If
End Sub
Public Function WhichWindows(lst) As Long
With lst
    .Clear
    Call EnumWindows(AddressOf WhichWindowsCallBack, .hwnd)
    WhichWindows = .ListCount
End With
End Function

Private Function WhichWindowsCallBack(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim lReturn     As Long
Dim lExStyle    As Long
Dim bNoOwner    As Boolean
Dim sWindowText As String

If hwnd <> frmShell.hwnd Then
    If IsWindowVisible(hwnd) Then
        If GetParent(hwnd) = 0 Then
            bNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
            
            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
                
                sWindowText = Space$(256)
                lReturn = GetWindowText(hwnd, sWindowText, Len(sWindowText))
                If lReturn Then
                  
                   sWindowText = Left$(sWindowText, lReturn)
                   lReturn = SendMessage(lParam, LB_ADDSTRING, 0, ByVal sWindowText)
                   Call SendMessage(lParam, LB_SETITEMDATA, lReturn, ByVal hwnd)
                End If
            End If
        End If
    End If
End If
WhichWindowsCallBack = True
End Function

Sub Restart()
Call ExitWindowsEx(EWX_REBOOT, 0)
End Sub

Sub ShutDown()
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End Sub
Sub ForceShutdown()
Call ExitWindowsEx(EWX_FORCE, 0)
End Sub
Sub LogOff()
Call ExitWindowsEx(EWX_LOGOFF, 0)
End Sub

Function AddASlash(ByVal path As String)
  If Right(path, 1) = "\" Then
    AddASlash = path
  Else
    AddASlash = path & "\"
  End If
End Function

Public Function FileExists(strPath As String) As Integer
  FileExists = Not (Dir(strPath) = "")
End Function

Sub DetermineIcon(extension, first5char, wholename, errcode)
N = N + 1
If errcode = 1 Then Exit Sub
If first5char = UCase$("start") Then
frmShell.txtcmd.ComboItems.Add N, , wholename, 8
frmShell.txtcmd.ComboItems.Item(N).Selected = True
Exit Sub
End If
Select Case extension
Case "exe"
iconkey = "exe"
Case "bat"
iconkey = "bat"
Case "url"
iconkey = "url"
Case "pif"
iconkey = "bat"
Case "ini"
iconkey = "ini"
Case "txt"
iconkey = "ini"
Case "com"
iconkey = "bat"
Case Else
iconkey = "xxx"
End Select
Select Case iconkey
Case "exe"
frmShell.txtcmd.ComboItems.Add , , wholename, 1
frmShell.txtcmd.ComboItems.Item(N).Selected = True
Case "bat"
frmShell.txtcmd.ComboItems.Add N, , wholename, 2
frmShell.txtcmd.ComboItems.Item(N).Selected = True
Case "url"
frmShell.txtcmd.ComboItems.Add N, , wholename, 3
frmShell.txtcmd.ComboItems.Item(N).Selected = True
Case "ini"
frmShell.txtcmd.ComboItems.Add N, , wholename, 6
frmShell.txtcmd.ComboItems.Item(N).Selected = True
Case "xxx"
frmShell.txtcmd.ComboItems.Add N, , wholename, 5
frmShell.txtcmd.ComboItems.Item(N).Selected = True
End Select
End Sub
Sub YAddIconX(iconkey)

End Sub
Sub ShellApp(wt, errcode)
On Error GoTo ErrHandlR
Select Case wt
'wt = whole_thing, a short form of the earlier variable
'Drive Handling Sub
'kind of arcane looking, but it works until F:\
'I'll need a new subroutine for this.
Case "A:\"
X = Shell("explorer " & wt, vbNormalFocus)
Exit Sub
Case "B:\"
X = Shell("explorer " & wt, vbNormalFocus)
Exit Sub
Case "C:\"
X = Shell("explorer " & wt, vbNormalFocus)
Exit Sub
Case "D:\"
X = Shell("explorer " & wt, vbNormalFocus)
Exit Sub
Case "E:\"
X = Shell("explorer " & wt, vbNormalFocus)
Exit Sub
Case "F:\"
X = Shell("explorer " & wt, vbNormalFocus)
Exit Sub
'more drive support will be added later...maybe handling network drives
'now for the tough part...
End Select
Select Case UCase$(Right$(wt, 3))
Case "pif"
X = Shell("start " & wt, vbNormalFocus)
Exit Sub
Case "ico"
X = Shell("pbrush " & wt, vbNormalFocus)
Exit Sub
Case "ini"
X = Shell("notepad " & wt, vbNormalFocus)
Exit Sub
Case "txt"
X = Shell("notepad " & wt, vbNormalFocus)
Exit Sub
Case "htm"
X = Shell("start " & wt, vbNormalFocus)
Exit Sub
End Select
'Note: for Web addresses, you can type "start" and your www address in
'the bar...these will be given an icon looking like a web page
X = Shell(wt, vbNormalFocus)
Exit Sub
'I'll add support for these types of files once I get errorhandling
'implemented: (*'s are already implemented)
'PIF ICO INI TXT HTM DOC RTF XLS PPT BMP GIF JPG
'LNK EML DDB CPL MOV MPG MP2 MP3 WMA SCF XML ASP
'UIN ZIP TAR GZ BIN IMG ISO CDR PUB
'These programs may or may not open with their associated programs.
'You could write it yourself...
ErrHandlR:
MsgBox "The program or document you typed in could not be opened."
Exit Sub
End Sub

Sub OpenSelProgram(ListIndex, xtensions As Boolean)
Dim X
'Here is where we take programs and shell them out.
Select Case ListIndex
Case 0
ShellPath ("calc")
Case 1
ShellPath ("charmap")
Case 2
ShellPath ("defrag")
Case 3
ShellPath ("hypertrm")
Case 4
ShellPath ("C:\Program Files\Internet Explorer\iexplore.exe")
Case 5
ShellPath ("C:\Program Files\Microsoft Office\office\msaccess.exe")
Case 6
ShellPath ("msbackup") ' this may not work, I'm unsure of the path
Case 7
ShellPath ("C:\Program Files\Microsoft Office\office\excel.exe")
Case 8
ShellPath ("C:\Program Files\Microsoft Office\office\frontpg.exe")
Case 9
ShellPath ("C:\Program Files\Microsoft Office\office\outlook.exe")
Case 10
ShellPath ("C:\Program Files\Microsoft Office\office\powerpnt.exe")
Case 11
ShellPath ("C:\Program Files\Microsoft Office\office\mspub.exe")
Case 12
ShellPath ("C:\Program Files\Microsoft Office\office\winword.exe")
Case 13
ShellPath ("C:\Program Files\MSWorks\msworks.exe")
Case 14
ShellPath ("command")
Case 15
ShellPath ("C:\Program Files\Napster\napster.exe")
Case 16
ShellPath ("C:\Program Files\Netscape\Communicator\Program\netscape.exe")
Case 17
ShellPath ("notepad")
Case 18
ShellPath ("C:\Program Files\Outlook Express\msimn.exe")
Case 19
ShellPath ("pbrush") ' this is improper, but it still works.
Case 20
ShellPath ("scandisk")
Case 21
ShellPath ("msinfo32")
Case 22
ShellPath ("explorer")
Case 23
ShellPath ("write") ' again, same as Paint, it's improper, but works
End Select
End Sub
Sub ShellPath(path)
On Error GoTo ErrorHandl
Dim X
X = Shell(path, vbNormalFocus)
Exit Sub
ErrorHandl:
If Err = 53 Then
MsgBox "When trying to launch your program, I couldn't find an EXE file." & vbNewLine & _
"You may have typed something in wrong, or Evergreen OS can't find the program." & vbNewLine & _
"Check to see that you can launch other programs. If you can, then there may be a problem with the OS. Please check for updates!", vbExclamation, "Couldn't Find Program: " & path
Exit Sub
Else
MsgBox "When trying to launch your program, I received an error. The error code was " & Err & vbNewLine & _
"Please check that you can launch other programs. If you can, you may need to update Evergreen OS.", vbExclamation, "Error: " & Err
Exit Sub
End If
End Sub
