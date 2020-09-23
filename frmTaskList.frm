VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTaskList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evergreen Shell Task Manager"
   ClientHeight    =   3495
   ClientLeft      =   10860
   ClientTop       =   7575
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox TaskList 
      Height          =   2985
      ItemData        =   "frmTaskList.frx":0000
      Left            =   0
      List            =   "frmTaskList.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   953
      ButtonWidth     =   1508
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Switch To"
            Key             =   "appactivate"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Send Keys"
            Key             =   "sendkeyz"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Terminate"
            Key             =   "abruptend"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "refreshlist"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start Shell"
            Key             =   "restart"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4440
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmTaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Call WhichWindows(TaskList)
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmShell.MenuImages.Tag = "False"
If frmShell.Visible = False Then End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "appactivate"
If TaskList.Text = "" Then MsgBox "You need to select a task before continuing.", vbInformation, "Please select a task to activate.": Exit Sub
f$ = TaskList.Text
AppActivate f$, True
Call WhichWindows(TaskList)
Case "sendkeyz"
If TaskList.Text = "" Then MsgBox "You need to select a task before continuing.", vbInformation, "Please select a task to send keystrokes.": Exit Sub
k$ = TaskList.Text
AppActivate k$
KYS = InputBox("Please enter the keystrokes you wish to send to " & k$ & ".", "Enter Keystrokes")
AppActivate k$
SendKeys KYS, True
Case "abruptend"
MsgBox "Terminate Task is not functional in this version of Evergreen Shell.", vbExclamation, "Warning: Function Not Supported"
Exit Sub
Case "refreshlist"
Call WhichWindows(TaskList)
Case "restart"
Load frmShell
frmShell.Show
Me.Hide
Unload Me
End Select
Call WhichWindows(TaskList)
End Sub
