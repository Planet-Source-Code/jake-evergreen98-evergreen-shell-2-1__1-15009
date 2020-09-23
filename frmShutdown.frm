VERSION 5.00
Begin VB.Form frmShutdown 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Termination Options"
   ClientHeight    =   2895
   ClientLeft      =   960
   ClientTop       =   1905
   ClientWidth     =   5730
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1998.18
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optPowersave 
      Caption         =   "Power save"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton optMSDOS 
      Caption         =   "MS-DOS mode"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton optExit 
      Caption         =   "Exit Shell"
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.OptionButton optLogoff 
      Caption         =   "Log off network"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1200
      Width           =   2055
   End
   Begin VB.OptionButton optRestart 
      Caption         =   "Restart"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.OptionButton optShutdown 
      Caption         =   "Shut down"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      Picture         =   "frmShutdown.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblDescription 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   5655
   End
   Begin VB.Label lblChoose 
      Caption         =   "Choose the termination action you want to perform, and then click the OK button to perform that action."
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.582
      Y2              =   1687.582
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.935
      Y2              =   1697.935
   End
End
Attribute VB_Name = "frmShutdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim M
Dim sdoption
Dim lflag
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2


Private Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)


    If SetOnTop Then
     
       lflag = HWND_TOPMOST
    Else
        lflag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lflag, myfrm.Left / Screen.TwipsPerPixelX, myfrm.Top / Screen.TwipsPerPixelY, myfrm.Width / Screen.TwipsPerPixelX, myfrm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub
Sub DisplayText(dispnum)
Select Case dispnum
Case 1 ' Shut down
M = "Shuts down your computer, after prompting you to save your files and close any open applications. "
M = M & "With most newer computers, this option will turn off the power automatically."
Case 2 ' Restart
M = "Restarts your computer and your operating system. This option is useful if you have "
M = M & "installed a new program or legacy hardware device. "
Case 3 ' Log off
M = "Logs you out of your selected network, after closing all programs. The login dialog box will "
M = M & "appear after all programs have been exited."
Case 4 ' Exit Shell
M = "Exits Evergreen Shell, so you can use other programs. The Task Manager will remain open, allowing you "
M = M & "to start other programs and restart Evergreen Shell."
Case 5 ' MS-DOS Mode
M = "Exits to MS-DOS mode, so you can perform system commands and other DOS operations. "
M = M & "You will probably not need to use this option unless you are working with old hardware or software."
Case 6 'Power save
M = "Enables the Advanced Power Management support in your computer. If you do not have APM support, this option will not function."
End Select
lblDescription.Caption = M
sdoption = dispnum
End Sub


Private Sub cmdCancel_Click()
frmShell.Picture = frmShell.PictureY.Picture
frmShell.PictureY.Picture = frmShell.PictureX.Picture
Unload Me
End Sub

Private Sub cmdOK_Click()
'The plot thickens...
'OK, here's what we have. We need APM support to function and a shell to exit.
Select Case sdoption
Case 4 ' Exit Shell - Easiest
frmShell.Hide
frmPrograms.Hide
frmOptions.Hide
frmAbout.Hide
Unload frmShell
AlwaysOnTop frmShutdown, False
frmTaskList.Show
Me.Hide
End Select
End Sub

Private Sub Form_LostFocus()
frmShutdown.Show
End Sub

Private Sub optExit_Click()
Call DisplayText(4)
End Sub

Private Sub optLogoff_Click()
Call DisplayText(3)
End Sub

Private Sub optMSDOS_Click()
Call DisplayText(5)
End Sub

Private Sub optPowersave_Click()
Call DisplayText(6)
End Sub

Private Sub optRestart_Click()
Call DisplayText(2)
End Sub

Private Sub optShutdown_Click()
Call DisplayText(1)
End Sub
