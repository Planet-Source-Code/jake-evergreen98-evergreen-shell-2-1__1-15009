VERSION 5.00
Begin VB.Form frmPrograms 
   BackColor       =   &H00004000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Evergreen Shell  [ Program Selection ]"
   ClientHeight    =   3015
   ClientLeft      =   2385
   ClientTop       =   3375
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrograms.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      BackColor       =   &H00800000&
      Caption         =   "Program Selector"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   3015
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6135
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Height          =   540
         Left            =   4080
         Picture         =   "frmPrograms.frx":000C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Top             =   1320
         Width           =   540
      End
      Begin VB.ListBox lstPrograms 
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   2400
         ItemData        =   "frmPrograms.frx":0E4E
         Left            =   120
         List            =   "frmPrograms.frx":0E9A
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblIcon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Open Selected Program"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   3360
         TabIndex        =   8
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label lblShell 
         BackStyle       =   0  'Transparent
         Caption         =   "2. then click on the icon."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblChoice 
         BackStyle       =   0  'Transparent
         Caption         =   "1. Choose a program from the list..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
   End
End
Attribute VB_Name = "frmPrograms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub fraMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lblIcon.ForeColor = 0
'lblExtensions.BackColor = &H8000000F
End Sub

Private Sub fraSample2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub lblExtensions_Click()
'lblExtensions.BackColor = &HFF&
End Sub

Private Sub lblIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lblIcon.ForeColor = &HFF0000
End Sub

Private Sub lstPrograms_DblClick()
MsgBox lstPrograms.ListIndex

End Sub

Private Sub Picture1_Click()
Call OpenSelProgram(lstPrograms.ListIndex, False)
Me.Hide
Unload Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lblIcon.ForeColor = &HFF0000
End Sub
