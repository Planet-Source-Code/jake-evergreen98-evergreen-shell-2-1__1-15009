VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Evergreen Shell Options"
   ClientHeight    =   3975
   ClientLeft      =   3975
   ClientTop       =   1980
   ClientWidth     =   6510
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
   ScaleHeight     =   3975
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      Caption         =   "Main Shell Options"
      Height          =   3135
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   6495
      Begin VB.CommandButton cmdChangeBackground 
         Caption         =   "Change Background Picture"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   6255
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6000
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   327681
         DialogTitle     =   "Select Background Picture..."
         Filter          =   "Bitmap Images (*.bmp)|*.bmp|JPEG Images (*.jpg; *.jpeg)|*.jpg;*.jpeg|GIF Images(*.gif)|*.gif"
      End
      Begin VB.Label Label1 
         Caption         =   "Background Picture: "
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   5775
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      MultiRow        =   -1  'True
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Main"
            Key             =   "keyMain"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   3480
      Width           =   1575
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDefault_Click()

End Sub

Private Sub cmdSelect_Click()

End Sub

Private Sub cmdCancel_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdChangeBackground_Click()
CommonDialog1.Action = 1

End Sub

Private Sub cmdOK_Click()
If CommonDialog1.filename <> "" Then
Set frmShell.Picture = LoadPicture(CommonDialog1.filename)
Open App.path & "\backgrnd.pcf" For Output As #6
Print #6, CommonDialog1.filename
Close #6
End If
frmShell.Label1.Caption = ""
frmShell.Label1.Visible = False
Me.Hide
End Sub

