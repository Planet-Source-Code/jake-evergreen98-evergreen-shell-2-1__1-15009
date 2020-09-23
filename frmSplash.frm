VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00008000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Loading Evergreen OS, please wait..."
   ClientHeight    =   570
   ClientLeft      =   3420
   ClientTop       =   5220
   ClientWidth     =   6495
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
   ScaleHeight     =   570
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2880
      Top             =   120
   End
   Begin VB.Label lblLoading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Evergreen OS G8..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
' this form serves no purpose - it's just a delay screen
' there might be some use for it in future.
Load frmOS
frmOS.Visible = True
frmOS.WindowState = 1
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
frmOS.WindowState = 2
Me.Hide
Unload Me
End Sub
