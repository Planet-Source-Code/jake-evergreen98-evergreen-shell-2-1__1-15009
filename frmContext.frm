VERSION 5.00
Begin VB.Form frmContext 
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   Caption         =   "Context Menu Popup"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdMore 
      Caption         =   "More..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblFutureUse 
      Caption         =   "This box is reserved for future application uses in Evergreen Shell. "
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   3120
      X2              =   3120
      Y1              =   0
      Y2              =   3600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   3120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -120
      X2              =   3120
      Y1              =   3600
      Y2              =   3600
   End
End
Attribute VB_Name = "frmContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Me.Hide
Unload Me
End Sub

Private Sub cmdMore_Click()
Dim CONE
CONE = Shell("notepad " & App.path & "\release.txt", vbNormalFocus)
End Sub

Private Sub Form_LostFocus()
Me.Hide
Unload Me
End Sub
