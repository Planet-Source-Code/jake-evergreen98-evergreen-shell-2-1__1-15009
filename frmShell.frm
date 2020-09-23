VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmShell 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evergreen Shell "
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   15270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShell.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10860
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList MenuImages 
      Left            =   1320
      Tag             =   "False"
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":310E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":5952
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":8196
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":A9DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":D21E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":FA62
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":122A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":14AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":1732E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":19B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":1C326
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":1EB6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":213AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":23BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":26436
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":28C7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":2B4BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":2DD02
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":304B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IconList 
      Left            =   1920
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":32CFA
            Key             =   "exe"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":3553E
            Key             =   "bat"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":36822
            Key             =   "url"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":39066
            Key             =   "rtftxt"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":3A34A
            Key             =   "xxx"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":3B62E
            Key             =   "ini"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":3C912
            Key             =   "unknown"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   10320
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   953
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1720
            MinWidth        =   1720
            Picture         =   "frmShell.frx":3DBF6
            Text            =   "Menu"
            TextSave        =   "Menu"
            Key             =   "menu"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
            Picture         =   "frmShell.frx":4043A
            Text            =   "Task Manager"
            TextSave        =   "Task Manager"
            Key             =   "taskman"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   11040
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":42C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":431DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":43736
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":43C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":441EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":4474A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":44CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":45202
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":4575E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":45CBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":4687E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":46DDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":47336
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":49B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":4A9CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":4D212
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":4E4F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":50CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":534EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":540C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":56906
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":57B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":5A3CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":5CC12
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":5DEF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":5E7D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShell.frx":61016
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   15210
      TabIndex        =   1
      Top             =   570
      Width           =   15270
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   14280
         TabIndex        =   6
         Top             =   120
         Width           =   855
      End
      Begin MSComctlLib.ImageCombo txtcmd 
         Height          =   330
         Left            =   1440
         TabIndex        =   4
         Top             =   120
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "IconList"
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "&Go!"
         Height          =   555
         Left            =   0
         Picture         =   "frmShell.frx":637CA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   1005
      ButtonWidth     =   1376
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Separator"
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Notepad"
            Key             =   "notepad"
            Description     =   "Notepad"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Controls"
            Key             =   "control"
            Description     =   "Controls"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Explorer"
            Key             =   "explorer"
            Description     =   "Explorer"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "DOS"
            Key             =   "command"
            Description     =   "DOS"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Separator"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Programs"
            Key             =   "programs"
            Description     =   "Programs"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tasks"
            Key             =   "tasks"
            Description     =   "Tasks"
            ImageIndex      =   27
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Key             =   "options"
            Description     =   "Options"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Separator"
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Internet"
            Key             =   "iexplore"
            Description     =   "Internet"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Outlook"
            Key             =   "msimn"
            Description     =   "Outlook"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Separator"
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sysedit"
            Key             =   "sysedit"
            Description     =   "Sysedit"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Regedit"
            Key             =   "regedit"
            Description     =   "Regedit"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MSConfig"
            Key             =   "msconfig"
            Description     =   "MSConfig"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Separator"
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "shutdown"
            Description     =   "Exit"
            ImageIndex      =   26
         EndProperty
      EndProperty
      Begin VB.PictureBox PictureY 
         BorderStyle     =   0  'None
         Height          =   15
         Left            =   11160
         ScaleHeight     =   15
         ScaleWidth      =   15
         TabIndex        =   8
         Top             =   120
         Width           =   15
      End
      Begin VB.PictureBox PictureX 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   10800
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   120
         Width           =   255
      End
      Begin VB.Line Line4 
         X1              =   3660
         X2              =   7230
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   8160
         Y1              =   720
         Y2              =   720
      End
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   600
      Left            =   -120
      TabIndex        =   5
      Top             =   7320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1058
      ButtonWidth     =   2514
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "MenuImages"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Programs"
            Key             =   "Programs"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Favorites"
            Key             =   "Favorites"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            Key             =   "Options"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "Help"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Shut Down"
            Key             =   "ShutDown"
            ImageIndex      =   18
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Line Line5 
      X1              =   15
      X2              =   15240
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   15255
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&Menu"
      Begin VB.Menu mnuAbout 
         Caption         =   "About Evergreen Shell..."
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitShell 
         Caption         =   "Exit Shell"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnuWindowsPrograms 
         Caption         =   "Programs Window"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuWindowsTaskman 
         Caption         =   "Task Manager"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuWindowsOptions 
         Caption         =   "Options Window"
         Shortcut        =   ^O
      End
   End
End
Attribute VB_Name = "frmShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Const GW_CHILD = 5
    Const GW_HWNDFIRST = 0

    Const GW_HWNDLAST = 1
    Const GW_HWNDNEXT = 2
    Const GW_HWNDPREV = 3
    Const GW_OWNER = 4
    
    
    Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)


    If SetOnTop Then
        lflag = HWND_TOPMOST
    Else
        lflag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lflag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Sub SetGlobalShow()
frmShell.MenuImages.Tag = "True"
End Sub
Sub SetGlobalHide()
frmShell.MenuImages.Tag = "False"
End Sub
Private Sub cmdGo_Click()
Toolbar2.Visible = False
wholename = frmShell.txtcmd.Text
If wholename = "" Then
MsgBox "You need to type in the name of a document or program to activate it.", vbInformation, "Enter Activation Name"
Exit Sub
End If
Call ShellApp(frmShell.txtcmd.Text, 0)
If errcode = 1 Then Exit Sub
Call DetermineIcon(Right$(txtcmd.Text, 3), Left$(txtcmd.Text, 5), wholename, errcode)
End Sub

Private Sub Form_Click()
Toolbar2.Visible = False
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
Me.WindowState = 2
End Sub

Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Me.WindowState = 2
End Sub

Private Sub Form_GotFocus()
If frmShell.MenuImages.Tag = "True" Then frmTaskList.Show: Call WhichWindows(frmTaskList.TaskList)
End Sub

Private Sub Form_Load()
On Error GoTo XYZHandl
Dim cmdtoshell
Dim TaskGoingUp
N = 0
Me.Caption = Me.Caption & App.Major & "." & App.Minor & " SR" & App.Revision
Toolbar2.Visible = False
Open App.path & "\backgrnd.pcf" For Input As #7
Input #7, Picfile
Close #7
Me.Picture = LoadPicture(Picfile)
Exit Sub
XYZHandl:
MsgBox "Warning! You do not have a picture file set in your preferences." & vbNewLine & "You must run the Options applet and choose a picture to create this file.", vbInformation, "Background Picture Not Selected"
Label1.Visible = True
Label1.Caption = "Warning! You do not have a picture file set in your preferences."
End Sub

Private Sub Form_Resize()
Me.WindowState = 2
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuExitShell_Click()
'What really would be nice here is a way to grey out everything, and add a shutdown dialog.
'Maybe a mask on the picture?
'Not happening anytime soon, sorry.
PictureY.Picture = Me.Picture
Me.BackColor = &H808080
Me.Picture = PictureX.Picture
frmShutdown.Show
AlwaysOnTop frmShutdown, True
If frmTaskList.Visible = True Then
AlwaysOnTop frmTaskList, False
End If
End Sub

Private Sub mnuWindowsOptions_Click()
frmOptions.Show
End Sub

Private Sub mnuWindowsPrograms_Click()
frmPrograms.Show
End Sub



Private Sub mnuWindowsTaskman_Click()
frmTaskList.Show
End Sub

Private Sub Picture1_Click()
Toolbar2.Visible = False
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
Call WhichWindows(frmTaskList.TaskList)
frmTaskList.Hide
Unload frmTaskList
If Panel.Key = "menu" Then
If Toolbar2.Visible = True Then Toolbar2.Visible = False: Exit Sub
If Toolbar2.Visible = False Then Toolbar2.Visible = True: Exit Sub
Exit Sub
End If
If Panel.Key = "taskman" Then frmTaskList.Show: Call WhichWindows(frmTaskList.TaskList): frmShell.MenuImages.Tag = "True": AlwaysOnTop frmTaskList, True
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'This subroutine has gone through a lot of garbage.
'Right now I'm using the ShellApp procedure found in the main BAS.
'NOTE: added Jan 28, 2001: will try and integrate ShellApp with ShellPath.
'It'll probably change though.
Select Case Button.Key
Case "menu"
PopupMenu Me.mnuMenu, , 0, 1300
Case "notepad"
ToShell = "notepad.exe"
Call ShellApp(ToShell, 0)
Case "control"
ToShell = "control.exe"
Call ShellApp(ToShell, 0)
Case "explorer"
ToShell = "explorer.exe"
Call ShellApp(ToShell, 0)
Case "command"
ToShell = "command.com"
Call ShellApp(ToShell, 0)
Case "programs"
frmPrograms.Show
Exit Sub
Case "tasks"
frmTaskList.Show
Exit Sub
Case "iexplore"
ToShell = "C:\program files\internet explorer\iexplore.exe"
Call ShellApp(ToShell, 0)
Case "sysedit"
ToShell = "sysedit.exe"
Call ShellApp(ToShell, 0)
Case "regedit"
ToShell = "regedit.exe"
Call ShellApp(ToShell, 0)
Case "msimn"
ToShell = "C:\program files\outlook express\msimn.exe"
Call ShellApp(ToShell, 0)
Case "options"
frmOptions.Show
Exit Sub
Case "msconfig"
ToShell = "msconfig.exe"
Call ShellApp(ToShell, 0)
Case "shutdown"
mnuExitShell_Click
Exit Sub
End Select
Call DetermineIcon(Right$(ToShell, 3), Left$(ToShell, 5), ToShell, 0)
End Sub


Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Programs"
frmPrograms.Show
Toolbar2.Visible = False
Case "Search"
Toolbar2.Visible = False
Case "MyFiles"
Toolbar2.Visible = False
Case "Options"
Toolbar2.Visible = False
frmOptions.Show
Case "Help"
Toolbar2.Visible = False
Case "Run"
Toolbar2.Visible = False
Case "ShutDown"
Toolbar2.Visible = False
End Select
End Sub

Private Sub txtcmd_Click()
Toolbar2.Visible = False
End Sub
