VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RS-232"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   21.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Anas Hyper Communicator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Anas Hyper Communicator.frx":0442
   ScaleHeight     =   7515
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6120
      OleObjectBlob   =   "Anas Hyper Communicator.frx":240484
      Top             =   6360
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   0
      OleObjectBlob   =   "Anas Hyper Communicator.frx":2406B8
      TabIndex        =   8
      Top             =   7080
      Width           =   8055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2445
      TabIndex        =   5
      Top             =   6480
      Width           =   1335
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1320
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   6360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   6480
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Anas Hyper Communicator.frx":24075A
      Left            =   4080
      List            =   "Anas Hyper Communicator.frx":24076A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5880
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Anas Hyper Communicator.frx":24077E
      Left            =   120
      List            =   "Anas Hyper Communicator.frx":2407A0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5880
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Anas Hyper Communicator.frx":2407E1
      Left            =   6000
      List            =   "Anas Hyper Communicator.frx":2407EE
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5880
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Anas Hyper Communicator.frx":240803
      Left            =   2040
      List            =   "Anas Hyper Communicator.frx":240813
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5880
      Width           =   1815
   End
   Begin VB.TextBox txtweight 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   720
      Width           =   7815
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      DragMode        =   1  'Automatic
      Height          =   495
      Left            =   405
      OleObjectBlob   =   "Anas Hyper Communicator.frx":24082F
      TabIndex        =   9
      Top             =   120
      Width           =   7215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   360
      OleObjectBlob   =   "Anas Hyper Communicator.frx":24088D
      TabIndex        =   10
      Top             =   5520
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   2280
      OleObjectBlob   =   "Anas Hyper Communicator.frx":2408F3
      TabIndex        =   11
      Top             =   5520
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   4320
      OleObjectBlob   =   "Anas Hyper Communicator.frx":240959
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      DragMode        =   1  'Automatic
      Height          =   375
      Left            =   6360
      OleObjectBlob   =   "Anas Hyper Communicator.frx":2409C1
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Private Sub Form_Initialize()
    Dim x As Long
    x = InitCommonControls
End Sub

Private Sub Combo_KeyPress(KeyAscii As Integer, combo As ComboBox)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{tab}"
        KeyAscii = 0
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Combo_KeyPress KeyAscii, Combo1
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    Combo_KeyPress KeyAscii, Combo2
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
    Combo_KeyPress KeyAscii, Combo3
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
    Combo_KeyPress KeyAscii, Combo4
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrorHandler
    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrorHandler
    Dim cmb As Control
    
    For Each cmb In Me.Controls
        If TypeOf cmb Is ComboBox Then
            cmb.Enabled = True
            cmb.ListIndex = 0
        End If
    Next
    
    Combo3.SetFocus
    Combo1.ListIndex = 3
    Combo4.ListIndex = 3
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

Private Sub Command1_Click()
    On Error GoTo ErrorHandler
    txtweight.Text = ""
    Call ConfigPort
    Timer1.Enabled = True
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

Public Sub ConfigPort()
    On Error GoTo ErrorHandler
    MSComm1.CommPort = Right(Combo3.Text, 1)
    MSComm1.Settings = Combo1.Text & "," & Mid$(Combo2.Text, 1, 1) & Trim$(Combo4.Text) & 1
    MSComm1.InputLen = 0
    MSComm1.PortOpen = True
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Form1.Skin1.LoadSkin App.Path & "\galaxy.skn"
    Form1.Skin1.ApplySkin Me.hWnd
    Me.MousePointer = vbHourglass
    Call AddPortstoCombo(Combo3, MSComm1)
    Me.MousePointer = vbDefault
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

Private Sub Timer1_Timer()
    On Error GoTo ErrorHandler
    txtweight.Text = Trim$(MSComm1.Input) & " " & txtweight.Text & vbCrLf
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub

