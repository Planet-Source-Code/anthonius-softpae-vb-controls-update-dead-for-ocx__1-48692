VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Test"
   ClientHeight    =   4890
   ClientLeft      =   2685
   ClientTop       =   1500
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7800
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmMain.frx":1272
      Left            =   5565
      List            =   "frmMain.frx":12A3
      TabIndex        =   11
      Text            =   "8"
      Top             =   120
      Width           =   705
   End
   Begin VB.CommandButton Command8 
      Caption         =   "  --"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   5.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7290
      TabIndex        =   10
      ToolTipText     =   "Right Justify"
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton Command7 
      Caption         =   " -- "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   5.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6960
      TabIndex        =   9
      ToolTipText     =   "Center Justify"
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton Command6 
      Caption         =   "--   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   5.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6630
      TabIndex        =   8
      ToolTipText     =   "Left Justify"
      Top             =   120
      Width           =   300
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3660
      TabIndex        =   7
      Text            =   "Tahoma"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3255
      TabIndex        =   6
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton Command4 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2910
      TabIndex        =   5
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton Command3 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2565
      TabIndex        =   4
      Top             =   120
      Width           =   300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   315
      Left            =   1155
      TabIndex        =   3
      Top             =   105
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load"
      Height          =   315
      Left            =   45
      TabIndex        =   2
      Top             =   105
      Width           =   1065
   End
   Begin Project1.RichEdit RichEdit1 
      Height          =   4065
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   7170
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ViewMode        =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Line :"
      Height          =   225
      Left            =   30
      TabIndex        =   1
      Top             =   4650
      Width           =   5715
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then RichEdit1.SelFontName = Combo1.Text
End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then RichEdit1.SelFontSize = Combo2.Text
End Sub

Private Sub Command1_Click()
    RichEdit1.LoadFromFile App.Path & "\test.rtf", SF_RTF
End Sub

Private Sub Command2_Click()
    RichEdit1.SaveToFile App.Path & "\test.rtf", SF_RTF
End Sub

Private Sub Command3_Click()
    RichEdit1.SelBold = Not RichEdit1.SelBold
End Sub

Private Sub Command4_Click()
    RichEdit1.SelItalic = Not RichEdit1.SelItalic
End Sub

Private Sub Command5_Click()
    RichEdit1.SelUnderline = Not RichEdit1.SelUnderline
End Sub

Private Sub Command6_Click()
    RichEdit1.SelAlignment = ercParaLeft
End Sub

Private Sub Command7_Click()
    RichEdit1.SelAlignment = ercParaCentre
End Sub

Private Sub Command8_Click()
    RichEdit1.SelAlignment = ercParaRight
End Sub

Private Sub Form_Load()
    
    Combo1.Clear
    For k = 1 To Screen.FontCount
        Combo1.AddItem Screen.Fonts(k - 1)
    Next
    Combo1.Text = "Tahoma"
    
    RichEdit1.SelFontName = Combo1.Text
    RichEdit1.SelFontSize = Combo2.Text
    
    RichEdit1.AutoURLDetect = True
    
End Sub

Private Sub RichEdit1_LinkOver(ByVal iType As Integer, ByVal lMin As Long, ByVal lMax As Long)
'    MsgBox "link : " & lMin & " : " & lMax
'    RichEdit1.SelStart = lMin
'    RichEdit1.SelLength = (lMax - lMin)
'    MsgBox RichEdit1.SelText
End Sub

Private Sub RichEdit1_SelChange(ByVal lMin As Long, ByVal lMax As Long, ByVal eSelType As ERECSelectionTypeConstants)
    Label1.Caption = "Current Line : " & RichEdit1.CurrentLine
    Combo1.Text = RichEdit1.SelFontName
    Combo2.Text = RichEdit1.SelFontSize
    DoEvents
End Sub
