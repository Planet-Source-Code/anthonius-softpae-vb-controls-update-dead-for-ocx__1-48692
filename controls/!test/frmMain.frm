VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Test - Flat Control"
   ClientHeight    =   5715
   ClientLeft      =   3015
   ClientTop       =   1635
   ClientWidth     =   6450
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
   ScaleHeight     =   5715
   ScaleWidth      =   6450
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmMain.frx":1272
      Left            =   4140
      List            =   "frmMain.frx":12A3
      TabIndex        =   11
      Text            =   "8"
      ToolTipText     =   "Font Size"
      Top             =   2550
      Width           =   690
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2475
      TabIndex        =   10
      Text            =   "Tahoma"
      ToolTipText     =   "Font"
      Top             =   2550
      Width           =   1605
   End
   Begin Project1.RichEdit RichEdit1 
      Height          =   2325
      Left            =   165
      TabIndex        =   7
      Top             =   2895
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   4101
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   -2147483630
      ViewMode        =   1
   End
   Begin Project1.List List1 
      Height          =   2085
      Left            =   3030
      TabIndex        =   0
      Top             =   330
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   3678
   End
   Begin Project1.Tree Tree1 
      Height          =   2085
      Left            =   180
      TabIndex        =   2
      Top             =   330
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   3678
      Border          =   1
      Indent          =   16
   End
   Begin Project1.Flater Flater1 
      Left            =   3480
      Top             =   -45
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.TextBox txtInPicture 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   165
      TabIndex        =   1
      Text            =   "Sample text"
      Top             =   5280
      Width           =   6135
   End
   Begin Project1.Flater Flater2 
      Left            =   3015
      Top             =   -60
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin Project1.ButtonXP ButtonXP2 
      Height          =   330
      Left            =   705
      TabIndex        =   3
      ToolTipText     =   "Save File"
      Top             =   2535
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":12E1
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP1 
      Height          =   330
      Left            =   180
      TabIndex        =   4
      ToolTipText     =   "Load File"
      Top             =   2535
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":143B
      Style           =   2
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP3 
      Height          =   330
      Left            =   1305
      TabIndex        =   5
      ToolTipText     =   "Bold"
      Top             =   2535
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":1595
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP4 
      Height          =   330
      Left            =   1680
      TabIndex        =   8
      ToolTipText     =   "Italic"
      Top             =   2535
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":1B2F
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP5 
      Height          =   330
      Left            =   2055
      TabIndex        =   9
      ToolTipText     =   "Underline"
      Top             =   2535
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":20C9
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP6 
      Height          =   330
      Left            =   5175
      TabIndex        =   12
      ToolTipText     =   "Justify Left"
      Top             =   2535
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":2663
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP7 
      Height          =   330
      Left            =   5550
      TabIndex        =   13
      ToolTipText     =   "Justify Center"
      Top             =   2535
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":2BFD
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP8 
      Height          =   330
      Left            =   5925
      TabIndex        =   14
      ToolTipText     =   "Justify Right"
      Top             =   2535
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "frmMain.frx":3197
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SOFTPAE - VB Controls"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   15
      Top             =   75
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.softpae.com"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   4740
      TabIndex        =   6
      Top             =   75
      Width           =   1530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFlat1 As FlatCtl
Dim cFlat2 As FlatCtl

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then RichEdit1.SelFontName = Combo1.Text
End Sub

Private Sub Combo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then RichEdit1.SelFontSize = Combo2.Text
End Sub

Private Sub ButtonXP1_Click()
    RichEdit1.LoadFromFile App.Path & "\test.rtf", SF_RTF
End Sub

Private Sub ButtonXP1_MenuClick()
    MsgBox "Button XP - MenuButton Clicked !"
End Sub

Private Sub ButtonXP2_Click()
    RichEdit1.SaveToFile App.Path & "\test.rtf", SF_RTF
End Sub

Private Sub ButtonXP3_Click()
    RichEdit1.SelBold = Not RichEdit1.SelBold
End Sub

Private Sub ButtonXP4_Click()
    RichEdit1.SelItalic = Not RichEdit1.SelItalic
End Sub

Private Sub ButtonXP5_Click()
    RichEdit1.SelUnderline = Not RichEdit1.SelUnderline
End Sub

Private Sub ButtonXP6_Click()
    RichEdit1.SelAlignment = ercParaLeft
End Sub

Private Sub ButtonXP7_Click()
    RichEdit1.SelAlignment = ercParaCentre
End Sub

Private Sub ButtonXP8_Click()
    RichEdit1.SelAlignment = ercParaRight
End Sub

Private Sub Form_Load()

    Set cFlat1 = New FlatCtl
    Set cFlat2 = New FlatCtl
    
    ' You can use control of class :))
    cFlat1.Attach txtInPicture
    'cFlat2.Attach Combo1
    
    Flater1.Attach Combo2
    Flater2.Attach Combo3
    
    Tree1.hImageList = Tree1.LoadList(App.Path & "\resource.bmp", &H0, False, 752, 16)
    
    Call Tree1.AddItem("", 0, "root", "Root", 0)
    Call Tree1.AddItem("root", tvwFirst, "today", "Private", 1)
    Call Tree1.AddItem("today", tvwLast, "notes", "Notes", 2)
    Call Tree1.AddItem("today", tvwLast, "calendar", "Calendar", 3)
    Call Tree1.AddItem("today", tvwLast, "contacts", "Contacts", 4)
    Call Tree1.AddItem("today", tvwLast, "favorites", "Favorites", 5)
    
    List1.ListImage = "bmp_01"
    For k = 1 To Screen.FontCount
        List1.AddItem Screen.Fonts(k - 1)
    Next
    
    Combo2.Clear
    For k = 1 To Screen.FontCount
        Combo2.AddItem Screen.Fonts(k - 1)
    Next
    Combo2.Text = "Tahoma"
    
    RichEdit1.SelFontName = Combo2.Text
    RichEdit1.SelFontSize = Combo3.Text
    
    RichEdit1.AutoURLDetect = True
    
    RichEdit1.LoadFromFile App.Path & "\test.rtf", SF_RTF
    
    Me.Show
    
End Sub

Private Sub RichEdit1_SelChange(ByVal lMin As Long, ByVal lMax As Long, ByVal eSelType As ERECSelectionTypeConstants)
    Combo2.Text = RichEdit1.SelFontName
    Combo3.Text = RichEdit1.SelFontSize
    DoEvents
End Sub

