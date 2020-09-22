VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Test - Listbox"
   ClientHeight    =   3300
   ClientLeft      =   1575
   ClientTop       =   1545
   ClientWidth     =   6585
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
   MaxButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6585
   Begin Project1.List List1 
      Height          =   2535
      Left            =   885
      TabIndex        =   0
      Top             =   375
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   4471
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
      Left            =   5010
      TabIndex        =   1
      Top             =   3105
      Width           =   1530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    List1.ListImage = "bmp_01"
    For k = 1 To Screen.FontCount
        List1.AddItem Screen.Fonts(k - 1)
    Next

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    List1.Top = 0: List1.Left = 0: List1.Height = Me.ScaleHeight: List1.Width = Me.ScaleWidth: DoEvents
    
End Sub

Private Sub List1_Click()
    Me.Caption = List1.List(List1.ListIndex)
End Sub

