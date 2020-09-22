VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Test - Treeview"
   ClientHeight    =   3825
   ClientLeft      =   2865
   ClientTop       =   1665
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4305
   Begin Project1.Tree Tree1 
      Height          =   2445
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   4313
      Border          =   1
      Indent          =   15
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Info :"
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
      Left            =   60
      TabIndex        =   3
      Top             =   2715
      Width           =   900
   End
   Begin VB.Label Label2 
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
      Left            =   2700
      TabIndex        =   2
      Top             =   3600
      Width           =   1530
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   600
      Left            =   60
      TabIndex        =   1
      Top             =   2940
      Width           =   4200
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_DblClick()
    Tree1.Visible = Not Tree1.Visible
End Sub

Private Sub Form_Load()
    
    Tree1.hImageList = Tree1.LoadList(App.Path & "\resource.bmp", &H0, False, 752, 16)
    
    Call Tree1.AddItem("", 0, "root", "Zoznam", 0)
    Call Tree1.AddItem("root", tvwFirst, "today", "Private", 1)
    Call Tree1.AddItem("today", tvwLast, "notes", "Poznámky", 2)
    Call Tree1.AddItem("today", tvwLast, "calendar", "Kalendár", 3)
    Call Tree1.AddItem("today", tvwLast, "contacts", "Adresár", 4)
    Call Tree1.AddItem("today", tvwLast, "favorites", "Ob¾úbené", 5)

End Sub

Private Sub Tree1_ItemClick(ByVal Button As Long, ByVal ItemText As String, ByVal ItemKey As String, ByVal ItemIndex As Long)
    Label1.Caption = "Button = " & Button & ", Item = " & ItemText & ", Key = " & ItemKey
End Sub
