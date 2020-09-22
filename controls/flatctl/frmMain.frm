VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Test - Flat Control"
   ClientHeight    =   1800
   ClientLeft      =   3810
   ClientTop       =   1785
   ClientWidth     =   5115
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
   ScaleHeight     =   1800
   ScaleWidth      =   5115
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1695
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3285
      TabIndex        =   0
      Top             =   1185
      Width           =   1395
   End
   Begin Project1.Flater Flater1 
      Left            =   30
      Top             =   1380
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
      Left            =   1695
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   255
      Width           =   3000
   End
   Begin Project1.Flater Flater2 
      Left            =   480
      Top             =   1380
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cFlat1 As FlatCtl
Dim cFlat2 As FlatCtl

Private Sub Form_Load()

    Set cFlat1 = New FlatCtl
    Set cFlat2 = New FlatCtl
    
    ' You can use control of class :))
    'cFlat1.Attach txtInPicture
    'cFlat2.Attach Combo1
    
    Flater1.Attach txtInPicture
    Flater2.Attach Combo1
    
    Me.Show
    
End Sub
