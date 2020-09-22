VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Test - Button XP"
   ClientHeight    =   2505
   ClientLeft      =   2445
   ClientTop       =   1635
   ClientWidth     =   4485
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4485
   Begin Project1.ButtonXP ButtonXP2 
      Height          =   330
      Left            =   540
      TabIndex        =   1
      Top             =   15
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "Form1.frx":1272
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP1 
      Height          =   330
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "Form1.frx":13CC
      Style           =   2
      TextOffset      =   0
      ButtonSize      =   330
   End
   Begin Project1.ButtonXP ButtonXP3 
      Height          =   330
      Left            =   885
      TabIndex        =   2
      Top             =   15
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   582
      HighlightBorderColor=   -2147483640
      HighlightBackColor=   13876918
      PressColor      =   11899782
      ButtonIcon      =   "Form1.frx":1526
      TextOffset      =   0
      ButtonSize      =   330
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
      Left            =   2895
      TabIndex        =   3
      Top             =   2280
      Width           =   1530
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonXP1_Click()
    MsgBox "Button Clicked !"
End Sub

Private Sub ButtonXP1_MenuClick()
    MsgBox "MenuButton Clicked !"
End Sub
