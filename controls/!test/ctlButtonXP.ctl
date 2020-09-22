VERSION 5.00
Begin VB.UserControl ButtonXP 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   90
   ToolboxBitmap   =   "ctlButtonXP.ctx":0000
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   465
      Picture         =   "ctlButtonXP.ctx":0312
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   1530
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   75
      Picture         =   "ctlButtonXP.ctx":045C
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   1605
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "ButtonXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SRCCOPY = &HCC0020

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Private Declare Function WindowFromPoint& Lib "user32" (ByVal lpPointX As Long, ByVal lpPointY As Long)

Public Enum eStyles
    Standart = 1
    Extended = 2
    UserSize = 3
End Enum

Dim TempColor As OLE_COLOR
Dim DefinedBorderColor As OLE_COLOR
Dim DefinedBackColor As OLE_COLOR
Dim DefinedPressColor As OLE_COLOR
Dim m_CheckedColor As OLE_COLOR
Dim DefinedImageDisplacement As Integer
Dim m_bEnabled As Boolean
Dim m_Index As Long
Dim m_ImgPos As Long
Dim m_Checked As Boolean
Dim m_Label As String
Dim m_TextPos As Long
Dim m_ButtonSize As Long
Dim m_BackColor As OLE_COLOR

Dim bSmall As Boolean

Dim m_Style As eStyles

Public Event Click()
Public Event MenuClick()
Public Event MouseDown(ByVal cx As Long, ByVal cy As Long)
Public Event MouseUp(ByVal cx As Long, ByVal cy As Long)

'**************************************************************************************

Private Const OFFSET_P1   As Long = 9
Private Const OFFSET_P2   As Long = 22
Private Const OFFSET_P3   As Long = 37
Private Const OFFSET_P4   As Long = 51
Private Const OFFSET_P5   As Long = 69
Private Const OFFSET_P6   As Long = 141
Private Const OFFSET_P7   As Long = 146
Private Const OFFSET_P8   As Long = 154
Private Const OFFSET_P9   As Long = 169
Private Const OFFSET_PA   As Long = 183
Private Const OFFSET_PB   As Long = 201
Private Const OFFSET_PC   As Long = 250
Private Const OFFSET_PD   As Long = 260
Private Const ARRAY_LB    As Long = 1

Private Type tCode
    Buf(ARRAY_LB To 272)    As Byte
End Type

Private Type tCodeBuf
    Code                    As tCode
End Type

Private CodeBuf           As tCodeBuf
Private nBreakGate        As Long
Private nMsgCntB          As Long
Private nMsgCntA          As Long
Private aMsgTblB()        As WinSubHook.eMsg
Private aMsgTblA()        As WinSubHook.eMsg
Private hWndSubclass      As Long
Private nWndProcSubclass  As Long
Private nWndProcOriginal  As Long

Implements WinSubHook.iSubclass

Public Property Get Caption() As String
    Caption = m_Label
End Property

Public Property Let Caption(ByVal NewValue As String)
    m_Label = NewValue
    UserControl.PropertyChanged "Caption"
    UserControl.Refresh
End Property

Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property

Public Property Let Checked(ByVal NewValue As Boolean)
    m_Checked = NewValue
    UserControl.PropertyChanged "Checked"
    UserControl.Refresh
End Property

Public Property Get CheckedColor() As OLE_COLOR
    CheckedColor = m_CheckedColor
End Property

Public Property Let CheckedColor(ByVal NewValue As OLE_COLOR)
    m_CheckedColor = NewValue
    UserControl.PropertyChanged "CheckedColor"
    UserControl.Refresh
End Property

Public Property Get TextOffset() As Long
    TextOffset = m_TextPos
End Property

Public Property Let TextOffset(ByVal NewValue As Long)
    m_TextPos = NewValue
    UserControl.PropertyChanged "TextOffset"
    UserControl.Refresh
End Property

Public Property Get ImagePosition() As Long
    ImagePosition = m_ImgPos
End Property

Public Property Let ImagePosition(ByVal NewValue As Long)
    m_ImgPos = NewValue
    UserControl.PropertyChanged "ImagePosition"
    UserControl.Refresh
End Property

Public Property Get Style() As eStyles
    Style = m_Style
End Property

Public Property Let Style(ByVal NewValue As eStyles)
    m_Style = NewValue
    UserControl.PropertyChanged "Style"
    UserControl_Resize
    UserControl.Refresh
End Property
'
'Public Property Get Index() As Long
'    Index = m_Index
'End Property
'
'Public Property Let Index(ByVal NewValue As Long)
'    m_Index = NewValue
'    UserControl.PropertyChanged "Index"
'End Property

Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    m_bEnabled = NewValue
    UserControl.Enabled = m_bEnabled
    UserControl.PropertyChanged "Enabled"
    UserControl.Refresh
End Property

Public Property Get ButtonIcon() As StdPicture
    Set ButtonIcon = Picture1.Picture
End Property

Public Property Set ButtonIcon(ByVal NewValue As StdPicture)
    Set Picture1.Picture = NewValue
    UserControl.PropertyChanged "ButtonIcon"
    UserControl.Refresh
End Property

Public Property Get ButtonSize() As Long
    ButtonSize = m_ButtonSize
End Property

Public Property Let ButtonSize(ByVal NewValue As Long)
    m_ButtonSize = NewValue
    Call UserControl_Resize
    UserControl.PropertyChanged "ButtonSize"
    UserControl.Refresh
End Property

Public Property Get HighlightBackColor() As OLE_COLOR
    HighlightBackColor = DefinedBackColor
End Property

Public Property Let HighlightBackColor(ByVal NewValue As OLE_COLOR)
    DefinedBackColor = NewValue
    UserControl.PropertyChanged "HighlightBackColor"
    UserControl.Refresh
End Property

Public Property Get PressColor() As OLE_COLOR
    PressColor = DefinedPressColor
End Property

Public Property Let PressColor(ByVal NewValue As OLE_COLOR)
    DefinedPressColor = NewValue
    UserControl.PropertyChanged "PressColor"
    UserControl.Refresh
End Property

Public Property Get HighlightBorderColor() As OLE_COLOR
    HighlightBorderColor = DefinedBorderColor
End Property

Public Property Let HighlightBorderColor(ByVal NewValue As OLE_COLOR)
    DefinedBorderColor = NewValue
    UserControl.PropertyChanged "HighlightBorderColor"
    UserControl.Refresh
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_BackColor = NewValue
    UserControl.BackColor = m_BackColor
    UserControl.PropertyChanged "BackColor"
    UserControl.Refresh
End Property

Public Property Get ImageDisplacement() As Integer
    ImageDisplacement = DefinedImageDisplacement
End Property

Public Property Let ImageDisplacement(ByVal NewValue As Integer)
    DefinedImageDisplacement = NewValue
    UserControl.PropertyChanged "ImageDisplacement"
    UserControl.Refresh
End Property

Private Sub Class_Initialize()

    Const OPS As String = "558BEC83C4F85756BE_patch1_33C08945FC8945F8B90000000083F900746183F9FF740CBF000000008B450CF2AF755033C03D_patch4_740B833E007542C70601000000BA_patch5_8B0283F8000F84A50000008D4514508D4510508D450C508D4508508D45FC508D45F8508B0252FF5020C706000000008B45F883F8007570FF7514FF7510FF750CFF750868_patch6_E8_patch7_8945FCB90000000083F900744D83F9FF740CBF000000008B450CF2AF753C33C03D_patchA_740B833E00752EC70601000000BA_patchB_8B0283F8007425FF7514FF7510FF750CFF75088D45FC508B0252FF501CC706000000005E5F8B45FCC9C2100068_patchC_6AFCFF7508E8_patchD_33C08945FCEBE190"
    
    Dim i As Long, j As Long, nIDE As Long

    With CodeBuf.Code

        j = 1

        For i = ARRAY_LB To UBound(.Buf)

            .Buf(i) = Val("&H" & Mid$(OPS, j, 2))
            j = j + 2

        Next i

        nWndProcSubclass = VarPtr(.Buf(ARRAY_LB))

    End With
  
    nIDE = InIDE

    Call PatchVal(OFFSET_P1, VarPtr(nBreakGate))
    Call PatchVal(OFFSET_P4, nIDE)
    Call PatchRel(OFFSET_P7, AddrFunc("CallWindowProcA"))
    Call PatchVal(OFFSET_PA, nIDE)
    Call PatchRel(OFFSET_PD, AddrFunc("SetWindowLongA"))

End Sub

Private Sub Class_Terminate()

    If hWndSubclass <> 0 Then
        Call UnSubclass
    End If

End Sub

Public Sub AddMsg(uMsg As WinSubHook.eMsg, When As WinSubHook.eMsgWhen)

    If When = WinSubHook.MSG_BEFORE Then
        Call AddMsgSub(uMsg, aMsgTblB, nMsgCntB, When)
    Else
        Call AddMsgSub(uMsg, aMsgTblA, nMsgCntA, When)
    End If

End Sub

Public Function CallOrigWndProc(ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long) As Long

    If hWndSubclass <> 0 Then
        CallOrigWndProc = WinSubHook.CallWindowProc(nWndProcOriginal, hWndSubclass, uMsg, wParam, lParam)
    Else
        Debug.Assert False
    End If

End Function

Public Sub DelMsg(uMsg As WinSubHook.eMsg, When As WinSubHook.eMsgWhen)

    If When = WinSubHook.MSG_BEFORE Then
        Call DelMsgSub(uMsg, aMsgTblB, nMsgCntB, When)
    Else
        Call DelMsgSub(uMsg, aMsgTblA, nMsgCntA, When)
    End If

End Sub

Public Sub Subclass(hwnd As Long, Owner As WinSubHook.iSubclass)

    Debug.Assert (hWndSubclass = 0)
    Debug.Assert IsWindow(hwnd)
  
    hWndSubclass = hwnd
    nWndProcOriginal = WinSubHook.SetWindowLong(hwnd, WinSubHook.GWL_WNDPROC, nWndProcSubclass)
    Debug.Assert nWndProcOriginal
  
    Call PatchVal(OFFSET_P5, ObjPtr(Owner))
    Call PatchVal(OFFSET_P6, nWndProcOriginal)
    Call PatchVal(OFFSET_PB, ObjPtr(Owner))
    Call PatchVal(OFFSET_PC, nWndProcOriginal)

End Sub

Public Sub UnSubclass()

    If hWndSubclass <> 0 Then

        Call PatchVal(OFFSET_P2, 0)
        Call PatchVal(OFFSET_P8, 0)
        Call WinSubHook.SetWindowLong(hWndSubclass, WinSubHook.GWL_WNDPROC, nWndProcOriginal)
        hWndSubclass = 0
        nMsgCntB = 0
        nMsgCntA = 0

    End If

End Sub

Private Sub AddMsgSub(uMsg As WinSubHook.eMsg, aMsgTbl() As WinSubHook.eMsg, nMsgCnt As Long, When As WinSubHook.eMsgWhen)

    Dim nEntry  As Long, nOff1   As Long, nOff2   As Long
  
    If uMsg = WinSubHook.ALL_MESSAGES Then

        nMsgCnt = -1

    Else

        For nEntry = ARRAY_LB To nMsgCnt

            Select Case aMsgTbl(nEntry)

                Case -1
                    aMsgTbl(nEntry) = uMsg
                    Exit Sub

                Case uMsg
                    Exit Sub

            End Select

        Next nEntry

        ReDim Preserve aMsgTbl(ARRAY_LB To nEntry)
        nMsgCnt = nEntry
        aMsgTbl(nEntry) = uMsg

    End If
  
    If When = WinSubHook.MSG_BEFORE Then
        nOff1 = OFFSET_P2
        nOff2 = OFFSET_P3
    Else
        nOff1 = OFFSET_P8
        nOff2 = OFFSET_P9
    End If

    Call PatchVal(nOff1, nMsgCnt)
    Call PatchVal(nOff2, AddrMsgTbl(aMsgTbl))

End Sub

Private Sub DelMsgSub(uMsg As WinSubHook.eMsg, aMsgTbl() As WinSubHook.eMsg, nMsgCnt As Long, When As WinSubHook.eMsgWhen)

    Dim nEntry As Long
  
    If uMsg = WinSubHook.ALL_MESSAGES Then

        nMsgCnt = 0

        If When = WinSubHook.MSG_BEFORE Then
            nEntry = OFFSET_P2
        Else
            nEntry = OFFSET_P8
        End If

        Call PatchVal(nEntry, 0)

    Else

        For nEntry = ARRAY_LB To nMsgCnt

            If aMsgTbl(nEntry) = uMsg Then
                aMsgTbl(nEntry) = -1
                Exit For
            End If

        Next nEntry

    End If

End Sub

Private Function AddrFunc(sProc As String) As Long
    AddrFunc = WinSubHook.GetProcAddress(WinSubHook.GetModuleHandle("user32"), sProc)
End Function

Private Function AddrMsgTbl(aMsgTbl() As WinSubHook.eMsg) As Long
    On Error Resume Next
    AddrMsgTbl = VarPtr(aMsgTbl(ARRAY_LB))
    On Error GoTo 0
End Function

Private Sub PatchVal(nOffset As Long, nValue As Long)
    Call WinSubHook.CopyMemory(ByVal (nWndProcSubclass + nOffset), nValue, 4)
End Sub

Private Sub PatchRel(nOffset As Long, nTargetAddr As Long)
    Call WinSubHook.CopyMemory(ByVal (nWndProcSubclass + nOffset), nTargetAddr - nWndProcSubclass - nOffset - 4, 4)
End Sub

Private Function InIDE() As Long

    Static Value As Long
  
    If Value = 0 Then

        Value = 1
        Debug.Assert True Or InIDE()
        InIDE = Value - 1

    End If
  
    Value = 0

End Function

Private Sub iSubclass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)

    'not used in this

End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)

    Dim cx As Long
    Dim cy As Long
    Dim lngHandle As Long
    Static bState As Boolean
    Static bClick As Boolean

    Select Case uMsg

        Case WM_SHOWWINDOW
            
            If wParam = 1 Then
                'If UserControl.Ambient.UserMode = True Then Call SetTimer(hwnd, 101, 100, ByVal 0&)
                If InIDE = False Then Call SetTimer(hwnd, 101, 100, ByVal 0&)
                'Call SetTimer(hwnd, 101, 100, ByVal 0&)
                bState = False
            ElseIf wParam = 0 Then
                Call KillTimer(hwnd, 101)
            End If
        
        Case WM_LBUTTONDOWN
        
            cx = (lParam And &HFFFF&)
            cy = (lParam \ &H10000)
            
            If cy > 20 Then
                bSmall = True
            Else
                bSmall = False
            End If
        
            Call DrawButton(2)

            RaiseEvent MouseDown(cx, cy)
            
            bClick = True

        Case WM_LBUTTONUP
            
            cx = (lParam And &HFFFF&)
            cy = (lParam \ &H10000)
            
            If cx > UserControl.ScaleWidth - 9 Then
                bSmall = True
            Else
                bSmall = False
            End If
            
            Call DrawButton(1)
            
            RaiseEvent MouseUp(cx, cy)
            
            If bClick = True And m_Style = 2 Then
                If bSmall = True Then
                    RaiseEvent MenuClick
                Else
                    RaiseEvent Click
                End If
            Else
                RaiseEvent Click
            End If
        
            bClick = False
            
        Case WM_TIMER
        
            lngHandle = HandleFromPos
            If lngHandle = hwnd And bState = False Then
                bState = True
                Call DrawButton(1)
            ElseIf lngHandle <> hwnd And bState = True Then
                bState = False
                Call DrawButton(0)
            End If
            
        Case WM_PAINT
            
            Call DrawButton(0)
    
    End Select
    bHandled = True

End Sub

Private Sub UserControl_Initialize()

    On Error Resume Next

    Call Class_Initialize
    Call Subclass(UserControl.hwnd, Me)
    
    'add messages to subclass
    Call AddMsg(WM_SHOWWINDOW, MSG_BEFORE)
    Call AddMsg(WM_TIMER, MSG_BEFORE)
    Call AddMsg(WM_LBUTTONDOWN, MSG_BEFORE)
    Call AddMsg(WM_LBUTTONUP, MSG_BEFORE)

    HighlightBackColor = &HD3BEB6
    HighlightBorderColor = &H80000008
    PressColor = &HB59386
    m_CheckedColor = &HE8D8D2
    DefinedImageDisplacement = 1
    m_bEnabled = True
    m_Style = 1
    m_Index = 0
    m_ImgPos = 4
    m_Checked = False
    m_Label = ""
    m_BackColor = &HD1D8DB
    m_ButtonSize = 285
    
    If m_Style <> 3 Then
    
        UserControl.Height = m_ButtonSize
        If m_Style = 1 Then
            UserControl.Width = m_ButtonSize
        ElseIf m_Style = 2 Then
            UserControl.Width = m_ButtonSize + 165
        End If
        
    End If
    
    Call DrawButton(0)

End Sub

Private Sub UserControl_Paint()
    Call DrawButton(0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    HighlightBorderColor = PropBag.ReadProperty("HighlightBorderColor", &H6B2408)
    HighlightBackColor = PropBag.ReadProperty("HighlightBackColor", &HD6BEB5)
    PressColor = PropBag.ReadProperty("PressColor", &HB59285)
    m_BackColor = PropBag.ReadProperty("BackColor", &HD1D8DB)
    UserControl.BackColor = m_BackColor
    DefinedImageDisplacement = PropBag.ReadProperty("DefinedImageDisplacement", 1)
    Set Picture1.Picture = PropBag.ReadProperty("ButtonIcon", Nothing)
    m_bEnabled = PropBag.ReadProperty("Enabled", True)
    m_ImgPos = PropBag.ReadProperty("ImagePosition", 4)
    m_Style = PropBag.ReadProperty("Style", 1)
    UserControl.Enabled = m_bEnabled
    m_Checked = PropBag.ReadProperty("Checked", False)
    m_CheckedColor = PropBag.ReadProperty("CheckedColor", &HE8D8D2)
    m_Label = PropBag.ReadProperty("Caption", "")
    m_TextPos = PropBag.ReadProperty("TextOffset", 5)
    m_ButtonSize = PropBag.ReadProperty("ButtonSize", 285)

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next

    If m_Style <> 3 Then
    
        UserControl.Height = m_ButtonSize
        If m_Style = 1 Then
            UserControl.Width = m_ButtonSize
        ElseIf m_Style = 2 Then
            UserControl.Width = m_ButtonSize + 165
        End If
        
    End If
    
End Sub

Private Sub UserControl_Show()
    Call DrawButton(0)
End Sub

Private Sub UserControl_Terminate()

    Call Class_Terminate

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    PropBag.WriteProperty "HighlightBorderColor", HighlightBorderColor, &H6B2408
    PropBag.WriteProperty "HighlightBackColor", HighlightBackColor, &HD6BEB5
    PropBag.WriteProperty "PressColor", PressColor, &HB59285
    PropBag.WriteProperty "DefinedImageDisplacement", DefinedImageDisplacement, 1
    PropBag.WriteProperty "ButtonIcon", Picture1.Picture
    PropBag.WriteProperty "Enabled", m_bEnabled, True
    PropBag.WriteProperty "Style", m_Style, 1
    PropBag.WriteProperty "BackColor", m_BackColor, &HD1D8DB
    PropBag.WriteProperty "ImagePosition", m_ImgPos, 4
    PropBag.WriteProperty "Checked", m_Checked, False
    PropBag.WriteProperty "CheckedColor", m_CheckedColor, &HE8D8D2
    PropBag.WriteProperty "Caption", m_Label, ""
    PropBag.WriteProperty "TextOffset", m_TextPos, 5
    PropBag.WriteProperty "ButtonSize", m_ButtonSize, 285

End Sub

Private Sub DisableHDC(SourceDC As Long, SourceWidth As Long, SourceHeight As Long)

    Const BLACK = 0
    Const DARKGREY = &H808080
    Const WHITE = &HFFFFFF
    
    Dim i As Long
    Dim j As Long
    Dim PixelColor As Long
    Dim BackgroundColor As Long
    Dim MemoryDC As Long
    Dim MemoryBitmap As Long
    Dim OldBitmap As Long
    Dim BooleanArray() As Boolean
    
    ReDim BooleanArray(SourceWidth, SourceHeight)
    
    MemoryDC = CreateCompatibleDC(SourceDC)
    MemoryBitmap = CreateCompatibleBitmap(SourceDC, SourceWidth, SourceHeight)
    OldBitmap = SelectObject(MemoryDC, MemoryBitmap)
    BitBlt MemoryDC, 0, 0, SourceWidth, SourceHeight, SourceDC, 0, 0, SRCCOPY
    BackgroundColor = GetBkColor(SourceDC)

    For i = 0 To SourceWidth
        For j = 0 To SourceHeight
            PixelColor = GetPixel(MemoryDC, i, j)
            If PixelColor <> BackgroundColor Then
                If PixelColor = BLACK Or Not PixelColor = WHITE Then
                    BooleanArray(i, j) = True
                    SetPixel MemoryDC, i, j, DARKGREY
                Else
                    SetPixel MemoryDC, i, j, BackgroundColor
                End If
            End If
        Next
    Next

    For i = 0 To SourceWidth - 1
        For j = 0 To SourceHeight - 1
            If BooleanArray(i, j) = True Then
                If BooleanArray(i + 1, j + 1) = False Then
                    SetPixel MemoryDC, i + 1, j + 1, WHITE
                End If
            End If
        Next
    Next

    BitBlt SourceDC, 0, 0, SourceWidth, SourceHeight, MemoryDC, 0, 0, SRCCOPY

    SelectObject MemoryDC, OldBitmap
    DeleteObject MemoryBitmap
    DeleteDC MemoryDC

End Sub

Private Function HandleFromPos() As Long
    
    Dim lngRet As Long
    Dim poiCursorPos As POINTAPI

    lngRet = GetCursorPos&(poiCursorPos)
    HandleFromPos = WindowFromPoint(poiCursorPos.x, poiCursorPos.y)
    
End Function

Private Sub DrawButton(ByVal State As Long)
    
    On Error Resume Next
    
    Dim btnRect As RECT
    Dim hPen As Long
    Dim hBrush As Long
    Dim xoff As Long
    Dim intLen As Long
    
    'LockWindowUpdate UserControl.hwnd
    
    intLen = UserControl.TextWidth(m_Label)
    
    If m_bEnabled = False Then
    
        UserControl.BackColor = m_BackColor
        Picture1.BackColor = UserControl.BackColor
        Picture2.BackColor = UserControl.BackColor
        Call DisableHDC(Picture1.hdc, Picture1.Width, Picture1.Height)
        Call BitBlt(UserControl.hdc, m_ImgPos, m_ImgPos, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, SRCCOPY)
        
        Call SetTextColor(UserControl.hdc, &H808080)
        
        intLen = UserControl.TextWidth(m_Label)
    
        Call TextOut(UserControl.hdc, 20 + m_TextPos, 3, m_Label, Len(m_Label))
       
        Exit Sub
        
    End If
    
    xoff = 0
    
    If m_Style = 2 Then xoff = 10
    
    If State = 0 Then
        If m_Checked = False Then
            hPen = CreatePen(0, 1, m_BackColor)
        Else
            hPen = CreatePen(0, 1, DefinedBorderColor)
        End If
    ElseIf State = 1 Then
        hPen = CreatePen(0, 1, DefinedBorderColor)
    ElseIf State = 2 Then
        hPen = CreatePen(0, 1, DefinedBorderColor)
    Else
        hPen = CreatePen(0, 1, m_BackColor)
    End If
    
    DeleteObject SelectObject(UserControl.hdc, hPen)
    
    If State = 0 Then
        If m_Checked = False Then
            UserControl.BackColor = m_BackColor
            Picture1.BackColor = UserControl.BackColor
            Picture2.BackColor = UserControl.BackColor
            hBrush = CreateSolidBrush(GetBkColor(UserControl.hdc))
        Else
            Picture1.BackColor = m_CheckedColor
            Picture2.BackColor = m_CheckedColor
            UserControl.BackColor = m_CheckedColor
            hBrush = CreateSolidBrush(GetBkColor(UserControl.hdc))
        End If
    ElseIf State = 1 Then
        UserControl.BackColor = m_BackColor
        Picture1.BackColor = DefinedBackColor
        Picture2.BackColor = DefinedBackColor
        hBrush = CreateSolidBrush(DefinedBackColor)
    ElseIf State = 2 Then
        UserControl.BackColor = m_BackColor
        Picture1.BackColor = DefinedPressColor
        Picture2.BackColor = DefinedPressColor
        hBrush = CreateSolidBrush(DefinedPressColor)
    Else
        UserControl.BackColor = m_BackColor
        Picture1.BackColor = UserControl.BackColor
        Picture2.BackColor = UserControl.BackColor
        hBrush = CreateSolidBrush(GetBkColor(UserControl.hdc))
    End If
    
    SelectObject UserControl.hdc, hBrush

    Call GetClientRect(UserControl.hwnd, btnRect)
    
    If m_Style = 2 Then
        Call Rectangle(UserControl.hdc, btnRect.Left, btnRect.Top, btnRect.Right - xoff + 1, btnRect.Bottom)
        Call Rectangle(UserControl.hdc, btnRect.Right - xoff, btnRect.Top, btnRect.Right, btnRect.Bottom)
    Else
        Call Rectangle(UserControl.hdc, btnRect.Left, btnRect.Top, btnRect.Right, btnRect.Bottom)
    End If
    
    Call SetTextColor(UserControl.hdc, &H0)
    
    If State = 0 Then
        Call BitBlt(UserControl.hdc, m_ImgPos, m_ImgPos, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, SRCCOPY)
        Call TextOut(UserControl.hdc, 20 + m_TextPos, 3, m_Label, Len(m_Label))
        'Call TextOut(UserControl.hdc, ((UserControl.ScaleWidth - 20) / 2) - (intLen / 2) + 20 + m_TextPos, 3, m_Label, Len(m_Label))
    ElseIf State = 1 Then
        Call BitBlt(UserControl.hdc, m_ImgPos - DefinedImageDisplacement, m_ImgPos - DefinedImageDisplacement, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, SRCCOPY)
        Call TextOut(UserControl.hdc, 20 + m_TextPos - 1, 2, m_Label, Len(m_Label))
    ElseIf State = 2 Then
        Call BitBlt(UserControl.hdc, m_ImgPos, m_ImgPos, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, SRCCOPY)
        Call TextOut(UserControl.hdc, 20 + m_TextPos, 3, m_Label, Len(m_Label))
    Else
        Call BitBlt(UserControl.hdc, m_ImgPos, m_ImgPos, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, SRCCOPY)
        Call TextOut(UserControl.hdc, 20 + m_TextPos, 3, m_Label, Len(m_Label))
    End If
    
    'If m_Style = 2 Then Call BitBlt(UserControl.hDC, 26, 1, 8, 16, Picture2.hDC, 5, 0, SRCCOPY)
    If m_Style = 2 Then Call BitBlt(UserControl.hdc, UserControl.ScaleWidth - 9, 1, 8, 16, Picture2.hdc, 5, (UserControl.ScaleHeight / 2) - 6 - ((UserControl.ScaleHeight - 16) / 2), SRCCOPY)

    DeleteObject hPen
    DeleteObject hBrush

    'LockWindowUpdate Null
    UpdateWindow UserControl.hwnd

End Sub
