VERSION 5.00
Begin VB.UserControl Flater 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   27
   ToolboxBitmap   =   "ctlFlatCtl.ctx":0000
   Begin VB.Image Image1 
      Height          =   240
      Left            =   90
      Picture         =   "ctlFlatCtl.ctx":0312
      Top             =   90
      Width           =   240
   End
End
Attribute VB_Name = "Flater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private m_hWnd             As Long
Private m_hWndEdit         As Long
Private m_hWndParent       As Long
Private m_bSubclass        As Boolean
Private m_bMouseOver       As Boolean

Private Enum EDrawStyle
    FC_DRAWNORMAL = &H1
    FC_DRAWRAISED = &H2
    FC_DRAWPRESSED = &H4
End Enum

Private m_bLBtnDown As Boolean
Private m_bCombo As Boolean

Private Const WM_COMMAND = &H111
Private Const WM_PAINT = &HF
Private Const WM_TIMER = &H113
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const SM_CXHTHUMB = 10

Private Const WM_SETFOCUS = &H7
Private Const WM_KILLFOCUS = &H8
Private Const WM_MOUSEACTIVATE = &H21

Private Const GWL_STYLE = (-16)
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const CBS_DROPDOWN = &H2&
Private Const CBS_DROPDOWNLIST = &H3&
Private Const CBN_DROPDOWN = 7
Private Const CBN_CLOSEUP = 8
Private Const CB_GETDROPPEDSTATE = &H157
Private Const GW_CHILD = 5

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left     As Long
    Top      As Long
    Right    As Long
    Bottom   As Long
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long

Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)

Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2

Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'********************************************************************************

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

Private Sub subInitialize()

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

Public Sub Subclass(ByVal hwnd As Long, Owner As WinSubHook.iSubclass)

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
    
    Select Case uMsg
    
    Case WM_PAINT
    
        bDown = DroppedDown()
        bFocus = (m_hWnd = GetFocus() Or m_hWndEdit = GetFocus() Or bDown)
        OnPaint (bFocus), bDown

        If (bFocus) Then

            OnTimer False

        End If
        
        bHandled = False
    
    End Select

End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)

    On Error Resume Next
    

    Dim bDown As Boolean
    Dim bFocus As Boolean

    bHandled = False
    
    Select Case uMsg

        Case WM_COMMAND

            If (m_hWnd = lParam) Then

                Select Case wParam \ &H10000

                    Case CBN_CLOSEUP
                        
                        OnPaint (m_hWnd = GetFocus() Or m_hWndEdit = GetFocus() Or bDown), bDown

                End Select

                OnTimer False

            End If

'        Case WM_PAINT
'
'            bDown = DroppedDown()
'            bFocus = (m_hWnd = GetFocus() Or m_hWndEdit = GetFocus() Or bDown)
'            OnPaint (bFocus), bDown
'
'            If (bFocus) Then
'
'                OnTimer False
'
'            End If
'
'            bHandled = False

        Case WM_SETFOCUS
        
            OnPaint True, False
            OnTimer False

        Case WM_KILLFOCUS
        
            OnPaint False, False

        Case WM_MOUSEMOVE

            If Not (m_bMouseOver) Then

                bDown = DroppedDown()

                If Not (m_hWnd = GetFocus() Or m_hWndEdit = GetFocus() Or bDown) Then

                    OnPaint True, False
                    m_bMouseOver = True
                    
                    SetTimer m_hWnd, 1, 10, 0

                End If

            End If

        Case WM_TIMER

            OnTimer True

            If Not (m_bMouseOver) Then

                OnPaint False, False

            End If

    End Select
    'bHandled = True

End Sub

Private Sub UserControl_Initialize()

    On Error Resume Next

    Call subInitialize
    'Call Subclass(UserControl.hwnd, Me)
    
    UserControl.Width = 420
    UserControl.Height = 420

End Sub

Private Sub UserControl_Paint()

    Dim R As RECT
    
    SetRect R, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    
    'DrawEdge UserControl.hdc, R, EDGE_RAISED, BF_RECT
    DrawEdge UserControl.hdc, R, EDGE_BUMP, BF_RECT
    
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 420
    UserControl.Height = 420
End Sub

Public Sub Attach(ByRef objthis As Object)

    Dim lStyle As Long
    Dim lhWnd As Long
    
    pRelease
    
    On Error Resume Next
    lhWnd = objthis.hwnd
    
    If lhWnd <> 0 Then
        Call Subclass(lhWnd, Me)
    Else
        Call UnSubclass
    End If

    If (Err.Number <> 0) Then

        Err.Raise vbObjectError + 1048 + 1, _
           App.EXEName & ".cFlatControl", _
           "Incorrect control type passed to 'Attach' parameter - must be a control with a hWnd property."
        Exit Sub

    End If

    m_bCombo = False

    If TypeName(objthis) = "ImageCombo" Then

        m_hWndParent = lhWnd
        lhWnd = FindWindowEx(lhWnd, 0&, "ComboBox", ByVal 0&)
        m_bCombo = True

    ElseIf TypeName(objthis) = "ComboBox" Then

        m_hWndParent = GetParent(objthis.hwnd)
        m_bCombo = True

    ElseIf TypeName(objthis) = "OwnerDrawComboList" Then

        m_hWndParent = lhWnd
        m_bCombo = True

    Else

        lStyle = GetWindowLong(lhWnd, GWL_STYLE)

        If ((lStyle And CBS_DROPDOWN) = CBS_DROPDOWN) Or ((lStyle And CBS_DROPDOWNLIST) = CBS_DROPDOWNLIST) Then

            m_hWndParent = objthis.Parent.hwnd
            m_bCombo = True

        Else


            With objthis

                .Move .Left + 2 * Screen.TwipsPerPixelX, .Top + 2 * Screen.TwipsPerPixelY, .Width - 4 * Screen.TwipsPerPixelX, .Height - 4 * Screen.TwipsPerPixelY

            End With

        End If

    End If

    pAttach lhWnd

End Sub

Private Sub pAttach(ByRef hWndA As Long)

    Dim lStyle As Long
    m_hWnd = hWndA

    If (m_hWnd <> 0) Then

        lStyle = GetWindowLong(m_hWnd, GWL_STYLE)

        If (lStyle And CBS_DROPDOWN) = CBS_DROPDOWN Then

            m_hWndEdit = GetWindow(m_hWnd, GW_CHILD)

        End If

        'Call AddMsg(WM_PAINT, MSG_BEFORE)
        Call AddMsg(WM_PAINT, MSG_AFTER)
        Call AddMsg(WM_MOUSEACTIVATE, MSG_BEFORE)
        Call AddMsg(WM_SETFOCUS, MSG_BEFORE)
        Call AddMsg(WM_KILLFOCUS, MSG_BEFORE)
        Call AddMsg(WM_MOUSEMOVE, MSG_BEFORE)
        Call AddMsg(WM_TIMER, MSG_BEFORE)

        If (m_hWndEdit <> 0) Then

            Call AddMsg(WM_MOUSEACTIVATE, MSG_BEFORE)
            Call AddMsg(WM_SETFOCUS, MSG_BEFORE)
            Call AddMsg(WM_KILLFOCUS, MSG_BEFORE)
            Call AddMsg(WM_MOUSEMOVE, MSG_BEFORE)

        End If

        If (m_bCombo) Then

            Call AddMsg(WM_COMMAND, MSG_BEFORE)
            
        End If

        m_bSubclass = True

    End If

End Sub

Private Sub pRelease()

    If (m_bSubclass) Then

        Call DelMsg(WM_PAINT, MSG_BEFORE)
        Call DelMsg(WM_SETFOCUS, MSG_BEFORE)
        Call DelMsg(WM_KILLFOCUS, MSG_BEFORE)
        Call DelMsg(WM_MOUSEACTIVATE, MSG_BEFORE)
        Call DelMsg(WM_MOUSEMOVE, MSG_BEFORE)
        Call DelMsg(WM_TIMER, MSG_BEFORE)

        If (m_hWndEdit <> 0) Then

            Call DelMsg(WM_MOUSEACTIVATE, MSG_BEFORE)
            Call DelMsg(WM_SETFOCUS, MSG_BEFORE)
            Call DelMsg(WM_KILLFOCUS, MSG_BEFORE)
            Call DelMsg(WM_MOUSEMOVE, MSG_BEFORE)

        End If

        If (m_bCombo) Then

            Call DelMsg(WM_COMMAND, MSG_BEFORE)

        End If

    End If

    m_hWnd = 0: m_hWndEdit = 0: m_hWndParent = 0

End Sub

Private Sub Draw(ByVal dwStyle As EDrawStyle, clrTopLeft As OLE_COLOR, clrBottomRight As OLE_COLOR)

    If m_hWnd = 0 Then Exit Sub

    If (m_bCombo) Then

        DrawCombo dwStyle, clrTopLeft, clrBottomRight

    Else

        DrawEdit dwStyle, clrTopLeft, clrBottomRight

    End If

End Sub

Private Sub DrawEdit(ByVal dwStyle As EDrawStyle, clrTopLeft As OLE_COLOR, clrBottomRight As OLE_COLOR)

    Dim rcItem As RECT
    Dim rcItem2 As RECT
    Dim pDC As Long
    Dim hWndFocus As Long
    Dim tP As POINTAPI
    Dim hWndP As Long
    
    hWndP = GetParent(m_hWnd)
    GetWindowRect m_hWnd, rcItem
    tP.x = rcItem.Left: tP.y = rcItem.Top
    ScreenToClient hWndP, tP
    rcItem.Left = tP.x: rcItem.Top = tP.y
    tP.x = rcItem.Right: tP.y = rcItem.Bottom
    ScreenToClient hWndP, tP
    rcItem.Right = tP.x: rcItem.Bottom = tP.y
    InflateRect rcItem, 2, 2
    pDC = GetDC(hWndP)
    Draw3DRect pDC, rcItem, clrTopLeft, clrBottomRight
    InflateRect rcItem, -1, -1

    If (IsWindowEnabled(m_hWnd) = 0) Then

        Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight

    Else

        Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace

    End If

    If (IsWindowEnabled(m_hWnd) = 0) Then

        DeleteDC pDC
        Exit Sub

    End If

    Select Case dwStyle

        Case FC_DRAWNORMAL
            '      rcItem.Top = rcItem.Top - 1
            '      rcItem.Bottom = rcItem.Bottom + 1
            '      Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight
            '      rcItem.Left = rcItem.Left - 1
            '      rcItem.Right = rcItem.Right
            '      Draw3DRect pDC, rcItem, vbWindowBackground, vbButtonShadow

        Case FC_DRAWRAISED, FC_DRAWPRESSED
            InflateRect rcItem, -1, -1
            Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight
            InflateRect rcItem, -1, -1
            Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight
            InflateRect rcItem, -1, -1
            Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight
            'Case FC_DRAWPRESSED
            '   rcItem.Top = rcItem.Top - 1
            '   rcItem.Bottom = rcItem.Bottom
            '   Draw3DRect pDC, rcItem, vbButtonShadow, vb3DHighlight

    End Select

    DeleteDC pDC  'ReleaseDC(pDC);

End Sub

Private Function Draw3DRect(ByVal hdc As Long, ByRef rcItem As RECT, ByVal oTopLeftColor As OLE_COLOR, ByVal oBottomRightColor As OLE_COLOR)

    Dim hPen As Long
    Dim hPenOld As Long
    Dim tP As POINTAPI
    hPen = CreatePen(PS_SOLID, 1, TranslateColor(oTopLeftColor))
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, rcItem.Left, rcItem.Bottom - 1, tP
    LineTo hdc, rcItem.Left, rcItem.Top
    LineTo hdc, rcItem.Right - 1, rcItem.Top
    SelectObject hdc, hPenOld
    DeleteObject hPen

    If (rcItem.Left <> rcItem.Right) Then

        hPen = CreatePen(PS_SOLID, 1, TranslateColor(oBottomRightColor))
        hPenOld = SelectObject(hdc, hPen)
        LineTo hdc, rcItem.Right - 1, rcItem.Bottom - 1
        LineTo hdc, rcItem.Left, rcItem.Bottom - 1
        SelectObject hdc, hPenOld
        DeleteObject hPen

    End If

End Function

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long

    If OleTranslateColor(clr, hPal, TranslateColor) Then

        TranslateColor = -1

    End If

End Function

Private Sub DrawCombo(ByVal dwStyle As EDrawStyle, clrTopLeft As OLE_COLOR, clrBottomRight As OLE_COLOR)

    Dim rcItem As RECT
    Dim rcItem2 As RECT
    Dim pDC As Long
    Dim hWndFocus As Long
    Dim tP As POINTAPI
    GetClientRect m_hWnd, rcItem
    
    pDC = GetDC(m_hWnd)
    
    Draw3DRect pDC, rcItem, clrTopLeft, clrBottomRight
    
    InflateRect rcItem, -1, -1

    If (IsWindowEnabled(m_hWnd) = 0) Then

        Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight

    Else
        
        Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace

    End If

    InflateRect rcItem, -1, -1
    rcItem.Left = rcItem.Right - Offset()
    
    Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace
    
    InflateRect rcItem, -1, -1
    
    Draw3DRect pDC, rcItem, vbButtonFace, vbButtonFace

    If (IsWindowEnabled(m_hWnd) = 0) Then

        DeleteDC pDC
        Exit Sub

    End If

    Select Case dwStyle

        Case FC_DRAWNORMAL
            'rcItem.top -= 1;
            'rcItem.bottom += 1;
            'pDC->Draw3dRect(rcItem, ::GetSysColor(COLOR_BTNHIGHLIGHT),
            '   ::GetSysColor(COLOR_BTNHIGHLIGHT));
            'rcItem.left -= 1;
            'pDC->Draw3dRect(rcItem, ::GetSysColor(COLOR_BTNHIGHLIGHT),
            '   ::GetSysColor(COLOR_BTNHIGHLIGHT));
            'break;
            rcItem.Top = rcItem.Top - 1
            rcItem.Bottom = rcItem.Bottom + 1
            Draw3DRect pDC, rcItem, vb3DHighlight, vb3DHighlight
            rcItem.Left = rcItem.Left - 1
            rcItem.Right = rcItem.Left
            Draw3DRect pDC, rcItem, vbWindowBackground, &H0

        Case FC_DRAWRAISED
            'rcItem.top -= 1;
            'rcItem.bottom += 1;
            'pDC->Draw3dRect(rcItem, ::GetSysColor(COLOR_BTNHIGHLIGHT),
            '   ::GetSysColor(COLOR_BTNSHADOW));
            'break;
            rcItem.Top = rcItem.Top - 1
            rcItem.Bottom = rcItem.Bottom + 1
            rcItem.Right = rcItem.Right + 1
            Draw3DRect pDC, rcItem, vb3DHighlight, vbButtonShadow

        Case FC_DRAWPRESSED
            'rcItem.top -= 1;
            'rcItem.bottom += 1;
            'rcItem.OffsetRect(1,1);
            'pDC->Draw3dRect(rcItem, ::GetSysColor(COLOR_BTNSHADOW),
            '   ::GetSysColor(COLOR_BTNHIGHLIGHT));
            'break;
            rcItem.Left = rcItem.Left - 1
            rcItem.Top = rcItem.Top - 2
            OffsetRect rcItem, 1, 1
            Draw3DRect pDC, rcItem, vbButtonShadow, vb3DHighlight
            '}

    End Select

    DeleteDC pDC

End Sub

Private Function Offset() As Long

    Offset = GetSystemMetrics(SM_CXHTHUMB)

End Function

Public Property Get DroppedDown() As Boolean

    If (m_bCombo) And (m_hWnd <> 0) Then

        DroppedDown = (SendMessageLong(m_hWnd, CB_GETDROPPEDSTATE, 0, 0) <> 0)

    End If

End Property

Private Sub OnPaint(ByVal bFocus As Boolean, ByVal bDropped As Boolean)

    'used for paint

    If bFocus Then

        If (bDropped) Then

            Draw FC_DRAWPRESSED, vbButtonShadow, vb3DHighlight

        Else

            Draw FC_DRAWRAISED, vbButtonShadow, vb3DHighlight

        End If

    Else

        Draw FC_DRAWNORMAL, vbButtonFace, vbButtonFace

    End If

End Sub

Private Sub OnTimer(ByVal bCheckMouse As Boolean)

    Dim bOver As Boolean
    Dim rcItem As RECT
    Dim tP As POINTAPI

    If (bCheckMouse) Then

        bOver = True
        GetCursorPos tP
        GetWindowRect m_hWnd, rcItem

        If (PtInRect(rcItem, tP.x, tP.y) = 0) Then

            bOver = False

        End If

    End If

    If Not (bOver) Then

        KillTimer m_hWnd, 1
        m_bMouseOver = False

    End If

End Sub

Private Sub UserControl_Terminate()
    
    pRelease
    
    If hWndSubclass <> 0 Then
        Call UnSubclass
    End If
    
End Sub
