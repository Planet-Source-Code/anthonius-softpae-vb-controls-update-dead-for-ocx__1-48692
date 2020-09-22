VERSION 5.00
Begin VB.UserControl RichEdit 
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5040
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ToolboxBitmap   =   "ctlRichEdit.ctx":0000
End
Attribute VB_Name = "RichEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************
' *
' * Programmer Name  : Anton Piták
' * Web Site         : www.softpae.sk
' * E-Mail           : anthonius@softpae.sk
' * Date             : 21.09.2003
' * Time             : 13:15
' * Module Name      : RichEdit
' * Module Filename  : ctlRichEdit.ctl
' *
' **********************************************************************

Private Type CHARFORMAT
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To LF_FACESIZE - 1) As Byte
    wPad2 As Integer
End Type

Private Type CHARFORMAT2
    cbSize As Integer
    wPad1 As Integer
    dwMask As Long
    dwEffects As Long
    yHeight As Long
    yOffset As Long
    crTextColor As Long
    bCharSet As Byte
    bPitchAndFamily As Byte
    szFaceName(0 To LF_FACESIZE - 1) As Byte
    wPad2 As Integer ' 60
    wWeight As Integer
    sSpacing As Integer
    crBackColor As Long
    lLCID As Long
    dwReserved As Long
    sStyle As Integer
    wKerning As Integer
    bUnderlineType As Byte
    bAnimation As Byte
    bRevAuthor As Byte
    bReserved1 As Byte
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type EDITSTREAM
    dwCookie As Long
    dwError As Long
    pfnCallback As Long
End Type

Private Type NMHDR_RICHEDIT
    hwndFrom As Long
    wPad1 As Integer
    idfrom As Integer
    code As Integer
    wPad2 As Integer
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As Long
End Type

Private Type FINDTEXTEX_A
    chrg As CHARRANGE
    lpstrText As Long
    chrgText As CHARRANGE
End Type

Private Type ENLINK
    NMHDR As NMHDR_RICHEDIT
    msg As Integer
    wPad1 As Integer
    wParam As Integer
    wPad2 As Integer
    lParam As Integer
    chrg As CHARRANGE
End Type

Private Type MSGFILTER
    NMHDR As NMHDR_RICHEDIT
    msg As Integer
    wPad1 As Integer
    wParam As Integer
    wPad2 As Integer
    lParam As Long
End Type

Private Type SelChange
    NMHDR As NMHDR_RICHEDIT
    chrg As CHARRANGE
    seltyp As Long
End Type

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * LF_FACESIZE
End Type

Public Enum ERECControlVersion
    eRICHED32
    eRICHED20
End Enum

Public Enum RE_EventNotificationMask
    ENM_NONE = &H0
    ENM_CHANGE = &H1
    ENM_UPDATE = &H2
    ENM_SCROLL = &H4
    ENM_KEYEVENTS = &H10000
    ENM_MOUSEEVENTS = &H20000
    ENM_REQUESTRESIZE = &H40000
    ENM_SELCHANGE = &H80000
    ENM_DROPFILES = &H100000
    ENM_PROTECTED = &H200000
    ENM_CORRECTTEXT = &H400000
    ENM_SCROLLEVENTS = &H8
    ENM_DRAGDROPDONE = &H10
    ENM_IMECHANGE = &H800000
    ENM_LANGCHANGE = &H1000000
    ENM_OBJECTPOSITIONS = &H2000000
    ENM_LINK = &H4000000
End Enum

Public Enum ERECViewModes
   ercDefault = 0
   ercWordWrap = 1
   ercWYSIWYG = 2
End Enum

Public Enum ERECFileTypes
    SF_TEXT = &H1
    SF_RTF = &H2
End Enum

Public Enum ERECSelectionTypeConstants
   SEL_EMPTY = &H0
   SEL_TEXT = &H1
   SEL_OBJECT = &H2
   SEL_MULTICHAR = &H4
   SEL_MULTIOBJECT = &H8
End Enum

Public Enum ERECSetFormatRange
   ercSetFormatAll = SCF_ALL
   ercSetFormatSelection = SCF_SELECTION
   ercSetFormatWord = SCF_WORD Or SCF_SELECTION
End Enum

Public Enum ERECFindTypeOptions
   FR_DEFAULT = &H0
   FR_DOWN = &H1
   FR_WHOLEWORD = &H2
   FR_MATCHCASE = &H4&
End Enum

Public Enum ERECParagraphAlignmentConstants
   ercParaLeft = PFA_LEFT
   ercParaCentre = PFA_CENTER
   ercParaRight = PFA_RIGHT
   ercParaJustify = PFA_JUSTIFY
End Enum

Public Enum ERECTextTypes
   ercTextNormal
   ercTextSuperscript
   ercTextSubscript
End Enum

Public Enum eBorder
    [None] = 0
    [Fixed Single] = 1
End Enum

Private m_hWnd As Long
Private m_hFont As Long
Private m_lplf As LOGFONT
Private m_fOwned As Boolean

Private cSelBold As Boolean
Private cSelAlignment As ERECParagraphAlignmentConstants
Private cSelColor As OLE_COLOR
Private cSelFontName As String
Private cSelFontSize As Byte
Private cSelItalic As Boolean
Private cSelUnderline As Boolean

Private m_Border As eBorder
Private m_eCharFormatRange As ERECSetFormatRange
Private m_eVersion As ERECControlVersion
Private m_eLastFindMode As ERECFindTypeOptions
Private m_sLastFindText As String
Private m_bLastFindNext As Boolean

Private m_MaxLimit As Long
Private m_eViewMode As ERECViewModes

Private Const RICHEDIT_CLASS = "RichEdit20A"
Private Const RICHEDIT_CLASS10 = "RICHEDIT" ' Richedit 1.0

Public Event SelChange(ByVal lMin As Long, ByVal lMax As Long, ByVal eSelType As ERECSelectionTypeConstants)
Public Event LinkOver(ByVal iType As Integer, ByVal lMin As Long, ByVal lMax As Long)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event DblClick(X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Changed()
Public Event VScroll()
Public Event HScroll()

Private Declare Function win32_SetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function win32_GetFocus Lib "user32" Alias "GetFocus" () As Long
Private Declare Function SendMessageBuf Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

'end your declaration here !!
'***********************************************************************************

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
    code                    As tCode
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

    With CodeBuf.code

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

Public Sub Subclass(hWnd As Long, Owner As WinSubHook.iSubclass)

    Debug.Assert (hWndSubclass = 0)
    Debug.Assert IsWindow(hWnd)
  
    hWndSubclass = hWnd
    nWndProcOriginal = WinSubHook.SetWindowLong(hWnd, WinSubHook.GWL_WNDPROC, nWndProcSubclass)
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

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)

    'funkcia volaná po hlavnej (systémovej) wnd funkcii
    'not used in this

End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)

    On Error Resume Next

    Dim tNMH As NMHDR_RICHEDIT
    Dim tSC As SelChange
    Dim tEN As ENLINK
    Dim tMF As MSGFILTER
    Dim tP As POINTL
    Dim X As Single, Y As Single
    Dim iKeyCode As Integer, iKeyAscii As Integer, iShift As Integer
    Dim bDefault As Boolean
    Dim bDoIt As Boolean
    Dim ID As Long
    Dim cButton As Integer

   If (uMsg = WM_NOTIFY) Then
   
      CopyMemory tNMH, ByVal lParam, Len(tNMH)
      
      If (tNMH.hwndFrom = m_hWnd) Then
      
         'If tNMH.wPad1 = ENM_CHANGE Then MsgBox ("a")
         
         Select Case tNMH.code

         Case EN_SELCHANGE
            CopyMemory tSC, ByVal lParam, Len(tSC)
            RaiseEvent SelChange(tSC.chrg.cpMin, tSC.chrg.cpMax, tSC.seltyp)
            If Me.SelStart <> tSC.chrg.cpMax - tSC.chrg.cpMin Then
                RaiseEvent Changed
            End If
        Case EN_LINK
            CopyMemory tEN, ByVal lParam, Len(tEN)
            RaiseEvent LinkOver(tEN.msg, tEN.chrg.cpMin, tEN.chrg.cpMax)
         Case EN_MSGFILTER
         
            bDefault = True
            CopyMemory tMF, ByVal lParam, Len(tMF)

            Select Case tMF.msg

            Case 515, 518
                GetCursorPos tP
                ScreenToClient m_hWnd, tP
                X = tP.X * Screen.TwipsPerPixelX
                Y = tP.Y * Screen.TwipsPerPixelY
                RaiseEvent DblClick(X, Y)
            Case 33, 516
                iShift = GetShiftState()
                If tMF.msg = 33 Then
                    cButton = 1
                ElseIf tMF.msg = 516 Then
                    cButton = 2
                Else
                    cButton = 0
                End If
                GetCursorPos tP
                ScreenToClient m_hWnd, tP
                X = tP.X * Screen.TwipsPerPixelX
                Y = tP.Y * Screen.TwipsPerPixelY
                RaiseEvent MouseDown(cButton, iShift, X, Y)
            Case 514, 517
                iShift = GetShiftState()
                If tMF.msg = 514 Then
                    cButton = 1
                ElseIf tMF.msg = 517 Then
                    cButton = 2
                Else
                    cButton = 0
                End If
                GetCursorPos tP
                ScreenToClient m_hWnd, tP
                X = tP.X * Screen.TwipsPerPixelX
                Y = tP.Y * Screen.TwipsPerPixelY
                RaiseEvent MouseUp(cButton, iShift, X, Y)
            Case 512
                iShift = GetShiftState()
                If KeyIsPressed(vbKeyLButton) = True Then
                    cButton = 1
                ElseIf KeyIsPressed(vbKeyRButton) = True Then
                    cButton = 2
                Else
                    cButton = 0
                End If
                GetCursorPos tP
                ScreenToClient m_hWnd, tP
                X = tP.X * Screen.TwipsPerPixelX
                Y = tP.Y * Screen.TwipsPerPixelY
                RaiseEvent MouseMove(cButton, iShift, X, Y)
            Case 256
                iShift = GetShiftState()
                iKeyCode = tMF.wParam
                RaiseEvent KeyDown(iKeyCode, iShift)
            Case 258
                iShift = GetShiftState()
                iKeyAscii = tMF.wParam
                RaiseEvent KeyPress(iKeyAscii)
            Case 257
                iShift = GetShiftState()
                iKeyCode = tMF.wParam
                RaiseEvent KeyUp(iKeyCode, iShift)
            Case Else
               ' Debug.Print "Something Different:", tMF.msg, tMF.wParam, tMF.lParam, tMF.wPad1, tMF.wPad2
            End Select
         End Select

      End If

    ElseIf (uMsg = WM_VSCROLL) Then
        RaiseEvent VScroll

    ElseIf (uMsg = WM_HSCROLL) Then
        RaiseEvent HScroll
    
    ElseIf (uMsg = WM_SETFOCUS) Then
        UserControl.SetFocus
        If (wParam <> 0) Then
           SendMessageLong wParam, WM_KILLFOCUS, m_hWnd, 0
        End If
        win32_SetFocus m_hWnd
   
    End If
    
    bHandled = True

End Sub

Private Sub UserControl_EnterFocus()

    Call win32_SetFocus(m_hWnd)

End Sub

Private Sub UserControl_Initialize()

    On Error Resume Next
    
    Dim dwMask As Long
    
    With Me
        If .SelFontSize <= 0 Then .SelFontSize = .Font.Size
        If .SelFontName = "" Then .SelFontName = .Font.Name
    End With
    
    cSelColor = vbBlack
    cSelBold = False
    cSelFontName = "MS Sans Serif"
    cSelFontSize = 10
    cSelItalic = False
    cSelUnderline = False
    
    Call Create

    Call subInitialize
    Call Subclass(UserControl.hWnd, Me)
    
    'add messages to subclass
    Call AddMsg(WM_NOTIFY, MSG_BEFORE)
    Call AddMsg(WM_SETFOCUS, MSG_BEFORE)
    Call AddMsg(WM_VSCROLL, MSG_BEFORE)
    Call AddMsg(WM_HSCROLL, MSG_BEFORE)
    
    dwMask = ENM_KEYEVENTS Or ENM_MOUSEEVENTS
    dwMask = dwMask Or ENM_SELCHANGE
    dwMask = dwMask Or ENM_DROPFILES
    dwMask = dwMask Or ENM_SCROLL
    dwMask = dwMask Or ENM_CHANGE
    dwMask = dwMask Or ENM_UPDATE
    dwMask = dwMask Or ENM_LINK
    dwMask = dwMask Or ENM_PROTECTED
    
    SendMessageLong m_hWnd, EM_SETEVENTMASK, 0, dwMask
    
    m_MaxLimit = 0
    m_eCharFormatRange = ercSetFormatSelection
    
    Me.ViewMode = ercWordWrap
    Me.Text = "RichEdit"

End Sub

Private Sub UserControl_InitProperties()
    m_eCharFormatRange = ercSetFormatAll
    Set Font = UserControl.Ambient.Font
    m_eCharFormatRange = ercSetFormatSelection
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    On Error Resume Next

    If (UserControl.Ambient.UserMode) Then
        m_eVersion = PropBag.ReadProperty("Version", eRICHED32)
    Else
        UseVersion = PropBag.ReadProperty("Version", eRICHED32)
    End If
    
    Border = PropBag.ReadProperty("Border", 1)
    m_eCharFormatRange = ercSetFormatSelection
    
    Dim sFnt As New StdFont
    
    Set Font = PropBag.ReadProperty("Font")
    m_eCharFormatRange = ercSetFormatSelection
    
    BackColor = PropBag.ReadProperty("BackColor", vbWindowBackground)
    ForeColor = PropBag.ReadProperty("ForeColor", vbWindowText)
    Text = PropBag.ReadProperty("Text", "")
    ViewMode = PropBag.ReadProperty("ViewMode", ercWordWrap)
   
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    On Error Resume Next

    PropBag.WriteProperty "Version", UseVersion, eRICHED32
    m_eCharFormatRange = ercSetFormatAll
    PropBag.WriteProperty "Font", Font
    PropBag.WriteProperty "BackColor", BackColor, vbWindowBackground
    PropBag.WriteProperty "ForeColor", ForeColor, vbWindowText
    PropBag.WriteProperty "Text", m_sText, ""
    PropBag.WriteProperty "ViewMode", ViewMode
    PropBag.WriteProperty "Border", Border, 1

End Sub

Private Sub UserControl_Resize()

    Call MoveWindow(m_hWnd, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1)

End Sub

Private Sub UserControl_Terminate()

    On Error Resume Next

    If hWndSubclass <> 0 Then
        Call UnSubclass
    End If
    
    Call Destroy
    
End Sub

'***********************************************************************************
'add your properties and metthots here !!

Private Function Create() As Long

    Dim hRelib As Long
    
    If (m_hWnd = 0) Then
    
        m_eVersion = eRICHED32
        
        hRelib = LoadLibrary("RICHED32.DLL")
        If hRelib = 0 Then
            hRelib = LoadLibrary("RICHED20.DLL")
            m_eVersion = eRICHED20
        End If
    
        'm_hWnd = CreateWindowEx(WS_EX_CLIENTEDGE, RICHEDIT_CLASS, vbNullString, WS_CHILD Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_TABSTOP Or WS_HSCROLL Or WS_VSCROLL Or ES_AUTOHSCROLL Or ES_AUTOVSCROLL Or ES_LEFT Or ES_MULTILINE Or ES_WANTRETURN, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
        'm_hWnd = CreateWindowEx(WS_EX_CLIENTEDGE, RICHEDIT_CLASS, vbNullString, WS_CHILD Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_TABSTOP Or WS_VSCROLL Or ES_AUTOHSCROLL Or ES_AUTOVSCROLL Or ES_LEFT Or ES_MULTILINE Or ES_WANTRETURN, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
        m_hWnd = CreateWindowEx(0&, RICHEDIT_CLASS, vbNullString, WS_CHILD Or WS_CLIPCHILDREN Or WS_CLIPSIBLINGS Or WS_TABSTOP Or WS_VSCROLL Or ES_AUTOHSCROLL Or ES_AUTOVSCROLL Or ES_LEFT Or ES_MULTILINE Or ES_WANTRETURN Or ES_NOHIDESEL Or ES_SAVESEL Or ES_SELECTIONBAR, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hWnd, 0, App.hInstance, ByVal 0&)
    
        With m_lplf
            .lfCharSet = DEFAULT_CHARSET
            .lfFaceName = ""
            Call lstrcpy(.lfFaceName, "MS Sans Serif")
            .lfHeight = 8
        End With
    
        If (m_hWnd <> 0) Then
            m_fOwned = True
            'Call SetFontEx
            ShowWindow m_hWnd, 1
            Create = m_hWnd
        End If
    
    End If

End Function

Private Sub SetFontEx()

    Dim pDC As Long
    Dim intDefault As Long
    Dim lOldSize As Long
    
    If m_hWnd <> 0 Then
    
        If m_hFont <> 0 Then Call DeleteObject(m_hFont)
        
        With m_lplf
            .lfCharSet = DEFAULT_CHARSET
            pDC = GetDC(m_hWnd)
            If pDC Then
                lOldSize = .lfHeight
                If .lfHeight = 0 Then
                    intDefault = 8
                Else
                    intDefault = 0
                End If
                .lfHeight = -MulDiv(.lfHeight + intDefault, GetDeviceCaps(pDC, LOGPIXELSY), 72)
                Call ReleaseDC(m_hWnd, pDC)
            End If
        End With
        
        m_hFont = CreateFontIndirect(m_lplf)
        
        If (m_hFont <> 0) Then
            Call SendMessage(m_hWnd, WM_SETFONT, m_hFont, ByVal 1&)
        End If
        
        m_lplf.lfHeight = lOldSize
    
    End If

End Sub

Public Function GetTextRange(ByVal chStart As Long, ByVal chLen As Long) As String

    Dim LPTR As TEXTRANGE
    Dim sBuffer As String
    Dim lReturn As Long
    Dim chEnd As Long

    chEnd = chStart + chLen

    LPTR.chrg.cpMin = chStart
    LPTR.chrg.cpMax = chEnd
    
    If (chEnd - chStart <= 0) Then Exit Function
    
    sBuffer = String$(chEnd - chStart, vbNullChar)
    
    LPTR.lpstrText = sBuffer
    
    lReturn = SendMessage(m_hWnd, EM_GETTEXTRANGE, 0, LPTR)
    If (lReturn > 0) Then
        GetTextRange = LPTR.lpstrText
    End If

End Function

Public Function SetUndoLimit(ByVal MaxUndos As Long) As Boolean
    
    SetUndoLimit = (SendMessage(m_hWnd, EM_SETUNDOLIMIT, MaxUndos, ByVal 0&) = MaxUndos)

End Function

Public Function CharFromPos(X As Long, Y As Long) As Long
    
    CharFromPos = SendMessage(m_hWnd, EM_CHARFROMPOS, 0, ByVal MAKELONG(X, Y))

End Function

Public Sub UnSelect()

    Dim ucharg As CHARRANGE
    
    ucharg.cpMax = 0
    ucharg.cpMin = 0
    Call SendMessage(m_hWnd, EM_EXSETSEL, 0, ucharg)

End Sub

Public Sub ClearUndoBuffer()
    
    Call SendMessage(m_hWnd, EM_EMPTYUNDOBUFFER, 0, ByVal 0&)

End Sub

Public Sub Paste()
    
    Call SendMessage(m_hWnd, WM_PASTE, 0, ByVal 0&)

End Sub

Public Sub Copy()
    
    Call SendMessage(m_hWnd, WM_COPY, 0, ByVal 0&)

End Sub

Public Sub Cut()
    
    Call SendMessage(m_hWnd, WM_CUT, 0, ByVal 0&)

End Sub

Public Sub Undo()
    
    Call SendMessage(m_hWnd, EM_UNDO, 0, ByVal 0&)

End Sub

Public Sub Redo()
    
    Call SendMessage(m_hWnd, EM_REDO, 0, ByVal 0&)

End Sub

Public Sub SelectAll()

    Dim ucharg As CHARRANGE
    
    ucharg.cpMax = -1
    ucharg.cpMin = 0
    Call SendMessage(m_hWnd, EM_EXSETSEL, 0, ucharg)

End Sub

Public Function LineFromChar(CharPos As Long) As Long
    
    LineFromChar = SendMessage(m_hWnd, EM_EXLINEFROMCHAR, 0, ByVal CharPos)

End Function

Public Sub GetSelection(ByRef lStart As Long, ByRef lEnd As Long)

    Dim tCR As CHARRANGE
    
    SendMessage m_hWnd, EM_EXGETSEL, 0, tCR
    lStart = tCR.cpMin
    lEnd = tCR.cpMax

End Sub

Public Property Get SelectedText() As String

    Dim sBuff As String
    Dim lStart As Long
    Dim lEnd As Long
    Dim lR As Long
    
    GetSelection lStart, lEnd
    
    If (lEnd > lStart) Then
        sBuff = String$(lEnd - lStart + 1, 0)
        lR = SendMessageStr(m_hWnd, EM_GETSELTEXT, 0, sBuff)
        If (lR > 0) Then
            SelectedText = Left$(sBuff, lR)
        End If
    End If

End Property

Public Sub InsertContents(ByVal eType As ERECFileTypes, ByRef sText As String)

    Dim tStream As EDITSTREAM
    Dim lR As Long
   
   tStream.dwCookie = m_hWnd
   tStream.pfnCallback = GetAddressLong(AddressOf LoadCallBack)
   tStream.dwError = 0
   StreamText = sText
   
   lR = SendMessage(m_hWnd, EM_STREAMIN, eType Or SFF_SELECTION, tStream)

End Sub

Public Function FindText(ByVal sText As String, Optional ByVal eOptions As ERECFindTypeOptions = FR_DEFAULT, Optional ByVal bFindNext As Boolean = True, Optional ByVal bFIndInSelection As Boolean = False, Optional ByRef lMin As Long, Optional ByRef lMax As Long) As Long

    Dim tEx1 As FINDTEXTEX_A
    Dim tCR As CHARRANGE
    Dim lR As Long
    Dim lJunk As Long
    Dim b() As Byte

    m_sLastFindText = sText
    m_eLastFindMode = eOptions
    m_bLastFindNext = bFindNext
    
    lMin = -1: lMax = -1
    If (bFIndInSelection) Then
        GetSelection tCR.cpMax, tCR.cpMax
    Else
        If (bFindNext) Then
            GetSelection tCR.cpMin, lJunk
            If (lJunk >= tCR.cpMin) Then
                tCR.cpMin = lJunk + 1
            End If
            tCR.cpMax = -1
        Else
            tCR.cpMin = 0
            tCR.cpMax = -1
        End If
    End If
    
    b = StrConv(sText, vbFromUnicode)
    
    ReDim Preserve b(0 To UBound(b) + 1) As Byte
    b(UBound(b)) = 0
    tEx1.lpstrText = VarPtr(b(0))
    LSet tEx1.chrg = tCR
    
    lR = SendMessage(m_hWnd, EM_FINDTEXTEX, eOptions, tEx1)
    
    LSet tCR = tEx1.chrgText
    If (lR <> -1) Then
        lMax = tCR.cpMax
        lMin = lMax - Len(sText)
    End If
    FindText = lR
   
End Function

Public Property Get BorderStyle() As Long
   BorderStyle = m_bBorder
End Property

Public Property Get Border() As eBorder
    Border = m_Border
End Property

Public Property Let Border(ByVal NewValue As eBorder)
    
    m_Border = NewValue
    
    UserControl.BorderStyle = m_Border
    
    UserControl.PropertyChanged "Border"
    UserControl.Refresh
    
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal oColor As OLE_COLOR)

    UserControl.BackColor = oColor
    If (m_hWnd <> 0) Then
        SendMessageLong m_hWnd, EM_SETBKGNDCOLOR, 0, TranslateColor(oColor)
    End If
    PropertyChanged "BackColor"
    
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal oColor As OLE_COLOR)

    UserControl.ForeColor = oColor
    If (m_hWnd <> 0) Then
        SetFont UserControl.Font, TranslateColor(oColor), , , ercSetFormatAll
    End If
    PropertyChanged "ForeColor"

End Property

Public Property Get LastFindText() As String
   LastFindText = m_sLastFindText
End Property

Public Property Get LastFindMode() As ERECFindTypeOptions
   LastFindMode = m_eLastFindMode
End Property

Public Property Get LastFindNext() As Boolean
   LastFindNext = m_bLastFindNext
End Property

Public Property Get Font() As StdFont

    If (m_eCharFormatRange = ercSetFormatAll) Or (m_hWnd = 0) Then
        Set Font = UserControl.Font
    Else
        Dim sFnt As New StdFont
        Set Font = GetFont(True)
    End If
    
End Property

Public Property Set Font(ByRef sFnt As StdFont)

    With UserControl.Font
        .Name = sFnt.Name
        .Size = sFnt.Size
        .Bold = sFnt.Bold
        .Italic = sFnt.Italic
        .Underline = sFnt.Underline
        .Strikethrough = sFnt.Strikethrough
        .Charset = sFnt.Charset
    End With
    
    If (m_hWnd <> 0) Then
        SetFont sFnt, , , , m_eCharFormatRange
    End If
    PropertyChanged "Font"
    
End Property

Public Property Get CurrentLine() As Long
    
    On Error Resume Next

    Dim lStart As Long, lEnd As Long
    
    CurrentLine = 0
    
    GetSelection lStart, lEnd
    
    CurrentLine = SendMessageLong(m_hWnd, EM_EXLINEFROMCHAR, 0, lStart)
    
    CurrentLine = CurrentLine + 1
    
End Property

Public Property Get LineCount() As Long
    
    LineCount = SendMessage(m_hWnd, EM_GETLINECOUNT, 0, ByVal 0&)

End Property

Public Property Let MaxLength(ByVal rValue As Long)
    
    m_MaxLimit = rValue
    Call SendMessage(m_hWnd, EM_EXLIMITTEXT, 0, ByVal rValue)

End Property

Public Property Get MaxLength() As Long
    
    MaxLength = m_MaxLimit

End Property

Public Property Let SelRTF(sSelRTF As String)
    Call InsertContents(SF_RTF, sSelRTF)
End Property

Public Property Get SelText() As String
    SelText = Me.SelectedText
End Property

Public Property Let SelText(sSelTxt As String)
    Call InsertContents(SF_TEXT, sSelTxt)
End Property

Public Property Let SelStart(ByVal rValue As Long)

    Dim ucharg As CHARRANGE
    
    ucharg.cpMin = rValue
    ucharg.cpMax = rValue
    Call SendMessage(m_hWnd, EM_EXSETSEL, 0, ucharg)
    
End Property

Public Property Get SelStart() As Long

    Dim ucharg As CHARRANGE
    
    Call SendMessage(m_hWnd, EM_EXGETSEL, 0, ucharg)
    SelStart = ucharg.cpMin

End Property

Public Property Let SelLength(ByVal rValue As Long)

    Dim ucharg As CHARRANGE

    ucharg.cpMin = Me.SelStart
    ucharg.cpMax = ucharg.cpMin + rValue
    
    Call SendMessage(m_hWnd, EM_EXSETSEL, 0, ucharg)
    
End Property

Public Property Get SelLength() As Long

    Dim ucharg As CHARRANGE

    Call SendMessage(m_hWnd, EM_EXGETSEL, 0, ucharg)
    SelLength = ucharg.cpMax - ucharg.cpMin
    
End Property

Public Property Get SelAlignment() As ERECParagraphAlignmentConstants

    Dim tP As PARAFORMAT
    Dim tP2 As PARAFORMAT2
    Dim lR As Long
    
    If (m_eVersion = eRICHED32) Then
        tP.dwMask = PFM_ALIGNMENT
        tP.cbSize = Len(tP)
        lR = SendMessage(m_hWnd, EM_GETPARAFORMAT, 0, tP)
        SelAlignment = tP.wAlignment
    Else
        tP2.dwMask = PFM_ALIGNMENT
        tP2.cbSize = Len(tP2)
        lR = SendMessage(m_hWnd, EM_GETPARAFORMAT, 0, tP2)
        SelAlignment = tP2.wAlignment
    End If

End Property

Public Property Let SelAlignment(ByVal eAlign As ERECParagraphAlignmentConstants)

    Dim tP As PARAFORMAT
    Dim tP2 As PARAFORMAT2
    Dim lR As Long
    
    If (m_eVersion = eRICHED32) Then
        If (eAlign = ercParaJustify) Then
            'Unsupported
        Else
            tP.dwMask = PFM_ALIGNMENT
            tP.cbSize = Len(tP)
            tP.wAlignment = eAlign
            lR = SendMessage(m_hWnd, EM_SETPARAFORMAT, 0, tP)
        End If
    Else
        tP2.dwMask = PFM_ALIGNMENT
        tP2.cbSize = Len(tP2)
        tP2.wAlignment = eAlign
        lR = SendMessage(m_hWnd, EM_SETPARAFORMAT, 0, tP2)
    End If

End Property

Public Property Get SelBold() As Boolean

    Dim tCF As CHARFORMAT
    Dim lR As Long
    
    tCF.dwMask = CFM_BOLD
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
    FontBold = ((tCF.dwEffects And CFE_BOLD) = CFE_BOLD)
    
    SelBold = FontBold
    
End Property

Public Property Let SelBold(ByVal bBold As Boolean)

    Dim tCF As CHARFORMAT
    Dim lR As Long
    
    tCF.dwMask = CFM_BOLD
    If (bBold) Then
        tCF.dwEffects = CFE_BOLD
    End If
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    
End Property

Public Property Get SelItalic() As Boolean

    Dim tCF As CHARFORMAT
    Dim lR As Long
    
    tCF.dwMask = CFM_ITALIC
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
    FontItalic = ((tCF.dwEffects And CFE_ITALIC) = CFE_ITALIC)
    
    SelItalic = FontItalic
    
End Property

Public Property Let SelItalic(ByVal bItalic As Boolean)

    Dim tCF As CHARFORMAT
    Dim lR As Long
    
    tCF.dwMask = CFM_ITALIC
    If (bItalic) Then
        tCF.dwEffects = CFE_ITALIC
    End If
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    
End Property

Public Property Get SelUnderline() As Boolean

    Dim tCF As CHARFORMAT
    Dim lR As Long
    
    tCF.dwMask = CFM_UNDERLINE
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
    FontUnderline = ((tCF.dwEffects And CFE_UNDERLINE) = CFE_UNDERLINE)
    
    SelUnderline = FontUnderline
    
End Property

Public Property Let SelUnderline(ByVal bUnderline As Boolean)

    Dim tCF As CHARFORMAT
    Dim lR As Long
    
    tCF.dwMask = CFM_UNDERLINE
    If (bUnderline) Then
        tCF.dwEffects = CFE_UNDERLINE
    End If
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
   
End Property

Public Property Get SelStrikeThru() As Boolean

    Dim tCF As CHARFORMAT
    Dim lR As Long
    
    tCF.dwMask = CFM_STRIKEOUT
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
    SelStrikeThru = ((tCF.dwEffects And CFE_STRIKEOUT) = CFE_STRIKEOUT)
   
End Property

Public Property Let SelStrikeThru(ByVal bStrikeOut As Boolean)

    Dim tCF As CHARFORMAT
    Dim lR As Long
    
    tCF.dwMask = CFM_STRIKEOUT
    If (bStrikeOut) Then
        tCF.dwEffects = CFE_STRIKEOUT
    End If
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
   
End Property

Public Property Get SelFontName() As String

    Dim tCF As CHARFORMAT
    Dim lR As Long, i As Long
    Dim lColour As Long
    Dim sName As String
    
    tCF.dwMask = CFM_FACE
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
    For i = 1 To LF_FACESIZE
        sName = sName & Chr$(tCF.szFaceName(i - 1))
    Next i
    
    SelFontName = sName
    
End Property

Public Property Let SelFontName(sSelName As String)

    Dim tCF As CHARFORMAT
    Dim lR As Long, i As Long
    Dim lColour As Long
    
    tCF.dwMask = CFM_FACE
    tCF.cbSize = Len(tCF)
    For i = 1 To Len(sSelName)
        tCF.szFaceName(i - 1) = Asc(Mid$(sSelName, i, 1))
    Next i
    lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    DoEvents
   
End Property

Public Property Get SelFontSize() As Long
    
    Dim tCF As CHARFORMAT
    Dim lR As Long
    Dim lColour As Long
    
    tCF.dwMask = CFM_SIZE
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
    
    SelFontSize = tCF.yHeight / 20
   
End Property

Public Property Let SelFontSize(sSelSize As Long)

    Dim tCF As CHARFORMAT
    Dim lR As Long
    Dim lColour As Long
    
    tCF.dwMask = CFM_SIZE
    tCF.cbSize = Len(tCF)
    tCF.yHeight = sSelSize * 20
    lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
   
End Property

Public Property Get SelColor() As OLE_COLOR

    Dim tCF As CHARFORMAT
    Dim lR As Long
    Dim lColour As Long
    
    tCF.dwMask = CFM_COLOR
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
    SelColor = tCF.crTextColor
   
End Property

Public Property Let SelColor(ByVal oColour As OLE_COLOR)

    Dim tCF As CHARFORMAT
    Dim lR As Long
    Dim lColour As Long
    
    tCF.crTextColor = TranslateColor(oColour)
    tCF.dwMask = CFM_COLOR
    tCF.cbSize = Len(tCF)
    lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    
End Property

Public Property Get SelBackColor() As OLE_COLOR

    Dim tCF2 As CHARFORMAT2
    Dim lR As Long
    
    If (m_eVersion = eRICHED20) Then
        tCF2.dwMask = CFM_BACKCOLOR
        tCF2.cbSize = Len(tCF2)
        lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF2)
        SelBackColor = tCF2.crBackColor
    Else
        'unsuported
    End If

End Property

Public Property Let SelBackColor(ByVal oColor As OLE_COLOR)

    Dim tCF2 As CHARFORMAT2
    Dim lR As Long
    
    If (m_eVersion = eRICHED20) Then
        tCF2.dwMask = CFM_BACKCOLOR
        tCF2.crBackColor = TranslateColor(oColor)
        tCF2.cbSize = Len(tCF2)
        lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF2)
    Else
        'unsuported
    End If

End Property

Public Property Get SelLink() As Boolean

    Dim tCF2 As CHARFORMAT2
    Dim lR As Long

    If (m_eVersion = eRICHED20) Then
        tCF2.dwMask = CFM_LINK
        tCF2.cbSize = Len(tCF2)
        lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF2)
        SelLink = ((tCF2.dwEffects And CFE_LINK) = CFE_LINK)
    Else
        'unsuported
    End If

End Property

Public Property Let SelLink(ByVal bState As Boolean)

    Dim tCF2 As CHARFORMAT2
    Dim lR As Long
    
    If (m_eVersion = eRICHED20) Then
        tCF2.dwMask = CFM_LINK
        If (bState) Then
            tCF2.dwEffects = CFE_LINK
        End If
        tCF2.cbSize = Len(tCF2)
        lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF2)
    Else
        'unsuported
    End If

End Property

Public Property Get SelSuperScript() As Boolean

Dim tCF As CHARFORMAT
Dim tCF2 As CHARFORMAT2
Dim lR As Long

    If (m_eVersion = eRICHED32) Then
        tCF.dwMask = CFM_OFFSET
        tCF.cbSize = Len(tCF)
        lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
        SelSuperScript = (tCF.yOffset > 0)
    Else
        tCF2.dwMask = CFM_SUPERSCRIPT
        tCF2.cbSize = Len(tCF2)
        lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF2)
        SelSuperScript = ((tCF2.dwEffects And CFE_SUPERSCRIPT) = CFE_SUPERSCRIPT)
    End If

End Property

Public Property Get SelSubScript() As Boolean

    Dim tCF As CHARFORMAT
    Dim tCF2 As CHARFORMAT2
    Dim lR As Long
    
    If (m_eVersion = eRICHED32) Then
        tCF.dwMask = CFM_OFFSET
        tCF.cbSize = Len(tCF)
        lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF)
    Else
        tCF2.dwMask = CFM_SUBSCRIPT
        tCF2.cbSize = Len(tCF2)
        lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, m_eCharFormatRange, tCF2)
        SelSubScript = ((tCF2.dwEffects And CFE_SUBSCRIPT) = CFE_SUBSCRIPT)
    End If

End Property

Public Property Let SelSuperScript(ByVal bState As Boolean)

    Dim tCF As CHARFORMAT
    Dim tCF2 As CHARFORMAT2
    Dim lR As Long
    Dim Y As Long
    
    If (m_eVersion = eRICHED32) Then
        tCF.dwMask = CFM_SIZE
        tCF.cbSize = Len(tCF)
        lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, ercSetFormatSelection, tCF)
        Y = tCF.yHeight \ 2
        
        tCF.dwMask = CFM_OFFSET
        tCF.cbSize = Len(tCF)
        If (bState) Then
            tCF.yOffset = Y
        Else
            tCF.yOffset = 0
        End If
        lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    Else
        tCF2.dwMask = CFM_SUPERSCRIPT
        If (bState) Then
            tCF2.dwEffects = CFE_SUPERSCRIPT
        End If
        tCF2.cbSize = Len(tCF2)
        lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF2)
    End If

End Property

Public Property Let SelSubScript(ByVal bState As Boolean)
    
    Dim tCF As CHARFORMAT
    Dim tCF2 As CHARFORMAT2
    Dim lR As Long
    Dim Y As Long
    
    If (m_eVersion = eRICHED32) Then
        tCF.dwMask = CFM_SIZE
        tCF.cbSize = Len(tCF)
        lR = SendMessage(m_hWnd, EM_GETCHARFORMAT, ercSetFormatSelection, tCF)
        Y = tCF.yHeight \ -2
        
        tCF.dwMask = CFM_OFFSET
        tCF.cbSize = Len(tCF)
        If (bState) Then
            tCF.yOffset = Y
        Else
            tCF.yOffset = 0
        End If
        lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF)
    Else
        tCF2.dwMask = CFM_SUBSCRIPT
        If (bState) Then
            tCF2.dwEffects = CFE_SUBSCRIPT
        End If
        tCF2.cbSize = Len(tCF2)
        lR = SendMessage(m_hWnd, EM_SETCHARFORMAT, m_eCharFormatRange, tCF2)
    End If

End Property

Public Property Get CanRedo() As Boolean
    
    CanRedo = (SendMessage(m_hWnd, EM_CANREDO, 0, ByVal 0&) <> 0)

End Property

Public Property Get CanUndo() As Boolean
    
    CanUndo = (SendMessage(m_hWnd, EM_CANUNDO, 0, ByVal 0&) <> 0)

End Property

Public Property Get CanPaste(Optional Format As Long) As Boolean
    
    CanPaste = (SendMessage(m_hWnd, EM_CANPASTE, Format, ByVal 0&) <> 0)

End Property

Public Property Get ReadOnly() As Boolean
    
    ReadOnly = (GetWindowLong(m_hWnd, GWL_STYLE) And ES_READONLY)

End Property

Public Property Let ReadOnly(ByVal rValue As Boolean)
    
    Call SendMessage(m_hWnd, EM_SETREADONLY, Abs(rValue), ByVal 0&)

End Property

Public Property Get Text() As String

    Text = Me.Contents(SF_TEXT)
    
End Property

Public Property Let Text(ByVal rValue As String)
    
    Me.Contents(SF_TEXT) = rValue

End Property

Public Property Get TextRTF() As String

    TextRTF = Me.Contents(SF_RTF)
    
End Property

Public Property Let TextRTF(sTextRTF As String)

    Me.Contents(SF_RTF) = sTextRTF
    
End Property

Public Property Let Contents(ByVal eType As ERECFileTypes, ByRef sContents As String)

    Dim tStream As EDITSTREAM
    Dim lR As Long

    tStream.dwCookie = m_hWnd
    tStream.pfnCallback = GetAddressLong(AddressOf LoadCallBack)
    tStream.dwError = 0
    StreamText = sContents
    RichEdit = Me
    
    lR = SendMessage(m_hWnd, EM_STREAMIN, eType, tStream)
    ClearRichEdit
    
    SendMessageLong m_hWnd, EM_SETMODIFY, 0, 0

End Property

Public Property Get Contents(ByVal eType As ERECFileTypes) As String

    Dim tStream As EDITSTREAM

   tStream.dwCookie = m_hWnd
   tStream.pfnCallback = GetAddressLong(AddressOf SaveCallBack)
   tStream.dwError = 0
   
   ClearStreamText
   RichEdit = Me
   SendMessage m_hWnd, EM_STREAMOUT, eType, tStream
   ClearRichEdit

   Contents = StreamText()

End Property

Public Property Get AutoURLDetect() As Boolean
    
    AutoURLDetect = (SendMessage(m_hWnd, EM_GETAUTOURLDETECT, 0, ByVal 0&) <> 0)

End Property

Public Property Let AutoURLDetect(ByVal rValue As Boolean)
    
    Call SendMessage(m_hWnd, EM_AUTOURLDETECT, Abs(rValue), ByVal 0&)

End Property

Public Property Get hWndRTF() As Long

    hWndRTF = m_hWnd
    
End Property

Public Property Get ViewMode() As ERECViewModes

   ViewMode = m_eViewMode
   
End Property

Public Property Let ViewMode(ByVal eViewMode As ERECViewModes)

   If (eViewMode <> m_eViewMode) Then
      m_eViewMode = eViewMode
      Select Case m_eViewMode
      Case ercWYSIWYG
         ' todo...
         SendMessageLong m_hWnd, EM_SETTARGETDEVICE, Printer.hDC, Printer.Width
      Case ercWordWrap
         SendMessageLong m_hWnd, EM_SETTARGETDEVICE, 0, 0
      Case ercDefault
         SendMessageLong m_hWnd, EM_SETTARGETDEVICE, 0, 1
      End Select
   End If
   
End Property

Public Property Let UseVersion(ByVal eVersion As ERECControlVersion)

    On Error Resume Next

    If (UserControl.Ambient.UserMode) Then
        ' can't set at run time in this implementation.
    Else
        m_eVersion = eVersion
    End If
    
End Property

Public Property Get UseVersion() As ERECControlVersion
    UseVersion = m_eVersion
End Property

Public Property Get EventNotificationMask() As RE_EventNotificationMask
    EventNotificationMask = SendMessage(m_hWnd, EM_GETEVENTMASK, 0, ByVal 0&)
End Property

Public Property Let EventNotificationMask(ByVal rValue As RE_EventNotificationMask)
    Call SendMessage(m_hWnd, EM_SETEVENTMASK, 0, ByVal rValue)
End Property

Public Function LoadFromFile(ByVal sFile As String, ByVal eType As ERECFileTypes) As Boolean
    
    Dim hFile As Long
    Dim tOF As OFSTRUCT
    Dim tStream As EDITSTREAM
    Dim lR As Long
    
    m_eProgressType = ercLoad
    
    hFile = OpenFile(sFile, tOF, OF_READ)
    
    If (hFile <> 0) Then
        tStream.dwCookie = hFile
        tStream.pfnCallback = GetAddressLong(AddressOf LoadCallBack)
        tStream.dwError = 0
        
        RichEdit = Me
        FileMode = True
        
        lR = SendMessage(m_hWnd, EM_STREAMIN, eType, tStream)
        
        LoadFromFile = (lR <> 0)
        
        FileMode = False
        ClearRichEdit
        
        CloseHandle hFile
    End If

End Function

Public Function SaveToFile(ByVal sFile As String, ByVal eType As ERECFileTypes) As Boolean
   
    Dim tStream As EDITSTREAM
    Dim tOF As OFSTRUCT
    Dim hFile As Long
    Dim lR As Long
    
    hFile = OpenFile(sFile, tOF, OF_CREATE)
    
    If (hFile <> 0) Then
        tStream.dwCookie = hFile
        tStream.pfnCallback = GetAddressLong(AddressOf SaveCallBack)
        tStream.dwError = 0
        FileMode = True
        RichEdit = Me
        
        lR = SendMessage(m_hWnd, EM_STREAMOUT, eType, tStream)
        
        SaveToFile = (lR <> 0)
        
        FileMode = False
        ClearRichEdit
        
        CloseHandle hFile
    End If
       
End Function

Private Function GetShiftState() As Integer

    Dim iR As Integer
    Dim lR As Long
    Dim lKey As Long
    
    iR = iR Or (-vbShiftMask * KeyIsPressed(VK_SHIFT))
    iR = iR Or (-vbAltMask * KeyIsPressed(VK_MENU))
    iR = iR Or (-vbCtrlMask * KeyIsPressed(VK_CONTROL))
    
    GetShiftState = iR

End Function

Private Function KeyIsPressed(ByVal nVirtKeyCode As KeyCodeConstants) As Boolean

    Dim lR As Long
    
    lR = GetAsyncKeyState(nVirtKeyCode)
    
    If (lR And &H8000&) = &H8000& Then
        KeyIsPressed = True
    End If
    
End Function

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long

    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If

End Function

Public Function HIWORD(ByVal dwValue As Long) As Long
    Call CopyMemory(HIWORD, ByVal VarPtr(dwValue) + 2, 2)
End Function
  
Public Function LOWORD(ByVal dwValue As Long) As Long
    Call CopyMemory(LOWORD, dwValue, 2)
End Function

Private Function MAKELONG(ByVal wLow As Long, ByVal wHi As Long) As Long

    If (wHi And &H8000&) Then
        MAKELONG = (((wHi And &H7FFF&) * 65536) Or (wLow And &HFFFF&)) Or &H80000000
    Else
        MAKELONG = LOWORD(wLow) Or (&H10000 * LOWORD(wHi))
    End If

End Function

Private Function GetAddressLong(ByVal lAddr As Long) As Long
    GetAddressLong = lAddr
End Function

Private Sub Destroy()

    If (m_fOwned And (m_hWnd <> 0)) Then
        Call DestroyWindow(m_hWnd)
        Call DeleteObject(m_hFont)
        m_hFont = 0
        m_hWnd = 0
        m_fOwned = False
    End If

End Sub

Private Sub SetFont(ByRef fntThis As StdFont, Optional ByVal oColor As OLE_COLOR = vbWindowText, Optional ByVal eType As ERECTextTypes = ercTextNormal, Optional ByVal bHyperLink As Boolean = False, Optional ByVal eRange As ERECSetFormatRange = ercSetFormatSelection)
    
    On Error Resume Next
    
    Dim tCF As CHARFORMAT
    Dim tCF2 As CHARFORMAT2
    Dim dwEffects As Long
    Dim dwMask As Long
    Dim i As Long
    
    tCF.cbSize = Len(tCF)
    tCF.crTextColor = TranslateColor(oColor)
    dwMask = CFM_COLOR
    If fntThis.Bold Then
        dwEffects = dwEffects Or CFE_BOLD
    End If
    dwMask = dwMask Or CFM_BOLD
    If fntThis.Italic Then
        dwEffects = dwEffects Or CFE_ITALIC
    End If
    dwMask = dwMask Or CFM_ITALIC
    If fntThis.Strikethrough Then
        dwEffects = dwEffects Or CFE_STRIKEOUT
    End If
    dwMask = dwMask Or CFM_STRIKEOUT
    If fntThis.Underline Then
        dwEffects = dwEffects Or CFE_UNDERLINE
    End If
    dwMask = dwMask Or CFM_UNDERLINE
    
    If bHyperLink Then
        dwEffects = dwEffects Or CFE_LINK
    End If
    dwMask = dwMask Or CFM_LINK
    
    tCF.dwEffects = dwEffects
    tCF.dwMask = dwMask Or CFM_FACE Or CFM_SIZE
    
    For i = 1 To Len(fntThis.Name)
        tCF.szFaceName(i - 1) = Asc(Mid$(fntThis.Name, i, 1))
    Next i
    tCF.yHeight = (fntThis.Size * 20)
    If (eType = ercTextSubscript) Then
        tCF.yOffset = -tCF.yHeight \ 2
    End If
    If (eType = ercTextSuperscript) Then
        tCF.yOffset = tCF.yHeight \ 2
    End If
    
    If (m_eVersion = eRICHED32) Then
        SendMessage m_hWnd, EM_SETCHARFORMAT, eRange, tCF
    Else
        CopyMemory tCF2, tCF, Len(tCF)
        tCF2.cbSize = Len(tCF2)
        tCF.yOffset = 0
        If (eType = ercTextSubscript) Then
            tCF.dwEffects = tCF.dwEffects Or CFE_SUBSCRIPT
            tCF.dwMask = tCF.dwMask Or CFM_SUBSCRIPT
        End If
        If (eType = ercTextSuperscript) Then
            tCF.dwEffects = tCF.dwEffects Or CFE_SUPERSCRIPT
            tCF.dwMask = tCF.dwMask Or CFM_SUPERSCRIPT
        End If
        SendMessage m_hWnd, EM_SETCHARFORMAT, eRange, tCF2
    End If
    
    cSelColor = oColor
    cSelBold = fntThis.Bold
    cSelFontName = fntThis.Name
    cSelFontSize = fntThis.Size
    cSelItalic = fntThis.Italic
    cSelUnderline = fntThis.Underline
   
End Sub

Private Function GetFont(Optional ByVal bForSelection As Boolean = False, Optional ByRef oColor As OLE_COLOR, Optional ByRef bHyperLink As Boolean, Optional ByVal eType As ERECTextTypes = ercTextNormal) As StdFont
    
    On Error Resume Next

    Dim sFnt As New StdFont
    Dim tCF As CHARFORMAT
    Dim tCF2 As CHARFORMAT2
    Dim dwEffects As Long
    Dim dwMask As Long
    Dim i As Long
    Dim sName As String
    Dim a As Boolean
    
    tCF.cbSize = Len(tCF)
    dwMask = dwMask Or CFM_COLOR
    
    dwMask = dwMask Or CFM_BOLD
    dwMask = dwMask Or CFM_ITALIC
    dwMask = dwMask Or CFM_STRIKEOUT
    dwMask = dwMask Or CFM_UNDERLINE
    dwMask = dwMask Or CFM_LINK
    
    If (m_eVersion = eRICHED32) Then
        tCF.dwEffects = dwEffects
        tCF.dwMask = dwMask Or CFM_FACE Or CFM_SIZE
        SendMessage m_hWnd, EM_GETCHARFORMAT, Abs(bForSelection), tCF
    Else
        CopyMemory tCF2, tCF, Len(tCF)
        tCF2.cbSize = Len(tCF2)
        SendMessage m_hWnd, EM_GETCHARFORMAT, Abs(bForSelection), tCF2
    End If
    If (m_eVersion = eRICHED32) Then
        
        oColor = tCF.crTextColor
    
        For i = 1 To LF_FACESIZE
            sName = sName & Chr$(tCF.szFaceName(i - 1))
        Next i
        sFnt.Name = sName
        sFnt.Size = tCF.yHeight \ 20
        
        cSelBold = ((tCF2.dwEffects And CFE_BOLD) = CFE_BOLD)
        
        sFnt.Bold = ((tCF.dwEffects And CFE_BOLD) = CFE_BOLD)
        sFnt.Italic = ((tCF.dwEffects And CFE_ITALIC) = CFE_ITALIC)
        sFnt.Underline = ((tCF.dwEffects And CFE_UNDERLINE) = CFE_UNDERLINE)
        sFnt.Strikethrough = ((tCF.dwEffects And CFE_STRIKEOUT) = CFE_STRIKEOUT)
        bHyperLink = ((tCF.dwEffects And CFE_LINK) = CFE_LINK)
        If (tCF.yOffset = 0) Then
            eType = ercTextNormal
        ElseIf (tCF.yOffset < 0) Then
            eType = ercTextSubscript
        Else
            eType = ercTextSuperscript
        End If
    Else
        oColor = tCF2.crTextColor
        For i = 1 To LF_FACESIZE
            sName = sName & Chr$(tCF2.szFaceName(i - 1))
        Next i
        sFnt.Name = sName
        sFnt.Size = tCF2.yHeight \ 20
        sFnt.Bold = ((tCF2.dwEffects And CFE_BOLD) = CFE_BOLD)
        sFnt.Italic = ((tCF2.dwEffects And CFE_ITALIC) = CFE_ITALIC)
        sFnt.Underline = ((tCF2.dwEffects And CFE_UNDERLINE) = CFE_UNDERLINE)
        sFnt.Strikethrough = ((tCF2.dwEffects And CFE_STRIKEOUT) = CFE_STRIKEOUT)
        bHyperLink = ((tCF2.dwEffects And CFE_LINK) = CFE_LINK)
        eType = ercTextNormal
        If ((tCF2.dwEffects And CFE_SUPERSCRIPT) = CFE_SUPERSCRIPT) Then
            eType = ercTextSuperscript
        End If
        If ((tCF2.dwEffects And CFE_SUBSCRIPT) = CFE_SUBSCRIPT) Then
            eType = ercTextSubscript
        End If
    End If
    
    cSelColor = oColor
    cSelFontName = sFnt.Name
    cSelFontSize = sFnt.Size
    cSelItalic = sFnt.Italic
    cSelUnderline = sFnt.Underline
    
    Set GetFont = sFnt

End Function

Private Function IsRtf(ByRef sFileText As String) As Boolean
    IsRtf = False
    If (Left$(sFileText, 5) = "{\rtf") Then
        IsRtf = True
    End If
End Function

