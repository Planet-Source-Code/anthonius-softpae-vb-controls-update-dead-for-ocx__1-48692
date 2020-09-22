VERSION 5.00
Begin VB.UserControl List 
   ClientHeight    =   4155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
   ScaleHeight     =   4155
   ScaleWidth      =   6240
   ToolboxBitmap   =   "ctlListBox.ctx":0000
   Begin VB.ListBox List1 
      Height          =   3660
      Left            =   180
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   5865
   End
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

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

'*******************************************************************************
Dim mlWndProc  As Long
Dim mlSetStyle As Long
Dim m_hBmp As Long

Dim m_ListCount As Long
Dim m_ListIndex As Long
Dim m_MousePointer As Long
Dim m_ListImage As String
Dim m_Text As String
Dim m_List As String

Public Enum SELECTED_STATE
    NotSelected = 0&
    IsSelected = 1&
End Enum
  
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event ItemInfoCallback(ByVal Index As Long, ByVal State As SELECTED_STATE, Text As String, BackColor As Long, TextColor As Long)
Public Event Scroll()

Private Const SWP_SHOWWINDOW = &H40

' SetWindowPos Flags
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType    As Long
    CtlID      As Long
    itemID     As Long
    itemAction As Long
    itemState  As Long
    hwndItem   As Long
    hdc        As Long
    rcItem     As RECT
    ItemData   As Long
End Type

Private Type MEASUREITEMSTRUCT
    CtlType    As Long
    CtlID      As Long
    itemID     As Long
    itemWidth  As Long
    itemHeight As Long
    ItemData   As Long
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Const WM_DRAWITEM = &H2B

Private Const LB_GETTEXT = &H189
Private Const LB_SETITEMHEIGHT = &H1A0
'
' Window field offsets for GetWindowLong and GetWindowWord APIs.
'
Private Const GWL_WNDPROC = (-4)
Private Const GWL_HINSTANCE = (-6)
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_USERDATA = (-21)
Private Const GWL_ID = (-12)
'
' GetSysColor colors.
'
Private Const COLOR_SCROLLBAR = 0
Private Const COLOR_BACKGROUND = 1
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_MENU = 4
Private Const COLOR_WINDOW = 5
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_WINDOWTEXT = 8
Private Const COLOR_CAPTIONTEXT = 9
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_INACTIVEBORDER = 11
Private Const COLOR_APPWORKSPACE = 12
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_GRAYTEXT = 17
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_INACTIVECAPTIONTEXT = 19
Private Const COLOR_BTNHIGHLIGHT = 20
'
' Window Styles.
'
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_CHILD = &H40000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED = &H8000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_CAPTION = &HC00000  'WS_BORDER Or WS_DLGFRAME
Private Const WS_BORDER = &H800000
Private Const WS_DLGFRAME = &H400000
Private Const WS_VSCROLL = &H200000
Private Const WS_HSCROLL = &H100000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_GROUP = &H20000
Private Const WS_TABSTOP = &H10000

Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
'
' Common Window Styles.
'
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_CHILDWINDOW = (WS_CHILD)
'
' Extended Window Styles.
'
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_TRANSPARENT = &H20&
'
' Listbox Styles.
'
Private Const LBS_NOTIFY = &H1&
Private Const LBS_SORT = &H2&
Private Const LBS_NOREDRAW = &H4&
Private Const LBS_MULTIPLESEL = &H8&
Private Const LBS_OWNERDRAWFIXED = &H10&
Private Const LBS_OWNERDRAWVARIABLE = &H20&
Private Const LBS_HASSTRINGS = &H40&
Private Const LBS_USETABSTOPS = &H80&
Private Const LBS_NOINTEGRALHEIGHT = &H100&
Private Const LBS_MULTICOLUMN = &H200&
Private Const LBS_WANTKEYBOARDINPUT = &H400&
Private Const LBS_EXTENDEDSEL = &H800&
Private Const LBS_DISABLENOSCROLL = &H1000&
Private Const LBS_NODATA = &H2000&
Private Const LBS_STANDARD = (LBS_NOTIFY Or LBS_SORT Or WS_VSCROLL Or WS_BORDER)
'
' Owner draw control types.
'
Private Const ODT_MENU = 1
Private Const ODT_LISTBOX = 2
Private Const ODT_COMBOBOX = 3
Private Const ODT_BUTTON = 4
Private Const ODT_STATIC = 5
Private Const ODT_HEADER = 100
Private Const ODT_TAB = 101
Private Const ODT_LISTVIEW = 102
'
' Owner draw actions.
'
Private Const ODA_DRAWENTIRE = &H1
Private Const ODA_SELECT = &H2
Private Const ODA_FOCUS = &H4
'
' Owner draw state.
'
Private Const ODS_SELECTED = &H1
Private Const ODS_GRAYED = &H2
Private Const ODS_DISABLED = &H4
Private Const ODS_CHECKED = &H8
Private Const ODS_FOCUS = &H10
Private Const ODS_DEFAULT = &H20
Private Const ODS_COMBOBOXEDIT = &H1000
Private Const ODS_HOTLIGHT = &H40
Private Const ODS_INACTIVE = &H80

Private Const CF_BITMAP = 2
Private Const IMAGE_BITMAP = 0
Private Const IMAGE_ICON = 1
Private Const IMAGE_CURSOR = 2
Private Const IMAGE_ENHMETAFILE = 3
Private Const LR_LOADFROMFILE = &H10
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_LOADTRANSPARENT = &H20

Private Const RDW_INVALIDATE = &H1
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_ERASE = &H4

Private Const RDW_VALIDATE = &H8
Private Const RDW_NOINTERNALPAINT = &H10
Private Const RDW_NOERASE = &H20

Private Const RDW_NOCHILDREN = &H40
Private Const RDW_ALLCHILDREN = &H80

Private Const RDW_UPDATENOW = &H100
Private Const RDW_ERASENOW = &H200

Private Const RDW_FRAME = &H400
Private Const RDW_NOFRAME = &H800

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function LoadBitmapEx Lib "user32" Alias "LoadBitmapA" (ByVal hInstance As Long, lpBitmapName As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

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
    
    'funkcia volaná pred hlavnou (systémovou) wnd funkcii
    
    Dim hdc     As Long
    Dim lRet    As Long
    Dim hBmp    As Long
    Dim lFont   As Long
    Dim nTxtColor As Long
    Dim nBkColor As Long
    Dim sString As String
    Dim tSize   As SIZE
    Dim hPic    As StdPicture
    Dim DIS     As DRAWITEMSTRUCT

    Select Case uMsg
    
    Case WM_DRAWITEM
        
        Call CopyMemory(DIS, ByVal lParam, Len(DIS))

        With DIS
            
            sString = Space$(128)
            lRet = SendMessage(.hwndItem, LB_GETTEXT, .itemID, ByVal sString)
            sString = Left$(sString, lRet)
            
            'hbmpPicture =(HBITMAP)SendMessage(lpdis->hwndItem, LB_GETITEMDATA, lpdis->itemID, (LPARAM) 0);
            
            If .itemState And ODS_FOCUS Then
                eSelectedState = IsSelected
                nBkColor = GetSysColor(COLOR_HIGHLIGHT)
                nTxtColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
            Else
                eSelectedState = NotSelected
                nBkColor = GetSysColor(COLOR_WINDOW)
                nTxtColor = GetSysColor(COLOR_WINDOWTEXT)
            End If

            RaiseEvent ItemInfoCallback(.itemID, eSelectedState, sString, nBkColor, nTxtColor)
            
            If m_hBmp = 0 Then
                If InIDE Then
                    hBmp = LoadImage(App.hInstance, App.Path & "\" & m_ListImage & ".bmp", IMAGE_BITMAP, 16, 16, LR_LOADTRANSPARENT Or LR_LOADFROMFILE)
                Else
                    hBmp = LoadImage(App.hInstance, m_ListImage, IMAGE_BITMAP, 16, 16, LR_LOADTRANSPARENT)
                End If
                m_hBmp = hBmp
            Else
                hBmp = m_hBmp
            End If
            
            'hBmp = LoadImage(App.hInstance, m_ListImage, IMAGE_BITMAP, 16, 16, LR_LOADTRANSPARENT)
            hdc = CreateCompatibleDC(.hdc)
            
            Call SelectObject(hdc, hBmp)
            
            Call GetTextExtentPoint32(.hdc, sString, Len(sString), tSize)
            Call SendMessage(.hwndItem, LB_SETITEMHEIGHT, .itemID, ByVal 20&) 'tSize.cy + 5)
            
            crback = nBkColor
            Call SetBkColor(.hdc, nBkColor)
            Call SetTextColor(.hdc, nTxtColor)
            
'            If .itemState And ODS_FOCUS Then '16
'                crback = GetSysColor(COLOR_HIGHLIGHT)
'                Call SetBkColor(.hdc, GetSysColor(COLOR_HIGHLIGHT))
'                Call SetTextColor(.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT))
'            Else
'                crback = GetSysColor(COLOR_WINDOW)
'                Call SetBkColor(.hdc, GetSysColor(COLOR_WINDOW))
'                Call SetTextColor(.hdc, GetSysColor(COLOR_WINDOWTEXT))
'            End If
            
            .rcItem.Left = 18
            hbrback = CreateSolidBrush(crback)
            FillRect .hdc, .rcItem, hbrback
            DeleteObject hbrback
            .rcItem.Left = 0
            
        End With
        
        With DIS.rcItem
            Call BitBlt(DIS.hdc, .Left, .Top, .Right - .Left, .Bottom - .Top, hdc, 0, 0, vbSrcCopy)
            Call TextOut(DIS.hdc, .Left + 22, .Top + 3, sString, Len(sString))
        End With
        
        If (DIS.itemState And ODS_FOCUS) Then
            DIS.rcItem.Left = 18
            DrawFocusRect DIS.hdc, DIS.rcItem
            DIS.rcItem.Left = 0
        End If
        
        Call DeleteDC(hdc)
        'Call DeleteObject(hBmp)
    
    End Select
    bHandled = True

End Sub

Private Sub UserControl_Initialize()

    On Error Resume Next

    Call SetListboxHook(List1.hWnd, True)

    Call subInitialize
    Call Subclass(UserControl.hWnd, Me)
    
    'add messages to subclass
    Call AddMsg(WM_DRAWITEM, MSG_BEFORE)
    
    m_hBmp = 0

    If m_hBmp = 0 Then m_hBmp = LoadImage(App.hInstance, m_ListImage, IMAGE_BITMAP, 16, 16, LR_LOADTRANSPARENT)

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    List1.Top = 0
    List1.Left = 0
    List1.Height = UserControl.ScaleHeight
    List1.Width = UserControl.ScaleWidth

    UserControl.Height = List1.Height

End Sub

Private Sub UserControl_Terminate()

    On Error Resume Next
    
    Call DeleteObject(m_hBmp)

    If hWndSubclass <> 0 Then
        Call UnSubclass
    End If
    
End Sub

Public Function SetListboxHook(ByVal hWnd As Long, ByVal m_mode As Boolean) As Long

    On Error Resume Next

    If m_mode = True Then
        ret = GetWindowLong(hWnd, GWL_STYLE)
        ret = ret Or LBS_OWNERDRAWFIXED
        ret = SetWindowLong(hWnd, GWL_STYLE, ret)
        'mlWndProc = SetWindowLong(hWndParent, GWL_WNDPROC, AddressOf lbs_WndProc)
    Else
        'mlWndProc = SetWindowLong(hWndParent, GWL_WNDPROC, mlWndProc)
    End If
    
    RedrawWindow hWnd, ByVal 0&, 0, RDW_INVALIDATE
    'RedrawWindow hWndParent, ByVal 0&, 0, RDW_INVALIDATE

End Function

Public Property Get ListCount() As Long
    ListCount = List1.ListCount
End Property

Public Property Get ListIndex() As Long
    ListIndex = List1.ListIndex
End Property

Public Property Get ListImage() As String
    ListImage = m_ListImage
End Property

Public Property Let ListImage(ByVal n_ListImage As String)
    m_ListImage = n_ListImage
End Property

Public Property Get List(ByVal Index As Long) As String
    m_List = List1.List(Index)
    List = m_List
End Property

Public Property Let List(ByVal Index As Long, ByVal n_List As String)
    m_List = n_List
    List1.List(Index) = m_List
End Property

Public Property Get Text() As String
    m_Text = List1.Text
    Text = m_Text
End Property

Public Property Let Text(ByVal n_Text As String)
    m_Text = n_Text
    List1.Text = m_Text
End Property

Public Property Get MousePointer() As Long
    m_MousePointer = List1.MousePointer
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal n_MousePointer As Long)
    m_MousePointer = n_MousePointer
    List1.MousePointer = m_MousePointer
End Property

Public Property Get ItemData(ByVal Index As Long) As Long
  ItemData = List1.ItemData(Index)
End Property

Public Property Let ItemData(ByVal Index As Long, ByVal nNewVal As Long)
  List1.ItemData(Index) = nNewVal
End Property

Private Sub List1_Click()
    RaiseEvent Click
End Sub

Private Sub List1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub List1_Scroll()
    RaiseEvent Scroll
End Sub

Public Sub AddItem(ByVal Item As String, Optional Index As Variant = -1)

    If Index >= 0 Then
        List1.AddItem Item, Index
    Else
        List1.AddItem Item
    End If
    
    'nItem = SendMessage(hwndList, LB_ADDSTRING, 0, lpstr);
    'SendMessage(hwndList, LB_SETITEMDATA, nItem, hbmp);
        
End Sub

Public Sub RemoveItem(ByVal Index As Integer)
    List1.RemoveItem Index
End Sub

Public Sub Clear()
    List1.Clear
End Sub

