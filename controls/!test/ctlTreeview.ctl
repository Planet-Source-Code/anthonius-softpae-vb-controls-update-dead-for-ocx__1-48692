VERSION 5.00
Begin VB.UserControl Tree 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ctlTreeview.ctx":0000
End
Attribute VB_Name = "Tree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private hCont As Long
Private hTree As Long
Private iNodes As Long
Private m_Item As Long
Private m_hImageList As Long

Dim bInitialized As Boolean

Private m_Border As eBorder

Public Enum eBorder
    [None] = 0
    [Fixed Single] = 1
End Enum

'Public Enum TV_ExpandFlags
'  TVE_COLLAPSE = &H1
'  TVE_EXPAND = &H2
'  TVE_TOGGLE = &H3
'  TVE_EXPANDPARTIAL = &H4000
'  TVE_COLLAPSERESET = &H8000
'End Enum

Private Const SS_WHITERECT = &H6&

Private Const MAX_LEN = 32
Private Const ID_TREEVIEW = 1000

Public Event ItemClick(ByVal Button As Long, ByVal ItemText As String, ByVal ItemKey As String, ByVal ItemIndex As Long)
Public Event ItemDblClick(ByVal ItemText As String, ByVal ItemKey As String, ByVal ItemIndex As Long)
Public Event SelChanged(ByVal ItemText As String, ByVal ItemKey As String, ByVal ItemIndex As Long)
Public Event Click(ByVal cx As Long, ByVal cy As Long)
Public Event DblClick(ByVal cx As Long, ByVal cy As Long)

Private Const ICC_TREEVIEW_CLASSES = &H2

Private Type TNode
    hItem As Long
    hParent As Long
    Index As Long
    Key As String
    Text As String
    Image As Long
    Tag As Variant
End Type

Private TNodes() As TNode

Public Enum HitTestInfoConstants
    htAbove = &H100
    htBelow = &H200
    htBelowLast = &H1
    htItemPlusMinus = &H10
    htItemIcon = &H2
    htItemIndent = &H8
    htItemText = &H4
    htItemRight = &H20
    htItemState = &H40
    htLeft = &H800
    htRight = &H400
End Enum

Public Enum RelationConstants
    tvwSort
    tvwFirst
    tvwLast
    tvwChild
End Enum

Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const GWL_WNDPROC = (-4)
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private Const TVGN_CARET = &H9

' TreeView messages.
Private Const TV_FIRST = &H1100
Private Const TVM_CREATEDRAGIMAGE = (TV_FIRST + 18)
Private Const TVM_DELETEITEM = (TV_FIRST + 1)
Private Const TVM_EDITLABEL = (TV_FIRST + 14)
Private Const TVM_ENDEDITLABELNOW = (TV_FIRST + 22)
Private Const TVM_ENSUREVISIBLE = (TV_FIRST + 20)
Private Const TVM_EXPAND = (TV_FIRST + 2)
Private Const TVM_GETBKCOLOR = (TV_FIRST + 31)
Private Const TVM_GETBORDER = (TV_FIRST + 36)
Private Const TVM_GETCOUNT = (TV_FIRST + 5)
Private Const TVM_GETEDITCONTROL = (TV_FIRST + 15)
Private Const TVM_GETIMAGELIST = (TV_FIRST + 8)
Private Const TVM_GETINDENT = (TV_FIRST + 6)
Private Const TVM_GETISEARCHSTRINGA = (TV_FIRST + 23)
Private Const TVM_GETITEM = (TV_FIRST + 12)
Private Const TVM_GETITEMHEIGHT = (TV_FIRST + 28)
Private Const TVM_GETITEMRECT = (TV_FIRST + 4)
Private Const TVM_GETNEXTITEM = (TV_FIRST + 10)
Private Const TVM_GETSCROLLTIME = (TV_FIRST + 34)
Private Const TVM_GETTEXTCOLOR = (TV_FIRST + 32)
Private Const TVM_GETTOOLTIPS = (TV_FIRST + 25)
Private Const TVM_GETVISIBLECOUNT = (TV_FIRST + 16)
Private Const TVM_HITTEST = (TV_FIRST + 17)
Private Const TVM_INSERTITEM = (TV_FIRST + 0)
Private Const TVM_SELECTITEM = (TV_FIRST + 11)
Private Const TVM_SETBKCOLOR = (TV_FIRST + 29)
Private Const TVM_SETBORDER = (TV_FIRST + 35)
Private Const TVM_SETIMAGELIST = (TV_FIRST + 9)
Private Const TVM_SETINDENT = (TV_FIRST + 7)
Private Const TVM_SETINSERTMARK = (TV_FIRST + 26)
Private Const TVM_SETITEM = (TV_FIRST + 13)
Private Const TVM_SETITEMHEIGHT = (TV_FIRST + 27)
Private Const TVM_SETSCROLLTIME = (TV_FIRST + 33)
Private Const TVM_SETTEXTCOLOR = (TV_FIRST + 30)
Private Const TVM_SETTOOLTIPS = (TV_FIRST + 24)
Private Const TVM_SORTCHILDREN = (TV_FIRST + 19)
Private Const TVM_SORTCHILDRENCB = (TV_FIRST + 21)
Private Const TVM_SETLINECOLOR = (TV_FIRST + 40)
Private Const TVM_GETLINECOLOR = (TV_FIRST + 41)

' Treeview Notifications
Private Const TVN_FIRST = -400
Private Const TVN_BEGINLABELEDIT = (TVN_FIRST - 10)
Private Const TVN_BEGINDRAG = (TVN_FIRST - 7)
Private Const TVN_BEGINRDRAG = (TVN_FIRST - 8)
Private Const TVN_DELETEITEM = (TVN_FIRST - 9)
Private Const TVN_GETDISPINFO = (TVN_FIRST - 3)
Private Const TVN_GETINFOTIP = (TVN_FIRST - 13)
Private Const TVN_KEYDOWN = (TVN_FIRST - 12)
Private Const TVN_ENDLABELEDIT = (TVN_FIRST - 11)
Private Const TVN_ITEMEXPANDED = (TVN_FIRST - 6)
Private Const TVN_ITEMEXPANDING = (TVN_FIRST - 5)
Private Const TVN_SELCHANGED = (TVN_FIRST - 2)
Private Const TVN_SELCHANGING = (TVN_FIRST - 1)
Private Const TVN_SINGLEEXPAND = (TVN_FIRST - 15)

' TreeView specific styles.
Private Const TVS_CHECKBOXES = &H100
Private Const TVS_DISABLEDRAGDROP = &H10
Private Const TVS_EDITLABELS = &H8
Private Const TVS_FULLROWSELECT = &H1000
Private Const TVS_HASBUTTONS = &H1
Private Const TVS_HASLINES = &H2
Private Const TVS_INFOTIP = &H800
Private Const TVS_LINESATROOT = &H4
Private Const TVS_NOSCROLL = &H2000
Private Const TVS_NOTOOLTIPS = &H80
Private Const TVS_SHOWSELALWAYS = &H20
Private Const TVS_SINGLEEXPAND = &H400
Private Const TVS_TRACKSELECT = &H200

' Notification messages.
Private Const NM_FIRST = 0
Private Const NM_CLICK = (NM_FIRST - 2)
Private Const NM_CUSTOMDRAW = (NM_FIRST - 12)
Private Const NM_DBLCLK = (NM_FIRST - 3)
Private Const NM_KILLFOCUS = (NM_FIRST - 8)
Private Const NM_RETURN = (NM_FIRST - 4)
Private Const NM_LDOWN = (NM_FIRST - 20)
Private Const NM_RDOWN = (NM_FIRST - 21)
Private Const NM_RCLICK = (NM_FIRST - 5)

' Inserting stuff.
Private Const TVI_ROOT = &HFFFF0000
Private Const TVI_FIRST = &HFFFF0001
Private Const TVI_LAST = &HFFFF0002
Private Const TVI_SORT = &HFFFF0003

' Mask values.
Private Const TVIF_CHILDREN = &H40
Private Const TVIF_DI_SETITEM = &H1000
Private Const TVIF_HANDLE = &H10
Private Const TVIF_IMAGE = &H2
Private Const TVIF_INTEGRAL = &H80
Private Const TVIF_PARAM = &H4
Private Const TVIF_SELECTEDIMAGE = &H20
Private Const TVIF_STATE = &H8
Private Const TVIF_TEXT = &H1

' More mask values, of the state kind.
Private Const TVIS_BOLD = &H10
Private Const TVIS_CUT = &H4
Private Const TVIS_DROPHILITED = &H8
Private Const TVIS_EXPANDED = &H20
Private Const TVIS_EXPANDEDONCE = &H40
Private Const TVIS_EXPANDPARTIAL = &H80
Private Const TVIS_OVERLAYMASK = &HF00
Private Const TVIS_SELECTED = &H2
Private Const TVIS_STATEIMAGEMASK = &HF000
Private Const TVIS_USERMASK = &HF000

' Expanding stuff.
Private Const TVE_COLLAPSE = &H1
Private Const TVE_COLLAPSERESET = &H8000
Private Const TVE_EXPAND = &H2
Private Const TVE_EXPANDPARTIAL = &H4000
Private Const TVE_TOGGLE = &H3

' ImageList type values.
Private Const TVSIL_NORMAL = 0
Private Const TVSIL_STATE = 2

Private Const TVGN_PARENT = &H3

Private Const WS_BORDER = &H800000
Private Const WS_CHILD = &H40000000
Private Const WS_DISABLED = &H8000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_TABSTOP = &H10000
Private Const WS_EX_CLIENTEDGE = &H200

Private Const WM_SETFOCUS = &H7
Private Const WM_SETREDRAW = &HB
Private Const WM_MOUSEACTIVATE = &H21
Private Const WM_NOTIFY = &H4E
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONUP = &H202
Private Const WM_TIMER = &H113
Private Const WM_VSCROLL = &H115
Private Const SB_LINEDOWN = 1
Private Const SB_LINEUP = 0

' ImageList Declarations
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50

Private Const ILC_MASK = &H1
Private Const ILC_COLOR = &H0
Private Const ILC_COLORDDB = &HFE
Private Const ILC_COLOR4 = &H4
Private Const ILC_COLOR8 = &H8
Private Const ILC_COLOR16 = &H10
Private Const ILC_COLOR24 = &H18
Private Const ILC_COLOR32 = &H20

Private Const ILD_BLEND25 = &H2
Private Const ILD_BLEND50 = &H4
Private Const ILD_MASK = &H10
Private Const ILD_NORMAL = &H0
Private Const ILD_FOCUS = ILD_BLEND25
Private Const ILD_SELECTED = ILD_BLEND50
Private Const ILD_TRANSPARENT = &H1

Private Const IMAGE_BITMAP = 0
Private Const LR_DEFAULTCOLOR = &H0
Private Const LR_CREATEDIBSECTION = &H2000
Private Const LR_LOADTRANSPARENT = &H20
Private Const LR_VGACOLOR = &H80
Private Const LR_LOADFROMFILE = &H10

Private Const CLR_NONE = &HFFFFFFFF
Private Const CLR_DEFAULT = &HFF000000

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const LANG_NEUTRAL = &H0
Private Const SUBLANG_DEFAULT = &H1

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type DWORD
    LOWORD As Integer
    HIWORD As Integer
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TVHITTESTINFO
    PT As POINTAPI
    Flags As Long
    hItem As Long
End Type

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    Code As Long
End Type

Private Type TVITEM
    mask As Long
    hItem As Long
    State As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
End Type

Private Type TVITEMEX
    mask As Long
    hItem As Long
    State As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    iSelectedImage As Long
    cChildren As Long
    lParam As Long
    iIntegral As Long
End Type

Private Type TVDISPINFO
    hdr As NMHDR
    Item As TVITEM
End Type

Private Type TVINSERTSTRUCT
    hParent As Long
    hInsertAfter As Long
    Item As TVITEMEX
End Type

Private Type NMTREEVIEW
    hdr As NMHDR
    action As Long
    itemOld As TVITEM
    itemNew As TVITEM
    ptDrag As POINTAPI
End Type

Private Type TVKEYDOWN
    hdr As NMHDR
    wVKey As Long
    Flags As Long
End Type

Private Type ICCEx
    dwSize As Long          ' size of this structure
    dwICC As Long           ' flags indicating which classes to be initialized
End Type

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (icc As ICCEx) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Sub ImageList_EndDrag Lib "comctl32.dll" ()
Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal hIml As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "comctl32.dll" (ByVal hIml As Long, ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_LoadImage Lib "comctl32.dll" (ByVal hi As Long, ByVal lpbmp As String, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal Flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_AddMasked Lib "comctl32.dll" (ByVal hIml As Long, ByVal hbmImage As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_BeginDrag Lib "comctl32.dll" (ByVal himlTrack As Long, ByVal iTrack As Long, ByVal dxHotspot As Long, ByVal dyHotspot As Long) As Long
Private Declare Function ImageList_DragEnter Lib "comctl32.dll" (ByVal hwndLock As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ImageList_DragMove Lib "comctl32.dll" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ImageList_DragShowNolock Lib "comctl32.dll" (ByVal fShow As Long) As Long
Private Declare Function ImageList_DragLeave Lib "comctl32.dll" (ByVal hwndLock As Long) As Long
Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal hIml As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long


'*******************************************************************
'*******************************************************************

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

Private Sub AddMsg(uMsg As WinSubHook.eMsg, When As WinSubHook.eMsgWhen)

    If When = WinSubHook.MSG_BEFORE Then

        Call AddMsgSub(uMsg, aMsgTblB, nMsgCntB, When)

    Else

        Call AddMsgSub(uMsg, aMsgTblA, nMsgCntA, When)

    End If

End Sub

Private Function CallOrigWndProc(ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long) As Long

    If hWndSubclass <> 0 Then

        CallOrigWndProc = WinSubHook.CallWindowProc(nWndProcOriginal, hWndSubclass, uMsg, wParam, lParam)

    Else

        Debug.Assert False

    End If

End Function

Private Sub DelMsg(uMsg As WinSubHook.eMsg, When As WinSubHook.eMsgWhen)

    If When = WinSubHook.MSG_BEFORE Then

        Call DelMsgSub(uMsg, aMsgTblB, nMsgCntB, When)

    Else

        Call DelMsgSub(uMsg, aMsgTblA, nMsgCntA, When)

    End If

End Sub

Private Sub Subclass(hWnd As Long, Owner As WinSubHook.iSubclass)

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

Private Sub UnSubclass()

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
    'not used in this
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)

    On Error Resume Next
    
    Dim retval As Long
    Dim tHDR As NMHDR
    Dim tvInsert As TVINSERTSTRUCT, TVDISPINFO As TVDISPINFO
    Dim PT As POINTAPI, Bool As Boolean, h As Long
    Dim rc As RECT
    Dim lLen As Long, iPos As Long
    Dim sText As String
    Dim TVHT As TVHITTESTINFO
    Dim mItem As TVITEM
    Dim TVK As TVKEYDOWN
    Dim lastkey As Long
    Dim ax As Long

    Select Case uMsg
    
    Case WM_NOTIFY
    
        CopyMemory tHDR, ByVal lParam, Len(tHDR)
        retval = 0
        
        Select Case tHDR.Code
        
        Case NM_CLICK, NM_RCLICK
            
            GetCursorPos TVHT.PT
            ScreenToClient hTree, TVHT.PT
            SendMessage hTree, TVM_HITTEST, 0, TVHT
            
            If TVHT.hItem <> 0 Then
                For ax = 0 To iNodes - 1
                    If TNodes(ax).hItem = TVHT.hItem Then
                        m_Item = ax
                        RaiseEvent ItemClick(IIf(uMsg = NM_CLICK, 1, 2), TNodes(ax).Text, TNodes(ax).Key, TNodes(ax).Index)
                    End If
                Next
            End If
            
            RaiseEvent Click(TVHT.PT.x, TVHT.PT.y)
        
        Case NM_DBLCLK
        
            GetCursorPos TVHT.PT
            ScreenToClient hTree, TVHT.PT
            SendMessage hTree, TVM_HITTEST, 0, TVHT
            
            If TVHT.hItem <> 0 Then
                For ax = 0 To iNodes - 1
                    If TNodes(ax).hItem = TVHT.hItem Then
                        m_Item = ax
                        RaiseEvent ItemDblClick(TNodes(ax).Text, TNodes(ax).Key, TNodes(ax).Index)
                    End If
                Next
            End If
            
            RaiseEvent DblClick(TVHT.PT.x, TVHT.PT.y)
        
        Case TVN_SELCHANGED
        
            GetCursorPos TVHT.PT
            ScreenToClient hTree, TVHT.PT
            SendMessage hTree, TVM_HITTEST, 0, TVHT
            
            If TVHT.hItem <> 0 Then
                For ax = 0 To iNodes - 1
                    If TNodes(ax).hItem = TVHT.hItem Then
                        m_Item = ax
                        RaiseEvent SelChanged(TNodes(ax).Text, TNodes(ax).Key, TNodes(ax).Index)
                    End If
                Next
            End If
        
        Case TVN_KEYDOWN
        
            CopyMemory TVK, ByVal lParam, Len(TVK)
            
            lastkey = TVK.wVKey
            
            mItem.hItem = TNodes(m_Item).hItem
            SendMessage hTree, TVM_GETITEM, 0, mItem
            
            SendMessage hTree, TVM_EXPAND, TVE_EXPAND, mItem
            
        End Select

    End Select
    
    bHandled = True
    
End Sub

Private Sub UserControl_Initialize()

    Call Create(UserControl.hWnd)

    Call Class_Initialize
    Call Subclass(hCont, Me)
    
    'add messages to subclass
    Call AddMsg(WM_NOTIFY, MSG_BEFORE)

    m_Border = 0

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        m_Border = PropBag.ReadProperty("Border", 0)
        UserControl.BorderStyle = m_Border
        Indent = PropBag.ReadProperty("Indent", 5)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Border", m_Border, 0
    PropBag.WriteProperty "Indent", Indent, 5
End Sub

Private Sub UserControl_Resize()

    Call MoveWindow(hCont, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 1)
    Call MoveWindow(hTree, 0, 3, UserControl.ScaleWidth, UserControl.ScaleHeight - 6, 1)
    
End Sub

Private Sub UserControl_Terminate()
    Call Class_Terminate
End Sub

Private Sub Create(ByVal hParent As Long)

    On Error Resume Next
    
    Dim rcl   As RECT
    Dim hIml  As Long
    
    Call InitCommonControls
    
    hCont = CreateWindowEx(0&, "STATIC", "bTreeViewClass", WS_VISIBLE Or WS_CHILD Or SS_WHITERECT, 0, 0, 250, 250, hParent, 0, App.hInstance, 0)
    hTree = CreateWindowEx(0&, "SysTreeView32", "", WS_VISIBLE Or WS_CHILD Or TVS_HASLINES Or TVS_HASBUTTONS Or TVS_LINESATROOT, 0, 0, 250, 250, hCont, ID_TREEVIEW, App.hInstance, 0)

    If hTree = 0 Then
        MsgBox "TreeView CreateWindow Failed"
        Exit Sub
    End If
    
    ReDim Preserve TNodes(0)
    iNodes = 0
    
    SendMessageLong hTree, TVM_SETBKCOLOR, 0, &HFFFFFF ' property BackColor
    SendMessageLong hTree, TVM_SETTEXTCOLOR, 0, &H0 ' property FontColor
    SendMessage hTree, TVM_SETINDENT, 16, 0 ' property Indent
    SendMessage hTree, TVM_SETITEMHEIGHT, 16, 0

    Dim TVIN As TVINSERTSTRUCT, mRoot As Long, mParent As Long, i As Byte
    
    ' design time items
    If TNodes(0).hItem = 0 Then
    
        bInitialized = True
    
        TVIN.hParent = TVI_ROOT
        TVIN.hInsertAfter = TVI_FIRST
        TVIN.Item.pszText = "Root Item" & Chr(0)
        TVIN.Item.cchTextMax = 10
        TVIN.Item.mask = TVIF_TEXT
        mRoot = SendMessage(hTree, TVM_INSERTITEM, 0, TVIN)
        TVIN.hParent = mRoot
        TVIN.Item.pszText = "Parent Item" & Chr(0)
        TVIN.Item.cchTextMax = 12
        mParent = SendMessage(hTree, TVM_INSERTITEM, 0, TVIN)
        SendMessage hTree, TVM_EXPAND, ByVal TVE_EXPAND, ByVal mRoot
        
        For i = 1 To 2
            TVIN.hParent = mParent
            TVIN.Item.pszText = "Child Item" & Chr(0)
            TVIN.Item.cchTextMax = 11
            SendMessage hTree, TVM_INSERTITEM, 0, TVIN
        Next i
        SendMessage hTree, TVM_EXPAND, ByVal TVE_EXPAND, ByVal mParent
        
    End If

    'OldProc = GetWindowLong(hCont, GWL_WNDPROC)
    'ret = SetWindowLong(hCont, GWL_WNDPROC, AddressOf TvwProc)

End Sub

Public Function AddMasked(ByVal hIml As Long, ByVal hImage As Long, ByVal cMask As Long) As Long
    Dim lRet As Long
    lRet = ImageList_AddMasked(hIml, hImage, cMask)
    AddMasked = IIf(lRet <> -1, hIml, -1)
End Function

Public Function AddIcon(ByVal hIml As Long, ByVal hIcon As Long) As Long
    AddIcon = ImageList_ReplaceIcon(hIml, -1, hIcon)
End Function

Public Function CreateList(ByVal lColor As Long, ByVal InitialImages As Long, Optional ByVal TotalImages As Variant) As Long
    
    Dim cxSmIcon As Long, cySmIcon As Long
    Dim IccInit As Boolean
    Dim m_hIml As Long
    
    If IccInit = False Then InitCommonControls
      
    If IsNull(TotalImages) = True Then TotalImages = 0
    If lColor = -1 Then lColor = ILC_COLOR24 Or ILC_MASK
       
    cxSmIcon = GetSystemMetrics(SM_CXSMICON)
    cySmIcon = GetSystemMetrics(SM_CYSMICON)
    m_hIml = ImageList_Create(cxSmIcon, cySmIcon, lColor, InitialImages, TotalImages)
    CreateList = m_hIml
    
End Function

Public Function LoadList(ByVal m_Image As String, ByVal m_Color As Long, ByVal m_Res As Long, ByVal m_Width As Long, ByVal m_Height As Long) As Long
    
    Dim hBitmap As Long, hTemp As Long
    Dim m_Images As Long
    
    m_Images = m_Width / m_Height
    
    hBitmap = LoadImage(App.hInstance, m_Image, IMAGE_BITMAP, m_Width, m_Height, IIf(m_Res = True, LR_CREATEDIBSECTION, LR_LOADFROMFILE))
    hTemp = CreateList(ILC_COLOR24 Or ILC_MASK, m_Images, m_Images)
    LoadList = AddMasked(hTemp, hBitmap, m_Color)
    
End Function

Public Sub AddItem(hRelItem, Relation As RelationConstants, Key As String, Text As String, Optional Image As Long = -1, Optional SelectedImage As Long = -1)

    Dim TVIN As TVINSERTSTRUCT, hRel As Long, TVI As TVITEMEX
    Dim ax As Long
    
    If Relation = 0 Then Relation = tvwLast
    If hRelItem = 0 Then hRelItem = 0&

    If bInitialized = True Then
        Call DeleteAllItems
        bInitialized = False
    End If

    If TypeName(hRelItem) = "Long" Then
      
        If hRelItem = 0 Then hRelItem = 0&
  
        hRel = hRelItem

    ElseIf TypeName(hRelItem) = "String" Then
        For ax = 0 To iNodes - 1
            If TNodes(ax).Key = hRelItem Then
                hRel = TNodes(ax).hItem
                Exit For
            Else
                hRel = 0&
            End If
        Next
    End If
    
    TVIN.hParent = hRel
    
    If Image > 0 Then
    
        TVIN.Item.mask = TVIN.Item.mask Or TVIF_IMAGE
        If SelectedImage < 0 Then
            SelectedImage = Image
            TVIN.Item.mask = TVIN.Item.mask Or TVIF_SELECTEDIMAGE
        End If
        
    End If
    
    If SelectedImage > 0 Then
        TVIN.Item.mask = TVIN.Item.mask Or TVIF_SELECTEDIMAGE
    End If
    
    TVIN.Item.mask = TVIN.Item.mask Or TVIF_STATE Or TVIF_TEXT
    TVIN.Item.pszText = Text & Chr(0)
    TVIN.Item.cchTextMax = Len(Text) + 1
    
    If Image >= 0 Then
        TVIN.Item.iImage = Image
    End If
    
    If SelectedImage >= 0 Then
        TVIN.Item.iSelectedImage = SelectedImage
    End If
    
    'TVIN.Item.stateMask = TVIS_BOLD
    'TVIN.Item.state = TVIS_BOLD
    
    If Relation = tvwSort Then
        TVIN.hInsertAfter = TVI_SORT
    ElseIf Relation = tvwFirst Then
        TVIN.hInsertAfter = TVI_FIRST
    ElseIf Relation = tvwLast Then
        TVIN.hInsertAfter = TVI_LAST
    ElseIf Relation = tvwChild Then
        TVIN.hParent = SendMessageLong(hTree, TVM_GETNEXTITEM, TVGN_PARENT, hRel)
        TVIN.hInsertAfter = hRel
    End If
    
    hRel = SendMessage(hTree, TVM_INSERTITEM, 0, TVIN)
    
    If hRel <> 0 Then
        
        SendMessage hTree, TVM_GETITEM, hRel, TVI
        TVI.mask = TVIF_PARAM
        TVI.lParam = hRel
        SendMessage hTree, TVM_SETITEM, hRel, TVI
        
        ReDim Preserve TNodes(iNodes)
        
        TNodes(iNodes).hItem = hRel
        TNodes(iNodes).hParent = SendMessageLong(hTree, TVM_GETNEXTITEM, TVGN_PARENT, hRel)
        TNodes(iNodes).Index = iNodes
        TNodes(iNodes).Text = Text
        TNodes(iNodes).Key = Key
        TNodes(iNodes).Image = Image
        
        iNodes = iNodes + 1

        'If DoSort(hRel) Then
        '    SendMessageL hTree, TVM_SORTCHILDREN, 0, hRel
        'End If
    End If
    
    SendMessage hTree, TVM_EXPAND, ByVal TVE_EXPAND, ByVal TNodes(iNodes - 1).hParent
    
End Sub

Public Sub DeleteItem(ByVal Index As Variant)

    On Error Resume Next

    Dim ax As Long
    Dim dIndex As Long
    
    If TypeName(Index) = "Long" Then
      
        dIndex = Index

    ElseIf TypeName(Index) = "String" Then
        For ax = 0 To iNodes - 1
            If TNodes(ax).Key = Index Then
                dIndex = TNodes(ax).Index
                Exit For
            Else
                dIndex = -1
            End If
        Next
    End If

    Call SendMessage(hTree, TVM_DELETEITEM, 0, ByVal TNodes(dIndex).hItem)
    Call SendMessage(hTree, TVM_SELECTITEM, TVGN_CARET, ByVal TNodes(dIndex).hParent)
    
    TNodes(dIndex).hItem = TNodes(dIndex).hItem
    TNodes(dIndex).hParent = TNodes(dIndex).hParent
    TNodes(dIndex).Image = TNodes(dIndex).Image
    TNodes(dIndex).Index = TNodes(dIndex).Index
    TNodes(dIndex).Key = ""
    TNodes(dIndex).Tag = ""
    TNodes(dIndex).Text = ""
    
    TNodes(dIndex).hItem = -1
    TNodes(dIndex).hParent = -1
    TNodes(dIndex).Image = -1
    TNodes(dIndex).Index = -1
    TNodes(dIndex).Key = ""
    TNodes(dIndex).Tag = ""
    TNodes(dIndex).Text = ""
    
    iNodes = iNodes - 1
    
End Sub

Public Sub RenameItem(ByVal Index As Variant, ByVal Text As String)

    On Error Resume Next

    Dim ax As Long
    Dim dIndex As Long
    Dim utvi As TVITEM
    
    If TypeName(Index) = "Long" Then
      
        dIndex = Index

    ElseIf TypeName(Index) = "String" Then
        For ax = 0 To iNodes - 1
            If TNodes(ax).Key = Index Then
                dIndex = TNodes(ax).Index
                Exit For
            Else
                dIndex = -1
            End If
        Next
    End If

    If Len(Text) <> 0 Then
        
        TNodes(dIndex).Text = Text
        
        utvi.hItem = TNodes(dIndex).hItem
        utvi.mask = TVIF_TEXT
        utvi.cchTextMax = Len(Text)
        utvi.pszText = Text & Chr(0)
        Call SendMessage(hTree, TVM_SETITEM, 0, utvi)
        
    End If
    
End Sub

Public Sub DeleteAllItems()
    Call SendMessage(hTree, TVM_DELETEITEM, 0, ByVal TVI_ROOT)
    iNodes = 0
End Sub

Public Sub SetItemTag(ByVal Index As Variant, ByVal Value As Variant)
    
    On Error Resume Next

    Dim ax As Long
    Dim dIndex As Long
    
    If TypeName(Index) = "Long" Then
      
        dIndex = Index

    ElseIf TypeName(Index) = "String" Then
        For ax = 0 To iNodes - 1
            If TNodes(ax).Key = Index Then
                dIndex = TNodes(ax).Index
                Exit For
            Else
                dIndex = -1
            End If
        Next
    End If
    
    TNodes(dIndex).Tag = Value
    
End Sub

Public Function GetItemTag(ByVal Index As Variant) As Variant
    
    On Error Resume Next

    Dim ax As Long
    Dim dIndex As Long
    
    If TypeName(Index) = "Long" Then
      
        dIndex = Index

    ElseIf TypeName(Index) = "String" Then
        For ax = 0 To iNodes - 1
            If TNodes(ax).Key = Index Then
                dIndex = TNodes(ax).Index
                Exit For
            Else
                dIndex = -1
            End If
        Next
    End If
    
    GetItemTag = TNodes(dIndex).Tag
    
End Function

Public Function SetItemNode(ByVal Index As Variant, ByVal arrValue) As Variant
    
    On Error Resume Next

    Dim ax As Long
    Dim dIndex As Long
    
    If TypeName(Index) = "Long" Then
      
        dIndex = Index

    ElseIf TypeName(Index) = "String" Then
        For ax = 0 To iNodes - 1
            If TNodes(ax).Key = Index Then
                dIndex = TNodes(ax).Index
                Exit For
            Else
                dIndex = -1
            End If
        Next
    End If
    
    TNodes(dIndex).Index = arrValue(0)
    TNodes(dIndex).Image = arrValue(1)
    TNodes(dIndex).Key = arrValue(2)
    TNodes(dIndex).Text = arrValue(3)
    TNodes(dIndex).Tag = arrValue(4)
    
End Function

Public Function GetItemNode(ByVal Index As Variant) As Variant
    
    On Error Resume Next

    Dim ax As Long
    Dim dIndex As Long
    Dim arrGetItemNode(4)
    
    If TypeName(Index) = "Long" Then
      
        dIndex = Index

    ElseIf TypeName(Index) = "String" Then
        For ax = 0 To iNodes - 1
            If TNodes(ax).Key = Index Then
                dIndex = TNodes(ax).Index
                Exit For
            Else
                dIndex = -1
            End If
        Next
    End If
    
    arrGetItemNode(0) = TNodes(dIndex).Index
    arrGetItemNode(1) = TNodes(dIndex).Image
    arrGetItemNode(2) = TNodes(dIndex).Key
    arrGetItemNode(3) = TNodes(dIndex).Text
    arrGetItemNode(4) = TNodes(dIndex).Tag
    
    GetItemNode = arrGetItemNode
    
End Function

Public Sub Expand(ByVal Index As Long, ByVal Flags As Long)
    Call SendMessage(hTree, TVM_EXPAND, ByVal Flags, ByVal TNodes(Index).hItem)
End Sub

Private Sub SetStyle(ByVal hWnd As Long, ByVal Value As Long, ByVal Bool As Boolean)
    If hWnd <> 0 Then
        If Bool Then
            SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) Or Value
        Else
            SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) And (Not Value)
        End If
        'SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
    End If
End Sub

Public Property Get Border() As eBorder
    Border = m_Border
End Property

Public Property Let Border(ByVal NewValue As eBorder)
    
    m_Border = NewValue
    
    UserControl.BorderStyle = m_Border
    
    UserControl.PropertyChanged "Border"
    UserControl.Refresh
    
End Property

Public Property Let hImageList(ByVal hImg As Long)
    m_hImageList = hImg
    SendMessageLong hTree, TVM_SETIMAGELIST, TVSIL_NORMAL, m_hImageList
End Property

Public Property Get hImageList() As Long
    'm_hImageList = SendMessage(hTree, TVM_GETIMAGELIST, TVSIL_NORMAL, ByVal 0&)
    hImageList = m_hImageList
End Property

Public Property Get Indent() As Long
    Indent = SendMessage(hTree, TVM_GETINDENT, 0, ByVal 0&)
End Property

Public Property Let Indent(ByVal Value As Long)
    Call SendMessage(hTree, TVM_SETINDENT, Value, ByVal 0&)
End Property

Public Property Get ItemCount() As Long
    ItemCount = iNodes
End Property

Public Property Get ItemKey() As String
    ItemKey = TNodes(m_Item).Key
End Property

Public Property Get ItemIndex() As Long
    ItemIndex = TNodes(m_Item).Index
End Property

Public Property Get ItemTag() As Variant
    ItemTag = TNodes(m_Item).Tag
End Property

Public Property Let ItemTag(ByVal Value As Variant)
    TNodes(m_Item).Tag = Value
End Property

Public Property Get ItemText() As String
    ItemText = TNodes(m_Item).Text
End Property

Public Property Let ItemText(Text As String)

    Dim utvi As TVITEM

    'utvi.hItem = TNodes(m_Item).hItem
    'SendMessage hTree, TVM_GETITEM, 0, utvi

    If Len(Text) <> 0 Then
        
        TNodes(m_Item).Text = Text
        
        utvi.hItem = TNodes(m_Item).hItem
        utvi.mask = TVIF_TEXT
        utvi.cchTextMax = Len(Text)
        utvi.pszText = Text & Chr(0)
        Call SendMessage(hTree, TVM_SETITEM, 0, utvi)
        
    End If

End Property
