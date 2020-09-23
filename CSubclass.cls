VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSubclass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------------
'cSubclass - module-less, IDE safe, machine code window subclasser
'
'v1.00 20030107 First cut..........................................................................
'

Option Explicit

Private Const OFFSET_P1   As Long = 9                     'Callback gate address
Private Const OFFSET_P2   As Long = 22                    'Before table entry count
Private Const OFFSET_P3   As Long = 37                    'Before table address
Private Const OFFSET_P4   As Long = 51                    'In IDE ?
Private Const OFFSET_P5   As Long = 69                    'Owner object address for iSubclass_Before
Private Const OFFSET_P6   As Long = 141                   'Original WndProc address
Private Const OFFSET_P7   As Long = 146                   'CallWindowProc address
Private Const OFFSET_P8   As Long = 154                   'After table entry count
Private Const OFFSET_P9   As Long = 169                   'After table address
Private Const OFFSET_PA   As Long = 183                   'In IDE?
Private Const OFFSET_PB   As Long = 201                   'Owner object address for iSubclass_After
Private Const OFFSET_PC   As Long = 250                   'Original WndProc address
Private Const OFFSET_PD   As Long = 260                   'SetWindowLong address
Private Const ARRAY_LB    As Long = 1                     'Low bound of arrays

Private Type tCode
  Buf(ARRAY_LB To 272)    As Byte                         'Code buffer
  'Buf(ARRAY_LB To 3012)    As Byte                         'Code buffer
End Type

Private Type tCodeBuf
  Code                    As tCode                        'Subclass WndProc code
End Type

Private CodeBuf           As tCodeBuf                     'Subclass WndProc code instance
Private nBreakGate        As Long                         'Callback breakpoint gate
Private nMsgCntB          As Long                         'Before msg table entry count
Private nMsgCntA          As Long                         'After msg table entry count
Private aMsgTblB()        As eMsg              'Before msg table array
Private aMsgTblA()        As eMsg              'After msg table array
Private hWndSubclass      As Long                         'Handle of the window being subclassed
Private nWndProcSubclass  As Long                         'The address of our WndProc
Private nWndProcOriginal  As Long                         'The address of the existing WndProc

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Const GWL_WNDPROC As Long = (-4)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd As Long, ByVal fAccept As Long)
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_ADDSTRING As Long = &H180
Private Const LB_FINDSTRINGEXACT As Long = &H1A2

Private Function TrimNull(StrIn As String) As String

   Dim nul As Long

   'Truncates the input string at first null. If no nulls, perform ordinary Trim.
   nul = VBA.InStr(StrIn, vbNullChar)
   Select Case nul
   Case Is > 1
      TrimNull = VBA.Left$(StrIn, nul - 1)
   Case 1
      TrimNull = ""
   Case 0
      TrimNull = VBA.Trim$(StrIn)
   End Select

End Function

Private Function ListAddDisinctItem(ByVal hWndList As Long, sItemText As String) As Long

   If SendMessage(hWndList, LB_FINDSTRINGEXACT, ByVal -1, ByVal sItemText) = -1 Then
      ListAddDisinctItem = SendMessage(hWndList, LB_ADDSTRING, ByVal 0&, ByVal sItemText)
   Else
      ListAddDisinctItem = -1
   End If
End Function


'Private m_lDragListMessage As Long
'Private m_objDragList As VB.ListBox
'Private m_IDragListBox As IDragListBox
'Private m_colHotKey As Collection
'Private m_colSysMenu As CollectionEx
'-----------------------------
'Class creation/destruction

'Called automatically when the class instance is created.
Private Sub Class_Initialize()
  Dim i     As Long
  Dim j     As Long
  Dim nIDE  As Long
  Const OPS As String = _
"558BEC83C4F85756BE_patch1_33C08945FC8945F8B90000000083F900746183F9FF740CBF000000008B450CF2AF755033C03D_patch4_740B833E007542C70601000000BA_patch5_8B0283F8000F84A50000008D4514508D4510508D450C508D4508508D45FC508D45F8508B0252FF5020C706000000008B45F883F8007570FF7514FF7510FF750CFF750868_patch6_E8_patch7_8945FCB90000000083F900744D83F9FF740CBF000000008B450CF2AF753C33C03D_patchA_740B833E00752EC70601000000BA_patchB_8B0283F8007425FF7514FF7510FF750CFF75088D45FC508B0252FF501CC706000000005E5F8B45FCC9C2100068_patchC_6AFCFF7508E8_patchD_33C08945FCEBE190"

  'Convert the string from hexadecimal pairs to bytes and store in the opcode buffer
  With CodeBuf.Code
    j = 1                                                 'Set the character index to the start of the opcode string
    For i = ARRAY_LB To UBound(.Buf)                      'For each byte of the code buffer
      .Buf(i) = Val("&H" & Mid$(OPS, j, 2))               'Pull a pair of hex characters and convert to a byte
      j = j + 2                                           'Bump to the next pair of characters
    Next i                                                'Next byte of the code buffer
    
    nWndProcSubclass = VarPtr(.Buf(ARRAY_LB))             'Address of the cSubclass WndProc entry point
  End With
  
  nIDE = InIDE                                            'Determine whether we're running in the IDE or not
  
  'Patch the WndProc code with runtime values
  Call PatchVal(OFFSET_P1, VarPtr(nBreakGate))            'Breakpoint gate address
  Call PatchVal(OFFSET_P4, nIDE)                          'Whether we need check the breakpoint gate and the vtable
  Call PatchRel(OFFSET_P7, Addr("CallWindowProcA"))       'Address of the CallWindowProc api function
  Call PatchVal(OFFSET_PA, nIDE)                          'Whether we need check the breakpoint gate and the vtable
  Call PatchRel(OFFSET_PD, Addr("SetWindowLongA"))        'Address of the SetWindowLong api function
  'Call InitializeMessages
End Sub


Public Property Get hWnd() As Long
   hWnd = hWndSubclass
End Property

'Called automatically when the class instance is destroyed.
Private Sub Class_Terminate()
  If hWndSubclass <> 0 Then                               'If the Subclass thunk is active
    
    Call EndDropFiles(hWndSubclass)
    'Call DestroyHotKey(hWndSubclass)
    'Set m_objDragList = Nothing
    'Set m_colSysMenu = Nothing
    'Call HideSysTrayIcon
    
    Call UnSubclass                                       'UnSubclass
  End If
End Sub

'-----------------------------
'Public interface

'Call this method to add a message to the msg callback table. NB This method can be called at any time
Public Sub AddMsg(uMsg As eMsg, When As eMsgWhen)
  If When = MSG_BEFORE Then                    'If before
    Call AddMsgSub(uMsg, aMsgTblB, nMsgCntB, When)        'Add the message, pass the before table and before table message count variables
  Else                                                    'Else after
    Call AddMsgSub(uMsg, aMsgTblA, nMsgCntA, When)        'Add the message, pass the after table and after table message count variables
  End If
End Sub

'Call the original WndProc
Public Function CallOrigWndProc(ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long) As Long
  If hWndSubclass <> 0 Then
    CallOrigWndProc = CallWindowProc( _
    nWndProcOriginal, hWndSubclass, uMsg, wParam, lParam) 'Call the original WndProc
  Else
    Debug.Assert False                                    'LOGIC ERROR: The subclasser isn't active!
  End If
End Function

'Call this method to delete a message from the msg table. NB This method can be called at any time
Public Sub DelMsg(uMsg As eMsg, When As eMsgWhen)
  If When = MSG_BEFORE Then                    'If before
    Call DelMsgSub(uMsg, aMsgTblB, nMsgCntB, When)        'Delete the message, pass the before table and before message count variables
  Else                                                    'Else after
    Call DelMsgSub(uMsg, aMsgTblA, nMsgCntA, When)        'Delete the message, pass the after table and after message count variables
  End If
End Sub

'Call this method to subclass the passed window handle
'Call this method to subclass the passed window handle
Public Sub SubclassEx(hWnd As Long, OwnerObjPtr As Long, Optional bDummy As Boolean)
   Call pSubclass(hWnd, OwnerObjPtr)
End Sub
Public Sub Subclass(hWnd As Long, Owner As ISubclass, Optional bDummy As Boolean)
   Call pSubclass(hWnd, ObjPtr(Owner))
End Sub
Public Sub Subclass2(hWnd As Long, Owner As ISubclass2, Optional bDummy As Boolean)
   Call pSubclass(hWnd, ObjPtr(Owner))
End Sub
'Public Sub Subclass3(hWnd As Long, Owner As ISubclass3, Optional bDummy As Boolean)
'   Call pSubclass(hWnd, ObjPtr(Owner))
'End Sub
'Public Sub Subclass4(hWnd As Long, Owner As ISubclass4, Optional bDummy As Boolean)
'   Call pSubclass(hWnd, ObjPtr(Owner))
'End Sub
'Public Sub Subclass5(hWnd As Long, Owner As ISubclass5, Optional bDummy As Boolean)
'   Call pSubclass(hWnd, ObjPtr(Owner))
'End Sub
Private Sub pSubclass(hWnd As Long, OwnerPtr As Long)
  Debug.Assert (hWndSubclass = 0)                         'LOGIC ERROR: The subclasser is already active!
  Debug.Assert hWnd                                       'LOGIC ERROR: Invalid window handle
  Debug.Assert IsWindow(hWnd)                             'LOGIC ERROR: Invalid window handle
  
  hWndSubclass = hWnd                                     'Store the window handle
  nWndProcOriginal = SetWindowLong( _
                                  hWnd, _
                                  GWL_WNDPROC, _
                                  nWndProcSubclass)       'Set our WndProc in place of the original
  Debug.Assert nWndProcOriginal                           '??? You can't subclass a window outside of the current process
  
  Call PatchVal(OFFSET_P5, OwnerPtr)                 'Owner object address for iSubclass_Before
  Call PatchVal(OFFSET_P6, nWndProcOriginal)              'Original WndProc address for CallWindowProc
  Call PatchVal(OFFSET_PB, OwnerPtr)                 'Owner object address for iSubclass_After
  Call PatchVal(OFFSET_PC, nWndProcOriginal)              'Original WndProc address for SetWindowLong
End Sub
'Call this method to stop subclassing the window
Public Sub UnSubclass()
  If hWndSubclass <> 0 Then
    Call PatchVal(OFFSET_P2, 0)                           'Patch the code to ensure no further iSubclass_Before callbacks
    Call PatchVal(OFFSET_P8, 0)                           'Patch the code to ensure no further iSubclass_After callbacks
    Call SetWindowLong(hWndSubclass, _
                                  GWL_WNDPROC, _
                                  nWndProcOriginal)       'Restore the original WndProc
    hWndSubclass = 0                                      'Indicate the subclasser is inactive
    nMsgCntB = 0                                          'Message before count equals zero
    nMsgCntA = 0                                          'Message after count equals zero
  End If
End Sub

'-----------------------------
' Private subroutines

'Worker sub for AddMsg
Private Sub AddMsgSub(uMsg As eMsg, aMsgTbl() As eMsg, nMsgCnt As Long, When As eMsgWhen)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                  'If ALL_MESSAGES
    nMsgCnt = -1                                          'Indicates that all messages are to callback
  Else                                                    'Else a specific message number
    For nEntry = ARRAY_LB To nMsgCnt                      'For each existing entry. NB will skip if nMsgCnt = 0 or -1 (ALL_MESSAGES)
      Select Case aMsgTbl(nEntry)                         'Select on the message number stored in this table entry
      Case -1                                             'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                            'Re-use this entry
        Exit Sub                                          'Bail
      Case uMsg                                           'The msg is already in the table!
        Exit Sub                                          'Bail
      End Select
    Next nEntry                                           'Next entry
    
    'Make space for the new entry
    ReDim Preserve aMsgTbl(ARRAY_LB To nEntry)            'Increase the size of the table. NB nEntry = nMsgCnt + 1
    nMsgCnt = nEntry                                      'Bump the entry count
    aMsgTbl(nEntry) = uMsg                                'Store the message number in the table
  End If
  
  If When = MSG_BEFORE Then                    'If before
    Call PatchVal(OFFSET_P2, nMsgCnt)                     'Patch the Before table entry count
    Call PatchVal(OFFSET_P3, AddrMsgTbl(aMsgTbl))         'Patch the address of the Before message table. We need do this because there's no guarantee that the table existed at SubClass time, the table only gets created if a specific message number is added.
  Else                                                    'Else after
    Call PatchVal(OFFSET_P8, nMsgCnt)                     'Patch the After table entry count
    Call PatchVal(OFFSET_P9, AddrMsgTbl(aMsgTbl))         'Patch the address of the After message table. We need do this because there's no guarantee that the table existed at SubClass time, the table only gets created if a specific message number is added.
  End If
End Sub

'Worker sub for DelMsg
Private Sub DelMsgSub(uMsg As eMsg, aMsgTbl() As eMsg, nMsgCnt As Long, When As eMsgWhen)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                  'If deleting all messages (specific or ALL_MESSAGES)
    nMsgCnt = 0                                           'Message count is now zero
    If When = MSG_BEFORE Then                  'If before
      Call PatchVal(OFFSET_P2, 0)                         'Patch the before table message count
    Else                                                  'Else after
      Call PatchVal(OFFSET_P8, 0)                         'Patch the after table message count
    End If
  Else                                                    'Else deleteting a specific message
    For nEntry = ARRAY_LB To nMsgCnt                      'For each table entry
      If aMsgTbl(nEntry) = uMsg Then                      'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = -1                              'Mark the table slot as available
        Exit For                                          'Bail
      End If
    Next nEntry                                           'Next entry
  End If
End Sub

'Return the address of the passed function
Private Function Addr(sProc As String) As Long
  Addr = GetProcAddress(GetModuleHandle("user32"), sProc)
End Function

'Return the address of the low bound of the passed table array
Private Function AddrMsgTbl(aMsgTbl() As eMsg) As Long
  On Error Resume Next                                    'The table may not be dimensioned yet so we need protection
    AddrMsgTbl = VarPtr(aMsgTbl(ARRAY_LB))                'Get the address of the first element of the passed message table
  On Error GoTo 0                                         'Switch off error protection
End Function

'Patch the code offset with the passed value
Private Sub PatchVal(nOffset As Long, nValue As Long)
  Call CopyMemory(ByVal (nWndProcSubclass + nOffset), nValue, 4)
End Sub

'Patch the code offset with the relative address to the target address
Private Sub PatchRel(nOffset As Long, nTargetAddr As Long)
  Call CopyMemory(ByVal (nWndProcSubclass + nOffset), nTargetAddr - nWndProcSubclass - nOffset - 4, 4)
End Sub

'Return -1 if we're running in the IDE or 0 if were running compiled.
Private Function InIDE() As Long
  Static Value As Long
  
  If Value = 0 Then
    Value = 1
    Debug.Assert True Or InIDE()                          'This line won't exist in the compiled app
    InIDE = Value - 1
  End If

  Value = 0
End Function


'================ Subclass Functions ===============================
'Public Function OnMouseMove(HWnd As Long, _
'                        uMsg As Long, _
'                        wParam As Long, _
'                        lParam As Long) As MSG_MOUSEMOVE
'   '********************** WM_MOUSEMOVE ***********************
'   Static mbMouseOver As Boolean
'   Dim tme As TRACKMOUSEEVENTINFO
'
'   Select Case uMsg
'   Case WM_MOUSEMOVE
'      With OnMouseMove
'         .HWnd = HWnd
'         .Point = MAKEPOINT(lParam)
'         .Shift = GetShiftState()
'         .IsMouseEnter = Not mbMouseOver
'      End With
'      If Not mbMouseOver Then
'         mbMouseOver = True
'         With tme
'            .cbSize = LenB(tme)
'            .dwFlags = TME_HOVER Or TME_LEAVE
'            .hwndTrack = HWnd
'            .dwHoverTime = HOVER_DEFAULT
'         End With
'         Call TrackMouseEvent(tme)
'      End If
'
'   Case WM_MOUSELEAVE
'      mbMouseOver = False
'      With OnMouseMove
'         .IsMouseLeave = True
'         Call GetCursorPos(.Point)  'screen coordinates.
'         .HWnd = WindowFromPoint(.Point.x, .Point.y)
'         .Shift = GetShiftState()
'      End With
'
'   Case WM_MOUSEHOVER
'      With OnMouseMove
'         .IsMouseEnter = Not mbMouseOver
'         .Point = MAKEPOINT(GetMessagePos()) 'client coordinates.
'         .Shift = GetShiftState()
'         .HWnd = HWnd
'      End With
'      With tme
'         .cbSize = LenB(tme)
'         .dwFlags = TME_HOVER Or TME_LEAVE
'         .hwndTrack = HWnd
'         .dwHoverTime = HOVER_DEFAULT
'      End With
'      Call TrackMouseEvent(tme)
'   End Select
'End Function


'====== DragListBox: DRAGLISTMSGSTRING ==================================
'Private Sub InitializeMessages()
'   m_lDragListMessage = RegisterWindowMessage(DRAGLISTMSGSTRING)
'End Sub
'
'Public Property Get WM_DRAGLIST() As Long
'   WM_DRAGLIST = m_lDragListMessage
'End Property
'Public Sub AddMsgDragList(When As eMsgWhen, ByVal ListBox As ListBox, ByVal IDragListBox As IDragListBox, Optional DraggingIndicatorLineHeight As Long = 1)
'   Set m_objDragList = ListBox
'   Set m_IDragListBox = IDragListBox
'   If ObjPtr(m_objDragList) <> 0 And ObjPtr(m_IDragListBox) <> 0 Then
'      With ListBox
'         'Grow listbox item height for good visual effect during dragging
'         ListGrowItemHeight .HWnd, DraggingIndicatorLineHeight
'         'Make the listbox a drag listbox
'         Call MakeDragList(.HWnd)
'      End With
'      AddMsg m_lDragListMessage, When
'   End If
'End Sub
'
'Private Sub ListGrowItemHeight(ListBoxhWnd As Long, Optional dy As Long = 1)
'   'Grow item height bye one
'   Dim nItemHeigth As Long
'   nItemHeigth = SendMessageList(ListBoxhWnd, LB_GETITEMHEIGHT, 0&, ByVal 0&)
'   If nItemHeigth <> LB_ERR Then
'     Call SendMessageList(ListBoxhWnd, LB_SETITEMHEIGHT, 0&, ByVal nItemHeigth + dy)
'   End If
'End Sub
'
'Public Function OnDragListBox(lParam As Long) As Long
'
''Processes draglist box messages
'
'   Static nIdxDragStartItem As Long
'   Static nIdxPrevDragging As Long
'
'   Dim lpDragListInfo As DRAGLISTINFO
'   Dim nIdxDrop As Long 'Item index at which the user dropped the dragging item.
'   Dim hDC As Long 'Device context for drawing indicating line.
'   Dim nIdxCursor As Long 'Item under the cursor during dragging.
'   Dim rcCursor As RECT, rcPrev As RECT 'Current and previous items' rect
'
'   'Copy draglist info structure from the pointer (lParam)
'   CopyMemory lpDragListInfo, ByVal lParam, Len(lpDragListInfo)
'
'   Select Case lpDragListInfo.uNotification
'   Case DL_BEGINDRAG 'The drag operation starts. Return False to cancel
'      'Get the selected item
'      With lpDragListInfo
'         nIdxDragStartItem = LBItemFromPt(.HWnd, .ptCursor.x, .ptCursor.y, False)
'      End With 'LPDRAGLISTINFO
'      'Continue with the drag
'      OnDragListBox = 1
'
'   Case DL_CANCELDRAG  'The drag was canceled
'      With m_objDragList
'         .Refresh
'         .MousePointer = vbDefault
'         .Parent.Cls
'      End With
'      'Stop the drag
'      OnDragListBox = 0
'
'   Case DL_DRAGGING  'The item is being dragged
'      'Draw the insert icon
'      With lpDragListInfo
'
'         Dim bDrawInsertIcon As Boolean
'         Dim lLineColor As eColor
'         'Call DragListDrawInsertIcon(hWndSubclass, .hWnd, nIdxCursor)
'         lLineColor = clrGold
'         Call m_IDragListBox.OnDraggingSetIcon(hWndSubclass, .HWnd, bDrawInsertIcon, lLineColor)
'
'         'Get the index of the item under the cursor.
'         nIdxCursor = LBItemFromPt(.HWnd, .ptCursor.x, .ptCursor.y, True)
'         'RaiseEvent Dragging(nIdxDragStartItem, nIdxCursor)
'
'         'Get the rect of the previous and current items.
'         'These rects are used to draw indicating line.
'         Call SendMessage(.HWnd, LB_GETITEMRECT, nIdxPrevDragging, rcPrev)
'         Call SendMessage(.HWnd, LB_GETITEMRECT, nIdxCursor, rcCursor)
'
'         'Get the device context of the list box.
'         hDC = GetDC(.HWnd)
'         'Erase the indicating line for the previous item.
'         'Top should be subtracted by one for good visual.
'         'For this purpose, we have enlarged itemheight of the listbox by one,
'         'when the app starts.
'         With rcPrev
'            Call DrawLineEx(hDC, .Left, .Top - 1, .Right, .Top - 1, PS_SOLID, 2, TranslateColor(m_objDragList.BackColor))
'            'Call DrawLineEx(hDC, .Left, .Top - 1, .Right, .Top - 1, PS_SOLID, 2, GetBkColor(hDC))
'         End With 'RCPREV
'         'Draw the indicating line for the current item.
'         With rcCursor
'            Call DrawLineEx(hDC, .Left, .Top - 1, .Right, .Top - 1, PS_SOLID, 2, lLineColor)
'         End With 'RCCURSOR
'
'         'Release the device context
'         ReleaseDC .HWnd, hDC
'
'         'Save the current index.
'         nIdxPrevDragging = nIdxCursor
'
'         'Draw insert icon also for more visual effect. (Optional)
'         If bDrawInsertIcon Then
'            Call DragListDrawInsertIcon(hWndSubclass, .HWnd, nIdxCursor)
'         End If
'
'      End With 'LPDRAGLISTINFO
'
'      'Return one of:
'      'DL_STOPCURSOR: Changes the cursor to stop
'      'DL_COPYCURSOR: Changes the cursor to copy
'      'DL_MOVECURSOR: Changes the cursor to move
'      'lstMenu.MousePointer = vbCustom
'      OnDragListBox = eDragListBoxDraggingReturn.DL_CURSORSET
'
'      'eMousePointerConstants
'   Case DL_DROPPED
'      With lpDragListInfo
'         nIdxDrop = LBItemFromPt(.HWnd, .ptCursor.x, .ptCursor.y, True)
'         If nIdxDrop <> nIdxDragStartItem Then
'            Dim Cancel As Boolean
'            Call m_IDragListBox.OnDragFinish(hWndSubclass, .HWnd, nIdxDragStartItem, nIdxDrop, Cancel)
'            If Not Cancel Then 'Do the default moving action.
'               ListMoveToEx m_objDragList, nIdxDragStartItem, nIdxDrop
'               'MoveNodes nIdxDragStartItem, nIdxDrop
'            End If
'         End If
'      End With 'LPDRAGLISTINFO
'      With m_objDragList
'         'Erase the drag indicating line.
'         .Refresh
'         .MousePointer = vbDefault
'         'Erase the drag indicating curosor.
'         .Parent.Cls
'      End With 'M_OBJLISTBOX
'      OnDragListBox = 0
'   End Select
'
'End Function


'====== DropFiles: WM_DROPFILES ==================================
Public Sub AddMsgDropFiles(When As eMsgWhen, hWnd As Long)
   'Enable File drag & drop
   DragAcceptFiles hWnd, 1&
   AddMsg WM_DROPFILES, When
End Sub

Private Sub EndDropFiles(hWnd As Long)
   'Enable File drag & drop
   DragAcceptFiles hWnd, 0&
End Sub

Public Function OnDropFilesAddToListDistinct(hWndList As Long, wParam As Long, _
      Optional ExtensionFilter As String) As Collection
      
   Dim i As Long, bExists1 As Boolean, bExists As Boolean
   Dim sExt As String
   Dim pos As Long
   
   Set OnDropFilesAddToListDistinct = New Collection
   bExists1 = LenB(ExtensionFilter) = 0 Or ExtensionFilter = "*.*"
   With OnDropFiles(wParam)
      For i = 1 To .Count
         pos = InStrRev(.Item(i), ".")
         If pos > 0 Then
            sExt = Mid$(.Item(i), pos + 1)
         End If
         bExists = bExists1 Or (LenB(sExt) > 0 And InStr(1, ExtensionFilter, sExt, vbTextCompare) > 0)
         If bExists Then
            If ListAddDisinctItem(hWndList, CStr(.Item(i))) <> -1 Then
               OnDropFilesAddToListDistinct.Add .Item(i)
            End If
         End If
      Next i
   End With
End Function

Public Function OnDropFilesGetFirst(wParam As Long) As String
   'HDROP is wParam
   'will return number of files dropped on the form
   OnDropFilesGetFirst = VBA.Space$(256)
   If DragQueryFile(wParam, 0, OnDropFilesGetFirst, Len(OnDropFilesGetFirst)) > 0 Then
      OnDropFilesGetFirst = TrimNull(OnDropFilesGetFirst)
   Else
      OnDropFilesGetFirst = vbNullString
   End If
   DragFinish wParam
End Function

Public Function OnDropFiles(wParam As Long) As Collection
  
   Dim sFilename As String
   Dim lFileCount As Long
   Dim i As Long
   
   Set OnDropFiles = New Collection
   
   'will return number of files dropped on the form
   sFilename = VBA.Space$(256)
   lFileCount = DragQueryFile(wParam, -1, sFilename, Len(sFilename))
   
   If lFileCount > 0 Then
      For i = 0 To lFileCount - 1
         'sets filename to name of (i+1) th file
         'DragQueryFile wParam, i, filename,127
         DragQueryFile wParam, i, sFilename, Len(sFilename)
         OnDropFiles.Add TrimNull(sFilename)
      Next i
   End If

   DragFinish wParam
End Function

'Public Sub OnGetMinMaxInfoFixSize( _
'   ByVal HWnd As Long, ByVal lParam As Long, Width As Single, Heigth As Single)
'   Dim MinMax As MINMAXINFO
'
'   CopyMemory MinMax, ByVal lParam, LenB(MinMax)
'   With MinMax
'      .ptMinTrackSize.x = Width \ Screen.TwipsPerPixelX
'      .ptMinTrackSize.y = Heigth \ Screen.TwipsPerPixelY
'      .ptMaxTrackSize.x = .ptMinTrackSize.x
'      .ptMaxTrackSize.y = .ptMinTrackSize.y
'   End With
'   CopyMemory ByVal lParam, MinMax, LenB(MinMax)
'End Sub
'
'Public Sub OnGetMinMaxInfo(lParam As Long, lpMinMax As MINMAXINFO_VB)
'   Dim MinMax As MINMAXINFO
'   CopyMemory MinMax, ByVal lParam, LenB(MinMax)
'
'   'Set min/max values
'   With lpMinMax
'      If .MaxLeft > 0 Then MinMax.ptMaxPosition.x = .MaxLeft
'      If .MaxTop > 0 Then MinMax.ptMaxPosition.y = .MaxTop
'      If .MaxWidth > 0 Then MinMax.ptMaxSize.x = .MaxWidth
'      If .MaxHeight > 0 Then MinMax.ptMaxSize.y = .MaxHeight
'      If .MaxTrackWidth > 0 Then MinMax.ptMaxTrackSize.x = .MaxTrackWidth
'      If .MaxTrackHeight > 0 Then MinMax.ptMaxTrackSize.y = .MaxTrackHeight
'      If .MinTrackWidth > 0 Then MinMax.ptMinTrackSize.x = .MinTrackWidth
'      If .MinTrackHeight > 0 Then MinMax.ptMinTrackSize.y = .MinTrackHeight
'   End With
'
'   'Returning Value
'   CopyMemory ByVal lParam, MinMax, LenB(MinMax)
'End Sub
'
'Public Sub OnGetMinMaxInfoTwips(lParam As Long, MinMaxTwips As MINMAXINFO_VB_TWIPS)
'   Dim MinMax As MINMAXINFO
'   Static TwipX As Single, TwipY As Single
'
'   If TwipX = 0 Then
'      TwipX = Screen.TwipsPerPixelX
'      TwipY = Screen.TwipsPerPixelY
'   End If
'
'   CopyMemory MinMax, ByVal lParam, LenB(MinMax)
'
'   'Set min/max values
'   With MinMaxTwips
'      If .MaxLeft > 0 Then MinMax.ptMaxPosition.x = .MaxLeft \ TwipX
'      If .MaxTop > 0 Then MinMax.ptMaxPosition.y = .MaxTop \ TwipY
'      If .MaxWidth > 0 Then MinMax.ptMaxSize.x = .MaxWidth \ TwipX
'      If .MaxHeight > 0 Then MinMax.ptMaxSize.y = .MaxHeight \ TwipY
'      If .MaxTrackWidth > 0 Then MinMax.ptMaxTrackSize.x = .MaxTrackWidth \ TwipX
'      If .MaxTrackHeight > 0 Then MinMax.ptMaxTrackSize.y = .MaxTrackHeight \ TwipY
'      If .MinTrackWidth > 0 Then MinMax.ptMinTrackSize.x = .MinTrackWidth \ TwipX
'      If .MinTrackHeight > 0 Then MinMax.ptMinTrackSize.y = .MinTrackHeight \ TwipY
'   End With
'
'   'Returning Value
'   CopyMemory ByVal lParam, MinMax, LenB(MinMax)
'End Sub
'
'Public Function OnCommand(HWnd As Long, wParam As Long, lParam As Long) As MSG_COMMAND
'   With OnCommand
'      'lParam: Handle to the control sending the message if the message
'      'is from a control. Otherwise, this parameter is NULL.
'
'      'wParam
'      'The high-order word specifies the notification code if the message is
'      'from a control. If the message is from an accelerator, this value is 1.
'      'If the message is from a menu, this value is zero.
'      '
'      'The low-order word specifies the identifier of the menu item, control,
'      'or accelerator.
'      '
'      If lParam = 0 Then
'         .SourceType = HiWord(wParam)
'         .Handle = GetMenu(HWnd)
'      Else
'         .SourceType = mcs_Control
'         .NotificationCode = HiWord(wParam)
'         .Handle = lParam
'      End If
'      .ItemID = LoWord(wParam) And &HFFFF& 'Fix to unsigned integer
'   End With
'End Function
'
'Public Function OnMenuChar(wParam As Long, lParam As Long) As MSG_MENUCHAR
'   'WM_MENUCHAR
'   With OnMenuChar
'      .hMenu = lParam
'      .KeyAscii = LoWord(wParam)
'      .MenuType = HiWord(wParam)
'      .Shift = GetShiftState()
'      .HiReturn = MNC_CLOSE 'Set by default. MNC_CLOSE = 1
'      '.nLoReturn = .nLoReturn 'Menu Item ID
'      ''The low-order word is ignored if the high-order word contains 0 or 1.
'      '.Return = MAKELONG(.LoReturn, .HiReturn)
'   End With
'   'The WM_MENUCHAR message is sent when a menu is active and
'   'the user presses a key that does not correspond to any mnemonic or accelerator key.
'   'This message is sent to the window that owns the menu.
'
'   'wParam:    The low-order word specifies the character code
'   '                  that corresponds to the key the user pressed.
'   'wParam:   The high-order word specifies the active menu type.
'   '                 This parameter can be one of the following values.
'   'MF_POPUP: A drop-down menu, submenu, or shortcut menu.
'   'MF_SYSMENU: The window menu.
'
'   'lParam:     Handle to the active menu.
'   'An application that processes this message should return
'   'one of the following values in the high-order word of the return value.
'   'MNC_IGNORE = 0
'   '  Informs the system that it should discard the character the user pressed
'   '  and create a short beep on the system speaker.
'   'MNC_CLOSE = 1
'   '  Informs the system that it should close the active menu.
'   'MNC_EXECUTE = 2
'   '  Informs the system that it should choose the item specified in the low-order
'   '  word of the return value. The owner window receives a WM_COMMAND message.
'   'MNC_SELECT = 3
'   '  Informs the system that it should select the item specified in the low-order
'   '  word of the return value.
'End Function
'
'
'Public Function OnKeyUp(wParam As Long, lParam As Long) As MSG_KEYUP
'   'WM_KEYUP
'   'OnKeyUp(KeyCode As eVirtualKey, ByVal nRepCnt As Long, ByVal nFlags As Long)
'   With OnKeyUp
'      .KeyCode = wParam 'KeyCode As eVirtualKey
'      .KeyFlag = GetCharFlag(lParam) 'KeyFlag As CHARFLAG
'      .Shift = GetShiftState() 'Shift As eShiftConstants
'   End With
'End Function
'-------------------------------------------------------------------------------
'GetShiftState: ´­·¯Áø Å°ÀÇ Shift »óÅÂ¸¦ °¡Á®¿Â´Ù.
'Public Function GetShiftState() As eShiftConstants
'   Dim Shift As eShiftConstants
'   If (GetAsyncKeyState(VK_CONTROL) And KEY_PRESSED) = KEY_PRESSED Then
'      Shift = Shift + vbCtrlMask
'   End If
'   If (GetAsyncKeyState(VK_SHIFT) And KEY_PRESSED) = KEY_PRESSED Then
'      Shift = Shift + vbShiftMask
'   End If
'   If (GetAsyncKeyState(VK_MENU) And KEY_PRESSED) = KEY_PRESSED Then
'      Shift = Shift + vbAltMask
'   End If
'   GetShiftState = Shift
'End Function

'====== WM_HOTKEY ==================================

'Public Function AddMsgHotKey( _
'                         When As eMsgWhen, _
'                         ByVal HWnd As Long, _
'                         Optional ByVal Modifiers As eHotkeyModifier = eHotkeyModifier.MOD_NONE, _
'                         Optional ByVal Key As eVirtualKey = eVirtualKey.VK_F9) As Long
'
'   'Registers a hotkey to a window
'   On Error GoTo bye
'   If ObjPtr(m_colHotKey) = 0 Then
'      AddMsg WM_HOTKEY, When
'      Set m_colHotKey = New Collection
'   End If
'
'   Dim lngID As Long
'   'NOTE: Windows nt4 and 2000 use F12 internally, so its best not to use it as a hotkey
'   lngID = (Modifiers + 1) * 1000 + Key
'   If RegisterHotKey(HWnd, lngID, Modifiers, Key) <> 0 Then
'      m_colHotKey.Add lngID, CStr(lngID)
'      AddMsgHotKey = lngID
'   End If
'bye:
'End Function

'Public Function OnHotKey(wParam As Long) As MSG_HOTKEY
'   With OnHotKey
'      .ID = wParam
'      .Modifiers = (wParam \ 1000) - 1
'      .Key = wParam Mod 1000
'   End With
'End Function

'Private Sub DestroyHotKey(ByVal HWnd As Long)
'   Dim i As Long
'   If ObjPtr(m_colHotKey) = 0 Then
'      Exit Sub      '>---> Bottom
'   End If
'   With m_colHotKey
'      For i = 1 To .Count
'         Call UnregisterHotKey(HWnd, .Item(i))
'      Next i
'      Set m_colHotKey = New Collection
'   End With
'End Sub

'====== WM_SYSCOMMAND ==================================
'Public Function OnSysCommand(HWnd As Long, wParam As Long, lParam As Long) As MSG_SYSCOMMAND
'   'WM_SYSCOMMAND
'   With OnSysCommand
'      .hSysMenu = GetSystemMenu(HWnd, False)
'      .ItemID = LoWord(wParam) And &HFFFF& 'Fix to unsigned integer
'      .Point = MAKEPOINT(lParam)
'      .Caption = GetMenuCaption(.hSysMenu, .ItemID)
'      If ObjPtr(m_colSysMenu) Then
'         .IsCustomItem = m_colSysMenu.Exists(CStr(.ItemID))
'         If .IsCustomItem Then
'            .ShortCut = m_colSysMenu(CStr(.ItemID))
'         End If
'      End If
'   End With
'bye:
'End Function

'Public Function AddMsgSysMenu( _
'   When As eMsgWhen, HWnd As Long, ByVal ID As eMenu, _
'   ByVal Caption As String, Optional ShortCut As eVirtualKey = VK_F1) As Long
'
'   Dim hSysMenu As Long, strShortCut As String
'   If hSysMenu = 0 Then 'If objptr(m_colSysMenu) = 0 then
'      Set m_colSysMenu = New CollectionEx
'      AddMsg WM_SYSCOMMAND, When
'   End If
'   hSysMenu = GetSystemMenu(HWnd, False)
'
'   If Caption = ksMinus Then
'      AddMsgSysMenu = AppendMenuString(hSysMenu, MF_SEPARATOR, ID, ByVal Caption)
'   Else
'      If ShortCut <> 0 Then
'         strShortCut = eVirtualKeyDesc(ShortCut, True)
'         If LenB(strShortCut) Then
'            Caption = Caption & vbTab & strShortCut
'         End If
'      End If
'      AddMsgSysMenu = AppendMenuString(hSysMenu, MF_STRING, ID, ByVal Caption)
'      If AddMsgSysMenu Then
'         m_colSysMenu.Add ShortCut, CStr(ID)
'      End If
'   End If
'End Function

'============ WM_SYSTRAY ================================
'Public Sub ShowSysTrayIcon(hIcon As Long, Tooltip As String)
'   'Set the tray notify callback
'
'   Dim abTooltip() As Byte
'
'   Dim udtTray As NOTIFYICONDATA 'Tray structure
'   With udtTray
'      .cbSize = Len(udtTray) 'size of NOTIFYICONDATA type
'      '.hIcon = moForm.Icon 'or LoadPicture("c:\icon1.ico")
'      .hIcon = hIcon 'or LoadPicture("c:\icon1.ico")
'      .HWnd = hWndSubclass 'handle of window to receive messages
'      '.szTip = App.Title & Chr$(0) 'tooltip; must be terminated with chr$(0)
'      'ToolTip = StrConv(App.Title & Chr$(0), vbFromUnicode)
'      abTooltip = StrConv(Tooltip & vbNullChar, vbFromUnicode)
'      api.CopyMemory .szTip(0), abTooltip(0), UBound(abTooltip) + 1
'      .uCallbackMessage = WM_SYSTRAY 'message to process,
'      'must be unique to other windows messages.
'      'this message will be sent to our form by windows
'      'when our icon gets clicked, mouse movement, etc
'      'since we subclass our form, our WndProc function will get the message
'      .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP 'tray icon style
'      .uID = 1 'tray icon id; can be any number
'   End With
'   Shell_NotifyIcon NIM_ADD, udtTray  'add the icon to the tray
'End Sub

'Public Sub HideSysTrayIcon()
'   'Set the tray notify callback
'   Dim udtTray As NOTIFYICONDATA 'Tray structure
'   With udtTray
'      .cbSize = Len(udtTray) 'size of NOTIFYICONDATA type
'      .HWnd = hWndSubclass 'handle of window to receive messages
'      .uCallbackMessage = WM_SYSTRAY 'message to process,
'      .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP 'tray icon style
'      .uID = 1 'tray icon id
'   End With
'   Shell_NotifyIcon NIM_DELETE, udtTray 'remove icon from tray
'End Sub

