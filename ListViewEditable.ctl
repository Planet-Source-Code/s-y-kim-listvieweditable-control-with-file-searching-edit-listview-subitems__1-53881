VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.UserControl ListViewEditable 
   ClientHeight    =   2190
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   ControlContainer=   -1  'True
   PropertyPages   =   "ListViewEditable.ctx":0000
   ScaleHeight     =   2190
   ScaleWidth      =   5970
   ToolboxBitmap   =   "ListViewEditable.ctx":0053
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Æò¸é
      Height          =   285
      Left            =   3510
      TabIndex        =   4
      Top             =   90
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   4440
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   3
      Top             =   1440
      Width           =   1245
   End
   Begin VB.PictureBox pixSmall 
      Appearance      =   0  'Æò¸é
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  '¾øÀ½
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5070
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   210
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pixDummy 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000005&
      BorderStyle     =   0  '¾øÀ½
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5610
      Picture         =   "ListViewEditable.ctx":0365
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5190
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListViewEditable.ctx":06A7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1665
      Left            =   90
      TabIndex        =   2
      Top             =   240
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   2937
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuView 
         Caption         =   "&List View"
         Index           =   0
      End
      Begin VB.Menu mnuView 
         Caption         =   "Detailed &Report"
         Index           =   1
      End
      Begin VB.Menu zzmnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddItem 
         Caption         =   "&Add Item..."
      End
      Begin VB.Menu mnuRemoveItems 
         Caption         =   "&Remove Selected Items"
      End
      Begin VB.Menu mnuRemoveCheckedItems 
         Caption         =   "R&emove Checked Items"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu zzmnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemEdit 
         Caption         =   "Allow Item &Edit"
      End
      Begin VB.Menu zzmnuSep111 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSortAZ 
         Caption         =   "&Ascending Order"
      End
      Begin VB.Menu mnuSortZA 
         Caption         =   "&Descending Order"
      End
      Begin VB.Menu zzmnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu zzmnuOrderBy 
         Caption         =   "&Order By"
         Begin VB.Menu mnuOrder 
            Caption         =   "&Full Name"
            Index           =   0
         End
         Begin VB.Menu mnuOrder 
            Caption         =   "&Path"
            Index           =   1
         End
         Begin VB.Menu mnuOrder 
            Caption         =   "&Name"
            Index           =   2
         End
         Begin VB.Menu mnuOrder 
            Caption         =   "&Size"
            Index           =   3
         End
         Begin VB.Menu mnuOrder 
            Caption         =   "&Type"
            Index           =   4
         End
         Begin VB.Menu mnuOrder 
            Caption         =   "Created &Date"
            Index           =   5
         End
      End
      Begin VB.Menu zzmnuResizeHeaderBy 
         Caption         =   "&Resize by Header"
         Begin VB.Menu mnuResizeHeaders 
            Caption         =   "All Columns"
         End
         Begin VB.Menu zzmnuResize2 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnuResizeHeader 
            Caption         =   "&Full Name"
            Index           =   0
         End
         Begin VB.Menu mnuResizeHeader 
            Caption         =   "&Path"
            Index           =   1
         End
         Begin VB.Menu mnuResizeHeader 
            Caption         =   "&Name"
            Index           =   2
         End
         Begin VB.Menu mnuResizeHeader 
            Caption         =   "&Size"
            Index           =   3
         End
         Begin VB.Menu mnuResizeHeader 
            Caption         =   "&Type"
            Index           =   4
         End
         Begin VB.Menu mnuResizeHeader 
            Caption         =   "Created &Date"
            Index           =   5
         End
      End
      Begin VB.Menu zzmnuResizeBy 
         Caption         =   "Resize by &Text"
         Begin VB.Menu mnuResizeColumns 
            Caption         =   "All Columns"
         End
         Begin VB.Menu zzmnuResize1 
            Caption         =   "-"
            Index           =   0
         End
         Begin VB.Menu mnuResize 
            Caption         =   "&Full Name"
            Index           =   0
         End
         Begin VB.Menu mnuResize 
            Caption         =   "&Path"
            Index           =   1
         End
         Begin VB.Menu mnuResize 
            Caption         =   "&Name"
            Index           =   2
         End
         Begin VB.Menu mnuResize 
            Caption         =   "&Size"
            Index           =   3
         End
         Begin VB.Menu mnuResize 
            Caption         =   "&Type"
            Index           =   4
         End
         Begin VB.Menu mnuResize 
            Caption         =   "Created &Date"
            Index           =   5
         End
      End
      Begin VB.Menu zzzmnuCheck 
         Caption         =   "&Checking Items"
         Begin VB.Menu mnuCheck 
            Caption         =   "Check &All"
            Index           =   0
         End
         Begin VB.Menu mnuCheck 
            Caption         =   "&Uncheck All"
            Index           =   1
         End
         Begin VB.Menu mnuCheck 
            Caption         =   "&Invert Checkes"
            Index           =   2
         End
      End
      Begin VB.Menu zzmnuSelection 
         Caption         =   "Selecting I&tems"
         Begin VB.Menu mnuSelect 
            Caption         =   "Select &All"
            Index           =   0
         End
         Begin VB.Menu mnuSelect 
            Caption         =   "&Deselect All"
            Index           =   1
         End
         Begin VB.Menu mnuSelect 
            Caption         =   "&Invert Selections"
            Index           =   2
         End
      End
      Begin VB.Menu zzmnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu zzmnuSettings 
         Caption         =   "&Settings"
         Begin VB.Menu mnuFlatAppearance 
            Caption         =   "Flat &Appearance"
         End
         Begin VB.Menu mnuFullRowSelect 
            Caption         =   "Allow &Full Row Select"
         End
         Begin VB.Menu mnuMultiSelect 
            Caption         =   "Allow &Multi Select"
         End
         Begin VB.Menu mnuAutoDeselect 
            Caption         =   "Enable Auto &Deselect"
         End
         Begin VB.Menu mnuAutoTooltip 
            Caption         =   "Enable Auto &Tooltip"
         End
         Begin VB.Menu mnuColumnHighlight 
            Caption         =   "Enable Column &Highlighting"
         End
         Begin VB.Menu mnuSubiitemSelect 
            Caption         =   "Enable Selection via S&ubitem"
         End
         Begin VB.Menu mnuFlatHeader 
            Caption         =   "Flat &Header"
         End
         Begin VB.Menu mnuFlatScrollBar 
            Caption         =   "Flat &Scroll Bar"
         End
         Begin VB.Menu mnuCheckboxes 
            Caption         =   "Show Check&boxes"
         End
         Begin VB.Menu mnuGridLines 
            Caption         =   "Show &Grid Lines"
         End
         Begin VB.Menu mnuSolidBorder 
            Caption         =   "Single Line Border"
         End
      End
   End
End
Attribute VB_Name = "ListViewEditable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'##########################
Private Const LVM_FIRST  As Long = &H1000

'MSComctlLib.ListView Styles(LVS_)
Private Enum eListViewStyles
   LVS_ICON = &H0
   LVS_REPORT = &H1
   LVS_SMALLICON = &H2
   LVS_LIST = &H3
   LVS_TYPEMASK = &H3
   LVS_SINGLESEL = &H4
   LVS_SHOWSELALWAYS = &H8
   LVS_SORTASCENDING = &H10
   LVS_SORTDESCENDING = &H20
   LVS_SHAREIMAGELISTS = &H40
   LVS_NOLABELWRAP = &H80
   LVS_AUTOARRANGE = &H100
   LVS_EDITLABELS = &H200
   LVS_OWNERDATA = &H1000 'IE 3+ only
   LVS_NOSCROLL = &H2000

   LVS_TYPESTYLEMASK = &HFC00

   LVS_ALIGNTOP = &H0
   LVS_ALIGNLEFT = &H800
   LVS_ALIGNMASK = &HC00

   LVS_OWNERDRAWFIXED = &H400
   LVS_NOCOLUMNHEADER = &H4000
   LVS_NOSORTHEADER = &H8000&
End Enum

'--------------------------------------------------------------------------------
Private Enum eListViewMessages 'MSComctlLib.ListView Messages(LVM_)(Generic)
   LVM_GETBKCOLOR = (LVM_FIRST + 0)
   LVM_SETBKCOLOR = (LVM_FIRST + 1)
   LVM_GETIMAGELIST = (LVM_FIRST + 2)
   LVM_SETIMAGELIST = (LVM_FIRST + 3)
   LVM_GETITEMCOUNT = (LVM_FIRST + 4)

   LVM_DELETEITEM = (LVM_FIRST + 8)
   LVM_DELETEALLITEMS = (LVM_FIRST + 9)
   LVM_GETCALLBACKMASK = (LVM_FIRST + 10)
   LVM_SETCALLBACKMASK = (LVM_FIRST + 11)
   LVM_GETNEXTITEM = (LVM_FIRST + 12)

   LVM_SETITEMPOSITION = (LVM_FIRST + 15)
   LVM_GETITEMPOSITION = (LVM_FIRST + 16)

   LVM_HITTEST = (LVM_FIRST + 18)
   LVM_ENSUREVISIBLE = (LVM_FIRST + 19)
   LVM_SCROLL = (LVM_FIRST + 20)
   LVM_REDRAWITEMS = (LVM_FIRST + 21)
   LVM_ARRANGE = (LVM_FIRST + 22)

   LVM_GETEDITCONTROL = (LVM_FIRST + 24)

   LVM_DELETECOLUMN = (LVM_FIRST + 28)
   LVM_GETCOLUMNWIDTH = (LVM_FIRST + 29)
   LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)

   LVM_GETHEADER = (LVM_FIRST + 31)     'IE 3+ only

   LVM_CREATEDRAGIMAGE = (LVM_FIRST + 33)
   LVM_GETVIEWRECT = (LVM_FIRST + 34)
   LVM_GETTEXTCOLOR = (LVM_FIRST + 35)
   LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
   LVM_GETTEXTBKCOLOR = (LVM_FIRST + 37)
   LVM_SETTEXTBKCOLOR = (LVM_FIRST + 38)
   LVM_GETTOPINDEX = (LVM_FIRST + 39)
   LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
   LVM_GETORIGIN = (LVM_FIRST + 41)
   LVM_UPDATE = (LVM_FIRST + 42)
   LVM_SETITEMSTATE = (LVM_FIRST + 43)
   LVM_GETITEMSTATE = (LVM_FIRST + 44)
   LVM_SETITEMCOUNT = (LVM_FIRST + 47)
   LVM_SORTITEMS = (LVM_FIRST + 48)
   LVM_SETITEMPOSITION32 = (LVM_FIRST + 49)
   LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
   LVM_GETITEMSPACING = (LVM_FIRST + 51)

   LVM_SETICONSPACING = (LVM_FIRST + 53)     'IE 3+ only

   LVM_GETSUBITEMRECT = (LVM_FIRST + 56)
   LVM_SUBITEMHITTEST = (LVM_FIRST + 57)
   LVM_SETCOLUMNORDERARRAY = (LVM_FIRST + 58)
   LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
   LVM_SETHOTITEM = (LVM_FIRST + 60)
   LVM_GETHOTITEM = (LVM_FIRST + 61)
   LVM_SETHOTCURSOR = (LVM_FIRST + 62)
   LVM_GETHOTCURSOR = (LVM_FIRST + 63)
   LVM_APPROXIMATEVIEWRECT = (LVM_FIRST + 64)
   LVM_SETWORKAREA = (LVM_FIRST + 65)

   LVM_GETSELECTIONMARK = (LVM_FIRST + 66) 'Win32 and IE 4 only
   LVM_SETSELECTIONMARK = (LVM_FIRST + 67) 'Win32 and IE 4 only
   LVM_GETWORKAREA = (LVM_FIRST + 70)      'Win32 and IE 4 only
   LVM_SETHOVERTIME = (LVM_FIRST + 71)     'Win32 and IE 4 only
   LVM_GETHOVERTIME = (LVM_FIRST + 72)     'Win32 and IE 4 only

   '--------------------------------------------------------------------------------
   'MSComctlLib.ListView Messages(LVM_)(Win95)
   LVM_GETITEM = (LVM_FIRST + 5)
   LVM_SETITEM = (LVM_FIRST + 6)

   LVM_INSERTITEMA = (LVM_FIRST + 7)
   LVM_INSERTITEM = LVM_INSERTITEMA

   LVM_FINDITEMA = (LVM_FIRST + 13)
   LVM_FINDITEM = LVM_FINDITEMA

   LVM_GETSTRINGWIDTHA = (LVM_FIRST + 17)
   LVM_GETSTRINGWIDTH = LVM_GETSTRINGWIDTHA

   LVM_EDITLABELA = (LVM_FIRST + 23)
   LVM_EDITLABEL = LVM_EDITLABELA

   LVM_GETCOLUMNA = (LVM_FIRST + 25)
   LVM_GETCOLUMN = LVM_GETCOLUMNA

   LVM_SETCOLUMNA = (LVM_FIRST + 26)
   LVM_SETCOLUMN = LVM_SETCOLUMNA

   LVM_INSERTCOLUMNA = (LVM_FIRST + 27)
   LVM_INSERTCOLUMN = LVM_INSERTCOLUMNA

   LVM_GETITEMTEXTA = (LVM_FIRST + 45)
   LVM_GETITEMTEXT = LVM_GETITEMTEXTA

   LVM_SETITEMTEXTA = (LVM_FIRST + 46)
   LVM_SETITEMTEXT = LVM_SETITEMTEXTA

   LVM_GETISEARCHSTRINGA = (LVM_FIRST + 52)
   LVM_GETISEARCHSTRING = LVM_GETISEARCHSTRINGA

   LVM_SETBKIMAGEA = (LVM_FIRST + 68)   'Win32 and IE 4 only
   LVM_GETBKIMAGEA = (LVM_FIRST + 69)   'Win32 and IE 4 only
End Enum 'eListViewMessages 'MSComctlLib.ListView Messages(LVM_)(Generic)

'MSComctlLib.ListView Set Column Width Messages (LVSCW_)
Private Enum eListViewSetColumnWidthMessages 'MSComctlLib.ListView Set Column Width Messages (LVSCW_)
   LVSCW_AUTOSIZE = -1
   LVSCW_AUTOSIZE_USEHEADER = -2
End Enum

Private Enum eListViewFindItemRectMessages
   LVIR_BOUNDS = 0
   LVIR_ICON = 1
   LVIR_LABEL = 2
   LVIR_SELECTBOUNDS = 3
End Enum

Private Const BIF_RETURNONLYFSDIRS As Long = &H1
Private Const MAX_PATH As Long = 260

Private Type SHITEMID
   cb      As Long
   abID    As Byte
End Type

Private Type ITEMIDLIST
   mkid    As SHITEMID
End Type

Private Type BROWSEINFO
   hOwner          As Long
   pidlRoot        As Long
   pszDisplayName  As String
   lpszTitle       As String
   ulFlags         As Long
   lpfn            As Long
   lParam          As Long
   iImage          As Long
End Type

'To the Constant declarations add:
Private Const SHGFI_DISPLAYNAME  As Long = &H200
Private Const SHGFI_EXETYPE  As Long = &H2000
Private Const SHGFI_SYSICONINDEX  As Long = &H4000  'system icon index
Private Const SHGFI_LARGEICON  As Long = &H0        'large icon
Private Const SHGFI_SMALLICON  As Long = &H1        'small icon
Private Const ILD_TRANSPARENT  As Long = &H1        'display transparent
Private Const SHGFI_SHELLICONSIZE  As Long = &H4
Private Const SHGFI_TYPENAME  As Long = &H400
Private Const BASIC_SHGFI_FLAGS  As Long = SHGFI_TYPENAME Or _
                SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or _
                SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO
   hIcon          As Long
   iIcon          As Long
   dwAttributes   As Long
   szDisplayName  As String * MAX_PATH
   szTypeName     As String * 80
End Type

Private shinfo As SHFILEINFO

Private Type FILETIME
   dwLowDateTime     As Long
   dwHighDateTime    As Long
End Type

Private Type SYSTEMTIME
   wYear             As Integer
   wMonth            As Integer
   wDayOfWeek        As Integer
   wDay              As Integer
   wHour             As Integer
   wMinute           As Integer
   wSecond           As Integer
   wMilliseconds     As Integer
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes  As Long
   ftCreationTime    As FILETIME
   ftLastAccessTime  As FILETIME
   ftLastWriteTime   As FILETIME
   nFileSizeHigh     As Long
   nFileSizeLow      As Long
   dwReserved0       As Long
   dwReserved1       As Long
   cFileName         As String * MAX_PATH
   cAlternate        As String * 14
End Type

Private Enum eFileAttributes
   FILE_ATTRIBUTE_ARCHIVE = &H20
   FILE_ATTRIBUTE_COMPRESSED = &H800
   FILE_ATTRIBUTE_DEVICE = &H40
   FILE_ATTRIBUTE_DIRECTORY = &H10
   FILE_ATTRIBUTE_ENCRYPTED = &H4000
   FILE_ATTRIBUTE_HIDDEN = &H2
   FILE_ATTRIBUTE_NORMAL = &H80
   FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = &H2000
   FILE_ATTRIBUTE_OFFLINE = &H1000
   FILE_ATTRIBUTE_READONLY = &H1
   FILE_ATTRIBUTE_REPARSE_POINT = &H400
   FILE_ATTRIBUTE_SPARSE_FILE = &H200
   FILE_ATTRIBUTE_SYSTEM = &H4
   FILE_ATTRIBUTE_TEMPORARY = &H100
End Enum

Private Enum eListViewNotificationItemMessages
   LVNI_ALL = &H0
   LVNI_FOCUSED = &H1
   LVNI_SELECTED = &H2
   LVNI_CUT = &H4
   LVNI_DROPHILITED = &H8

   LVNI_ABOVE = &H100
   LVNI_BELOW = &H200
   LVNI_TOLEFT = &H400
   LVNI_TORIGHT = &H800
End Enum

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Enum eListViewHitTestMessages 'MSComctlLib.ListView Hit Test Messages (LVHT_)
   LVHT_NOWHERE = &H1
   LVHT_ONITEMICON = &H2
   LVHT_ONITEMLABEL = &H4
   LVHT_ONITEMSTATEICON = &H8
   LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

   LVHT_ABOVE = &H8
   LVHT_BELOW = &H10
   LVHT_TORIGHT = &H20
   LVHT_TOLEFT = &H40
End Enum 'eListViewHitTest 'MSComctlLib.ListView Hit Test Messages (LVHT_)

Private Type LVHITTESTINFO
   pt As POINTAPI
   flags As eListViewHitTestMessages
   iItem As Long
   iSubItem  As Long  'ie3+ only .. was NOT in win95.
   'Valid only for LVM_SUBITEMHITTEST
End Type

Private Type HD_HITTESTINFO
   pt  As POINTAPI
   flags  As Long
   iItem As Long
End Type

'HitTest positions
Private Const HHT_NOWHERE As Long = &H1
Private Const HHT_ONHEADER As Long = &H2
Private Const HHT_ONDIVIDER As Long = &H4
Private Const HHT_ONDIVOPEN As Long = &H8
Private Const HHT_ABOVE As Long = &H100
Private Const HHT_BELOW As Long = &H200
Private Const HHT_TORIGHT As Long = &H400
Private Const HHT_TOLEFT As Long = &H800

'header messages
Private Const HDM_FIRST           As Long = &H1200
Private Const HDM_GETITEMCOUNT    As Long = (HDM_FIRST + 0)
Private Const HDM_INSERTITEM      As Long = (HDM_FIRST + 1)
Private Const HDM_DELETEITEM      As Long = (HDM_FIRST + 2)
Private Const HDM_GETITEM         As Long = (HDM_FIRST + 3)
Private Const HDM_SETITEM         As Long = (HDM_FIRST + 4)
Private Const HDM_LAYOUT          As Long = (HDM_FIRST + 5)
Private Const HDM_HITTEST         As Long = (HDM_FIRST + 6)
Private Const HDM_GETITEMRECT     As Long = (HDM_FIRST + 7)
Private Const HDM_SETIMAGELIST    As Long = (HDM_FIRST + 8)
Private Const HDM_GETIMAGELIST    As Long = (HDM_FIRST + 9)
Private Const HDM_ORDERTOINDEX    As Long = (HDM_FIRST + 15)

Private Const LVIS_SELECTED As Long = &H2
Private Const LVIF_STATE As Long = &H8
Private Const LVIS_STATEIMAGEMASK As Long = &HF000

'listview, header
Private Const ICC_LISTVIEW_CLASSES  As Long = &H1
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 54)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE As Long = (LVM_FIRST + 55)

'notify messages
Private Const HDN_FIRST As Long = -300&  'header
Private Const HDN_ITEMCHANGING As Long = (HDN_FIRST - 0)
Private Const WM_NOTIFY As Long = &H4E&

Private Const HDS_BUTTONS As Long = &H2
Private Const GWL_STYLE As Long = (-16)
Private Const SWP_DRAWFRAME As Long = &H20
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FLAGS As Long = SWP_NOZORDER Or _
                SWP_NOSIZE Or _
                SWP_NOMOVE Or _
                SWP_DRAWFRAME

Private Type LV_ITEM
   mask As Long
   iItem As Long
   iSubItem As Long
   State As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   lParam As Long
   iIndent As Long
End Type

Private Type NMHDR
   hWndFrom As Long
   idfrom   As Long
   Code     As Long
End Type

Private Const WM_VSCROLL = &H115
Private Const SB_VERT = 1

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_DISABLENOSCROLL = &H8
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const SB_LINEUP = 0
Private Const SB_LINEDOWN = 1
Private Const SB_PAGEUP = 2
Private Const SB_PAGEDOWN = 3
Private Const SB_THUMBPOSITION = 4
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_BOTTOM = 7
Private Const SB_ENDSCROLL = 8

Private Type SCROLLINFO
   cbSize As Long
   fMask As Long
   nMin As Long
   nMax As Long
   nPage As Long
   nPos As Long
   nTrackPos As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpScrollInfo As SCROLLINFO) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal fRedraw As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "Kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "Kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FileTimeToSystemTime Lib "Kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ImageList_Draw Lib "comctl32" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal flags&) As Long
Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptX As Long, ByVal ptY As Long) As Long

'#########################################
' ---------- Custom Structure & Enumerations --------------------------
Private Type FILEINFO
   sFilename As String
   sSizeKB As String
   sType As String
   sCreatedDate As String
   hSmallIcon As Long
   iIcon As Long
   sImgKey As String
   bExeFile As Boolean
   bDOSExe As Boolean
End Type

'Private Enum eListColumnDataType
'   ldtString = 0
'   ldtNumber = 1
'   ldtDateTime = 2
'End Enum

'Public Enum eImageSizingTypes
'   sizeNone = 0
'   sizeCheckBox
'   sizeIcon
'End Enum
'
'Public Enum eLedgerColours
'   vbLedgerWhite = &HF9FEFF
'   vbLedgerGreen = &HD0FFCC
'   vbLedgerYellow = &HE1FAFF
'   vbLedgerRed = &HE1E1FF
'   vbLedgerGrey = &HE0E0E0
'   vbLedgerBeige = &HD9F2F7
'   vbLedgerSoftWhite = &HF7F7F7
'   vbLedgerPureWhite = &HFFFFFF
'End Enum

'ÀÌº¥Æ® ¼±¾ð:
Public Event AddItemRequested()
Public Event SearchDirChange(Directory As String)
Public Event SubitemClick(Index As Long, Subitem As Long, Button As Integer, Shift As Integer)
Public Event HeaderItemChanging()
Public Event VScroll()
Public Event HScroll()
Public Event MouseWheel()
Public Event AfterItemEdit(Cancel As Boolean, Index As Long, Subitem As Long, OldString As String, NewString As String) 'MappingInfo=ListView1,ListView1,-1,AfterLabelEdit
Public Event BeforeItemEdit(Cancel As Boolean, Index As Long, Subitem As Long, OldString As String, NewString As String) 'MappingInfo=ListView1,ListView1,-1,AfterLabelEdit
Public Event FileDragDropFinish(Count As Long, Files As Long, Folders As Long) 'MappingInfo=ListView1,ListView1,-1,BeforeLabelEdit
Public Event Resize()
Public Event BeforeScanStart(Cancel As Boolean, Path As String, FileSpec As String, Recursive As Boolean, IncludeFolder As Boolean)
Public Event ScanFinish(Stopped As Boolean, Path As String, FileSpec As String, Recursive As Boolean, IncludeFolder As Boolean)
'Public Event AfterLabelEdit(Cancel As Integer, NewString As String) 'MappingInfo=ListView1,ListView1,-1,AfterLabelEdit
'Public Event BeforeLabelEdit(Cancel As Integer) 'MappingInfo=ListView1,ListView1,-1,BeforeLabelEdit
Public Event Click() 'MappingInfo=ListView1,ListView1,-1,Click
Public Event ColumnClick(ByVal ColumnHeader As ColumnHeader) 'MappingInfo=ListView1,ListView1,-1,ColumnClick
Public Event DblClick() 'MappingInfo=ListView1,ListView1,-1,DblClick
Public Event ItemCheck(ByVal Item As ListItem) 'MappingInfo=ListView1,ListView1,-1,ItemCheck
Public Event ItemClick(ByVal Item As ListItem) 'MappingInfo=ListView1,ListView1,-1,ItemClick
Public Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=ListView1,ListView1,-1,KeyDown
Public Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=ListView1,ListView1,-1,KeyUp
Public Event KeyPress(KeyAscii As Integer) 'MappingInfo=ListView1,ListView1,-1,KeyPress
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=ListView1,ListView1,-1,MouseDown
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=ListView1,ListView1,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=ListView1,ListView1,-1,MouseUp
Public Event OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=ListView1,ListView1,-1,OLEDragDrop

' ---------- Custom Structure & Enumerations --------------------------
'#########################################

Private sc As CSubclass
Implements ISubclass
Private sc2 As CSubclass
Implements ISubclass2

Private m_bStop As Boolean

Private prevOrder As Integer

'Private m_lCurrrentHighlightColumn As Long
'Private m_enHighlightColor As eLedgerColours
'Private m_enDefaultColor As eLedgerColours
'Private m_enSizingType As eImageSizingTypes

'Default Property Values
Private Const m_def_AutoDeselect As Boolean = True
Private Const m_def_FileSpec = "*.*"
Private Const m_def_IncludeFolder  As Boolean = True
Private Const m_def_Path = "C:\"
Private Const m_def_Recursive As Boolean = True
Private Const m_def_SubitemSelect As Boolean = True
Private Const m_def_UpdateFrequency As Long = 100
Private Const m_def_CurrrentHighlightColumn As Long = 1
Private Const m_def_HighlightColor As Long = vbLedgerRed
Private Const m_def_DefaultColor As Long = vbLedgerPureWhite
Private Const m_def_SizingType As Long = sizeIcon
Private Const m_def_HighlightColumn As Boolean = True
Private Const m_def_FlatHeader As Boolean = False
Private Const m_def_AutoTooltip  As Boolean = True
Private Const m_def_AutoPopupMenu As Boolean = True
Private Const m_def_AllowItemEdit As Boolean = True
Private Const m_def_ColumnHeadersText As String = _
                     "|Text=FullName|Key=|Width=1299.969|Alignment=0|DataType=0|Level=0|" & vbCrLf & _
                     "|Text=Path|Key=|Width=1000.063|Alignment=0|DataType=0|Level=0|" & vbCrLf & _
                     "|Text=Name|Key=|Width=1399.748|Alignment=0|DataType=0|Level=0|" & vbCrLf & _
                     "|Text=Size|Key=|Width=799.9371|Alignment=1|DataType=1|Level=0|" & vbCrLf & _
                     "|Text=Type|Key=|Width=1000.063|Alignment=0|DataType=0|Level=0|" & vbCrLf & _
                     "|Text=Created Date|Key=|Width=1299.969|Alignment=0|DataType=2|Level=0|"
Private Const m_def_AllowFileDragDrop As Boolean = True

'Property Variables
Private m_bAutoDeselect As Boolean
Private m_strFileSpec As String
Private m_bIncludeFolder As Boolean
Private m_strPath As String
Private m_bRecursive As Boolean
Private m_bSubitemSelect As Boolean
Private m_lUpdateFrequency As Long
Private m_enHighlightColor As eLedgerColours
Private m_enDefaultColor As eLedgerColours
Private m_enSizingType As eImageSizingTypes
Private m_lCurrrentHighlightColumn As Long
Private m_bHighlightColumn As Boolean
Private m_bFlatHeader As Boolean
Private m_bAutoTooltip As Boolean
Private m_bAutoPopupMenu As Boolean
Private m_bAllowItemEdit As Boolean '@@ ItemEdit 04/05/17, 16:58:13
Private m_strColumnHeadersText As String
Private m_bAllowFileDragDrop As Boolean
Private Enum eFileNameParts
   efpBaseName = 0
   efpExtension
   efpPath
   efpPathUnqualified
   efpName
   efpPathPlusBaseName
   efpPathBaseName
   efpDrive
   efpDriveQualified
   efpConvToLocalName
   efpConvToShortName
   efpConvToLongName
End Enum

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, lpRect As RECT, ByVal uFormat As Long) As Long

Private Function CalcTextWidth(Text As String) As Long
   On Error Resume Next

   'Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
   Const DT_CALCRECT As Long = &H400& '(DrawText, UINT uFormat (text-drawing options), eDrawText)

   Dim rcText As RECT
   Dim hDC As Long
   hDC = GetDC(ListView1.hWnd)
   Call DrawText(hDC, Text, -1&, rcText, DT_CALCRECT)
   CalcTextWidth = rcText.Right
   ReleaseDC ListView1.hWnd, hDC
End Function



Public Sub AddFile(QualifiedPath As String, FileName As String)
Attribute AddFile.VB_Description = "Add a file."
'Add a file.

   Dim itmX As ListItem
   Dim imgX As ListImage
   Dim r As Long

   On Local Error GoTo AddFileItemViewError

   Dim fi As FILEINFO
   With fi
      .sFilename = QualifiedPath & FileName
      GetFileInfo fi

      Set itmX = ListView1.ListItems.Add(, , .sFilename)
      If Len(.sImgKey) <> 0 Then
         itmX.SmallIcon = ImageList1.ListImages(.sImgKey).Key
      End If
      itmX.SubItems(1) = QualifiedPath
      itmX.SubItems(2) = FileName
      itmX.SubItems(3) = .sSizeKB
      itmX.SubItems(4) = .sType
      itmX.SubItems(5) = .sCreatedDate
   End With 'FI

Exit Sub

AddFileItemViewError:
   'Debug.Print Err.Description
   With fi
      pixSmall.Picture = LoadPicture()
      r& = ImageList_Draw(.hSmallIcon, .iIcon, pixSmall.hDC, 0, 0, ILD_TRANSPARENT)
      pixSmall.Picture = pixSmall.Image
      Set imgX = ImageList1.ListImages.Add(, .sImgKey, pixSmall.Picture)
   End With 'FI
   Resume

End Sub

'Returns/sets whether a user can reorder columns in report view.
Public Property Get AllowColumnReorder() As Boolean
Attribute AllowColumnReorder.VB_Description = "Returns/sets whether a user can reorder columns in report view."

   AllowColumnReorder = ListView1.AllowColumnReorder

End Property

Public Property Let AllowColumnReorder(ByVal New_AllowColumnReorder As Boolean)

   ListView1.AllowColumnReorder() = New_AllowColumnReorder
   PropertyChanged "AllowColumnReorder"

End Property

Public Property Get Appearance() As AppearanceConstants
   Appearance = ListView1.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
   ListView1.Appearance() = New_Appearance
   mnuFlatAppearance.Checked = New_Appearance = ccFlat
   PropertyChanged "Appearance"
End Property

Public Property Get Arrange() As ListArrangeConstants
   Arrange = ListView1.Arrange
End Property

Public Property Let Arrange(ByVal New_Arrange As ListArrangeConstants)
   ListView1.Arrange() = New_Arrange
   PropertyChanged "Arrange"
End Property

'Returns/sets whether to deselect an selected item automatically when the user clicks its space area.
Public Property Get AutoDeselect() As Boolean
Attribute AutoDeselect.VB_Description = "Returns/sets whether to deselect an selected item automatically when the user clicks its space area."

   AutoDeselect = m_bAutoDeselect

End Property
Public Property Let AutoDeselect(ByVal bAutoDeselect As Boolean)

   m_bAutoDeselect = bAutoDeselect
   PropertyChanged "AutoDeselect"

End Property

Public Property Get AllowItemEdit() As Boolean
Attribute AllowItemEdit.VB_Description = "Returns/sets whether to allow the user to edit item and subitems."
   'Returns/sets whether to allow the user to edit item and subitems.
   AllowItemEdit = m_bAllowItemEdit
End Property

Public Property Let AllowItemEdit(ByVal New_AllowItemEdit As Boolean)
   m_bAllowItemEdit = New_AllowItemEdit
   mnuItemEdit.Checked = New_AllowItemEdit
   If m_bAllowItemEdit Then
      MoveEditBox
   Else
      txtEdit.Visible = False
   End If
   PropertyChanged "AllowItemEdit"
End Property


Public Property Get AutoPopupMenu() As Boolean
Attribute AutoPopupMenu.VB_Description = "Returns/sets whether to display the default popup menu on Mouse Up event of the mouse right button."
   'Returns/sets whether to display the default popup menu on Mouse Up event of the mouse right button.
   AutoPopupMenu = m_bAutoPopupMenu
End Property
Public Property Let AutoPopupMenu(ByVal New_AutoPopupMenu As Boolean)
   'If Ambient.UserMode = False Then Err.Raise 387
   m_bAutoPopupMenu = New_AutoPopupMenu
   PropertyChanged "AutoPopupMenu"
End Property

Public Property Get AutoTooltip() As Boolean
Attribute AutoTooltip.VB_Description = "Returns/sets whether to automatically set the ListView's tooltip text to the text of list item or subitem under the mouse pointer during the MouseMove event."
'Returns/sets whether to automatically set the ListView's tooltip text to the text of list item or subitem under the mouse pointer during the MouseMove event.
   AutoTooltip = m_bAutoTooltip
End Property

Public Property Let AutoTooltip(ByVal New_AutoTooltip As Boolean)
   m_bAutoTooltip = New_AutoTooltip
   mnuAutoTooltip.Checked = New_AutoTooltip
   PropertyChanged "AutoTooltip"
End Property


'Returns/sets a value which determines if the control displays a checkbox next to each item in the list.
Public Property Get Checkboxes() As Boolean
Attribute Checkboxes.VB_Description = "Returns/sets a value which determines if the control displays a checkbox next to each item in the list."

   Checkboxes = ListView1.Checkboxes

End Property

Public Property Let Checkboxes(ByVal New_Checkboxes As Boolean)

   ListView1.Checkboxes() = New_Checkboxes
   mnuCheckboxes.Checked = New_Checkboxes
   zzzmnuCheck.Enabled = New_Checkboxes
   mnuRemoveCheckedItems.Enabled = New_Checkboxes
   PropertyChanged "Checkboxes"
   'UpdateWindow ListView1.hWnd
   'DoEvents

End Property

'Returns whether the checkbox left to the specified item is checked or not,
'or checks or unchecks the check box left to specified item.
Public Property Get Checked(ByVal Index As Long) As Boolean
Attribute Checked.VB_Description = "Returns whether the checkbox left to the specified item is checked or not, or checks or unchecks the check box left to specified item."

   Dim r As Long

   r = SendMessage(ListView1.hWnd, LVM_GETITEMSTATE, Index - 1, ByVal LVIS_STATEIMAGEMASK)
   Checked = r And &H2000&

End Property

Public Property Let Checked(ByVal Index As Long, bState As Boolean)

   Dim lv As LV_ITEM

   With lv
      .mask = LVIF_STATE
      .State = IIf(bState, &H2000, &H1000)
      .stateMask = LVIS_STATEIMAGEMASK
   End With 'LV
   Call SendMessage(hWnd, LVM_SETITEMSTATE, Index - 1, lv)

End Property

'Checks or unchecks all items.
Public Property Let CheckedAll(ByVal bState As Boolean)
Attribute CheckedAll.VB_Description = "Checks or unchecks all items."

   Dim lv As LV_ITEM

   With lv
      .mask = LVIF_STATE
      .State = IIf(bState, &H2000, &H1000)
      .stateMask = LVIS_STATEIMAGEMASK
   End With 'LV

   Call SendMessage(ListView1.hWnd, LVM_SETITEMSTATE, -1, lv)

End Property

'Returns the count of checked items and also returns their index in the given array.
Public Function GetCheckedItems(nChecked() As Long) As Long
Attribute GetCheckedItems.VB_Description = "Returns the count of checked items and also returns their index in the given array."

   Dim numChecked As Long
   Dim hWnd As Long
   Dim lv As LV_ITEM
   Dim lvCount As Long
   Dim lvIndex As Long
   Dim r As Long

   hWnd = ListView1.hWnd
   lvCount = ListView1.ListItems.Count - 1

   Do

      r = SendMessage(hWnd, LVM_GETITEMSTATE, lvIndex, ByVal LVIS_STATEIMAGEMASK)
      If r And &H2000& Then
         ReDim Preserve nChecked(0 To numChecked)
         nChecked(numChecked) = lvIndex + 1
         numChecked = numChecked + 1
      End If

      lvIndex = lvIndex + 1

   Loop Until lvIndex > lvCount

   GetCheckedItems = numChecked

End Function

Public Property Get ColumnHeaderIcons() As Object

'Returns/sets the ImageList control to be used for ColumnHeader icons.

   Set ColumnHeaderIcons = ListView1.ColumnHeaderIcons

End Property

Public Property Let ColumnHeaderIcons(ByVal objColumnHeaderIcons As Object)

'Returns/sets the ImageList control to be used for ColumnHeader icons.

   ListView1.ColumnHeaderIcons = objColumnHeaderIcons

End Property

Public Property Set ColumnHeaderIcons(ByVal objColumnHeaderIcons As Object)

'Returns/sets the ImageList control to be used for ColumnHeader icons.

   Set ListView1.ColumnHeaderIcons = objColumnHeaderIcons

End Property

Public Property Get ColumnHeaders() As ColumnHeaders

'Returns a reference to a collection of ColumnHeader objects.

   Set ColumnHeaders = ListView1.ColumnHeaders

End Property

Public Property Get CurrrentHighlightColumn() As Long

   CurrrentHighlightColumn = m_lCurrrentHighlightColumn

End Property

Public Property Let CurrrentHighlightColumn(ByVal lCurrrentHighlightColumn As Long)

   m_lCurrrentHighlightColumn = lCurrrentHighlightColumn
   If m_bHighlightColumn Then
      HighlightColumn = True
   End If
   'Call DoHighlightColumn(m_enHighlightColor, m_enDefaultColor, m_lCurrrentHighlightColumn, m_enSizingType)
   PropertyChanged "CurrrentHighlightColumn"

End Property

Public Property Get DefaultColor() As eLedgerColours

   DefaultColor = m_enDefaultColor

End Property

Public Property Let DefaultColor(ByVal enDefaultColor As eLedgerColours)

   m_enDefaultColor = enDefaultColor
   If m_bHighlightColumn Then
      HighlightColumn = True
   End If
   'Call DoHighlightColumn(m_enHighlightColor, m_enDefaultColor, m_lCurrrentHighlightColumn, m_enSizingType)
   PropertyChanged "DefaultColor"

End Property

Public Property Get DisplayName() As String

   DisplayName = m_strPath & m_strFileSpec

End Property

Public Sub DoHighlightColumn( _
                             Optional clrHighlight As eLedgerColours = vbLedgerRed, _
                             Optional clrDefault As eLedgerColours = vbLedgerPureWhite, _
                             Optional nColumn As Long = 1, _
                             Optional nSizingType As eImageSizingTypes = sizeNone)

   m_enHighlightColor = clrHighlight
   m_enDefaultColor = clrDefault
   m_lCurrrentHighlightColumn = nColumn
   If m_lCurrrentHighlightColumn <= 0 Then
      m_lCurrrentHighlightColumn = 1
   End If
   m_enSizingType = nSizingType
   If nColumn = 1 Then
      m_enSizingType = sizeIcon
   End If

   With ListView1
      .Visible = False
      '.Checkboxes = False
      '.FullRowSelect = True
      Call SetHighlightColumn(ListView1, m_enHighlightColor, m_enDefaultColor, m_lCurrrentHighlightColumn, m_enSizingType, Picture1)
      .Refresh
      .Visible = True            '/* Restore visibility
   End With 'LISTVIEW1

End Sub

Public Property Get Enabled() As Boolean
   Enabled = ListView1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   ListView1.Enabled() = New_Enabled
   UserControl.Enabled = New_Enabled
   PropertyChanged "Enabled"
End Property

Public Property Get FileItemCount() As Long

   FileItemCount = ListView1.ListItems.Count

End Property

Public Property Let FileSpec(ByVal strFileSpec As String)

   m_strFileSpec = strFileSpec
   PropertyChanged "FileSpec"

End Property

Public Property Get FileSpec() As String

   FileSpec = m_strFileSpec

End Property

Public Property Get FlatHeader() As Boolean

   FlatHeader = m_bFlatHeader

End Property

Public Property Let FlatHeader(ByVal New_FlatHeader As Boolean)

   m_bFlatHeader = New_FlatHeader
   If m_bFlatHeader Then
      SetFlatHeader
   Else 'M_BFLATHEADER = FALSE
      Set3DHeader
   End If
   PropertyChanged "FlatHeader"

End Property

Public Property Let FlatScrollBar(ByVal New_FlatScrollBar As Boolean)

   ListView1.FlatScrollBar() = New_FlatScrollBar
   PropertyChanged "FlatScrollBar"

End Property

'Returns/sets whether the scrollbars appear flat.
Public Property Get FlatScrollBar() As Boolean

   FlatScrollBar = ListView1.FlatScrollBar

End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)

   ListView1.FullRowSelect() = New_FullRowSelect
   mnuFullRowSelect.Checked = New_FullRowSelect
   mnuAutoDeselect.Enabled = Not New_FullRowSelect
   PropertyChanged "FullRowSelect"

End Property

'Returns/sets whether selecting a column highlights the entire row.
Public Property Get FullRowSelect() As Boolean

   FullRowSelect = ListView1.FullRowSelect

End Property

Private Sub GetFileInfo(fi As FILEINFO)

   Dim wfd As WIN32_FIND_DATA
   Dim hFile As Long
   Dim pos As Long
   Dim hExeType As Long
   Dim sTypeLCase As String

   With fi
      hFile = FindFirstFile(.sFilename, wfd)
      If hFile > 0 Then

         '.hIcon = SHGetFileInfo(.sFileName, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS)
         .hSmallIcon = SHGetFileInfo(.sFilename, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
         .iIcon = shinfo.iIcon

         'Get type with Trim null
         .sType = shinfo.szTypeName
         pos = InStr(.sType, Chr$(0))
         If pos Then
            .sType = Left$(.sType, pos - 1)
         End If

         sTypeLCase = LCase$(.sType)

         .sImgKey = .sType
         If sTypeLCase = "application" Or sTypeLCase = "shortcut" Or _
            .sType = "ÀÀ¿ë ÇÁ·Î±×·¥" Or .sType = "¹Ù·Î °¡±â" Then

            .bExeFile = True

            If sTypeLCase = "application" Or .sType = "ÀÀ¿ë ÇÁ·Î±×·¥" Then
               hExeType = SHGetFileInfo(.sFilename, 0&, shinfo, Len(shinfo), SHGFI_EXETYPE)
               'Get the high word
               If hExeType And &H80000000 Then
                  hExeType = (hExeType \ 65535) - 1
               Else 'NOT HEXETYPE...
                  hExeType = hExeType \ 65535
               End If
            End If

            If hExeType > 0 Or sTypeLCase = "shortcut" Or .sType = "¹Ù·Î °¡±â" Then
               'Get the file name without path
               .sImgKey = Mid$(.sFilename, InStrRev(.sFilename, "\") + 1)
               .bDOSExe = False
            Else 'NOT HEXETYPE...
               .sImgKey = "DOSExeIcon"
               .bDOSExe = True
            End If
         Else 'NOT STYPELCASE...
            .bExeFile = False
            .bDOSExe = False
         End If

         .sSizeKB = Format$(((wfd.nFileSizeHigh + wfd.nFileSizeLow) / 1000) + 0.5, "#,###,###") & "KB"

         Dim ST As SYSTEMTIME
         If FileTimeToSystemTime(wfd.ftCreationTime, ST) Then
            .sCreatedDate = Format$(DateSerial(ST.wYear, ST.wMonth, ST.wDay), "Short Date")
         End If

      End If
      FindClose hFile
   End With 'FI

End Sub

'Returns/sets whether grid lines appear between rows and columns
Public Property Get GridLines() As Boolean

   GridLines = ListView1.GridLines

End Property

Public Property Let GridLines(ByVal New_GridLines As Boolean)

   ListView1.GridLines() = New_GridLines
   PropertyChanged "GridLines"

End Property

Public Property Get HideColumnHeaders() As Boolean
   HideColumnHeaders = ListView1.HideColumnHeaders
End Property

Public Property Let HideColumnHeaders(ByVal New_HideColumnHeaders As Boolean)
   ListView1.HideColumnHeaders() = New_HideColumnHeaders
   PropertyChanged "HideColumnHeaders"
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)

   ListView1.HideSelection() = New_HideSelection
   PropertyChanged "HideSelection"

End Property

'Determines whether the selected item will display as selected when the MSComctlLib.ListView loses focus
Public Property Get HideSelection() As Boolean

   HideSelection = ListView1.HideSelection

End Property

Public Property Get HighlightColor() As eLedgerColours

   HighlightColor = m_enHighlightColor

End Property

Public Property Let HighlightColor(ByVal enHighlightColor As eLedgerColours)

   m_enHighlightColor = enHighlightColor
   If m_bHighlightColumn Then
      HighlightColumn = True
   End If
   'Call DoHighlightColumn(m_enHighlightColor, m_enDefaultColor, m_lCurrrentHighlightColumn, m_enSizingType)

   PropertyChanged "HighlightColor"

End Property

Public Property Let HighlightColumn(ByVal New_HighlightColumn As Boolean)

'If Ambient.UserMode = False Then Err.Raise 387

   m_bHighlightColumn = New_HighlightColumn
   mnuColumnHighlight.Checked = New_HighlightColumn
   If New_HighlightColumn Then
      Call DoHighlightColumn(m_enHighlightColor, m_enDefaultColor, m_lCurrrentHighlightColumn, m_enSizingType)
   Else 'NEW_HIGHLIGHTCOLUMN = FALSE
      With ListView1
         .Visible = False
         Call SetHighlightColumn(ListView1, vbLedgerPureWhite, vbLedgerPureWhite, m_lCurrrentHighlightColumn, m_enSizingType, Picture1)
         .Refresh
         .Visible = True            '/* Restore visibility
      End With 'LISTVIEW1
   End If
   PropertyChanged "HighlightColumn"

End Property

Public Property Get HighlightColumn() As Boolean

   HighlightColumn = m_bHighlightColumn

End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)

   ListView1.HotTracking() = New_HotTracking
   PropertyChanged "HotTracking"

End Property

'Returns/sets whether hot tracking is enabled.
Public Property Get HotTracking() As Boolean

   HotTracking = ListView1.HotTracking

End Property

'Returns/sets whether hover selection is enabled.
Public Property Get HoverSelection() As Boolean

   HoverSelection = ListView1.HoverSelection

End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)

   ListView1.HoverSelection() = New_HoverSelection
   PropertyChanged "HoverSelection"

End Property

Public Property Get hWnd() As String

   hWnd = ListView1.hWnd

End Property

Public Property Get hWndCtrl() As String

   hWndCtrl = UserControl.hWnd

End Property

Public Property Get ImageListCount() As Long

   ImageListCount = ImageList1.ListImages.Count - 1

End Property

Public Property Get IncludeFolder() As Boolean

   IncludeFolder = m_bIncludeFolder

End Property

Public Property Let IncludeFolder(ByVal bIncludeFolder As Boolean)

   m_bIncludeFolder = bIncludeFolder
   PropertyChanged "IncludeFolder"

End Property

'*******************************************************************************
'
'-------------------------------------------------------------------------------

Private Function InitializeImageList( _
                                     ListView1 As MSComctlLib.ListView, ImageList1 As ImageList, pixDummy As PictureBox) As Boolean

   On Local Error GoTo InitializeError

   Set ListView1.SmallIcons = Nothing
   ImageList1.ListImages.Clear
   ImageList1.ListImages.Add , "dummy", pixDummy.Picture
   Set ListView1.SmallIcons = ImageList1

   InitializeImageList = True

Exit Function

InitializeError:

   InitializeImageList = False

End Function

Public Sub InvertAllChecks()

   Dim lv As LV_ITEM
   Dim lvCount As Long
   Dim lvIndex As Long
   Dim r As Long

   lvCount = ListView1.ListItems.Count - 1

   Do

      r = SendMessage(ListView1.hWnd, LVM_GETITEMSTATE, lvIndex, ByVal LVIS_STATEIMAGEMASK)

      With lv
         .mask = LVIF_STATE
         .stateMask = LVIS_STATEIMAGEMASK

         If r And &H2000& Then
            'it is checked, so set the state
            'to 'unchecked'
            .State = &H1000
         Else 'NOT R...
            .State = &H2000
         End If

      End With 'LV

      Call SendMessage(ListView1.hWnd, LVM_SETITEMSTATE, lvIndex, lv)
      lvIndex = lvIndex + 1

   Loop Until lvIndex > lvCount

End Sub

Public Sub InvertSelections()

   Dim i As Long

   With ListView1.ListItems
      For i = 1 To .Count
         With .Item(i)
            .Selected = Not .Selected
         End With '.ITEM(I)
      Next i
   End With 'LISTVIEW1.LISTITEMS

End Sub

'*******************************************************************************
' Modifies a numeric string to allow it to be sorted alphabetically
'-------------------------------------------------------------------------------

Private Function InvNumber(ByVal Number As String) As String

   Static i As Integer

   For i = 1 To Len(Number)
      Select Case Mid$(Number, i, 1)
      Case "-"
         Mid$(Number, i, 1) = " "
      Case "0"
         Mid$(Number, i, 1) = "9"
      Case "1"
         Mid$(Number, i, 1) = "8"
      Case "2"
         Mid$(Number, i, 1) = "7"
      Case "3"
         Mid$(Number, i, 1) = "6"
      Case "4"
         Mid$(Number, i, 1) = "5"
      Case "5"
         Mid$(Number, i, 1) = "4"
      Case "6"
         Mid$(Number, i, 1) = "3"
      Case "7"
         Mid$(Number, i, 1) = "2"
      Case "8"
         Mid$(Number, i, 1) = "1"
      Case "9"
         Mid$(Number, i, 1) = "0"
      End Select
   Next i
   InvNumber = Number

End Function

Public Function IsFlatHeader() As Boolean

   Dim style As Long
   Dim hHeader As Long

   'get the handle to the listview header
   hHeader = SendMessage(ListView1.hWnd, LVM_GETHEADER, 0, ByVal 0&)

   'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)

   IsFlatHeader = Not ((style And HDS_BUTTONS) = HDS_BUTTONS)

End Function

Private Sub ISubclass2_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long)

'**************************************
'Subclassing
'**************************************
   Static pt As POINTAPI
   Static HTI As HD_HITTESTINFO

   Select Case uMsg
   Case WM_LBUTTONUP
      If Not ((GetWindowLong(hWnd, GWL_STYLE) And HDS_BUTTONS) = HDS_BUTTONS) Then
         'Debug.Print Now() & "::WM_LBUTTONUP"
         'get the current cursor position in the header
         Call GetCursorPos(pt)
         Call ScreenToClient(hWnd, pt)

         'get the header's hit-test info
         With HTI
            .flags = HHT_ONHEADER Or HHT_ONDIVIDER
            .pt = pt
         End With 'HTI

         Call SendMessage(hWnd, HDM_HITTEST, 0&, HTI)

         If HTI.iItem > -1 And HTI.iItem < ListView1.ColumnHeaders.Count Then
            'RaiseEvent ColumnClick(ListView1.ColumnHeaders(HTI.iItem + 1))
            Call ListView1_ColumnClick(ListView1.ColumnHeaders(HTI.iItem + 1))
         End If
      End If
   Case Else
   End Select

End Sub

Private Sub ISubclass2_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As eMsg, wParam As Long, lParam As Long)

End Sub

Private Sub ISubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As eMsg, ByVal wParam As Long, ByVal lParam As Long)

'**************************************
'Subclassing
'**************************************

   Static nm As NMHDR
   
   Select Case uMsg
   Case WM_NOTIFY
      Call CopyMemory(nm, ByVal lParam, Len(nm))
      'react to the HDN_ code
      Select Case nm.Code
      Case HDN_ITEMCHANGING
         If m_bHighlightColumn Then
            Call DoHighlightColumn(m_enHighlightColor, m_enDefaultColor, m_lCurrrentHighlightColumn, m_enSizingType)
         End If
         If m_bAllowItemEdit Then '@@ ItemEdit 04/05/17, 16:58:13
            MoveEditBox
         End If
         RaiseEvent HeaderItemChanging
      Case Else
      End Select
   Case WM_VSCROLL
      If m_bAllowItemEdit Then '@@ ItemEdit 04/05/17, 16:58:13
         MoveEditBox
      End If
      RaiseEvent VScroll
   Case WM_HSCROLL
      If m_bAllowItemEdit Then '@@ ItemEdit 04/05/17, 16:58:13
         MoveEditBox
      End If
      RaiseEvent HScroll
   Case WM_MOUSEWHEEL
      If m_bAllowItemEdit Then '@@ ItemEdit 04/05/17, 16:58:13
         MoveEditBox
      End If
      RaiseEvent MouseWheel
   Case Else
   End Select

End Sub

Private Sub ISubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As eMsg, wParam As Long, lParam As Long)

End Sub

Public Property Get ListItems() As ListItems

   Set ListItems = ListView1.ListItems

End Property

Public Property Get ListView() As MSComctlLib.ListView

   Set ListView = ListView1

End Property

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)

   'RaiseEvent AfterLabelEdit(Cancel, NewString)

End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

   'RaiseEvent BeforeLabelEdit(Cancel)

End Sub

Private Sub ListView1_Click()

   RaiseEvent Click

End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)

   With ColumnHeader
      SortListView ListView1, .Index, Val(.Tag), IIf(ListView1.SortOrder = 0, False, True)
   End With 'COLUMNHEADER

   mnuSortAZ.Checked = ListView1.SortOrder = 0
   mnuSortZA.Checked = mnuSortAZ.Checked = False
   
   Dim i As Long
   For i = 0 To 5
      mnuOrder(i).Checked = ((ColumnHeader.Index - 1) = i)
   Next i
   prevOrder = ColumnHeader.Index - 1

   If m_bHighlightColumn Then
      Call DoHighlightColumn(m_enHighlightColor, m_enDefaultColor, ColumnHeader.Index, m_enSizingType)
      DoEvents
   End If
   
   '@@ ItemEdit 04/05/17, 16:58:13
   If m_bAllowItemEdit Then
      MoveEditBox
   End If
   RaiseEvent ColumnClick(ColumnHeader)

End Sub

'Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
'   RaiseEvent ColumnClick(ColumnHeader)
'End Sub

Private Sub ListView1_DblClick()

   RaiseEvent DblClick

End Sub

Private Sub ListView1_ItemCheck(ByVal Item As ListItem)

   RaiseEvent ItemCheck(Item)

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ListItem)

   RaiseEvent ItemClick(Item)

End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)

   RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)

   RaiseEvent KeyUp(KeyCode, Shift)

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

   If m_bAutoTooltip Then
      Dim HTI As LVHITTESTINFO
   
      '----------------------------------------
      'this gets the hittest info using the mouse co-ordinates
      With HTI
         .pt.x = (x \ Screen.TwipsPerPixelX)
         .pt.y = (y \ Screen.TwipsPerPixelY)
         .flags = LVHT_ONITEM
      End With 'HTI
   
      Call SendMessage(ListView1.hWnd, LVM_SUBITEMHITTEST, 0, HTI)
   
      '----------------------------------------
      'this determines whether the hit test returned a main or sub item
      If HTI.iItem >= 0 And HTI.iSubItem >= 0 Then
         On Local Error GoTo Bye
         With ListView1.ListItems(HTI.iItem + 1)
            If HTI.iSubItem = 0 Then
               ListView1.ToolTipText = .Text
            Else
               ListView1.ToolTipText = .SubItems(HTI.iSubItem)
            End If
         End With
      End If
   End If
Bye:

   RaiseEvent MouseMove(Button, Shift, x, y)
   
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

   If Button = 2 Then
      If m_bAutoPopupMenu Then
         PopupMenu mnuOptions
      End If
   End If

   RaiseEvent MouseUp(Button, Shift, x, y)


   Dim HTI As LVHITTESTINFO
  
   '----------------------------------------
   'this gets the hittest info using the mouse co-ordinates
   With HTI
      .pt.x = (x \ Screen.TwipsPerPixelX)
      .pt.y = (y \ Screen.TwipsPerPixelY)
      .flags = LVHT_ONITEM
   End With 'HTI

   Call SendMessage(ListView1.hWnd, LVM_SUBITEMHITTEST, 0, HTI)

   '----------------------------------------
   'this determines whether the hit test returned a main or sub item
   If HTI.iItem >= 0 And HTI.iSubItem >= 0 Then
      On Local Error GoTo Bye
      '----------------------------------------
      'this selects the current item if the ontrol's FullRowSelect is False, and
      'a SubItem was clicked. (One is added because the API is 0-based, and the
      'ListItems collection is 1-based).
      If HTI.iSubItem >= 0 And m_bSubitemSelect Then
         ListView1.ListItems(HTI.iItem + 1).Selected = True
         
         '@@ ItemEdit 04/05/17, 16:58:13
         txtEdit.Tag = HTI.iItem + 1 & ":" & HTI.iSubItem
         If m_bAllowItemEdit Then '@@ ItemEdit 04/05/17, 16:58:13
            MoveEditBox
         End If
         DoEvents
         
         RaiseEvent SubitemClick(HTI.iItem + 1, HTI.iSubItem, Button, Shift)
      End If

   ElseIf m_bAutoDeselect Then 'NOT HTI.IITEM...

      Set ListView1.SelectedItem = Nothing

   End If

Bye:
End Sub

'Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'   RaiseEvent MouseUp(Button, Shift, x, y)
'End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
   
   Dim strFileSpec As String
   Dim i As Long
   Dim numFiles As Long
   Dim iFiles As Long, iFolders As Long
   
   If m_bAllowFileDragDrop Then
      numFiles = Data.Files.Count
      For i = 1 To numFiles
         If (GetAttr(Data.Files(i)) And vbDirectory) = vbDirectory Then
           'InitializeImageList ListView1, ImageList1, pixDummy
           iFolders = iFolders + 1
           subFindFiles Data.Files(i), m_strFileSpec, m_bRecursive, True, m_bIncludeFolder
           LockWindowUpdate 0&
         Else
            strFileSpec = Replace(AfterRev(m_strFileSpec, "."), "*", "")
            iFiles = iFiles + 1
            If Len(strFileSpec) = 0 Then
               AddFile FileNamePart(Data.Files(i), efpPath), FileNamePart(Data.Files(i), efpName)
            ElseIf InStr(1, FileNamePart(Data.Files(i), efpExtension), strFileSpec, vbTextCompare) > 0 Then
               AddFile FileNamePart(Data.Files(i), efpPath), FileNamePart(Data.Files(i), efpName)
           End If
         End If
      Next i
      
      RaiseEvent FileDragDropFinish(numFiles, iFiles, iFolders)
      
   End If
   RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
   Exit Sub

End Sub

Private Sub mnuAddItem_Click()

   RaiseEvent AddItemRequested

End Sub

Private Sub mnuAutoDeselect_Click()

   mnuAutoDeselect.Checked = Not mnuAutoDeselect.Checked
   m_bAutoDeselect = mnuAutoDeselect.Checked

End Sub

Private Sub mnuAutoTooltip_Click()
   Me.AutoTooltip = Not mnuAutoTooltip.Checked
End Sub

Private Sub mnuCheck_Click(Index As Integer)

   Select Case Index
   Case 0
      Me.CheckedAll = True
   Case 1
      Me.CheckedAll = False
   Case 2
      Me.InvertAllChecks
   End Select

End Sub

Private Sub mnuCheckboxes_Click()

   Me.Checkboxes = Not mnuCheckboxes.Checked

End Sub

Private Sub mnuClear_Click()

   ListView1.ListItems.Clear

End Sub

Private Sub mnuColumnHighlight_Click()

   mnuColumnHighlight.Checked = Not mnuColumnHighlight.Checked
   Me.HighlightColumn = mnuColumnHighlight.Checked

End Sub

Private Sub mnuFlatAppearance_Click()
   mnuFlatAppearance.Checked = Not mnuFlatAppearance.Checked
   Me.Appearance = IIf(mnuFlatAppearance.Checked, ccFlat, cc3D)
End Sub

Private Sub mnuFlatHeader_Click()

   mnuFlatHeader.Checked = Not mnuFlatHeader.Checked
   FlatHeader = mnuFlatHeader.Checked

End Sub

Private Sub mnuFlatScrollBar_Click()

   mnuFlatScrollBar.Checked = Not mnuFlatScrollBar.Checked
   ListView1.FlatScrollBar() = mnuFlatScrollBar.Checked

End Sub

Private Sub mnuFullRowSelect_Click()

   Me.FullRowSelect = Not mnuFullRowSelect.Checked

End Sub

Private Sub mnuGridLines_Click()

   mnuGridLines.Checked = Not mnuGridLines.Checked
   ListView1.GridLines() = mnuGridLines.Checked

End Sub

Private Sub mnuItemEdit_Click()
   AllowItemEdit = Not mnuItemEdit.Checked
End Sub

Private Sub mnuMultiSelect_Click()

   mnuMultiSelect.Checked = Not mnuMultiSelect.Checked
   ListView1.MultiSelect = mnuMultiSelect.Checked

End Sub

Private Sub mnuOrder_Click(Index As Integer)

   Dim i As Long

   For i = 0 To 5
      mnuOrder(i).Checked = (Index = i)
   Next i

   ListView1_ColumnClick ListView1.ColumnHeaders(Index + 1)
   prevOrder = Index - 1

End Sub

Private Sub mnuRemoveCheckedItems_Click()

   Dim i As Long

   LockWindowUpdate ListView1.hWnd
   With ListView1.ListItems
      For i = .Count To 1 Step -1
         If .Item(i).Checked Then
            .Remove i
         End If
      Next i
   End With 'LISTVIEW1.LISTITEMS
   LockWindowUpdate 0&

End Sub

Private Sub mnuRemoveItems_Click()

   Dim i As Long

   LockWindowUpdate ListView1.hWnd
   With ListView1.ListItems
      For i = .Count To 1 Step -1
         If .Item(i).Selected Then
            .Remove i
         End If
      Next i
   End With 'LISTVIEW1.LISTITEMS
   LockWindowUpdate 0&

End Sub

Private Sub mnuResize_Click(Index As Integer)

   Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, Index, ByVal LVSCW_AUTOSIZE)

End Sub

Private Sub mnuResizeColumns_Click()

   Dim ColumnIndex As Long

   For ColumnIndex = 0 To ListView1.ColumnHeaders.Count - 1
      Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex, ByVal LVSCW_AUTOSIZE)
   Next ColumnIndex

End Sub

Private Sub mnuResizeHeader_Click(Index As Integer)

   Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, Index, ByVal LVSCW_AUTOSIZE_USEHEADER)

End Sub

Private Sub mnuResizeHeaders_Click()

   Dim ColumnIndex As Long

   For ColumnIndex = 0 To ListView1.ColumnHeaders.Count - 1
      Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next ColumnIndex

End Sub

Private Sub mnuSelect_Click(Index As Integer)

   Select Case Index
   Case 0
      Me.SelectedAll = True
   Case 1
      Me.SelectedAll = False
   Case 2
      Me.InvertSelections
   End Select

End Sub

Private Sub mnuSolidBorder_Click()
   BorderStyle = Abs(Not mnuSolidBorder.Checked)
End Sub

Private Sub mnuSortAZ_Click()

   With ListView1.ColumnHeaders(prevOrder + 1)
      SortListView ListView1, .Index, Val(.Tag), True
   End With 'LVFILELIST.COLUMNHEADERS(PREVORDER'LISTVIEW1.COLUMNHEADERS(PREVORDER

   mnuSortAZ.Checked = True
   mnuSortZA.Checked = False

End Sub

Private Sub mnuSortZA_Click()

   With ListView1.ColumnHeaders(prevOrder + 1)
      SortListView ListView1, .Index, Val(.Tag), False
   End With 'LVFILELIST.COLUMNHEADERS(PREVORDER'LISTVIEW1.COLUMNHEADERS(PREVORDER

   mnuSortAZ.Checked = False
   mnuSortZA.Checked = True

End Sub

Private Sub mnuSubiitemSelect_Click()

   mnuSubiitemSelect.Checked = Not mnuSubiitemSelect.Checked
   Me.SubitemSelect = mnuSubiitemSelect.Checked

End Sub

Private Sub mnuView_Click(Index As Integer)

   mnuView(ListView1.View - 2).Checked = False
   mnuView(Index).Checked = True

   ListView1.View = Index + 2
   ListView1.Sorted = True

End Sub

Public Property Get MousePointer() As MousePointerConstants
   MousePointer = ListView1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
   ListView1.MousePointer() = New_MousePointer
   PropertyChanged "MousePointer"
End Property

'Returns/sets a value indicating whether a user can make multiple selections in the MSComctlLib.ListView control and how the multiple selections can be made.
Public Property Get MultiSelect() As Boolean

   MultiSelect = ListView1.MultiSelect

End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)

   ListView1.MultiSelect() = New_MultiSelect
   mnuMultiSelect.Checked = New_MultiSelect
   PropertyChanged "MultiSelect"

End Property

Public Function NextItem(Optional Index As Long = -1, _
                         Optional flags As Long = LVNI_SELECTED) As Long

   NextItem = SendMessage(ListView1.hWnd, LVM_GETNEXTITEM, Index, ByVal flags)

End Function

Public Property Get Path() As String

   Path = m_strPath

End Property

Public Property Let Path(ByVal strPath As String)

   m_strPath = strPath
   PropertyChanged "Path"

End Property

Public Property Get Picture() As stdole.Picture

   Set Picture = ListView1.Picture

End Property

Public Property Let Picture(ByVal objPicture As stdole.Picture)

'Returns/sets the background picture for the control.

   ListView1.Picture = objPicture

End Property

Public Property Set Picture(ByVal objPicture As stdole.Picture)

'Returns/sets the background picture for the control.

   Set ListView1.Picture = objPicture

End Property

Public Property Let PictureAlignment(ByVal enmPictureAlignment As ListPictureAlignmentConstants)

'Returns/sets the picture alignment.

   ListView1.PictureAlignment = enmPictureAlignment

End Property

Public Property Get PictureAlignment() As ListPictureAlignmentConstants

'Returns/sets the picture alignment.

   PictureAlignment = ListView1.PictureAlignment

End Property

Public Property Let Recursive(ByVal bRecursive As Boolean)

   m_bRecursive = bRecursive
   PropertyChanged "Recursive"

End Property

Public Property Get Recursive() As Boolean

   Recursive = m_bRecursive

End Property

Public Sub Refresh()

   ListView1.Refresh

End Sub

Public Sub ResizeListViewColumn(ColumnIndex As Long)

   Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex, ByVal LVSCW_AUTOSIZE)

End Sub

Public Sub ResizeListViewColumnsByText()

   Dim ColumnIndex As Long

   For ColumnIndex = 0 To ListView1.ColumnHeaders.Count - 1
      Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex, ByVal LVSCW_AUTOSIZE)
   Next ColumnIndex

End Sub

Public Sub ResizeListViewHeader(ColumnIndex As Long)

   Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex, ByVal LVSCW_AUTOSIZE_USEHEADER)

End Sub

Public Sub ResizeListViewHeaders()

   Dim ColumnIndex As Long

   For ColumnIndex = 0 To ListView1.ColumnHeaders.Count - 1
      Call SendMessage(ListView1.hWnd, LVM_SETCOLUMNWIDTH, ColumnIndex, ByVal LVSCW_AUTOSIZE_USEHEADER)
   Next ColumnIndex

End Sub

Public Sub Scan()

   If m_strPath > "" And m_strFileSpec > "" Then
      'InitializeImageList ListView1, ImageList1, pixDummy
      Dim Cancel As Boolean
      RaiseEvent BeforeScanStart(Cancel, m_strPath, m_strFileSpec, m_bRecursive, m_bIncludeFolder)
      If Not Cancel Then
         subFindFiles m_strPath, m_strFileSpec, m_bRecursive, True, m_bIncludeFolder
         LockWindowUpdate 0&
         RaiseEvent ScanFinish(m_bStop, m_strPath, m_strFileSpec, m_bRecursive, m_bIncludeFolder)
      End If
   End If
   
End Sub

Public Function SelectDirectory(Optional DialogTitle As String = "Select the search directory for ", _
                                                Optional ByVal OwnerHwnd As Long) As String

   If OwnerHwnd = 0 Then
      OwnerHwnd = ListView1.hWnd
   End If
   SelectDirectory = subSelectDirectory(OwnerHwnd, DialogTitle & m_strFileSpec & ".")

End Function

Public Property Let SelectedAll(ByVal bState As Boolean)

   Dim lv As LV_ITEM

   With lv
      .mask = LVIF_STATE
      .State = bState
      .stateMask = LVIS_SELECTED
   End With 'LV

   'by setting wParam to -1, the call affects all
   'listitems. To just change a particular item,
   'pass its index as wParam.
   Call SendMessage(ListView1.hWnd, LVM_SETITEMSTATE, -1, lv)

End Property

Public Function SelectedCount() As Long

   SelectedCount = SendMessage(ListView1.hWnd, LVM_GETSELECTEDCOUNT, 0&, ByVal 0&)

End Function

Public Property Get SelectedItem() As IListItem

'Returns a reference to the currently selected ListItem or Node object.

   Set SelectedItem = ListView1.SelectedItem

End Property

Public Property Set SelectedItem(ByVal objSelectedItem As IListItem)

'Returns a reference to the currently selected ListItem or Node object.

   Set ListView1.SelectedItem = objSelectedItem

End Property

Public Function GetSelectedItems(nSelected() As Long) As Long

'Const LVNI_SELECTED = &H2

'Dim nSelected() As Long
   Dim Index As Long
   Dim numSelected As Long
   Dim cnt As Long
   Dim hWnd As Long

   hWnd = ListView1.hWnd
   numSelected = SendMessage(hWnd, LVM_GETSELECTEDCOUNT, 0&, ByVal 0&)

   'Debug.Print "numSelected=" & numSelected

   If numSelected <> 0 Then

      Index = -1
      ReDim nSelected(0 To numSelected - 1)

      Do

         'Get the next selected item
         Index = SendMessage(hWnd, LVM_GETNEXTITEM, Index, ByVal LVNI_SELECTED)

         If Index > -1 Then
            'Debug.Print "index + 1=" & index + 1
            nSelected(cnt) = Index + 1
            cnt = cnt + 1
         End If

      Loop Until Index = -1

      'debug only: print results to the list
      'For cnt = 0 To numSelected - 1
      '   Debug.Print nSelected(cnt) & "::" & ListView1.ListItems(nSelected(cnt)).Text
      'Next

   End If

   GetSelectedItems = numSelected

End Function

Private Sub Set3DHeader()

   Dim style As Long
   Dim hHeader As Long
   Dim bIs3DHeader As Boolean

   'get the handle to the listview header
   hHeader = SendMessage(ListView1.hWnd, LVM_GETHEADER, 0, ByVal 0&)

   'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)

   bIs3DHeader = ((style And HDS_BUTTONS) = HDS_BUTTONS)

   If Not bIs3DHeader Then

      'modify the style by toggling the HDS_BUTTONS style
      style = style Xor HDS_BUTTONS

      'set the new style and redraw the listview
      If style Then
         Call SetWindowLong(hHeader, GWL_STYLE, style)
         Call SetWindowPos(ListView1.hWnd, UserControl.hWnd, 0, 0, 0, 0, SWP_FLAGS)
      End If
   End If

End Sub

Private Sub SetFlatHeader()

   Dim style As Long
   Dim hHeader As Long
   Dim bIs3DHeader As Boolean

   'get the handle to the listview header
   hHeader = SendMessage(ListView1.hWnd, LVM_GETHEADER, 0, ByVal 0&)

   'get the current style attributes for the header
   style = GetWindowLong(hHeader, GWL_STYLE)

   bIs3DHeader = ((style And HDS_BUTTONS) = HDS_BUTTONS)

   If bIs3DHeader Then

      'modify the style by toggling the HDS_BUTTONS style
      style = style Xor HDS_BUTTONS

      'set the new style and redraw the listview
      If style Then
         Call SetWindowLong(hHeader, GWL_STYLE, style)
         Call SetWindowPos(ListView1.hWnd, UserControl.hWnd, 0, 0, 0, 0, SWP_FLAGS)
      End If
   End If

End Sub

Private Sub SetHighlightColumn(lv As MSComctlLib.ListView, _
                               clrHighlight As eLedgerColours, _
                               clrDefault As eLedgerColours, _
                               nColumn As Long, _
                               nSizingType As eImageSizingTypes, _
                               Picture1 As PictureBox)

   Dim cnt     As Long  '/* counter
   Dim cl      As Long  '/* columnheader left
   Dim cw      As Long  '/* columnheader width

   On Local Error GoTo SetHighlightColumn_Error

   If lv.View = lvwReport Then

      '/* set up the listview properties
      With lv
         .Picture = Nothing  '/* clear picture
         .Refresh
         .Visible = 1
         .PictureAlignment = lvwTile
      End With  ' lv'LV

      '/* set up the picture box properties
      With Picture1
         .AutoRedraw = False       '/* clear/reset picture
         .Picture = Nothing
         .BackColor = clrDefault
         .Height = 1
         .AutoRedraw = True        '/* assure image draws
         .BorderStyle = vbBSNone   '/* other attributes
         .ScaleMode = vbTwips
         .Top = .Top - 10000  '/* move it off screen
         .Visible = False
         .Height = 1               '/* only need a 1 pixel high picture
         .Width = Screen.Width

         '/* draw a box in the highlight colour
         '/* at location of the column passed
         With lv.ColumnHeaders(nColumn)
            cl = .Left
            cw = .Left + .Width
         End With 'LV.COLUMNHEADERS(NCOLUMN)
         Picture1.Line (cl, 0)-(cw, 210), clrHighlight, BF

         .AutoSize = True
      End With  'Picture1

      '/* set the lv picture to the
      '/* Picture1 image
      lv.Refresh
      lv.Picture = Picture1.Image

   Else 'NOT LV.VIEW...

      lv.Picture = Nothing

   End If  'lv.View = lvwReport

SetHighlightColumn_Exit:
   On Local Error GoTo 0

Exit Sub

SetHighlightColumn_Error:

   '/* clear the listview's picture and exit
   With lv
      .Picture = Nothing
      .Refresh
   End With 'LV

   Resume SetHighlightColumn_Exit

End Sub

Public Property Let SizingType(ByVal enSizingType As eImageSizingTypes)

   m_enSizingType = enSizingType
   If m_bHighlightColumn Then
      HighlightColumn = True
   End If
   'Call DoHighlightColumn(m_enHighlightColor, m_enDefaultColor, m_lCurrrentHighlightColumn, m_enSizingType)
   PropertyChanged "SizingType"

End Property

Public Property Get SizingType() As eImageSizingTypes

   SizingType = m_enSizingType

End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)

   ListView1.Sorted() = New_Sorted
   PropertyChanged "Sorted"

End Property

'Indicates whether the elements of a control are automatically sorted alphabetically.
Public Property Get Sorted() As Boolean

   Sorted = ListView1.Sorted

End Property

Public Property Let SortKey(ByVal intSortKey As Integer)

'Returns/sets the current sort key.

   ListView1.SortKey = intSortKey
   PropertyChanged "SortKey"

End Property

Public Property Get SortKey() As Integer

'Returns/sets the current sort key.

   SortKey = ListView1.SortKey

End Property

'*******************************************************************************
' Sort a MSComctlLib.ListView1 by String, Number, or DateTime
'
' Parameters:
'
'   MSComctlLib.ListView1    Reference to the MSComctlLib.ListView1 control to be sorted.
'   ColumnIndex       ColumnIndex of the column in the MSComctlLib.ListView1 to be sorted. The first
'               column in a MSComctlLib.ListView1 has an index value of 1.
'   DataType    Sets whether the data in the column is to be sorted
'               alphabetically, numerically, or by date.
'   Ascending   Sets the direction of the sort. True sorts A-Z (Ascending),
'               and False sorts Z-A (descending)
'-------------------------------------------------------------------------------

Private Sub SortListView(ListView1 As MSComctlLib.ListView, ByVal ColumnIndex As Integer, _
                         ByVal DataType As eListColumnDataType, ByVal Ascending As Boolean)

   On Error Resume Next
      Dim i As Integer
      Dim l As Long
      Dim strFormat As String

      ' Display the hourglass cursor whilst sorting

      Dim lngCursor As Long
      lngCursor = ListView1.MousePointer
      ListView1.MousePointer = vbHourglass

      ' Prevent the MSComctlLib.ListView1 control from updating on screen - this is to hide
      ' the changes being made to the listitems, and also to speed up the sort

      LockWindowUpdate ListView1.hWnd

      Dim blnRestoreFromTag As Boolean

      Select Case DataType
      Case ldtString

         ' Sort alphabetically. This is the only sort provided by the
         ' MS MSComctlLib.ListView1 control (at this time), and as such we don't really
         ' need to do much here

         blnRestoreFromTag = False
      Case ldtNumber

         ' Sort Numerically

         strFormat = String$(20, "0") & "." & String$(10, "0")

         ' Loop through the values in this column. Re-format the values so
         ' as they can be sorted alphabetically, having already stored their
         ' text values in the tag, along with the tag's original value

         With ListView1.ListItems
            If (ColumnIndex = 1) Then
               For l = 1 To .Count
                  With .Item(l)
                     .Tag = .Text & Chr$(0) & .Tag
                     If IsNumeric(.Text) Then
                        If CDbl(.Text) >= 0 Then
                           .Text = Format$(CDbl(.Text), strFormat)
                        Else 'NOT CDBL(.TEXT)...
                           .Text = "&" & InvNumber(Format$(0 - CDbl(.Text), strFormat))
                        End If
                     Else 'ISNUMERIC(.TEXT) = FALSE
                        .Text = Replace(.Text, ",", "")
                        .Text = Format$(Val(.Text), strFormat)
                     End If
                  End With '.ITEM(L)
               Next l
            Else 'NOT (COLUMNINDEX...
               For l = 1 To .Count
                  With .Item(l).ListSubItems(ColumnIndex - 1)
                     .Tag = .Text & Chr$(0) & .Tag
                     If IsNumeric(.Text) Then
                        If CDbl(.Text) >= 0 Then
                           .Text = Format$(CDbl(.Text), strFormat)
                        Else 'NOT CDBL(.TEXT)...
                           .Text = "&" & InvNumber(Format$(0 - CDbl(.Text), strFormat))
                        End If
                     Else 'ISNUMERIC(.TEXT) = FALSE
                        .Text = Replace(.Text, ",", "")
                        .Text = Format$(Val(.Text), strFormat)
                     End If
                  End With '.ITEM(L).LISTSUBITEMS(COLUMNINDEX
               Next l
            End If
         End With 'LISTVIEW.LISTITEMS'LISTVIEW1.LISTITEMS

         blnRestoreFromTag = True

      Case ldtDateTime

         ' Sort by date.

         strFormat = "YYYYMMDDHhNnSs"

         Dim dte As Date

         ' Loop through the values in this column. Re-format the dates so as they
         ' can be sorted alphabetically, having already stored their visible
         ' values in the tag, along with the tag's original value

         With ListView1.ListItems
            If (ColumnIndex = 1) Then
               For l = 1 To .Count
                  With .Item(l)
                     .Tag = .Text & Chr$(0) & .Tag
                     dte = CDate(.Text)
                     .Text = Format$(dte, strFormat)
                  End With '.ITEM(L)
               Next l
            Else 'NOT (COLUMNINDEX...
               For l = 1 To .Count
                  With .Item(l).ListSubItems(ColumnIndex - 1)
                     .Tag = .Text & Chr$(0) & .Tag
                     dte = CDate(.Text)
                     .Text = Format$(dte, strFormat)
                  End With '.ITEM(L).LISTSUBITEMS(COLUMNINDEX
               Next l
            End If
         End With 'LISTVIEW.LISTITEMS'LISTVIEW1.LISTITEMS

         blnRestoreFromTag = True

      End Select

      ' Sort the MSComctlLib.ListView1 Alphabetically

      ListView1.SortOrder = IIf(Ascending, lvwAscending, lvwDescending)
      ListView1.SortKey = ColumnIndex - 1
      ListView1.Sorted = True

      ' Restore the Text Values if required

      If blnRestoreFromTag Then

         ' Restore the previous values to the 'cells' in this column of the list
         ' from the tags, and also restore the tags to their original values

         With ListView1.ListItems
            If (ColumnIndex = 1) Then
               For l = 1 To .Count
                  With .Item(l)
                     i = InStr(.Tag, Chr$(0))
                     .Text = Left$(.Tag, i - 1)
                     .Tag = Mid$(.Tag, i + 1)
                  End With '.ITEM(L)
               Next l
            Else 'NOT (COLUMNINDEX...
               For l = 1 To .Count
                  With .Item(l).ListSubItems(ColumnIndex - 1)
                     i = InStr(.Tag, Chr$(0))
                     .Text = Left$(.Tag, i - 1)
                     .Tag = Mid$(.Tag, i + 1)
                  End With '.ITEM(L).LISTSUBITEMS(COLUMNINDEX
               Next l
            End If
         End With 'LISTVIEW.LISTITEMS'LISTVIEW1.LISTITEMS
      End If

      ' Unlock the list window so that the OCX can update it

      LockWindowUpdate 0&

      ' Restore the previous cursor

      ListView1.MousePointer = lngCursor

   On Error GoTo 0

End Sub

Public Property Let SortOrder(ByVal enmSortOrder As ListSortOrderConstants)

'Returns/sets whether or not the ListItems will be sorted in ascending or descending order.

   ListView1.SortOrder = enmSortOrder
   PropertyChanged "SortOrder"

End Property

Public Property Get SortOrder() As ListSortOrderConstants

'Returns/sets whether or not the ListItems will be sorted in ascending or descending order.

   SortOrder = ListView1.SortOrder

End Property

Public Property Get Stopped() As Boolean

   Stopped = m_bStop

End Property

Public Sub ResetStopFlag()

   m_bStop = False

End Sub

Public Sub StopScan()

   m_bStop = True

End Sub

Private Function subFindFiles( _
                              ByVal sDirSpec As String, _
                              ByVal sFileSpec As String, _
                              Optional ByVal bRecursive As Boolean = True, _
                              Optional bIsNewSearch As Boolean = True, _
                              Optional bIncludeFolder As Boolean = False) As Long

   Dim hFile As Long, hMatch As Long
   Dim wfd As WIN32_FIND_DATA
   Dim strDirName As String
   Static lFound As Long

   'Static m_bStop As Boolean
   Static counter As Long

   If bIsNewSearch Then
      lFound = 0
      counter = 0
      ListView1.ListItems.Clear
      m_bStop = False
      LockWindowUpdate ListView1.hWnd
      DoEvents
   End If

   RaiseEvent SearchDirChange(sDirSpec)
   sDirSpec = subQualifyPath(sDirSpec)

   'Scan Subdirs First
   If bRecursive Then
      'Set the search directory and Raise event
      hFile = FindFirstFile(sDirSpec & "*.*", wfd)
      hMatch = 99
      Do While hFile > 0 And hMatch > 0
         DoEvents
         If m_bStop Then
            Exit Do '>---> Loop
         End If
         If (wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) > 0 Then
            '-- Declare ¹öÀü: cFileName As String * 260
            strDirName = subTrimNull(wfd.cFileName)
            If strDirName <> "." And strDirName <> ".." Then
               subFindFiles subQualifyPath(sDirSpec & strDirName), sFileSpec, bRecursive, False, bIncludeFolder
            End If
         End If
         hMatch = FindNextFile(hFile, wfd)
      Loop
      FindClose hFile
   End If

   DoEvents
   If Not m_bStop Then

      Dim bFolderAdded As Boolean
      hFile = FindFirstFile(sDirSpec & sFileSpec, wfd)
      hMatch = 99
      Do While hFile > 0 And hMatch > 0
         DoEvents
         If m_bStop Then
            Exit Do '>---> Loop
         End If
         If Not (wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) > 0 Then
            If bIncludeFolder Then
               If Not bFolderAdded Then
                  lFound = lFound + 1
                  AddFile Left$(sDirSpec, Len(sDirSpec) - 1), ""
                  'ISearchFile.FileFound
                  bFolderAdded = True
               End If
            End If
            lFound = lFound + 1
            AddFile sDirSpec, subTrimNull(wfd.cFileName)
            counter = counter + 1

            If counter = UpdateFrequency Then
               LockWindowUpdate 0
               Call UpdateWindow(ListView1.hWnd)
               LockWindowUpdate ListView1.hWnd
               counter = 0
            End If
            'ISearchFile.FileFound
         End If
         hMatch = FindNextFile(hFile, wfd)
      Loop

   End If

Exit_Proc:
   FindClose hFile

   subFindFiles = lFound

End Function

Public Property Let SubitemSelect(ByVal bSubitemSelect As Boolean)

   m_bSubitemSelect = bSubitemSelect
   PropertyChanged "SubitemSelect"

End Property

Public Property Get SubitemSelect() As Boolean

   SubitemSelect = m_bSubitemSelect

End Property

Private Function subQualifyPath(Path As String) As String

'Add a backslash to the path
'assures passed path ends in a slash

   If VBA.Right$(Path, 1) <> "\" Then
      subQualifyPath = Path & "\"
   Else 'NOT VBA.RIGHT$(PATH,...
      subQualifyPath = Path
   End If

End Function

Private Function subSelectDirectory(OwnerHwnd As Long, DialogTitle As String) As String

   Dim bi As BROWSEINFO
   Dim IDL As ITEMIDLIST
   Dim pidl As Long
   Dim tmpPath As String
   Dim pos As Integer

   bi.hOwner = OwnerHwnd
   bi.pidlRoot = 0&
   bi.lpszTitle = DialogTitle
   bi.ulFlags = BIF_RETURNONLYFSDIRS

   'get the folder
   pidl = SHBrowseForFolder(bi)

   tmpPath = Space$(MAX_PATH)

   If SHGetPathFromIDList(ByVal pidl, ByVal tmpPath) Then
      pos = InStr(tmpPath, Chr$(0))
      tmpPath = Left$(tmpPath, pos - 1)
      If Right$(tmpPath, 1) = "\" Then
         subSelectDirectory = tmpPath
      Else 'NOT RIGHT$(TMPPATH,...
         subSelectDirectory = tmpPath & "\"
      End If
   Else 'NOT SHGETPATHFROMIDLIST(BYVAL...
      subSelectDirectory = ""
   End If

   Call CoTaskMemFree(pidl)

End Function

Private Function subTrimNull(StrIn As String) As String

   Dim nul As Long

   'Truncates the input string at first null. If no nulls, perform ordinary Trim.
   nul = VBA.InStr(StrIn, vbNullChar)
   Select Case nul
   Case Is > 1
      subTrimNull = VBA.Left$(StrIn, nul - 1)
   Case 1
      subTrimNull = ""
   Case 0
      subTrimNull = VBA.Trim$(StrIn)
   End Select

End Function

Public Property Let TextBackground(ByVal New_TextBackground As ListTextBackgroundConstants)

   ListView1.TextBackground() = New_TextBackground
   PropertyChanged "TextBackground"

End Property

'Returns/sets a value that determines if the text background is transparent or uses the MSComctlLib.ListView background color
Public Property Get TextBackground() As ListTextBackgroundConstants

   TextBackground = ListView1.TextBackground

End Property

Public Property Let TopIndex(ByVal Index As Long)

   Dim lvItemsPerPage As Long
   Dim lvNeededItems As Long
   Dim lvCurrentTopIndex As Long
   Dim hWnd As Long

   'determine if desired index + number
   'of items in view will exceed total
   'items in the control
   hWnd = ListView1.hWnd
   lvCurrentTopIndex = SendMessage(hWnd, LVM_GETTOPINDEX, 0&, ByVal 0&) + 1 '0-based!
   lvItemsPerPage = SendMessage(hWnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&)

   lvNeededItems = (Index - lvItemsPerPage)

   'is current index above or below
   'desired index?
   If lvCurrentTopIndex > Index Then

      'it is above the desired index, so
      'scroll up. The item will automatically
      'be positioned at the top
      ListView1.ListItems((Index)).EnsureVisible

   ElseIf (Index - lvCurrentTopIndex) >= lvItemsPerPage Then 'NOT LVCURRENTTOPINDEX...

      'it's below, so based on whether there
      'are sufficient items to set to the topindex ...
      If (Index + lvItemsPerPage) > ListView1.ListItems.Count Then

         'it is below but it can't be set to
         'the top as the control has insufficient
         'items, so just scroll to the end of listview
         ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible

      Else 'NOT (INDEX...

         'it is below, and since a listview
         'always moves the item just into view,
         'have it instead move to the top by
         'faking item we want to 'EnsureVisible'
         'the item lvItemsPerPage -1 below the actual
         'index of interest.
         ListView1.ListItems((Index + lvItemsPerPage) - 1).EnsureVisible

      End If

   End If

End Property

Public Property Get TopIndex() As Long

   TopIndex = SendMessage(ListView1.hWnd, LVM_GETTOPINDEX, 0&, ByVal 0&) + 1

End Property

Public Property Let UpdateFrequency(ByVal lUpdateFrequency As Long)

   m_lUpdateFrequency = lUpdateFrequency
   PropertyChanged "UpdateFrequency"

End Property

Public Property Get UpdateFrequency() As Long

   UpdateFrequency = m_lUpdateFrequency

End Property

Private Sub UserControl_Initialize()

'# used to highlight a column

   Picture1.Top = -1000
   
   With ListView1
'      With .ColumnHeaders
'         .Add(, , "FullName", 1300, 0).Tag = 0 'Left Alignment, String
'         .Add(, , "Path", 1000, 0).Tag = 0 'Left Alignment, String
'         .Add(, , "Name", 1400, 0).Tag = 0 'Left Alignment, String
'         .Add(, , "Size", 800, 1).Tag = 1 'Right Alignment, Number
'         .Add(, , "Type", 1000, 0).Tag = 0  'Left Alignment, String
'         .Add(, , "Created Date", 1300, 0).Tag = 2  'Left Alignment, Date
'      End With '.COLUMNHEADERS
      .SmallIcons = ImageList1
   End With 'LVFILELIST'LISTVIEW1
   
   InitializeImageList ListView1, ImageList1, pixDummy

End Sub

'»ç¿ëÀÚ Á¤ÀÇ ÄÁÆ®·Ñ¿¡ ´ëÇÑ ¼Ó¼ºÀ» ÃÊ±âÈ­ÇÕ´Ï´Ù.
Private Sub UserControl_InitProperties()

   m_bAutoDeselect = m_def_AutoDeselect
   m_strFileSpec = m_def_FileSpec
   m_bIncludeFolder = m_def_IncludeFolder
   m_strPath = m_def_Path
   m_bRecursive = m_def_Recursive
   m_bSubitemSelect = m_def_SubitemSelect
   m_lUpdateFrequency = m_def_UpdateFrequency
   m_enHighlightColor = m_def_HighlightColor
   m_enDefaultColor = m_def_DefaultColor
   m_enSizingType = m_def_SizingType
   m_lCurrrentHighlightColumn = m_def_CurrrentHighlightColumn
   m_bHighlightColumn = m_def_HighlightColumn
   m_bFlatHeader = m_def_FlatHeader
   m_bAutoTooltip = m_def_AutoTooltip
   m_bAutoPopupMenu = m_def_AutoPopupMenu
   m_bAllowItemEdit = m_def_AllowItemEdit '@@ ItemEdit 04/05/17, 16:58:13
   m_bAllowFileDragDrop = m_def_AllowFileDragDrop
   m_strColumnHeadersText = m_def_ColumnHeadersText
   
   Dim ColHeaders As ColHeaders
   Set ColHeaders = New ColHeaders
   With ColHeaders
      .CreateItemsFromTagText m_strColumnHeadersText, vbCrLf, "|"
      .SetColumnHeaders ListView1.ColumnHeaders
   End With
   Set ColHeaders = Nothing
   
   'With ListView1
   '   With .ColumnHeaders
   '      .Add(, , "FullName", 1300, 0).Tag = 0 'Left Alignment, String
   '      .Add(, , "Path", 1000, 0).Tag = 0 'Left Alignment, String
   '      .Add(, , "Name", 1400, 0).Tag = 0 'Left Alignment, String
   '      .Add(, , "Size", 800, 1).Tag = 1 'Right Alignment, Number
   '      .Add(, , "Type", 1000, 0).Tag = 0  'Left Alignment, String
   '      .Add(, , "Created Date", 1300, 0).Tag = 2  'Left Alignment, Date
   '   End With '.COLUMNHEADERS
   'End With 'LVFILELIST'LISTVIEW1
   
End Sub

'ÀúÀå¼Ò¿¡¼­ ¼Ó¼º°ªÀ» ·ÎµåÇÕ´Ï´Ù.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

   ListView1.Appearance = PropBag.ReadProperty("Appearance", 1)
   mnuFlatAppearance.Checked = ListView1.Appearance = ccFlat
   
   ListView1.AllowColumnReorder = PropBag.ReadProperty("AllowColumnReorder", True)
   ListView1.Arrange = PropBag.ReadProperty("Arrange", 0)
   ListView1.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
   mnuSolidBorder.Checked = ListView1.BorderStyle
   'UserControl.BorderStyle = ListView1.BorderStyle
   
   ListView1.Checkboxes = PropBag.ReadProperty("Checkboxes", False)
   zzzmnuCheck.Enabled = ListView1.Checkboxes
   mnuCheckboxes.Checked = ListView1.Checkboxes
   mnuRemoveCheckedItems.Enabled = ListView1.Checkboxes
   
   ListView1.Enabled = PropBag.ReadProperty("Enabled", True)
   UserControl.Enabled = ListView1.Enabled
   
   ListView1.FlatScrollBar = PropBag.ReadProperty("FlatScrollBar", False)
   mnuFlatScrollBar.Checked = ListView1.FlatScrollBar

   ListView1.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
   mnuFullRowSelect.Checked = ListView1.FullRowSelect

   ListView1.GridLines = PropBag.ReadProperty("GridLines", False)
   mnuGridLines.Checked = ListView1.GridLines

   ListView1.HideColumnHeaders = PropBag.ReadProperty("HideColumnHeaders", False)
   
   ListView1.HideSelection = PropBag.ReadProperty("HideSelection", False)
   ListView1.HotTracking = PropBag.ReadProperty("HotTracking", False)
   ListView1.HoverSelection = PropBag.ReadProperty("HoverSelection", False)
   
   'ListView1.LabelEdit = PropBag.ReadProperty("LabelEdit", 1)
   ListView1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
   
   ListView1.MultiSelect = PropBag.ReadProperty("MultiSelect", True)
   mnuMultiSelect.Checked = ListView1.MultiSelect
   
   'ListView1.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
   'ListView1.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
   ListView1.SortKey = PropBag.ReadProperty("SortKey", 0)
   ListView1.SortOrder = PropBag.ReadProperty("SortOrder", lvwAscending)
   ListView1.Sorted = PropBag.ReadProperty("Sorted", False)
   ListView1.TextBackground = PropBag.ReadProperty("TextBackground", 0)
   ListView1.View = PropBag.ReadProperty("View", 3)
   mnuView(0).Checked = ListView1.View = lvwList
   mnuView(1).Checked = ListView1.View = lvwReport
   
   ListView1.Visible = PropBag.ReadProperty("Visible", True)
   
   m_strColumnHeadersText = PropBag.ReadProperty("ColumnHeadersText", m_def_ColumnHeadersText)
   Dim ColHeaders As ColHeaders
   Set ColHeaders = New ColHeaders
   With ColHeaders
      .CreateItemsFromTagText m_strColumnHeadersText, vbCrLf, "|"
      .SetColumnHeaders ListView1.ColumnHeaders
   End With
   Set ColHeaders = Nothing
     
   
   m_bAutoDeselect = PropBag.ReadProperty("AutoDeselect", m_def_AutoDeselect)
   mnuAutoDeselect.Checked = m_bAutoDeselect
   mnuAutoDeselect.Enabled = Not ListView1.FullRowSelect

   m_strFileSpec = PropBag.ReadProperty("FileSpec", m_def_FileSpec)
   m_bIncludeFolder = PropBag.ReadProperty("IncludeFolder", m_def_IncludeFolder)
   m_strPath = PropBag.ReadProperty("Path", m_def_Path)
   m_bRecursive = PropBag.ReadProperty("Recursive", m_def_Recursive)
   m_bSubitemSelect = PropBag.ReadProperty("SubitemSelect", m_def_SubitemSelect)
   mnuSubiitemSelect.Checked = m_bSubitemSelect

   m_lUpdateFrequency = PropBag.ReadProperty("UpdateFrequency", m_def_UpdateFrequency)
   
   m_bHighlightColumn = PropBag.ReadProperty("HighlightColumn", m_def_HighlightColumn)
   mnuColumnHighlight.Checked = m_bHighlightColumn
   m_lCurrrentHighlightColumn = PropBag.ReadProperty("CurrrentHighlightColumn", m_def_CurrrentHighlightColumn)
   m_enSizingType = PropBag.ReadProperty("SizingType", m_def_SizingType)
   m_enHighlightColor = PropBag.ReadProperty("HighlightColor", m_def_HighlightColor)
   m_enDefaultColor = PropBag.ReadProperty("DefaultColor", m_def_DefaultColor)
   HighlightColumn = m_bHighlightColumn
   'Call DoHighlightColumn(m_enHighlightColor, m_enDefaultColor, m_lCurrrentHighlightColumn, m_enSizingType)

   m_bFlatHeader = PropBag.ReadProperty("FlatHeader", m_def_FlatHeader)
   If m_bFlatHeader Then
      Call SetFlatHeader
   End If
   
   m_bAutoTooltip = PropBag.ReadProperty("AutoTooltip", m_def_AutoTooltip)
   mnuAutoTooltip.Checked = m_bAutoTooltip
   
   m_bAutoPopupMenu = PropBag.ReadProperty("AutoPopupMenu", m_def_AutoPopupMenu)
   
   m_bAllowFileDragDrop = PropBag.ReadProperty("AllowFileDragDrop", m_def_AllowFileDragDrop)
      
   
   '@@ ItemEdit 04/05/17, 16:58:13
   m_bAllowItemEdit = PropBag.ReadProperty("AllowItemEdit", m_def_AllowItemEdit)
   mnuItemEdit.Checked = m_bAllowItemEdit
   
   Dim bRunMode As Boolean
   On Error Resume Next
      bRunMode = UserControl.Ambient.UserMode
      If bRunMode Then
         Set sc = New CSubclass
         With sc
            .AddMsg WM_NOTIFY, MSG_AFTER
            .AddMsg WM_VSCROLL, MSG_AFTER
            .AddMsg WM_HSCROLL, MSG_AFTER
            .AddMsg WM_MOUSEWHEEL, MSG_AFTER
            .Subclass ListView1.hWnd, Me
         End With 'SC

         Dim hHeader As Long
         hHeader = SendMessage(ListView1.hWnd, LVM_GETHEADER, 0, ByVal 0&)
         If hHeader Then
            Set sc2 = New CSubclass
            With sc2
               .AddMsg WM_LBUTTONUP, MSG_AFTER
               .AddMsg WM_NOTIFY, MSG_AFTER
               .Subclass2 hHeader, Me
            End With 'SC2
         End If
      End If

   On Error GoTo 0

End Sub

Private Sub UserControl_Resize()

   On Error GoTo Bye
   If ListView1.FlatScrollBar Then
      ListView1.Move 0, 0, Width - 500, Height
   Else 'LISTVIEW1.FLATSCROLLBAR = FALSE
      ListView1.Move 0, 0, Width, Height
   End If
   
   RaiseEvent Resize
Exit Sub

Bye:

End Sub

Private Sub UserControl_Terminate()

   If ObjPtr(sc) Then
      Set sc = Nothing
   End If

   If ObjPtr(sc2) Then
      Set sc2 = Nothing
   End If

End Sub

'¼Ó¼º°ªÀ» ÀúÀå¼Ò¿¡ ±â·ÏÇÕ´Ï´Ù.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

   Call PropBag.WriteProperty("Appearance", ListView1.Appearance, 1)
   Call PropBag.WriteProperty("Arrange", ListView1.Arrange, 0)
   Call PropBag.WriteProperty("AllowColumnReorder", ListView1.AllowColumnReorder, True)
   Call PropBag.WriteProperty("Checkboxes", ListView1.Checkboxes, False)
   Call PropBag.WriteProperty("Enabled", ListView1.Enabled, True)
   Call PropBag.WriteProperty("FlatScrollBar", ListView1.FlatScrollBar, False)
   Call PropBag.WriteProperty("FullRowSelect", ListView1.FullRowSelect, False)
   Call PropBag.WriteProperty("GridLines", ListView1.GridLines, False)
   Call PropBag.WriteProperty("HideColumnHeaders", ListView1.HideColumnHeaders, False)
   Call PropBag.WriteProperty("HideSelection", ListView1.HideSelection, False)
   Call PropBag.WriteProperty("HotTracking", ListView1.HotTracking, False)
   Call PropBag.WriteProperty("HoverSelection", ListView1.HoverSelection, False)
   Call PropBag.WriteProperty("MousePointer", ListView1.MousePointer, 0)
   Call PropBag.WriteProperty("MultiSelect", ListView1.MultiSelect, True)
   Call PropBag.WriteProperty("SortKey", ListView1.SortKey, 0)
   Call PropBag.WriteProperty("SortOrder", ListView1.SortOrder, lvwAscending)
   Call PropBag.WriteProperty("Sorted", ListView1.Sorted, False)
   Call PropBag.WriteProperty("TextBackground", ListView1.TextBackground, 0)
   Call PropBag.WriteProperty("View", ListView1.View, 3)
   Call PropBag.WriteProperty("Visible", ListView1.Visible, True)
   'Call PropBag.WriteProperty("LabelEdit", ListView1.LabelEdit, 1)
   Call PropBag.WriteProperty("BorderStyle", ListView1.BorderStyle, 1)
   'Call PropBag.WriteProperty("OLEDragMode", ListView1.OLEDragMode, 0)
   'Call PropBag.WriteProperty("OLEDropMode", ListView1.OLEDropMode, 0)
   
   Call PropBag.WriteProperty("ColumnHeadersText", m_strColumnHeadersText, m_def_ColumnHeadersText)
   Call PropBag.WriteProperty("AutoDeselect", m_bAutoDeselect, m_def_AutoDeselect)
   Call PropBag.WriteProperty("FileSpec", m_strFileSpec, m_def_FileSpec)
   Call PropBag.WriteProperty("IncludeFolder", m_bIncludeFolder, m_def_IncludeFolder)
   Call PropBag.WriteProperty("Path", m_strPath, m_def_Path)
   Call PropBag.WriteProperty("Recursive", m_bRecursive, m_def_Recursive)
   Call PropBag.WriteProperty("SubitemSelect", m_bSubitemSelect, m_def_SubitemSelect)
   Call PropBag.WriteProperty("UpdateFrequency", m_lUpdateFrequency, m_def_UpdateFrequency)
   Call PropBag.WriteProperty("HighlightColor", m_enHighlightColor, m_def_HighlightColor)
   Call PropBag.WriteProperty("DefaultColor", m_enDefaultColor, m_def_DefaultColor)
   Call PropBag.WriteProperty("SizingType", m_enSizingType, m_def_SizingType)
   Call PropBag.WriteProperty("CurrrentHighlightColumn", CurrrentHighlightColumn, m_def_CurrrentHighlightColumn)
   Call PropBag.WriteProperty("HighlightColumn", m_bHighlightColumn, m_def_HighlightColumn)
   Call PropBag.WriteProperty("FlatHeader", m_bFlatHeader, m_def_FlatHeader)
   Call PropBag.WriteProperty("AutoTooltip", m_bAutoTooltip, m_def_AutoTooltip)
   Call PropBag.WriteProperty("AutoPopupMenu", m_bAutoPopupMenu, m_def_AutoPopupMenu)
   '@@ ItemEdit 04/05/17, 16:58:13
   Call PropBag.WriteProperty("AllowItemEdit", m_bAllowItemEdit, m_def_AllowItemEdit)
   Call PropBag.WriteProperty("AllowFileDragDrop", m_bAllowFileDragDrop, m_def_AllowFileDragDrop)
   
End Sub

'Returns/sets the current view of the MSComctlLib.ListView control.
Public Property Get View() As ListViewConstants

   View = ListView1.View

End Property

Public Property Let View(ByVal New_View As ListViewConstants)

   ListView1.View() = New_View
   PropertyChanged "View"

End Property

Public Property Get Visible() As Boolean

   Visible = ListView1.Visible

End Property

Public Property Let Visible(ByVal boolVisible As Boolean)

   ListView1.Visible = boolVisible
   PropertyChanged "Visible"

End Property

Public Property Get VisibleCount() As Long

   VisibleCount = SendMessage(ListView1.hWnd, LVM_GETCOUNTPERPAGE, 0&, ByVal 0&)

End Property



Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
   On Error GoTo Bye
   Select Case KeyCode
   Case vbKeyReturn
      Call txtEdit_LostFocus
      KeyCode = 0
      ListView1.SetFocus
   End Select
Bye:
End Sub
Private Sub txtEdit_LostFocus()
   On Error GoTo Bye
   Dim Index As Long, Subitem As Long
   Dim Cancel As Boolean
   Dim strOldString As String
   Dim strNewString As String
   With txtEdit
      Index = Val(Split(.Tag, ":")(0))
      Subitem = Val(Split(.Tag, ":")(1))
      strNewString = .Text
      
      If Subitem = 0 Then
         strOldString = ListView1.ListItems(Index).Text
         
         If strOldString <> strNewString Then
            RaiseEvent BeforeItemEdit(Cancel, Index, Subitem, strOldString, strNewString)
            If Not Cancel Then
               ListView1.ListItems(Index).Text = strNewString
               RaiseEvent AfterItemEdit(Cancel, Index, Subitem, strOldString, strNewString)
               If Cancel Then
                  ListView1.ListItems(Index).Text = strOldString
               End If
            End If
         End If
         
      Else
          strOldString = ListView1.ListItems(Index).SubItems(Subitem)
         
         If strOldString <> strNewString Then
            RaiseEvent BeforeItemEdit(Cancel, Index, Subitem, strOldString, strNewString)
            If Not Cancel Then
               ListView1.ListItems(Index).SubItems(Subitem) = strNewString
               RaiseEvent AfterItemEdit(Cancel, Index, Subitem, strOldString, strNewString)
               If Cancel Then
                  ListView1.ListItems(Index).SubItems(Subitem) = strOldString
               End If
            End If
         End If
         
      End If
      .Visible = False
   End With
Bye:
End Sub

Private Sub MoveEditBox()
   Dim Index As Long, Subitem As Long
   Dim Text As String
   
   On Error GoTo Bye
   
   With txtEdit
      If Len(.Tag) = 0 Then
         Exit Sub
      End If
      Index = Val(Split(.Tag, ":")(0))
      Subitem = Val(Split(.Tag, ":")(1))
   End With
   
   If Subitem = 0 Then
      Text = ListView1.ListItems(Index).Text
   Else
      Text = ListView1.ListItems(Index).SubItems(Subitem)
   End If
   
   Dim lpItemRect As RECT
   Dim lpClientRect As RECT
   Dim hHeader As Long
   Static lHeaderHeight As Long
    
   'Get the height of the header once.
   If lHeaderHeight = 0 Then
      hHeader = SendMessage(ListView1.hWnd, LVM_GETHEADER, 0, ByVal 0&)
      If hHeader Then
         GetClientRect hHeader, lpItemRect
      End If
      lHeaderHeight = lpItemRect.Bottom
   End If

   'Get the client rect
   GetClientRect ListView1.hWnd, lpClientRect
   'Adjust the top and bottom of the client rect for list items.
   'This new rect's top  start from HeaderHeight + 1.
   With lpClientRect
      .Top = lHeaderHeight + 1
      .Bottom = .Bottom - 1
   End With
   
   'Get the item rect
   With lpItemRect
      .Top = Subitem  'The one-based index of the subitem.
      .Left = LVIR_LABEL
   End With
   SendMessage ListView1.hWnd, LVM_GETSUBITEMRECT, Index - 1, lpItemRect
   
   Dim lTextWidth As Long
   lTextWidth = CalcTextWidth(Text)
   With lpItemRect
      If lTextWidth > .Right - .Left Then
         .Right = .Left + lTextWidth
      End If
   End With
   
   If RectIsOverlayed(lpClientRect, lpItemRect) Then
      'Converr Pixels to Twips.
      lpItemRect = GetTwipsRect(lpItemRect)
      With txtEdit
         .Alignment = ListView1.ColumnHeaders(Subitem + 1).Alignment
         Dim dLeft As Long
         
         If ListView1.Appearance = cc3D Then
            If Subitem = 0 Then
               dLeft = 50
            Else
               dLeft = 100
            End If
         Else
            If Subitem = 0 Then
               dLeft = 30
            Else
               dLeft = 70
            End If
         End If
         
         Dim dTop As Long
         If ListView1.BorderStyle = ccFixedSingle Then
            dTop = 50
         Else
            dTop = 30
         End If
         
         'Adjust the movement in consideration with spacings in the List item rectanges.
         If .Alignment = lvwColumnRight Then
            .Move ListView1.Left + lpItemRect.Left + dLeft, ListView1.Top + lpItemRect.Top + dTop, _
                                 lpItemRect.Right - lpItemRect.Left - 100, lpItemRect.Bottom - lpItemRect.Top - 100
         Else
            .Move ListView1.Left + lpItemRect.Left + dLeft, ListView1.Top + lpItemRect.Top + dTop, _
                                 lpItemRect.Right - lpItemRect.Left, lpItemRect.Bottom - lpItemRect.Top - 100
         End If
         .Text = Text
         .Visible = True
      End With
   Else
      txtEdit.Visible = False
   End If
  
   DoEvents
   Exit Sub
Bye:
   
End Sub

Private Function RectIsOverlayed(lpOwner As RECT, lpTarget As RECT) As Boolean
   With lpTarget
      RectIsOverlayed = PtInRect(lpOwner, .Left, .Top) Or _
                                    PtInRect(lpOwner, .Left, .Bottom) Or _
                                    PtInRect(lpOwner, .Right, .Top) Or _
                                    PtInRect(lpOwner, .Right, .Bottom)
   End With
End Function

Private Function GetTwipsRect(lpPixelRect As RECT) As RECT
   Static m_TwipsPerPixelX As Single
   Static m_TwipsPerPixelY As Single
   
   If m_TwipsPerPixelX = 0 Then
      m_TwipsPerPixelX = Screen.TwipsPerPixelX
   End If
   If m_TwipsPerPixelY = 0 Then
      m_TwipsPerPixelY = Screen.TwipsPerPixelY
   End If
   
   With GetTwipsRect
      .Left = lpPixelRect.Left * m_TwipsPerPixelX
      .Top = lpPixelRect.Top * m_TwipsPerPixelY
      .Right = lpPixelRect.Right * m_TwipsPerPixelX
      .Bottom = lpPixelRect.Bottom * m_TwipsPerPixelY
   End With
End Function


Public Function Cell(ByVal RowIndex As Long, ByVal ColumnIndex As Long) As String
   On Error GoTo Bye
   With ListView1.ListItems(RowIndex)
      Select Case ColumnIndex
      Case Is <= 1
         Cell = .Text
      Case Else
         Cell = .SubItems(ColumnIndex - 1)
      End Select
   End With
Bye:
End Function

Public Function Row(ByVal Index As Long, Optional Delimiter As String = vbTab, Optional IncludeIndex As Boolean) As String
   On Error GoTo Bye
   Dim i As Long
   With ListView1.ListItems(Index)
      For i = 1 To ListView1.ColumnHeaders.Count
         Select Case i
         Case 1
            If IncludeIndex Then
               Row = Index & Delimiter & .Text
            Else
               Row = .Text
            End If
         Case Else
            Row = Row & Delimiter & .SubItems(i - 1)
         End Select
      Next
   End With
Bye:
End Function

Public Function Rows( _
                              RowDelimiter As String, ColumnDelimiter As String, IncludeIndex As Boolean, _
                              ParamArray vIndexes() As Variant) As String
   On Error GoTo Bye
   Dim i As Long
   Dim asItems() As String
   ReDim asItems(UBound(vIndexes))
   For i = 0 To UBound(vIndexes)
      asItems(i) = Row(vIndexes(i), ColumnDelimiter, IncludeIndex)
   Next
   Rows = VBA.Join(asItems, RowDelimiter)
Bye:
End Function

Public Function RowsArray( _
                              Indexes() As Long, _
                              Optional RowDelimiter As String = vbCrLf, _
                              Optional ColumnDelimiter As String = vbTab, _
                              Optional IncludeIndex As Boolean = True) As String
   On Error GoTo Bye
   Dim i As Long
   Dim asItems() As String
   ReDim asItems(LBound(Indexes) To UBound(Indexes))
   For i = LBound(Indexes) To UBound(Indexes)
      asItems(i) = Row(Indexes(i), ColumnDelimiter, IncludeIndex)
   Next
   RowsArray = VBA.Join(asItems, RowDelimiter)
Bye:
End Function
Public Function Col(ByVal Index As Long, Optional Delimiter As String = vbCrLf, Optional IncludeIndex As Boolean) As String
   On Error GoTo Bye
   Dim i As Long
   Dim asItems() As String
   
   With ListView1.ListItems
      ReDim asItems(1 To .Count)
      If Index <= 1 Then
         For i = 1 To .Count
            asItems(i) = .Item(i).Text
            If IncludeIndex Then
               asItems(i) = i & vbTab & asItems(i)
            End If
         Next i
      Else
         For i = 1 To .Count
            asItems(i) = .Item(i).SubItems(Index - 1)
            If IncludeIndex Then
               asItems(i) = i & vbTab & asItems(i)
            End If
         Next i
      End If
   End With
   Col = VBA.Join(asItems, Delimiter)
Bye:
End Function

Public Function Cols( _
                              RowDelimiter As String, ColumnDelimiter As String, IncludeIndex As Boolean, _
                              ParamArray vIndexes() As Variant) As String
   On Error GoTo Bye
   Dim i As Long, j As Long
   Dim asItems() As String
   Dim asLine() As String
   
   ReDim asLine(UBound(vIndexes))
   With ListView1.ListItems
      
      ReDim asItems(1 To .Count)
      
      For i = 1 To .Count
      
         With .Item(i)
            For j = 0 To UBound(vIndexes)
               If vIndexes(j) <= 1 Then
                  asLine(j) = .Text
               Else
                  asLine(j) = .SubItems(vIndexes(j) - 1)
               End If
            Next
         End With
         asItems(i) = VBA.Join(asLine, ColumnDelimiter)
         If IncludeIndex Then
            asItems(i) = i & ColumnDelimiter & asItems(i)
         End If
         
      Next
   
   End With
   
   Cols = VBA.Join(asItems, RowDelimiter)
Bye:
End Function


'Public Property Get LabelEdit() As ListLabelEditConstants
'   LabelEdit = ListView1.LabelEdit
'End Property
'Public Property Let LabelEdit(ByVal New_LabelEdit As ListLabelEditConstants)
'   ListView1.LabelEdit() = New_LabelEdit
'   PropertyChanged "LabelEdit"
'End Property

Public Property Get BorderStyle() As MSComctlLib.BorderStyleConstants
   BorderStyle = ListView1.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As MSComctlLib.BorderStyleConstants)
   ListView1.BorderStyle() = New_BorderStyle
   'UserControl.BorderStyle = New_BorderStyle
   mnuSolidBorder.Checked = New_BorderStyle
   PropertyChanged "BorderStyle"
End Property

'Public Property Get OLEDragMode() As OLEDragConstants
'   OLEDragMode = ListView1.OLEDragMode
'End Property
'Public Property Let OLEDragMode(ByVal New_OLEDragMode As OLEDragConstants)
'   ListView1.OLEDragMode() = New_OLEDragMode
'   PropertyChanged "OLEDragMode"
'End Property

'Public Property Get OLEDropMode() As OLEDropConstants
'   OLEDropMode = ListView1.OLEDropMode
'End Property
'Public Property Let OLEDropMode(ByVal New_OLEDropMode As OLEDropConstants)
'   ListView1.OLEDropMode() = New_OLEDropMode
'   PropertyChanged "OLEDropMode"
'End Property


'Public Property Get ColumnHeadersText() As String
'   ColumnHeadersText = m_strColumnHeadersText
'End Property

Public Property Let ColumnHeadersText(ByVal New_ColumnHeadersText As String)
   m_strColumnHeadersText = New_ColumnHeadersText
   PropertyChanged "ColumnHeadersText"
End Property

Public Property Get AllowFileDragDrop() As Boolean
   AllowFileDragDrop = m_bAllowFileDragDrop
End Property

Public Property Let AllowFileDragDrop(ByVal New_AllowFileDragDrop As Boolean)
   m_bAllowFileDragDrop = New_AllowFileDragDrop
   PropertyChanged "AllowFileDragDrop"
End Property


Private Function FileNamePart( _
                         ByVal FullName As String, _
                         ByVal ePortions As eFileNameParts) As String

'This function is used to parse keys peices of info from a
'filename that is passed into it.

'FullName = "D:\AdvVB\project\vbAdvanced_comp.vbp"
'efpBaseName = vbAdvanced_comp
'efpExtension = vbp
'efpPath=D:\AdvVB\project\
'efpPathUnqualified=D:\AdvVB\project
'efpName = vbAdvanced_comp.vbp
'efpPathPlusBaseName=D:\AdvVB\project\vbAdvanced_comp.vbp
'efpPathBaseName = project
'efpDrive = D
'efpDriveQualified = D:

   Dim lFirstPeriod As Long, lFirstBackSlash As Long
   Dim strQualifiedPath As String, strBaseName As String, sExt As String
   Dim sRet As String, sDrive As String

   Select Case ePortions
   Case efpDriveQualified, efpDrive
      lFirstPeriod = VBA.InStrRev(FullName, ":")
      If lFirstPeriod Then
         If ePortions = efpDriveQualified Then
            sDrive = VBA.Left$(FullName, lFirstPeriod)
         Else
            If lFirstPeriod > 1 Then
               sDrive = VBA.Left$(FullName, lFirstPeriod - 1)
            End If
         End If
         lFirstBackSlash = VBA.InStrRev(sDrive, "/")
         If lFirstBackSlash = 0 Then
            lFirstBackSlash = VBA.InStrRev(sDrive, "\")
         End If
         If lFirstBackSlash Then
            sDrive = VBA.Mid$(sDrive, lFirstBackSlash + 1)
         End If
         FileNamePart = sDrive
      End If
      Exit Function
   Case efpConvToLocalName
      lFirstPeriod = VBA.InStrRev(FullName, ":")
      If lFirstPeriod Then
         sDrive = VBA.Left$(FullName, lFirstPeriod)
         lFirstBackSlash = VBA.InStrRev(sDrive, "/")
         If lFirstBackSlash = 0 Then
            lFirstBackSlash = VBA.InStrRev(sDrive, "\")
         End If
         If lFirstBackSlash Then
            sDrive = VBA.Mid$(sDrive, lFirstBackSlash + 1)
         End If
         FullName = sDrive & VBA.Mid$(FullName, lFirstPeriod + 1)
      End If
      FileNamePart = VBA.Replace(FullName, "/", "\")
      lFirstPeriod = VBA.InStr(1, FileNamePart, "\\")
      If lFirstPeriod Then
         FileNamePart = VBA.Mid$(FileNamePart, lFirstPeriod + 2)
      End If
      Exit Function
   
   Case efpConvToShortName
      'sRet = VBA.Space$(1024)
      'GetShortPathName FullName, sRet, Len(sRet)
      'FileNamePart = TrimNull(sRet)
      FileNamePart = FullName
      Exit Function
   
   Case efpConvToLongName
      'FileNamePart = ProperCaseDirectory(FileNamePart(FullName, efpPath)) _
      '               & "\" & VBA.Dir$(FullName)
      FileNamePart = FullName
      Exit Function
   End Select
   
   '**** File name parts

   lFirstPeriod = VBA.InStrRev(FullName, ".")
   lFirstBackSlash = VBA.InStrRev(FullName, "\")
   If lFirstBackSlash = 0 Then
      lFirstBackSlash = VBA.InStrRev(FullName, "/")
   End If
   
   If lFirstBackSlash > 0 Then
      strQualifiedPath = VBA.Left$(FullName, lFirstBackSlash)
   End If
   
   If lFirstPeriod > 0 And lFirstPeriod > lFirstBackSlash Then
      sExt = VBA.Mid$(FullName, lFirstPeriod + 1)
      strBaseName = VBA.Mid$(FullName, lFirstBackSlash + 1, lFirstPeriod - lFirstBackSlash - 1)
   Else 'NOT LFIRSTPERIOD...
      strBaseName = VBA.Mid$(FullName, lFirstBackSlash + 1)
   End If
   
   Select Case ePortions
   Case efpBaseName
      FileNamePart = strBaseName
   Case efpExtension
      FileNamePart = sExt
   Case efpPath
      FileNamePart = strQualifiedPath
   Case efpPathUnqualified
      If Len(strQualifiedPath) Then
         FileNamePart = VBA.Left$(strQualifiedPath, Len(strQualifiedPath) - 1)
      End If
   Case efpName
      If Len(sExt) Then
         FileNamePart = strBaseName & "." & sExt
      Else
         FileNamePart = strBaseName
      End If
   Case efpPathPlusBaseName
      'If Len(sExt) Then
         FileNamePart = strQualifiedPath & strBaseName
      'Else
      '   FileNamePart = strQualifiedPath & strBaseName
      'End If
   Case efpPathBaseName
      If Len(strQualifiedPath) Then
         FullName = VBA.Left$(strQualifiedPath, Len(strQualifiedPath) - 1)
         FileNamePart = FileNamePart(FullName, efpBaseName)
      End If
   End Select

End Function



Private Function AfterRev( _
                         ByVal Source As String, ByVal Target As String, _
                         Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbTextCompare, _
                         Optional Return_pos As Long, _
                         Optional bReturnWholeTextIfNotFound As Boolean) As String

   Return_pos = VBA.InStrRev(Source, Target, -1, Compare)
   If Return_pos = 0 Then
      If bReturnWholeTextIfNotFound Then
         AfterRev = Source
      Else 'BRETURNWHOLETEXTIFNOTFOUND = FALSE
         AfterRev = ""
      End If
   Else 'NOT RETURN_POS...
      AfterRev = VBA.Mid$(Source, Return_pos + Len(Target))
   End If

End Function

