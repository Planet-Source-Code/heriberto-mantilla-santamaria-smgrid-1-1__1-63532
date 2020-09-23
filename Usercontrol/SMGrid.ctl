VERSION 5.00
Begin VB.UserControl SMGrid 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   4125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "SMGrid.ctx":0000
End
Attribute VB_Name = "SMGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2005       *'
'******************************************************'
'*                     Version 1.0e                   *'
'******************************************************'
'* Control:       SMGrid Control                      *'
'******************************************************'
'* Author:        Heriberto Mantilla Santamaría       *'
'******************************************************'
'* Description:   Emulate a FlexGrid Control          *'
'******************************************************'
'* Started on:    Wednesday, 10-nov-2004.             *'
'* Release date:  Friday, 18-mar-2005.                *'
'******************************************************'
'* Comments: This control was necessary to develop it *'
'*           for a program of a thesis of grade of my *'
'*           University, it's evolution was stopped by*'
'*           a lot of time, although it isn't comple- *'
'*           tely ended, but it's a beginning.        *'
'******************************************************'
'* Credits/Thanks: Richard Mewett                     *'
'*                 (GetColFromX Function)             *'
'*                 (Flat Border)                      *'
'*                 [CodeId = 61438].                  *'
'*                                                    *'
'*                 Paul Caton (self-subclassing)      *'
'*                 I don't need say nothing.          *'
'*                 [CodeId = 54117].                  *'
'*                                                    *'
'*                 Carles P.V.                        *'
'*                 (API's AlphaBlend emulation)       *'
'*                 [CodeId = 59786].                  *'
'*                                                    *'
'*                 Steve McMahon (vbalScrollButton)   *'
'*                                                    *'
'*                 MArio Florez (Create GrayIcon)     *'
'*                                                    *'
'*                 Kristian S. Stangeland             *'
'*                 (Save Array's)                     *'
'*                                                    *'
'*                 fred.cpp and Habin                 *'
'*                 (for suggestions and debugging)    *'
'*----------------------------------------------------*'
'* Added                                              *'
'*++++++++++++++++++++++++++++++++++++++++++++++++++++*'
'* BorderStyles, Styles Buttons, Enabled, Values.     *'
'*                                                    *'
'* Changed individuals Item properties.               *'
'*                                                    *'
'* Remove the vb scrollbar's.                         *'
'*++++++++++++++++++++++++++++++++++++++++++++++++++++*'
'*                                                    *'
'* Note:     Comments, suggestions, doubts or bug     *'
'*           reports are wellcome to these e-mail     *'
'*           addresses:                               *'
'*                                                    *'
'*                  heri_05-hms@mixmail.com or        *'
'*                  hcammus@hotmail.com               *'
'******************************************************'
'* Now my website is available but alone the version  *'
'* in Spanish.                                        *'
'-----------------------------------------------------*'
'* WebSite:  http://www.geocities.com/hackprotm/      *'
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2005       *'
'******************************************************'
Option Explicit
  
'*******************************************************'
'* Subclasser Declarations Paul Caton                  *'
  
 '-uSelfSub declarations---------------------------------------------------------------------------
 Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
 End Enum

 Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
 End Enum

 Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                      As Long
  dwFlags                     As TRACKMOUSEEVENT_FLAGS
  hwndTrack                   As Long
  dwHoverTime                 As Long
 End Type

 Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
 Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
 Private Const CODE_LEN      As Long = 240                                   'Thunk length in bytes
 Private Const WNDPROC_OFF   As Long = &H30                                  'WndProc execution offset
 Private Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))    'Bytes to allocate per thunk, data + code + msg tables
 Private Const PAGE_RWX      As Long = &H40                                  'Allocate executable memory
 Private Const MEM_COMMIT    As Long = &H1000                                'Commit allocated memory
 Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
 Private Const IDX_SHUTDOWN  As Long = 1                                     'Shutdown flag data index
 Private Const IDX_HWND      As Long = 2                                     'hWnd data index
 Private Const IDX_EBMODE    As Long = 3                                     'EbMode data index
 Private Const IDX_CWP       As Long = 4                                     'CallWindowProc data index
 Private Const IDX_SWL       As Long = 5                                     'SetWindowsLong data index
 Private Const IDX_FREE      As Long = 6                                     'VirtualFree data index
 Private Const IDX_ME        As Long = 7                                     'Owner data index
 Private Const IDX_WNDPROC   As Long = 8                                     'Original WndProc data index
 Private Const IDX_CALLBACK  As Long = 9                                     'zWndProc data index
 Private Const IDX_BTABLE    As Long = 10                                    'Before table data index
 Private Const IDX_ATABLE    As Long = 11                                    'After table data index
 Private Const IDX_EBX       As Long = 14                                    'Data code index
 
 Private z_Code(29)          As Currency                                     'Thunk machine-code initialised here
 Private z_Data(552)         As Long                                         'Array whose data pointer is re-mapped to arbitary memory addresses
 Private z_DataDataPtr       As Long                                         'Address of z_Data()'s SafeArray data pointer
 Private z_DataOrigData      As Long                                         'Address of z_Data()'s original data
 Private z_hWnds             As Collection                                   'hWnd/thunk-address collection
 
 Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
 Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
 Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
 Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
 Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
 Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
 Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
 Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
 
 Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
 Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
 Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
 Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

 Public Event MouseEnter()
 Public Event MouseLeave()
 
 Private bTrack                As Boolean
 Private bTrackUser32          As Boolean
 Private bInCtrl               As Boolean
'*******************************************************'
 
 Private Type Tamano
  Height  As Long
  Width   As Long
  Visible As Boolean
 End Type
 
 Private Type isStyleI
  Colors()     As OLE_COLOR  '* Item Colors.
  Enabled()    As Boolean    '* Item Enabled.
  Item()       As String     '* Item Text.
  tPicture()   As StdPicture '* Image.
  Style()      As String     '* Item Style.
  TextColors() As OLE_COLOR  '* Itext ForeColor.
  TextList()   As String     '* List of ComboBox.
  Values()     As Boolean    '* True/False.
 End Type
 
 Private Type isStyle
  Col          As Long
  Item         As isStyleI
  Row          As Long
  tFont        As StdFont    '* Item Font.
 End Type
 
 Private Type POINTAPI
  X           As Long
  Y           As Long
 End Type
 
 Private Type RECT
  hLeft        As Long
  hTop         As Long
  hRight       As Long
  hBottom      As Long
 End Type
   
 Private Type xTamano
  lStartX      As Long
  lStartY      As Long
 End Type
 
 Public Enum BorderStyleEnum
  BorderNone = &H0
  BorderBump = &H1
  BorderEtched = &H2
  BorderSunken = &H3
  BorderRaised = &H4
  BorderFlat = &H5
  BorderFlatFlat = &H6
  BorderPicture = &H7
 End Enum
    
 Private Const BF_BOTTOM = &H8
 Private Const BF_FLAT = &H4000
 Private Const BF_LEFT = &H1
 Private Const BF_RIGHT = &H4
 Private Const BF_SOFT = &H1000
 Private Const BF_TOP = &H2
 Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
 Private Const COLOR_BTNFACE = 15
 Private Const COLOR_WINDOW = 5
 Private Const DEFAULTHEIGHT = 380
 Private Const DEFAULTWIDTH = 350
 Private Const DFC_BUTTON = 4
 Private Const DFCS_BUTTONCHECK = &H0
 Private Const DFCS_BUTTONRADIO = &H4
 Private Const DFCS_BUTTON3STATE = &H10
 Private Const DFCS_CHECKED         As Long = &H400
 Private Const DFCS_INACTIVE        As Long = &H100
 Private Const BDR_RAISEDINNER = &H4
 Private Const BDR_RAISEDOUTER = &H1
 Private Const BDR_SUNKENINNER = &H8
 Private Const BDR_SUNKENOUTER = &H2
 Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
 Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
 Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
 Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
 Private Const DT_BOTTOM            As Long = &H8
 Private Const DT_CALCRECT          As Long = &H400
 Private Const DT_CENTER            As Long = &H1
 Private Const DT_EDITCONTROL       As Long = &H2000
 Private Const DT_RIGHT             As Long = &H2
 Private Const DT_TOP               As Long = &H0
 Private Const DT_VCENTER           As Long = &H4
 Private Const DT_WORD_ELLIPSIS     As Long = &H40000
 Private Const DT_WORDBREAK         As Long = &H10
 Private Const WM_LBUTTONDOWN       As Long = &H201
 Private Const WM_MOUSEMOVE         As Long = &H200
 Private Const WM_MOUSELEAVE        As Long = &H2A3
 Private Const WM_MOUSEWHEEL        As Long = &H20A
 Private Const WM_RBUTTONDOWN       As Long = &H204
 Private Const WM_THEMECHANGED      As Long = &H31A
 
 '* Richard Mewett.
 Private Const WM_KILLFOCUS         As Long = &H8
 Private Const WM_GETMINMAXINFO     As Long = &H24
 Private Const WM_WINDOWPOSCHANGED  As Long = &H47
 Private Const WM_WINDOWPOSCHANGING As Long = &H46
 Private Const WM_SIZE              As Long = &H5
 Private Const WM_CTLCOLORSCROLLBAR As Long = &H137
 
 '************************************************************
 '* Scrollbar's Steve McMahon (steve@dogma.demon.co.uk)
 Private lStyle              As Long
 Private m_bNoFlatScrollBars As Boolean
 Private m_hWnd              As Long
 Private m_hWNdH             As Long
 Private m_hWNdV             As Long
 Private m_lSmallChange      As Long
 Private Value1              As Long
 Private Value2              As Long
  
 Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
 Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
 Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" (ByVal hWnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
 Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" (ByVal hWnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
 Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" (ByVal hWnd As Long, ByVal index As Long, ByVal newValue As Long, ByVal fRedraw As Boolean) As Long
 Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" (ByVal hWnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long
 Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
 Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long
 Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
 Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
 Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
 Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hWnd As Long) As Long
   
 Private Const FSB_ENCARTA_MODE      As Long = 1
 Private Const FSB_FLAT_MODE         As Long = 2
 Private Const FSB_REGULAR_MODE      As Long = 0
 Private Const SB_BOTTOM             As Long = 7
 Private Const SB_CTL                As Long = 2
 Private Const SB_ENDSCROLL          As Long = 8
 Private Const SB_HORZ               As Long = 0
 Private Const SB_LEFT               As Long = 6
 Private Const SB_LINEDOWN           As Long = 1
 Private Const SB_LINELEFT           As Long = 0
 Private Const SB_LINERIGHT          As Long = 1
 Private Const SB_LINEUP             As Long = 0
 Private Const SB_PAGEDOWN           As Long = 3
 Private Const SB_PAGELEFT           As Long = 2
 Private Const SB_PAGERIGHT          As Long = 3
 Private Const SB_PAGEUP             As Long = 2
 Private Const SB_RIGHT              As Long = 7
 Private Const SB_THUMBTRACK         As Long = 5
 Private Const SB_TOP                As Long = 6
 Private Const SB_VERT               As Long = 1
 Private Const SBS_HORZ              As Long = &H0&
 Private Const SBS_VERT              As Long = &H1&
 Private Const SIF_RANGE             As Long = &H1
 Private Const SIF_PAGE              As Long = &H2
 Private Const SIF_POS               As Long = &H4
 Private Const SIF_TRACKPOS          As Long = &H10
 Private Const SIF_ALL As Long = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)
 Private Const WM_HSCROLL            As Long = &H114
 Private Const WM_VSCROLL            As Long = &H115
 Private Const WS_CHILD              As Long = &H40000000
 Private Const WSB_PROP_HSTYLE       As Long = &H200&
 Private Const WS_VISIBLE            As Long = &H10000000
  
 ' Scroll bar stuff
 Private Type SCROLLINFO
  cbSize    As Long
  fMask     As Long
  nMin      As Long
  nMax      As Long
  nPage     As Long
  nPos      As Long
  nTrackPos As Long
 End Type
 
 Public Event Change()
 Public Event Scroll()
 Public Event DblClick()
 Public Event ThemeChanged()
 '************************************************************
 
 Private WithEvents txtEdit  As TextBox
Attribute txtEdit.VB_VarHelpID = -1
 Private WithEvents cmbEdit  As ComboBox
Attribute cmbEdit.VB_VarHelpID = -1
 Private WithEvents picEdit  As PictureBox
Attribute picEdit.VB_VarHelpID = -1
 
 Private Const MOUSEEVENTF_LEFTDOWN = &H2 '* For generating a mousedown event to replace double click.
 Private Const Version      As String = "SMGrid 1.1 By HACKPRO TM"
 
 Private b_AutoSizeColumn   As Boolean
 Private bDrawTheme         As Boolean
 Private ClickHeader        As Boolean
 Private ColItem            As Long
 Private Cols               As Long
 Private ColsV              As Long
 Private FirstTime          As Boolean
 Private Headers()          As Tamano
 Private hTheme             As Long             '* hTheme Handle.
 Private isEnabled          As Boolean
 Private isLine             As Boolean
 Private isXp               As Boolean
 Private Items()            As isStyle
 Private LastButton         As Integer
 Private LastCol            As Long
 Private LastRow            As Long
 Private lHScroll           As Boolean
 Private lVScroll           As Boolean
 Private mColumnHeadingH    As Single
 Private mFont              As Font
 Private mHeaderHot         As Boolean
 Private MouseX             As Single
 Private MouseY             As Single
 Private m_bAlphaBlendSel   As Boolean
 Private m_bViewStyle       As Boolean
 Private m_eBorderStyle     As BorderStyleEnum
 Private m_lBackColor       As OLE_COLOR
 Private m_lBackgroundPic   As StdPicture
 Private m_lBorderColor     As OLE_COLOR
 Private m_lFullSelect      As Boolean
 Private m_lHeadersColor    As OLE_COLOR
 Private m_lSelectBackColor As OLE_COLOR
 Private m_sTextHeaders     As String
 Private m_Text             As Variant
 Private m_txtRect          As RECT
 Private m_Buttons          As RECT
 Private NoChanged          As Boolean
 Private RowItem            As Long
 Private ShowItems          As Integer
 Private tmpC1              As Integer
 Private tmpC2              As Integer
 Private TheForm            As Object
 Private TheText            As Object
 Private Tamano()           As xTamano
 Private TotalItems         As Long
 Private xRow()             As Long
 Private yRow()             As Long
  
 Private Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
 Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
 Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
 Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
 Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
 Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lhDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
 Private Declare Function DrawThemeEdge Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pDestRect As RECT, ByVal uEdge As Long, ByVal uFlags As Long, pContentRect As RECT) As Long
 Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
 Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
  
 '* XP detection.
 Private Declare Function GetVersion Lib "kernel32" () As Long
 
 Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
 Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
 Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
 Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
 Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
 Private Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
 Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
 Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
 Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
    
 Public Event Click(ByVal Row As Long, ByVal Col As Long, ByVal Style As String)
Attribute Click.VB_MemberFlags = "200"
 
 '* For Create GrayIcon --> MArio Florez.
 Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
 Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
 Private Declare Function CreateIconIndirect Lib "user32.dll" (ByRef piconinfo As ICONINFO) As Long
 Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
 Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
 Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
 Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
 Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
 Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
 Private Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
 Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
 Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
 
 ' Type - GetObjectAPI.lpObject
 Private Type BITMAP
  bmType       As Long    'LONG   // Specifies the bitmap type. This member must be zero.
  bmWidth      As Long    'LONG   // Specifies the width, in pixels, of the bitmap. The width must be greater than zero.
  bmHeight     As Long    'LONG   // Specifies the height, in pixels, of the bitmap. The height must be greater than zero.
  bmWidthBytes As Long    'LONG   // Specifies the number of bytes in each scan line. This value must be divisible by 2, because Windows assumes that the bit values of a bitmap form an array that is word aligned.
  bmPlanes     As Integer 'WORD   // Specifies the count of color planes.
  bmBitsPixel  As Integer 'WORD   // Specifies the number of bits required to indicate the color of a pixel.
  bmBits       As Long    'LPVOID // Points to the location of the bit values for the bitmap. The bmBits member must be a long pointer to an array of character (1-byte) values.
 End Type

 ' Type - CreateIconIndirect / GetIconInfo
 Private Type ICONINFO
  fIcon    As Long 'BOOL    // Specifies whether this structure defines an icon or a cursor. A value of TRUE specifies an icon; FALSE specifies a cursor.
  xHotspot As Long 'DWORD   // Specifies the x-coordinate of a cursor’s hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
  yHotspot As Long 'DWORD   // Specifies the y-coordinate of the cursor’s hot spot. If this structure defines an icon, the hot spot is always in the center of the icon, and this member is ignored.
  hbmMask  As Long 'HBITMAP // Specifies the icon bitmask bitmap. If this structure defines a black and white icon, this bitmask is formatted so that the upper half is the icon AND bitmask and the lower half is the icon XOR bitmask. Under this condition, the height should be an even multiple of two. If this structure defines a color icon, this mask only defines the AND bitmask of the icon.
  hbmColor As Long 'HBITMAP // Identifies the icon color bitmap. This member can be optional if this structure defines a black and white icon. The AND bitmask of hbmMask is applied with the SRCAND flag to the destination; subsequently, the color bitmap is applied (using XOR) to the destination by using the SRCINVERT flag.
 End Type
 
 '* Declares for Unicode support --> Richard Wells.
 Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
 Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
 Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
 
 Private Const VER_PLATFORM_WIN32_NT = 2
 
 Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion      As Long
  dwMinorVersion      As Long
  dwBuildNumber       As Long
  dwPlatformId        As Long
  szCSDVersion        As String * 128 '* Maintenance string for PSS usage.
 End Type
 
 Private mWindowsNT   As Boolean
 
 '* By: Kristian S. Stangeland
 Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
 
 '================================================
 ' Name:          (API's AlphaBlend emulation)
 ' Class:         cTile.cls
 ' Author:        Carles P.V.
 ' Dependencies:
 ' Last revision: 2003.03.28
 '================================================
 
 '-- API:
 Private Declare Function CreateDIBPatternBrushPt Lib "gdi32" (lpPackedDIB As Any, ByVal iUsage As Long) As Long
 Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
 Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
 Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
 Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
 Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, ByVal wUsage As Long) As Long
 
 Private Const DIB_RGB_COLORS As Long = 0
 Private Const OBJ_BITMAP     As Long = 7
  
 '-- Public Enums.:
 Public Enum HatchBrushStyleCts
  [brHorizontal] = &H0
  [brVertival] = &H1
  [brDownwardDiagonal] = &H2
  [brUpwardDiagonal] = &H3
  [brCross] = &H4
  [brDiagonalCross] = &H5
 End Enum
 
 Private Type BITMAPINFOHEADER
  biSize          As Long
  biWidth         As Long
  biHeight        As Long
  biPlanes        As Integer
  biBitCount      As Integer
  biCompression   As Long
  biSizeImage     As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed       As Long
  biClrImportant  As Long
 End Type
 
 '-- Private Variables:
 Private m_hBrush As Long ' Pattern brush

Public Property Get ActualRow() As Long
Attribute ActualRow.VB_MemberFlags = "400"
 ActualRow = TotalItems
End Property

'---------------------------------------------------------------------------------------
' Procedure : AutoSizeColumn
' DateTime  : 29/10/05 12:57
' Author    : HACKPRO TM
' Purpose   : Resize Columns.
'---------------------------------------------------------------------------------------
Public Property Get AutoSizeColumn() As Boolean
 AutoSizeColumn = b_AutoSizeColumn
End Property

'---------------------------------------------------------------------------------------
' Procedure : AutoSizeColumn
' DateTime  : 29/10/05 09:02
' Author    : HACKPRO TM
' Purpose   : Resize Columns.
'---------------------------------------------------------------------------------------
Public Property Let AutoSizeColumn(ByVal bAutoSizeColumn As Boolean)
 b_AutoSizeColumn = bAutoSizeColumn
 Call UserControl.PropertyChanged("AutoSizeColumn")
 Call Refresh
End Property

'---------------------------------------------------------------------------------------
' Procedure : BackColor
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Set the BackColor.
'---------------------------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR
 BackColor = m_lBackColor
End Property

'---------------------------------------------------------------------------------------
' Procedure : BackColor
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Set the BackColor.
'---------------------------------------------------------------------------------------
Public Property Let BackColor(ByVal lBackColor As OLE_COLOR)
 m_lBackColor = ConvertSystemColor(lBackColor)
 Call PropertyChanged("BackColor")
 If (Ambient.UserMode = True) Then Call Refresh
End Property

Public Property Get BackgroundPicture() As StdPicture
 Set BackgroundPicture = m_lBackgroundPic
End Property

Public Property Set BackgroundPicture(ByVal lPicture As StdPicture)
 Set m_lBackgroundPic = lPicture
 Call PropertyChanged("BackgroundPicture")
 Call Refresh
End Property

'---------------------------------------------------------------------------------------
' Procedure : BorderColor
' DateTime  : 27/11/05 09:26
' Author    : HACKPRO TM
' Purpose   : Set the BorderColor.
'---------------------------------------------------------------------------------------
Public Property Get BorderColor() As OLE_COLOR
 BorderColor = m_lBorderColor
End Property

'---------------------------------------------------------------------------------------
' Procedure : BorderColor
' DateTime  : 27/11/05 09:26
' Author    : HACKPRO TM
' Purpose   : Set the BorderColor.
'---------------------------------------------------------------------------------------
Public Property Let BorderColor(ByVal lBorderColor As OLE_COLOR)
 m_lBorderColor = ConvertSystemColor(lBorderColor)
 Call PropertyChanged("BorderColor")
End Property

'---------------------------------------------------------------------------------------
' Procedure : BorderStyle
' DateTime  : 03/07/05 16:43
' Author    : HACKPRO TM
' Purpose   : Set the Border Style.
'---------------------------------------------------------------------------------------
Public Property Get BorderStyle() As BorderStyleEnum
 BorderStyle = m_eBorderStyle
End Property

'---------------------------------------------------------------------------------------
' Procedure : BorderStyle
' DateTime  : 03/07/05 16:43
' Author    : HACKPRO TM
' Purpose   : Set the Border Style.
'---------------------------------------------------------------------------------------
Public Property Let BorderStyle(ByVal eBorderStyle As BorderStyleEnum)
 m_eBorderStyle = eBorderStyle
 Call PropertyChanged("BorderStyle")
 If (Ambient.UserMode = False) Then Call Refresh
End Property

'---------------------------------------------------------------------------------------
' Procedure : ColumnHeadingHeight
' DateTime  : 28/10/05 16:20
' Author    : HACKPRO TM
' Purpose   : Set the height for the columns.
'---------------------------------------------------------------------------------------
Public Property Get ColumnHeadingHeight() As Single
 ColumnHeadingHeight = mColumnHeadingH
End Property

Public Property Let ColumnHeadingHeight(ByVal newValue As Single)
 mColumnHeadingH = newValue
 Call PropertyChanged("ColumnHeadingHeight")
 Call ReDraw
End Property

'---------------------------------------------------------------------------------------
' Procedure : ColPos
' DateTime  : 09/07/05 14:29
' Author    : HACKPRO TM
' Purpose   : Columna.
'---------------------------------------------------------------------------------------
Public Property Get ColPos()
Attribute ColPos.VB_MemberFlags = "400"
 ColPos = ColItem
End Property

'---------------------------------------------------------------------------------------
' Procedure : Enabled
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Enabled/Disabled control.
'---------------------------------------------------------------------------------------
Public Property Get Enabled() As Boolean
 Enabled = isEnabled
End Property

'---------------------------------------------------------------------------------------
' Procedure : Enabled
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Enabled/Disabled control.
'---------------------------------------------------------------------------------------
Public Property Let Enabled(ByVal isValue As Boolean)
 isEnabled = isValue
 UserControl.Enabled = isEnabled
 Call PropertyChanged("Enabled")
End Property

Public Property Get Font() As Font
 Set Font = mFont
End Property

Public Property Set Font(ByVal newValue As StdFont)
 Set mFont = newValue
 Set UserControl.Font = mFont
 Call PropertyChanged("Font")
 Call Refresh
End Property

Public Property Get FlatScrollbars() As Boolean
 FlatScrollbars = m_bNoFlatScrollBars
End Property

Public Property Let FlatScrollbars(ByVal isValue As Boolean)
 m_bNoFlatScrollBars = isValue
 Call PropertyChanged("FlatScrollbars")
 Call Refresh
End Property

Public Property Get FullSelection() As Boolean
 FullSelection = m_lFullSelect
End Property

Public Property Let FullSelection(ByVal isValue As Boolean)
 m_lFullSelect = isValue
 Call PropertyChanged("FullSelection")
End Property

'---------------------------------------------------------------------------------------
' Procedure : HeadersColor
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : BackColor of the Headers.
'---------------------------------------------------------------------------------------
Public Property Get HeadersColor() As OLE_COLOR
 HeadersColor = m_lHeadersColor
End Property

'---------------------------------------------------------------------------------------
' Procedure : HeadersColor
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : BackColor of the Headers.
'---------------------------------------------------------------------------------------
Public Property Let HeadersColor(ByVal lHeadersColor As OLE_COLOR)
 m_lHeadersColor = ConvertSystemColor(lHeadersColor)
 Call PropertyChanged("HeadersColor")
 If (Ambient.UserMode = True) Then Call Refresh
End Property
 
Public Property Get HeaderHotTrack() As Boolean
 HeaderHotTrack = mHeaderHot
End Property

Public Property Let HeaderHotTrack(ByVal bState As Boolean)
 mHeaderHot = bState
 Call PropertyChanged("HeaderHotTrack")
End Property
 
'---------------------------------------------------------------------------------------
' Procedure : ListCount
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Total items of grid.
'---------------------------------------------------------------------------------------
Public Property Get ListCount() As Long
Attribute ListCount.VB_MemberFlags = "400"
 ListCount = TotalItems
End Property

Public Property Get ListViewStyle() As Boolean
 ListViewStyle = m_bViewStyle
End Property

Public Property Let ListViewStyle(ByVal l_ListView As Boolean)
 m_bViewStyle = l_ListView
 Call PropertyChanged("ListViewStyle")
End Property

'---------------------------------------------------------------------------------------
' Procedure : RowColPicture
' DateTime  : 27/11/05 22:53
' Author    : HACKPRO TM
' Purpose   : Set a Picture in Row with Col.
'---------------------------------------------------------------------------------------
Public Sub RowColPicture(ByVal Col As Long, ByVal Row As Long, ByVal tPicture As StdPicture)
On Error Resume Next
 Set Items(Col).Item.tPicture(Row) = tPicture
On Error GoTo 0
End Sub

Public Sub RowHeight(ByVal Row As Long, lHeight As Long)
On Error Resume Next
 yRow(Row + 1) = ScaleY(lHeight, UserControl.Parent.ScaleMode, vbPixels)
On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RowPos
' DateTime  : 09/07/05 14:29
' Author    : HACKPRO TM
' Purpose   : Fila.
'---------------------------------------------------------------------------------------
Public Property Get RowPos()
Attribute RowPos.VB_MemberFlags = "400"
 RowPos = RowItem
End Property

'---------------------------------------------------------------------------------------
' Procedure : SelectBackColor
' DateTime  : 27/11/05 09:26
' Author    : HACKPRO TM
' Purpose   : Set the Selection Cell BackColor.
'---------------------------------------------------------------------------------------
Public Property Get SelectBackColor() As OLE_COLOR
 SelectBackColor = m_lSelectBackColor
End Property

'---------------------------------------------------------------------------------------
' Procedure : SelectBackColor
' DateTime  : 27/11/05 09:26
' Author    : HACKPRO TM
' Purpose   : Set the Selection Cell BackColor.
'---------------------------------------------------------------------------------------
Public Property Let SelectBackColor(ByVal lSelectBackColor As OLE_COLOR)
 m_lSelectBackColor = ConvertSystemColor(lSelectBackColor)
 Call PropertyChanged("SelectBackColor")
End Property

Public Property Get SelectionAlphaBlend() As Boolean
 SelectionAlphaBlend = m_bAlphaBlendSel
End Property

Public Property Let SelectionAlphaBlend(ByVal bState As Boolean)
 If Not (m_bAlphaBlendSel = bState) Then
  m_bAlphaBlendSel = bState
  Call PropertyChanged("SelectionAlphaBlend")
 End If
End Property

'---------------------------------------------------------------------------------------
' Procedure : TextHeaders
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Text of the Headers of grid.
'---------------------------------------------------------------------------------------
Public Property Get TextHeaders() As String
 TextHeaders = m_sTextHeaders
End Property

'---------------------------------------------------------------------------------------
' Procedure : TextHeaders
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Text of the Headers of grid.
'---------------------------------------------------------------------------------------
Public Property Let TextHeaders(ByVal sTextHeaders As String)
 m_sTextHeaders = sTextHeaders
 If (Ambient.UserMode = False) Then TotalItems = 1
 m_Text = Split(m_sTextHeaders, "|")
 Cols = UBound(m_Text)
 If (Cols >= 0) Then
  Dim i As Long
  
  ReDim Preserve Headers(Cols)
  For i = 0 To Cols
   Headers(i).Visible = True
  Next
 End If
 Call PropertyChanged("TextHeaders")
 Call Refresh
End Property

'---------------------------------------------------------------------------------------
' Procedure : AddItem
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Add a new item.
'---------------------------------------------------------------------------------------
Public Sub AddItem(ByVal Item As String, Optional ByVal Colors As String = "", Optional ByVal TextColors As String = "", Optional ByVal Style As String = "", Optional ByVal TextList As String = "", Optional ByVal Values As String = "", Optional ByVal Enabled As String = "", Optional ByVal tFont As StdFont)
 Dim iCarac As Variant, i As Integer, tValor As Boolean, iCols As Long
 
 TotalItems = TotalItems + 1
 ReDim Preserve Items(TotalItems)
 iCarac = Split(m_sTextHeaders, "|")
 iCols = UBound(iCarac)
 Items(TotalItems).Col = iCols
 ReDim Preserve xRow(TotalItems)
 ReDim Preserve yRow(TotalItems)
On Error Resume Next
 xRow(TotalItems) = ScaleX(DEFAULTWIDTH, UserControl.Parent.ScaleMode, vbPixels)
 yRow(TotalItems) = ScaleY(DEFAULTHEIGHT, UserControl.Parent.ScaleMode, vbPixels)
 iCarac = Split(Item, "|")
 If (UBound(iCarac) >= 0) Or (Trim$(iCarac) = "") Then
  ReDim Preserve Items(TotalItems).Item.Colors(iCols)
  ReDim Preserve Items(TotalItems).Item.Enabled(iCols)
  ReDim Preserve Items(TotalItems).Item.Item(iCols)
  ReDim Preserve Items(TotalItems).Item.tPicture(iCols)
  ReDim Preserve Items(TotalItems).Item.Style(iCols)
  ReDim Preserve Items(TotalItems).Item.TextColors(iCols)
  ReDim Preserve Items(TotalItems).Item.TextList(iCols)
  ReDim Preserve Items(TotalItems).Item.Values(iCols)
  For i = 0 To Items(TotalItems).Col
   Items(TotalItems).Item.Colors(i) = ConvertSystemColor(UserControl.BackColor)
   Items(TotalItems).Item.Enabled(i) = True
   If (tFont = "") Then
    Set Items(TotalItems).tFont = mFont
   Else
    Set Items(TotalItems).tFont = tFont
   End If
   Items(TotalItems).Item.Item(i) = ""
   Set Items(TotalItems).Item.tPicture(i) = Nothing
   Items(TotalItems).Item.TextList(i) = ""
   Items(TotalItems).Item.Style(i) = ""
   Items(TotalItems).Item.TextColors(i) = &H0
   Items(TotalItems).Item.Values(i) = False
  Next
 End If
 For i = 0 To UBound(iCarac)
  Items(TotalItems).Item.Item(i) = iCarac(i)
 Next
 iCarac = Split(Colors, "|")
 For i = 0 To UBound(iCarac)
  Items(TotalItems).Item.Colors(i) = IIf(iCarac(i) = "", ConvertSystemColor(UserControl.BackColor), iCarac(i))
 Next
 iCarac = Split(TextColors, "|")
 For i = 0 To UBound(iCarac) - 1
  Items(TotalItems).Item.TextColors(i) = IIf(iCarac(i) = "", &H0, iCarac(i))
 Next
 iCarac = Split(Style, "|")
 For i = 0 To UBound(iCarac)
  If (InControl(iCarac(i)) = True) Then
   Items(TotalItems).Item.Style(i) = iCarac(i)
  Else
   Items(TotalItems).Item.Style(i) = ""
  End If
 Next
 iCarac = Split(Values, "|")
 For i = 0 To UBound(iCarac)
  If ((iCarac(i) <> "T") And (iCarac(i) <> "F")) Or (iCarac(i) = "F") Then
   tValor = False
  ElseIf (iCarac(i) = "T") Then
   tValor = True
  End If
  Items(TotalItems).Item.Values(i) = tValor
 Next
 iCarac = Split(Enabled, "|")
 For i = 0 To UBound(iCarac)
  If ((iCarac(i) <> "T") And (iCarac(i) <> "F")) Or (iCarac(i) = "T") Then
   tValor = True
  ElseIf (iCarac(i) = "F") Then
   tValor = False
  End If
  Items(TotalItems).Item.Enabled(i) = tValor
 Next
 iCarac = Split(TextList, "|")
 For i = 0 To UBound(iCarac)
  Items(TotalItems).Item.TextList(i) = iCarac(i)
 Next
 Items(TotalItems).Row = TotalItems
End Sub

'---------------------------------------------------------------------------------------
' Procedure : APIFillRect
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Pinta el rectángulo de un objeto.
'---------------------------------------------------------------------------------------
Private Sub APIFillRect(ByVal hDC As Long, ByRef rc As RECT, ByVal Color As Long, Optional ByVal Selected As Boolean = False)
 Dim NewBrush As Long
 
 NewBrush& = CreateSolidBrush(Color&)
 Call FillRect(hDC&, rc, NewBrush&)
 Call DeleteObject(NewBrush&)
 If (Selected = True) Then
  NewBrush& = CreateSolidBrush(m_lBorderColor)
  Call FrameRect(hDC&, rc, NewBrush&)
  Call DeleteObject(NewBrush&)
 End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : APILine
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Pinta líneas de forma sencilla y rápida.
'---------------------------------------------------------------------------------------
Private Sub APILine(ByVal whDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal lColor As Long)
 Dim PT As POINTAPI, hPen As Long, hPenOld As Long
 
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(whDC, hPen)
 Call MoveToEx(whDC, x1, y1, PT)
 Call LineTo(whDC, x2, y2)
 Call SelectObject(whDC, hPenOld)
 Call DeleteObject(hPen)
End Sub

Private Function APIRectangle(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional ByVal lColor As OLE_COLOR = -1) As Long
 Dim hPen As Long, hPenOld As Long
 Dim PT   As POINTAPI
 
 hPen = CreatePen(0, 1, lColor)
 hPenOld = SelectObject(hDC, hPen)
 Call MoveToEx(hDC, X, Y, PT)
 Call LineTo(hDC, X + W, Y)
 Call LineTo(hDC, X + W, Y + H)
 Call LineTo(hDC, X, Y + H)
 Call LineTo(hDC, X, Y)
 Call SelectObject(hDC, hPenOld)
 Call DeleteObject(hPen)
End Function

Private Function BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR, Optional ByVal Alpha As Long = 128) As Long
 Dim lCFrom As Long, lCTo   As Long
 Dim lSrcR  As Long, lSrcG  As Long
 Dim lSrcB  As Long, lDstR  As Long
 Dim lDstG  As Long, lDstB  As Long
 
 lCFrom = ConvertSystemColor(oColorFrom)
 lCTo = ConvertSystemColor(oColorTo)
 lSrcR = lCFrom And &HFF
 lSrcG = (lCFrom And &HFF00&) \ &H100&
 lSrcB = (lCFrom And &HFF0000) \ &H10000
 lDstR = lCTo And &HFF
 lDstG = (lCTo And &HFF00&) \ &H100&
 lDstB = (lCTo And &HFF0000) \ &H10000
 BlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))
End Function

'---------------------------------------------------------------------------------------
' Procedure : BorderObject
' DateTime  : 10/07/05 12:48
' Author    : HACKPRO TM
' Purpose   : Cambia el valor de un item.
'---------------------------------------------------------------------------------------
Private Sub BorderObject()
 m_txtRect.hLeft = 0
 m_txtRect.hTop = 0
 m_txtRect.hRight = UserControl.ScaleWidth
 m_txtRect.hBottom = UserControl.ScaleHeight
 Call DrawEdge(hDC, m_txtRect, EDGE_SUNKEN, BF_RECT)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChangedBackColor
' DateTime  : 03/07/05 15:39
' Author    : HACKPRO TM
' Purpose   : Color de fondo de la celda.
'---------------------------------------------------------------------------------------
Public Sub ChangedBackColor(ByVal Row As Integer, ByVal Col As Integer, ByVal Color As OLE_COLOR)
On Error Resume Next
 Items(Col).Item.Colors(Row) = Color
 Call Refresh
On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChangedEnabled
' DateTime  : 03/07/05 15:39
' Author    : HACKPRO TM
' Purpose   : Cambia el valor de un item.
'---------------------------------------------------------------------------------------
Public Sub ChangedEnabled(ByVal Row As Integer, ByVal Col As Integer, ByVal Value As Boolean)
On Error Resume Next
 Items(Col).Item.Enabled(Row) = Value
 Call Refresh
On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChangedForeColor
' DateTime  : 03/07/05 15:39
' Author    : HACKPRO TM
' Purpose   : Color del texto de la celda.
'---------------------------------------------------------------------------------------
Public Sub ChangedForeColor(ByVal Row As Integer, ByVal Col As Integer, ByVal Color As OLE_COLOR)
On Error Resume Next
 Items(Col).Item.TextColors(Row) = Color
 Call Refresh
On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChangedItem
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Cambia el valor de un item.
'---------------------------------------------------------------------------------------
Public Sub ChangedItem(ByVal Row As Integer, ByVal Col As Integer, ByVal Item As String)
On Error Resume Next
 Items(Col).Item.Item(Row) = Item
 Call Refresh
On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChangedStyle
' DateTime  : 03/07/05 15:39
' Author    : HACKPRO TM
' Purpose   : Cambia el estilo de un item.
'---------------------------------------------------------------------------------------
Public Sub ChangedStyle(ByVal Row As Integer, ByVal Col As Integer, ByVal Item As String)
On Error Resume Next
 Items(Col).Item.Style(Row) = Item
 Call Refresh
On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ChangedValue
' DateTime  : 03/07/05 15:39
' Author    : HACKPRO TM
' Purpose   : Cambia el valor de un item.
'---------------------------------------------------------------------------------------
Public Sub ChangedValue(ByVal Row As Integer, ByVal Col As Integer, ByVal Value As Boolean)
 Dim i As Integer
 
On Error Resume Next
 If (Items(Col).Item.Style(Row) = "O") Then
  For i = 1 To TotalItems
   Items(i).Item.Values(Row) = False
  Next
 End If
 Items(Col).Item.Values(Row) = Value
 Call Refresh
On Error GoTo 0
End Sub

Public Sub Clear()
On Error Resume Next
 TotalItems = 0
 ReDim Items(1)
 Call DestroyWindow(m_hWNdH)
 Call DestroyWindow(m_hWNdV)
 Erase xRow
 Erase yRow
 Erase Items
On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ColumnWidth
' DateTime  : 29/10/05 01:04
' Author    : HACKPRO TM
' Purpose   : Width of the Column.
'---------------------------------------------------------------------------------------
Public Function ColumnWidth(ByVal Row As Long, ByVal Width As Long) As Long
On Error Resume Next
 xRow(Row) = ScaleX(Width, UserControl.Parent.ScaleMode, vbPixels)
On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : ConvertSystemColor
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Convierte un Long en un color del sistema.
'---------------------------------------------------------------------------------------
Private Function ConvertSystemColor(ByVal theColor As Long) As Long
 Call OleTranslateColor(theColor, 0, ConvertSystemColor)
End Function

'---------------------------------------------------------------------------------------
' Procedure : DrawCaption
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Paint Text of the object.
'---------------------------------------------------------------------------------------
Private Sub DrawCaption(ByVal lCaption As String, Optional ByVal lColor As OLE_COLOR = &H0, Optional ByVal iControl As Boolean = False)
 Dim iAlign As String, RemF As Boolean, fAlign As Long
 
 isLine = False
 If (lCaption = "-") Then
  Call DrawHGradient(m_lBorderColor, m_lBackColor, 1, m_txtRect.hTop - 5, UserControl.ScaleWidth - 35, m_txtRect.hTop - 4)
  Exit Sub
 End If
 iAlign = Mid$(lCaption, 1, 1)
 RemF = True
 fAlign = 0
 Select Case iAlign
  Case "^" '* Center Align.
   fAlign = DT_CENTER
  Case "~" '* Right Align.
   fAlign = DT_RIGHT
   If (iControl = True) Then
    m_txtRect.hRight = m_txtRect.hRight - 22
   Else
    m_txtRect.hRight = m_txtRect.hRight - 5
   End If
  Case Else
   RemF = False
   lCaption = " " & lCaption
 End Select
 If (RemF = True) Then lCaption = Mid$(Trim$(lCaption), 2, Len(lCaption))
 Call SetTextColor(hDC, lColor)
 '*************************************************************************
 '* Draws the text with Unicode support based on OS version.              *
 '* Thanks to Richard Mewett.                                             *
 '*************************************************************************
 If (mWindowsNT = True) Then
  Call DrawTextW(UserControl.hDC, StrPtr(lCaption), Len(lCaption), m_txtRect, DT_TOP Or DT_VCENTER Or DT_WORD_ELLIPSIS Or DT_WORDBREAK Or DT_EDITCONTROL Or fAlign)
 Else
  Call DrawTextA(UserControl.hDC, lCaption, Len(lCaption), m_txtRect, DT_TOP Or DT_VCENTER Or DT_WORD_ELLIPSIS Or DT_WORDBREAK Or DT_EDITCONTROL Or fAlign)
 End If
 If (iControl = True) And (fAlign = DT_RIGHT) Then
  m_txtRect.hRight = m_txtRect.hRight + 22
 ElseIf (fAlign = DT_RIGHT) Then
  m_txtRect.hRight = m_txtRect.hRight + 5
 End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DrawHeaders
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Paint the text headers.
'---------------------------------------------------------------------------------------
Private Function DrawHeaders(Optional ByVal Value As Long = 0, Optional ByVal iPos As Long = -1, Optional ByVal hPos As Integer = 0) As Long
 Dim lBack  As OLE_COLOR, i   As Long, Fo As Integer
 Dim xRight As Long, hVisible As Long
 
On Error Resume Next
 m_txtRect.hRight = 0
 m_txtRect.hTop = 1
 m_txtRect.hBottom = ScaleY(mColumnHeadingH, UserControl.Parent.ScaleMode, vbPixels)
 m_txtRect.hLeft = 0
 If (Cols > 0) Then
  hVisible = UserControl.ScaleWidth \ (Cols * UserControl.TextWidth("Qr"))
  If (hVisible > Cols) Then hVisible = Cols
 End If
 For i = Value To hVisible
  If (Headers(i).Visible = True) Then
   If (b_AutoSizeColumn = True) Then
    m_txtRect.hRight = m_txtRect.hRight + UserControl.TextWidth(m_Text(i)) + 10
   ElseIf (Ambient.UserMode = False) Then
    If (UBound(xRow) = 0) Then m_txtRect.hRight = m_txtRect.hRight + UserControl.TextWidth(m_Text(i)) + 10
   Else
    m_txtRect.hRight = m_txtRect.hRight + xRow(i)
   End If
   lBack = ConvertSystemColor(UserControl.BackColor)
   Call APIFillRect(hDC, m_txtRect, m_lHeadersColor)
   Fo = 0
   If (i = iPos) Then Fo = hPos
   bDrawTheme = DrawTheme("Header", 1, Fo, m_txtRect)
   If (bDrawTheme = False) Then
    If (Fo = 3) Then
     Call DrawEdge(hDC, m_txtRect, EDGE_ETCHED, BF_RECT)
    Else
     Call DrawEdge(hDC, m_txtRect, EDGE_RAISED, BF_RECT)
    End If
   Else
    Call DrawThemeEdge(hTheme, hDC, 1, 0, m_txtRect, BDR_RAISEDINNER, BF_RECT, m_txtRect)
   End If
   xRight = m_txtRect.hRight
   Headers(i).Width = m_txtRect.hRight
   Headers(i).Height = m_txtRect.hBottom
   If (i = iPos) And (hPos = 3) Then
    m_txtRect.hTop = 6
    m_txtRect.hLeft = m_txtRect.hLeft + 2
   Else
    m_txtRect.hTop = 4
   End If
   Call DrawCaption(m_Text(i))
   m_txtRect.hLeft = xRight
   m_txtRect.hTop = 2
  Else
   Headers(i).Width = m_txtRect.hRight
   Headers(i).Height = m_txtRect.hBottom
  End If
 Next
 DrawHeaders = hVisible + i
 If (hVisible + 1 > Cols) And (hPos = 0) Then
  lBack = ConvertSystemColor(UserControl.BackColor)
  m_txtRect.hRight = m_txtRect.hRight + UserControl.ScaleWidth
  Call APIFillRect(hDC, m_txtRect, m_lHeadersColor)
  bDrawTheme = DrawTheme("Header", 1, Fo, m_txtRect)
  If (bDrawTheme = False) Then
   Call DrawEdge(hDC, m_txtRect, EDGE_RAISED, BF_RECT)
  Else
   Call DrawThemeEdge(hTheme, hDC, 1, 0, m_txtRect, BDR_RAISEDINNER, BF_RECT, m_txtRect)
  End If
 End If
 Call CloseThemeData(hTheme)
On Error GoTo 0
End Function

Private Sub DrawHGradient(ByVal lEndColor As Long, ByVal lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal x2 As Long, ByVal y2 As Long)
 Dim dR As Single, dG As Single, dB As Single
 Dim sR As Single, sG As Single, sB As Single
 Dim eR As Single, eG As Single, eB As Single
 Dim lh As Long, lw   As Long, ni   As Long
  
 '* Draw a Horizontal Gradient in the current hDC.
 lh = y2 - Y
 lw = x2 - X
 sR = (lStartcolor And &HFF)
 sG = (lStartcolor \ &H100) And &HFF
 sB = (lStartcolor And &HFF0000) / &H10000
 eR = (lEndColor And &HFF)
 eG = (lEndColor \ &H100) And &HFF
 eB = (lEndColor And &HFF0000) / &H10000
 dR = (sR - eR) / lw
 dG = (sG - eG) / lw
 dB = (sB - eB) / lw
 For ni = 0 To lw
  Call APILine(UserControl.hDC, X + ni, Y, X + ni, y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
 Next
End Sub

Private Sub DrawScrollBar(Optional ByVal iType As Integer = 0)
 Dim lMajor As Long, lMinor As Long
 Dim lR     As Long
 
 isXp = False
 Call GetWindowsVersion(lMajor, lMinor)
 If (lMajor > 5) Then
  isXp = True
 ElseIf (lMajor = 5) And (lMinor >= 1) Then
  isXp = True
 End If
 lStyle = WS_CHILD Or WS_VISIBLE
 If Not (m_hWNdV) And (iType = 1) Then
  lStyle = lStyle Or SB_VERT And Not SBS_HORZ
  m_hWNdV = CreateWindowEx(0, "SCROLLBAR", "", lStyle, UserControl.ScaleWidth - 19, 2, 17, UserControl.ScaleHeight - 22, hWnd, 0, App.hInstance, ByVal 0&)
  If (isXp = True) And (m_bNoFlatScrollBars = False) Then
   Call ShowScrollBar(m_hWNdV, SB_CTL, 0)
  Else
   Call FlatSB_ShowScrollBar(m_hWNdV, SB_CTL, False)
  End If
 ElseIf Not (m_hWNdH) And (iType = 2) Then
  lStyle = lStyle Or SB_HORZ And Not SBS_VERT
  m_hWNdH = CreateWindowEx(0, "SCROLLBAR", "", lStyle, 2, UserControl.ScaleHeight - 20, UserControl.ScaleWidth - 4, 18, hWnd, 0, App.hInstance, ByVal 0&)
  If (isXp = True) And (m_bNoFlatScrollBars = False) Then
   Call ShowScrollBar(m_hWNdH, SB_CTL, 0)
  Else
   Call FlatSB_ShowScrollBar(m_hWNdH, SB_CTL, False)
  End If
 End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DrawPoint
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Pinta tres puntos.
'---------------------------------------------------------------------------------------
Private Sub DrawPoint(Optional ByVal iColor3 As OLE_COLOR = &H0)
 iColor3 = ConvertSystemColor(iColor3)
 tmpC1 = m_Buttons.hRight - m_Buttons.hLeft - 3
 tmpC2 = m_Buttons.hBottom - m_Buttons.hTop + 1
 tmpC1 = m_Buttons.hLeft + tmpC1 / 2 + 1
 tmpC2 = m_Buttons.hTop + tmpC2 / 2 - 1
 Call APILine(UserControl.hDC, tmpC1 - 3, tmpC2, tmpC1 - 3, tmpC2 + 1, iColor3)
 Call APILine(UserControl.hDC, tmpC1, tmpC2, tmpC1, tmpC2 + 1, iColor3)
 Call APILine(UserControl.hDC, tmpC1 + 3, tmpC2, tmpC1 + 3, tmpC2 + 1, iColor3)
End Sub

'---------------------------------------------------------------------------------------
' Function  : DrawTheme
' DateTime  : 03/08/05 13:38
' Author    : HACKPRO TM
' Purpose   : Try to open Uxtheme.dll.
'---------------------------------------------------------------------------------------
Private Function DrawTheme(sClass As String, ByVal iPart As Long, ByVal iState As Long, rtRect As RECT, Optional ByVal CloseTheme As Boolean = False) As Boolean
 Dim lResult As Long '* Temp Variable.
 
 '* If a error occurs then or we are not running XP or the visual style is Windows Classic.
On Error GoTo NoXP
 '* Get out hTheme Handle.
 hTheme = OpenThemeData(UserControl.hWnd, StrPtr(sClass))
 '* Did we get a theme handle?.
 If (hTheme) Then
  '* Yes! Draw the control Background.
  lResult = DrawThemeBackground(hTheme, UserControl.hDC, iPart, iState, rtRect, rtRect)
  '* If drawing was successful, return true, or false If not.
  DrawTheme = IIf(lResult, False, True)
 Else
  '* No, we couldn't get a hTheme, drawing failed.
  DrawTheme = False
 End If
 '* Close theme.
 If (CloseTheme = True) Then Call CloseThemeData(hTheme)
 '* Exit the function now.
 Exit Function
NoXP:
 '* An Error was detected, drawing Failed.
 DrawTheme = False
On Error GoTo 0
End Function

Private Sub DrawVGradient(ByVal lEndColor As Long, ByVal lStartcolor As Long, ByVal X As Long, ByVal Y As Long, ByVal x2 As Long, ByVal y2 As Long)
 Dim dR As Single, dG As Single, dB As Single, ni As Long
 Dim sR As Single, sG As Single, sB As Single
 Dim eR As Single, eG As Single, eB As Single
 
 '* Draw a Vertical Gradient in the current hDC.
 sR = (lStartcolor And &HFF)
 sG = (lStartcolor \ &H100) And &HFF
 sB = (lStartcolor And &HFF0000) / &H10000
 eR = (lEndColor And &HFF)
 eG = (lEndColor \ &H100) And &HFF
 eB = (lEndColor And &HFF0000) / &H10000
 dR = (sR - eR) / y2
 dG = (sG - eG) / y2
 dB = (sB - eB) / y2
 For ni = 0 To y2
  Call APILine(UserControl.hDC, X, Y + ni, x2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB)))
 Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DrawXpArrow
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Dibuja la flecha estilo Xp.
'---------------------------------------------------------------------------------------
Private Sub DrawXpArrow(Optional ByVal iColor3 As OLE_COLOR = &H0)
 Dim tmpC1 As Long, tmpC2 As Long
 
 iColor3 = ConvertSystemColor(iColor3)
 tmpC1 = m_Buttons.hRight - m_Buttons.hLeft - 1
 tmpC2 = m_Buttons.hBottom - m_Buttons.hTop + 1
 tmpC1 = m_Buttons.hLeft + tmpC1 / 2 + 1
 tmpC2 = m_Buttons.hTop + tmpC2 / 2
 Call APILine(UserControl.hDC, tmpC1 - 5, tmpC2 - 2, tmpC1, tmpC2 + 3, iColor3)
 Call APILine(UserControl.hDC, tmpC1 - 4, tmpC2 - 2, tmpC1, tmpC2 + 2, iColor3)
 Call APILine(UserControl.hDC, tmpC1 - 4, tmpC2 - 3, tmpC1, tmpC2 + 1, iColor3)
 Call APILine(UserControl.hDC, tmpC1 + 3, tmpC2 - 2, tmpC1 - 2, tmpC2 + 3, iColor3)
 Call APILine(UserControl.hDC, tmpC1 + 2, tmpC2 - 2, tmpC1 - 2, tmpC2 + 2, iColor3)
 Call APILine(UserControl.hDC, tmpC1 + 2, tmpC2 - 3, tmpC1 - 2, tmpC2 + 1, iColor3)
End Sub

Private Function FindFirstEnabled(Optional ByVal BeginF As Boolean = True, Optional ByVal ForCols As Boolean = True) As Long
 Dim iPos As Long
 
On Error Resume Next
 If (ForCols = True) Then '* Find in the grid for Vertical Pos.
  If (BeginF = True) Then '* Find First Item.
   For iPos = 1 To TotalItems
    If (Items(iPos).Item.Enabled(RowItem) = True) Then Exit For
   Next
  Else                     '* Find Last Item.
   For iPos = TotalItems To 1 Step -1
    If (Items(iPos).Item.Enabled(RowItem) = True) Then Exit For
   Next
  End If
 Else '* Find in the grid for Horizontal Pos.
  If (BeginF = True) Then '* Find First Item.
   For iPos = 0 To Cols
    If (Headers(iPos).Visible = True) Then
     If (Items(ColItem).Item.Enabled(iPos) = True) Then Exit For
    End If
   Next
  Else                     '* Find Last Item.
   For iPos = Cols To 0 Step -1
    If (Headers(iPos).Visible = True) Then
     If (Items(ColItem).Item.Enabled(iPos) = True) Then Exit For
    End If
   Next
  End If
 End If
 FindFirstEnabled = iPos
On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : FormShow
' DateTime  : 09/07/05 13:35
' Author    : HACKPRO TM
' Purpose   : Establece los controles para Button.
'---------------------------------------------------------------------------------------
Private Sub FormShow(ByRef isObject As Object, ByRef isText As Object)
On Error Resume Next
 isText.Text = ItemText(RowItem, ColItem)
 Call isObject.Show(1)
On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetColFromX
' DateTime  : 09/07/05 09:19
' Author    : Richard Mewett (Thanks)
' Purpose   : Position Col X.
'---------------------------------------------------------------------------------------
Private Function GetColFromX(ByVal X As Single) As Integer
 Dim nCol As Integer
    
On Error Resume Next
 GetColFromX = -1
 For nCol = Value2 To UBound(Headers)
  If (X > Tamano(nCol).lStartX) And (X <= Tamano(nCol).lStartX + Headers(nCol).Width) Then
   GetColFromX = nCol
   Exit For
  End If
 Next
On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetControlVersion
' DateTime  : 09/07/05 18:00
' Author    : HACKPRO TM
' Purpose   : Control Version.
'---------------------------------------------------------------------------------------
Public Function GetControlVersion() As String
 GetControlVersion = Version & " © " & Year(Now)
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetRowFromY
' DateTime  : 09/07/05 09:19
' Author    : Richard Mewett (Thanks)
' Purpose   : Position Col Y.
'---------------------------------------------------------------------------------------
Private Function GetRowFromY(ByVal Y As Single) As Integer
 Dim nCol As Integer
    
On Error Resume Next
 GetRowFromY = -1
 For nCol = Value1 To Value1 + ShowItems
  If (Y <= Tamano(nCol).lStartY) Then
   GetRowFromY = nCol
   Exit For
  End If
 Next
 If (nCol = Value1) Then GetRowFromY = 0
On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetWindowsVersion
' DateTime  : 09/07/05 18:00
' Author    : HACKPRO TM
' Purpose   : OS Version.
'---------------------------------------------------------------------------------------
Private Sub GetWindowsVersion(Optional ByRef lMajor = 0, Optional ByRef lMinor = 0, Optional ByRef lRevision = 0, Optional ByRef lBuildNumber = 0)
 Dim lR As Long
 
 lR = GetVersion()
 lBuildNumber = (lR And &H7F000000) \ &H1000000
 If (lR And &H80000000) Then lBuildNumber = lBuildNumber Or &H80
 lRevision = (lR And &HFF0000) \ &H10000
 lMinor = (lR And &HFF00&) \ &H100
 lMajor = (lR And &HFF)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HideCol
' DateTime  : 26/11/05 19:07
' Author    : HACKPRO TM
' Purpose   : Hide or Show any Col.
'---------------------------------------------------------------------------------------
Public Sub HideCol(ByVal index As Long, ByVal bVisible As Boolean)
 Headers(index).Visible = bVisible
End Sub

'---------------------------------------------------------------------------------------
' Procedure : InControl
' DateTime  : 09/07/05 16:29
' Author    : HACKPRO TM
' Purpose   : Verifica que sea un control como B, T, ...
'---------------------------------------------------------------------------------------
Private Function InControl(ByVal iControl As String) As Boolean
 InControl = False
 Select Case iControl
  Case "T":  InControl = True
  Case "Ch": InControl = True
  Case "C":  InControl = True
  Case "O":  InControl = True
  Case "B":  InControl = True
  Case "Bk": InControl = True
  Case Else: InControl = True
 End Select
End Function

Private Function InFocusControl(ByVal ObjecthWnd As Long) As Boolean
 Dim mPos As POINTAPI, oRect As RECT
 
 '* Verifies if the mouse is on the object or if one makes clic outside of him.
 Call GetCursorPos(mPos)
 Call GetWindowRect(ObjecthWnd, oRect)
 If (mPos.X >= oRect.hLeft) And (mPos.X <= oRect.hRight) And (mPos.Y >= oRect.hTop) And (mPos.Y <= oRect.hBottom) Then
  InFocusControl = True
 End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ItemText
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Muestra el item actual.
'---------------------------------------------------------------------------------------
Public Function ItemText(ByVal Col As Integer, ByVal Row As Integer) As String
 Dim iPos As String
 
On Error Resume Next
 If (Col = -1) Or (Row = -1) Then
  ItemText = ""
  Exit Function
 End If
 iPos = Mid$(Items(Col).Item.Item(Row), 1, 1)
 If (iPos = "^") Or (iPos = "~") Then
  ItemText = Mid$(Items(Col).Item.Item(Row), 2, Len(Items(Col).Item.Item(Row)))
 Else
  ItemText = Items(Col).Item.Item(Row)
 End If
On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : ItemStyle
' DateTime  : 03/07/05 17:38
' Author    : HACKPRO TM
' Purpose   : Devuelve el valor contenido.
'---------------------------------------------------------------------------------------
Public Function ItemStyle(ByVal Col As Integer, ByVal Row As Integer) As String
On Error Resume Next
 ItemStyle = Items(Col).Item.Style(Row)
On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : ItemValue
' DateTime  : 03/07/05 17:38
' Author    : HACKPRO TM
' Purpose   : Devuelve el valor contenido.
'---------------------------------------------------------------------------------------
Public Function ItemValue(ByVal Col As Integer, ByVal Row As Integer) As Boolean
On Error Resume Next
 ItemValue = Items(Col).Item.Values(Row)
On Error GoTo 0
End Function

'---------------------------------------------------------------------------------------
' Procedure : ObjectForm
' DateTime  : 09/07/05 13:29
' Author    : HACKPRO TM
' Purpose   : Establece el Form y Text cuando style es B.
'---------------------------------------------------------------------------------------
Public Function ObjectForm(ByRef isObject As Object, ByRef isText As Object)
On Error Resume Next
 Set TheForm = isObject
 Set TheText = isText
On Error GoTo 0
End Function

'* Offset a color.
Private Function OffSetColor(ByVal lColor As OLE_COLOR, ByVal lOffset As Long) As OLE_COLOR
 Dim lRed  As OLE_COLOR, lGreen As OLE_COLOR
 Dim lBlue As OLE_COLOR, lR     As OLE_COLOR
 Dim lG    As OLE_COLOR, lB     As OLE_COLOR
   
 lR = (lColor And &HFF)
 lG = ((lColor And 65280) \ 256)
 lB = ((lColor) And 16711680) \ 65536
 lRed = (lOffset + lR)
 lGreen = (lOffset + lG)
 lBlue = (lOffset + lB)
 If (lRed > 255) Then lRed = 255
 If (lRed < 0) Then lRed = 0
 If (lGreen > 255) Then lGreen = 255
 If (lGreen < 0) Then lGreen = 0
 If (lBlue > 255) Then lBlue = 255
 If (lBlue < 0) Then lBlue = 0
 OffSetColor = RGB(lRed, lGreen, lBlue)
End Function

'---------------------------------------------------------------------------------------
' Procedure : ReDraw
' DateTime  : 12/07/05 21:45
' Author    : HACKPRO TM
' Purpose   : Redibuja el control.
'---------------------------------------------------------------------------------------
Public Sub ReDraw()
 Call Refresh(, True)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Refresh
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Refresca el control.
'---------------------------------------------------------------------------------------
Private Sub Refresh(Optional ByVal UpDown As Boolean = False, Optional ByVal isNew As Boolean = False, Optional ByVal isDblClick As Boolean = False)
 Dim CountItems As Integer, i         As Long, tmpRect As RECT, iCtrl As Boolean, isOk As Boolean
 Dim iPos       As String, xRight     As Long, j       As Long, iHeig As Long, l_nBack As OLE_COLOR
 Dim lColor     As OLE_COLOR, mColor  As Long, isOpt   As Integer, nFont As StdFont, k As Long
 Dim lScrollH   As Boolean, lTile     As Boolean
 
On Error Resume Next
 ColsV = 0
 If (SafeUBound(VarPtrArray(Headers)) >= 0) Then
  For i = 0 To Cols - 1
   If (Headers(i).Visible = True) Then ColsV = ColsV + 1
  Next
 End If
 Set nFont = New StdFont
 UserControl.BackColor = GetSysColor(COLOR_WINDOW)
 iHeig = ScaleY(mColumnHeadingH, UserControl.Parent.ScaleMode, vbPixels)
 ShowItems = TotalHeight
 m_hWnd = m_hWNdV
 If (isNew = True) Then Call ScrollVisible(False, "V")
 If (TotalItems >= ShowItems) Then
  If (isNew = True) Then Max = (TotalItems - ShowItems) + 2
  If (Ambient.UserMode = True) And (Max >= 0) Then Call ScrollVisible(True, "V")
 Else
  Max = TotalItems
  ShowItems = TotalItems
  Call ScrollVisible(False, "V")
 End If
 If (lVScroll = True) Then
  Value1 = Value
 Else
  Value1 = 0
 End If
 SmallChange = 3
 LargeChange = 3
 UserControl.Cls
 UserControl.BackColor = m_lBackColor
 Call SetRect(m_txtRect, 4, 4, UserControl.ScaleWidth - 4, UserControl.ScaleHeight - 4)
 m_hWnd = m_hWNdH
 CountItems = DrawHeaders(Value2)
 If (lHScroll = True) Then
  Value2 = Value
 Else
  Value2 = 0
 End If
 SmallChange = 3
 LargeChange = 3
 If (isNew = True) Then Call ScrollVisible(False, "H")
 If (TotalItems > 0) Then
  If (TotalWidth > UserControl.ScaleWidth) Then
   If (isNew = True) Then Max = ColsV - 2
   If (Ambient.UserMode = True) And (Max >= 0) Then Call ScrollVisible(True, "H")
  End If
 End If
 lScrollH = False
 If (TotalItems < ShowItems) Then
  Call MoveWindow(m_hWnd, 2, UserControl.ScaleHeight - 20, UserControl.ScaleWidth - 2, 18, 1)
 Else
  Call MoveWindow(m_hWnd, 2, UserControl.ScaleHeight - 20, UserControl.ScaleWidth - 4, 18, 1)
  lScrollH = True
 End If
 m_hWnd = m_hWNdV
 If Not (TotalWidth > UserControl.ScaleWidth) And (lScrollH = False) Then
  Call MoveWindow(m_hWnd, UserControl.ScaleWidth - 19, 2, 17, UserControl.ScaleHeight - 2, 1)
 ElseIf (lHScroll = False) Then
  Call MoveWindow(m_hWnd, UserControl.ScaleWidth - 19, 2, 17, UserControl.ScaleHeight - 2, 1)
 Else
  Call MoveWindow(m_hWnd, UserControl.ScaleWidth - 19, 2, 17, UserControl.ScaleHeight - 22, 1)
 End If
 iHeig = ScaleY(mColumnHeadingH, UserControl.Parent.ScaleMode, vbPixels)
 m_Buttons.hTop = iHeig + 3
 m_Buttons.hBottom = iHeig + 18
 Set nFont = UserControl.Font
 ReDim Tamano(1)
 isOk = False
 For i = Value1 To (ShowItems + Value1)
  ReDim Preserve Tamano(i)
  m_txtRect.hLeft = 2
  m_txtRect.hRight = 0
  If (i < TotalItems) Then
   iHeig = yRow(i + 1)
  Else
   iHeig = ScaleY(mColumnHeadingH, UserControl.Parent.ScaleMode, vbPixels)
  End If
  Call OffsetRect(m_txtRect, 0, iHeig - 2)
  Tamano(i).lStartY = m_txtRect.hTop
  For j = Value2 To CountItems
   k = j
   If (Headers(k).Visible = True) Then
    k = i + 1
    If (j = CountItems) Then
     m_txtRect.hRight = UserControl.ScaleWidth
    Else
     m_txtRect.hRight = Headers(j).Width
     m_Buttons.hLeft = Headers(j).Width - 17
     m_Buttons.hRight = Headers(j).Width - 2
    End If
    lColor = ConvertSystemColor(UserControl.BackColor)
    mColor = &H0
    If (j <= UBound(Items(k).Item.Colors)) Then lColor = ConvertSystemColor(Items(k).Item.Colors(j))
    If (j <= UBound(Items(k).Item.TextColors)) Then mColor = ConvertSystemColor(Items(k).Item.TextColors(j))
    If (m_bAlphaBlendSel = True) Then
     l_nBack = BlendColor(m_lSelectBackColor, lColor, 120)
    Else
     l_nBack = m_lSelectBackColor
    End If
    If (lTile = False) And (m_eBorderStyle = &H7) Then
     Call CreatePatternFromStdPicture(m_lBackgroundPic)
     Call Tile(UserControl.hDC, 0, m_txtRect.hTop, UserControl.ScaleWidth, UserControl.ScaleHeight, True)
     lTile = True
    Else 'If Not (m_eBorderStyle = &H7) Then
     If (m_bViewStyle = True) And (isOk = False) And (m_lFullSelect = True) And (k = ColItem) Then
      Call SetRect(tmpRect, 3, m_txtRect.hTop, UserControl.ScaleWidth - 2, m_txtRect.hBottom)
      Call APIFillRect(UserControl.hDC, tmpRect, l_nBack, True)
      isOk = True
     ElseIf (m_bViewStyle = True) And (m_lFullSelect = False) And (k = ColItem) And (j = RowItem) Then
      Call APIFillRect(UserControl.hDC, m_txtRect, l_nBack, True)
     ElseIf (isOk = False) And (lTile = False) Then
      Call APIFillRect(UserControl.hDC, m_txtRect, lColor)
     End If
    End If
    Select Case m_eBorderStyle
     Case &H1: Call DrawEdge(hDC, m_txtRect, EDGE_BUMP, BF_RECT)
     Case &H2: Call DrawEdge(hDC, m_txtRect, EDGE_ETCHED, BF_RECT)
     Case &H3: Call DrawEdge(hDC, m_txtRect, EDGE_BUMP, BF_RECT)
     Case &H4: Call DrawEdge(hDC, m_txtRect, EDGE_SUNKEN, BF_RECT)
     Case &H5: Call DrawEdge(hDC, m_txtRect, EDGE_SUNKEN, BF_RECT Or BF_FLAT) '* Of Richard Mewett.
     Case &H6: Call DrawEdge(hDC, m_txtRect, BDR_RAISEDINNER, BF_RECT)
    End Select
    xRight = m_txtRect.hRight
    m_txtRect.hTop = m_txtRect.hTop + 3
    iCtrl = False
    If (Items(k).Item.Style(j) <> "") Then iCtrl = True
    If Not (Items(k).Item.tPicture(j) Is Nothing) Then
     Call RenderIconGrayscale(UserControl.hDC, Items(k).Item.tPicture(j).Handle, 8, m_txtRect.hTop, 16, 16, False)
     m_txtRect.hLeft = m_txtRect.hLeft + 22
    End If
    Set UserControl.Font = Items(k).tFont
    Call DrawCaption(Items(k).Item.Item(j), mColor, iCtrl)
    Set UserControl.Font = nFont
    iPos = ""
    If (j <= UBound(Items(k).Item.Style)) Then iPos = Items(k).Item.Style(j)
    bDrawTheme = False
    isOpt = 0
    If (Items(k).Item.Values(j) = True) And (UpDown = True) Then isOpt = 3
    If (iPos = "B") Then      '* Button.
     If (Items(k).Item.Enabled(j) = False) Then isOpt = 4
     bDrawTheme = DrawTheme("Button", 1, isOpt, m_Buttons)
     If (bDrawTheme = False) Then
      Call APIFillRect(hDC, m_Buttons, GetSysColor(COLOR_BTNFACE))
      If (ColItem = j) And (UpDown = True) Then
       Call DrawEdge(hDC, m_Buttons, EDGE_SUNKEN, BF_RECT)
      Else
       Call DrawEdge(hDC, m_Buttons, EDGE_RAISED, BF_RECT)
      End If
     End If
     Call DrawPoint(IIf(isEnabled = False, &H80000011, &H80000012))
    ElseIf (k = ColItem) And (j = RowItem) And (isDblClick = True) And ((iPos = "C") Or (iPos = "D")) Then  '* ComboBox or Calendar.
     
     isDblClick = False
    ElseIf (iPos = "Ch") Then  '* CheckBox.
     isOpt = (Items(k).Item.Values(j) * -5)
     If (Items(k).Item.Enabled(j) = False) Then isOpt = 4
     bDrawTheme = DrawTheme("Button", 3, isOpt, m_Buttons)
     If (bDrawTheme = False) Then
      m_Buttons.hTop = m_Buttons.hTop + 1
      m_Buttons.hBottom = m_Buttons.hBottom - 1
      m_Buttons.hLeft = m_Buttons.hLeft + 1
      isOpt = 0
      If (Items(k).Item.Enabled(j) = False) Then isOpt = DFCS_INACTIVE
      If (Items(k).Item.Values(j) = False) Then
       Call DrawFrameControl(hDC, m_Buttons, DFC_BUTTON, DFCS_BUTTONCHECK Or isOpt)
      Else
       Call DrawFrameControl(hDC, m_Buttons, DFC_BUTTON, DFCS_BUTTONCHECK Or DFCS_CHECKED Or isOpt)
      End If
      m_Buttons.hBottom = m_Buttons.hBottom + 1
      m_Buttons.hTop = m_Buttons.hTop - 1
      m_Buttons.hLeft = m_Buttons.hLeft - 1
     End If
    ElseIf (iPos = "O") Then  '* OptionBox.
     isOpt = IIf(Items(k).Item.Values(j) = True, 6, 1)
     If (Items(k).Item.Enabled(j) = False) Then isOpt = 4
     bDrawTheme = DrawTheme("Button", 2, isOpt, m_Buttons)
     If (bDrawTheme = False) Then
      m_Buttons.hTop = m_Buttons.hTop + 1
      m_Buttons.hBottom = m_Buttons.hBottom - 1
      m_Buttons.hLeft = m_Buttons.hLeft + 1
      isOpt = 0
      If (Items(k).Item.Enabled(j) = False) Then isOpt = DFCS_INACTIVE
      If (Items(k).Item.Values(j) = False) Then
       Call DrawFrameControl(hDC, m_Buttons, DFC_BUTTON, DFCS_BUTTONRADIO Or isOpt)
      Else
       Call DrawFrameControl(hDC, m_Buttons, DFC_BUTTON, DFCS_BUTTONRADIO Or DFCS_CHECKED Or isOpt)
      End If
      m_Buttons.hBottom = m_Buttons.hBottom + 1
      m_Buttons.hTop = m_Buttons.hTop - 1
      m_Buttons.hLeft = m_Buttons.hLeft - 1
     End If
    ElseIf (k = ColItem) And (j = RowItem) And (isDblClick = True) And (iPos = "T") Then '* TextBox.
     FirstTime = True
     With txtEdit
      Set .Font = Items(k).tFont
      .Text = ItemText(k, j)
      .ForeColor = Items(k).Item.TextColors(j)
      .BackColor = Items(k).Item.Colors(j)
      If (b_AutoSizeColumn = True) Then
       Call .Move(m_txtRect.hLeft, m_txtRect.hTop - 3, m_txtRect.hRight, yRow(j + 1) - 1)
      Else
       Call .Move(m_txtRect.hLeft, m_txtRect.hTop - 3, xRow(j) + 3, yRow(j + 1) - 1)
      End If
      .Tag = Mid$(Items(k).Item.Item(j), 1, 1)
      .Visible = True
     End With
    ElseIf (iPos = "UD") Then '* UpDown.
     
    End If
    m_txtRect.hLeft = xRight
    m_txtRect.hTop = m_txtRect.hTop - 3
   End If
  Next
  m_Buttons.hTop = m_Buttons.hTop + (iHeig - 2)
  m_Buttons.hBottom = m_Buttons.hTop + (iHeig - 7)
  Tamano(k).lStartX = Tamano(k - 1).lStartX + ScaleY(m_txtRect.hBottom, UserControl.Parent.ScaleMode, vbPixels)
 Next
 Call BorderObject
 'Columns = Cols
On Error GoTo 0
End Sub

Private Sub RuntimeControls()
On Error Resume Next
 Set txtEdit = UserControl.Controls.Add("VB.TextBox", "txtEdit")
 With txtEdit
  .Visible = False
  .BackColor = m_lBackColor
  .Text = ""
 End With
On Error GoTo 0
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ShiftColorOXP
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Shift a color.
'---------------------------------------------------------------------------------------
Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
 Dim cRed   As Long, cBlue  As Long
 Dim Delta  As Long, cGreen As Long

 cBlue = ((theColor \ &H10000) Mod &H100)
 cGreen = ((theColor \ &H100) Mod &H100)
 cRed = (theColor And &HFF)
 Delta = &HFF - Base
 cBlue = Base + cBlue * Delta \ &HFF
 cGreen = Base + cGreen * Delta \ &HFF
 cRed = Base + cRed * Delta \ &HFF
 If (cRed > 255) Then cRed = 255
 If (cGreen > 255) Then cGreen = 255
 If (cBlue > 255) Then cBlue = 255
 ShiftColorOXP = cRed + 256& * cGreen + 65536 * cBlue
End Function

'---------------------------------------------------------------------------------------
' Procedure : ScrollVisible
' DateTime  : 10/07/05 09:24
' Author    : HACKPRO TM
' Purpose   : Set visible the Scrollbar's.
'---------------------------------------------------------------------------------------
Private Sub ScrollVisible(ByVal bState As Boolean, ByVal WhatScroll As String)
 Dim m_hWnd As Long, lR As Long
 
 If (WhatScroll = "V") Then
  m_hWnd = m_hWNdV
  lVScroll = bState
 ElseIf (WhatScroll = "H") Then
  m_hWnd = m_hWNdH
  lHScroll = bState
 End If
 If (m_hWnd = 0) Then Exit Sub
 If (m_bNoFlatScrollBars = True) Then
  'If (m_hWnd) Then Call UninitializeFlatSB(m_hWnd)
  'lR = InitialiseFlatSB(m_hWnd)
  If (Err.Number <> 0) Then m_bNoFlatScrollBars = False
  lR = FlatSB_SetScrollProp(m_hWNdV, WSB_PROP_HSTYLE, FSB_ENCARTA_MODE, True)
 End If
 If (m_bNoFlatScrollBars = False) Then
  Call ShowScrollBar(m_hWnd, SB_CTL, Abs(bState))
 Else
  Call FlatSB_ShowScrollBar(m_hWnd, SB_CTL, Abs(bState))
 End If
End Sub

Private Function TotalHeight() As Long
 Dim i As Long, tTotal As Long
 
 If (TotalItems <= 0) Then TotalHeight = 0: Exit Function
 tTotal = 0
 If (SafeUBound(VarPtrArray(Headers)) >= 0) Then
  For i = 0 To TotalItems
   If (tTotal > UserControl.ScaleHeight) Then Exit For
   tTotal = tTotal + yRow(i)
  Next
 End If
 TotalHeight = i
End Function

Private Function TotalWidth() As Long
 Dim i As Long, tTotal As Long, uPos As Long
 
 tTotal = 0
 uPos = 0
 If (SafeUBound(VarPtrArray(Headers)) >= 0) Then
  For i = 0 To Cols
   If (Headers(i).Visible = True) Then
    If (ColsV = Cols + 1) Then
     tTotal = tTotal + Headers(uPos).Width
    Else
     tTotal = Headers(uPos).Width
    End If
    uPos = uPos + 1
   End If
  Next
 End If
 TotalWidth = tTotal
End Function

'---------------------------------------------------------------------------------------
' Procedure : Wait
' DateTime  : 03/07/05 09:02
' Author    : HACKPRO TM
' Purpose   : Espera un x tiempo.
'---------------------------------------------------------------------------------------
Private Sub Wait(ByVal Segundos As Single)
 Dim ComienzoSeg As Single, sumSeg As Long
 Dim FinSeg      As Single
 
 ComienzoSeg = Timer
 FinSeg = ComienzoSeg + Segundos
 sumSeg = 0
 Do While (FinSeg > Timer)
  DoEvents
  If (ComienzoSeg > Timer) Then FinSeg = FinSeg - 24 * 60 * 60
  If (sumSeg > 20) Then Exit Do
  sumSeg = sumSeg + 1
 Loop
End Sub

Private Sub txtEdit_Change()
On Error Resume Next
 If (FirstTime = True) Then
  FirstTime = False
  Exit Sub
 End If
 If (txtEdit.Tag = "^") Or (txtEdit.Tag = "~") Then
  Call ChangedItem(ColItem, RowItem, txtEdit.Tag & txtEdit.Text)
 Else
  Call ChangedItem(ColItem, RowItem, txtEdit.Text)
 End If
On Error GoTo 0
End Sub

Private Sub txtEdit_GotFocus()
 txtEdit.SelStart = 0
 txtEdit.SelLength = Len(txtEdit.Text)
End Sub

Private Sub UserControl_DblClick()
 Dim iStyle As String
 
On Error Resume Next
 If (LastButton = vbLeftButton) And (InFocusControl(UserControl.hWnd) = True) Then
  RowItem = GetColFromX(MouseX)  '* Col From X.
  ColItem = GetRowFromY(MouseY)  '* Row From Y.
  ClickHeader = False
  If (ColItem = 0) Then
   ClickHeader = True
   Call Refresh(True)
   Call DrawHeaders(Value2, RowItem, 3)
   Call BorderObject
   Call Wait(0.04)
   ColItem = -1
   RowItem = -1
   Exit Sub
  End If
  iStyle = Items(ColItem).Item.Style(RowItem)
  If (Items(ColItem).Item.Enabled(RowItem) = False) Then
   ColItem = -1
   RowItem = -1
   Call Refresh(True)
   Exit Sub
  End If
  If (iStyle = "B") Then Items(ColItem).Item.Values(RowItem) = True
  Call Refresh(True)
  If (iStyle = "T") Or (iStyle = "D") Or (iStyle = "C") Then '* Add the Item's.
   Call Wait(0.04)
   Items(ColItem).Item.Values(RowItem) = False
   Call Refresh(False, , True)
  End If
 Else
  RaiseEvent DblClick
 End If
On Error GoTo 0
End Sub

Private Sub UserControl_Initialize()
 Dim OS As OSVERSIONINFO

 '* Get the operating system version for text drawing purposes.
 OS.dwOSVersionInfoSize = Len(OS)
 Call GetVersionEx(OS)
 mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
 Call RuntimeControls
End Sub

Private Sub UserControl_InitProperties()
 b_AutoSizeColumn = True
 isEnabled = True
 m_bAlphaBlendSel = True
 m_bNoFlatScrollBars = False
 m_bViewStyle = False
 m_lBackColor = GetSysColor(COLOR_BTNFACE)
 m_lBorderColor = &HFF9534
 m_lFullSelect = False
 m_lHeadersColor = GetSysColor(COLOR_BTNFACE)
 m_lSelectBackColor = ConvertSystemColor(vbHighlight)
 mColumnHeadingH = 320
 mHeaderHot = True
 m_sTextHeaders = ""
 ReDim Items(1)
 Set m_lBackgroundPic = Nothing
 Set BackgroundPicture = Nothing
 TotalItems = 0
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim mNew  As Long, iPos   As Long, isR As Boolean
 Dim FindF As Long, iStyle As String

 '* Select each Key.
On Error Resume Next
 If (lVScroll = True) Then m_hWnd = m_hWNdV
 isR = False
 Select Case KeyCode
  Case vbKeyF2
   iStyle = Items(ColItem).Item.Style(RowItem)
   If (Items(ColItem).Item.Enabled(RowItem) = False) Then
    ColItem = -1
    RowItem = -1
    Call Refresh(True)
    Exit Sub
   End If
   If (iStyle = "B") Then Items(ColItem).Item.Values(RowItem) = True
   Call Refresh(True)
   If (iStyle = "T") Or (iStyle = "D") Or (iStyle = "C") Then '* Add the Item's.
    Call Wait(0.04)
    Items(ColItem).Item.Values(RowItem) = False
    Call Refresh(False, , True)
   End If
   Exit Sub
  Case vbKeyUp
   mNew = ColItem - 1
   If (lVScroll = True) And (mNew >= 1) Then Value = Value - 1
  Case vbKeyDown
   mNew = ColItem + 1
   If (lVScroll = True) And (ShowItems <= mNew) Then Value = Value + 1
  Case vbKeyEnd
   mNew = FindFirstEnabled(False)
   If (lVScroll = True) Then Value = mNew
  Case vbKeyHome
   mNew = FindFirstEnabled()
   If (lVScroll = True) Then Value = 0
  Case vbKeyPageDown
   mNew = (ColItem + ShowItems)
   If (lVScroll = True) Then Value = mNew - 1
  Case vbKeyPageUp
   mNew = (ColItem - ShowItems)
   If (lVScroll = True) Then Value = mNew - 1
  Case vbKeyLeft
   If (lHScroll = True) Then m_hWnd = m_hWNdH
   mNew = RowItem - 1
   If (lHScroll = True) And (mNew >= 0) Then
    Value = mNew
    Call Refresh
   End If
   isR = True
  Case vbKeyRight
   If (lHScroll = True) Then m_hWnd = m_hWNdH
   mNew = RowItem + 1
   If (lHScroll = True) And (ColsV >= mNew) Then
    Value = Value + 1
    Call Refresh
   End If
   isR = True
  Case Else
   Exit Sub
 End Select
 If (isR = False) Then
  If (mNew > TotalItems) Then mNew = TotalItems
  If (mNew < 0) Then mNew = 0
 Else
  If (mNew > ColsV) Then mNew = ColsV
  If (mNew < 0) Then mNew = 0
 End If
 '* Refrech Control.
 If Not (mNew = ColItem) And Not (mNew = -1) And (isR = False) Then
  ColItem = mNew
  If (Items(ColItem).Item.Enabled(RowItem) = False) Then
   If (KeyCode = vbKeyUp) Or (KeyCode = vbKeyPageDown) Then
    FindF = FindFirstEnabled()
    If (ColItem > FindF) Then
     For iPos = ColItem - 1 To 1 Step -1
      If (Items(iPos).Item.Enabled(RowItem) = True) Then Exit For
      mNew = ColItem - 1
     Next
    Else
     iPos = FindF
    End If
   ElseIf (KeyCode = vbKeyDown) Or (KeyCode = vbKeyPageUp) Then
    For iPos = ColItem + 1 To TotalItems
     If (Items(iPos).Item.Enabled(RowItem) = True) Then Exit For
     mNew = ColItem + 1
    Next
   End If
   If (TotalItems < iPos) Then
    ColItem = mNew - 1
   Else
    ColItem = iPos
   End If
  End If
 ElseIf Not (mNew = RowItem) And Not (mNew = -1) And (isR = True) Then
  RowItem = mNew
  If (Headers(RowItem).Visible = False) Or (Items(ColItem).Item.Enabled(RowItem) = False) Then
   If (KeyCode = vbKeyLeft) Then
    FindF = FindFirstEnabled(, False)
    If (RowItem > FindF) Then
     For iPos = RowItem - 1 To 0 Step -1
      If (Headers(iPos).Visible = True) And (Items(ColItem).Item.Enabled(iPos) = True) Then Exit For
      mNew = RowItem - 1
     Next
    Else
     iPos = FindF
    End If
   ElseIf (KeyCode = vbKeyRight) Then
    For iPos = RowItem + 1 To Cols
     If (Headers(iPos).Visible = True) And (Items(ColItem).Item.Enabled(iPos) = True) Then Exit For
     mNew = RowItem + 1
    Next
   End If
   RowItem = iPos
  End If
 End If
 txtEdit.Visible = False
 Call Refresh
On Error GoTo 0
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim iStyle As String, isCol  As Long
 
On Error Resume Next
 If (Button <> vbLeftButton) Then Exit Sub
 LastButton = Button
 txtEdit.Visible = False
 RowItem = GetColFromX(X)  '* Col From X.
 ColItem = GetRowFromY(Y)  '* Row From Y.
 ClickHeader = False
 If (ColItem = 0) Then
  Call Refresh(True)
  Call DrawHeaders(Value2, RowItem, 3)
  Call BorderObject
  Call Wait(0.04)
  ColItem = -1
  RowItem = -1
  ClickHeader = True
  Exit Sub
 End If
 If (m_lFullSelect = True) Then
  If (Items(ColItem).Item.Enabled(RowItem) = False) Then
   ColItem = -1
   RowItem = -1
  End If
  Call Refresh(True)
  If (RowItem <> -1) And (ColItem <> -1) Then RaiseEvent Click(RowItem, ColItem, iStyle)
  Exit Sub
 End If
 iStyle = Items(ColItem).Item.Style(RowItem)
 If (Items(ColItem).Item.Enabled(RowItem) = False) Then
  ColItem = -1
  RowItem = -1
  Call Refresh(True)
  Exit Sub
 End If
 If (iStyle = "B") Or (iStyle = "C") Or (iStyle = "D") Then
  Items(ColItem).Item.Values(RowItem) = True
 ElseIf (iStyle = "O") Then '* Option.
  For isCol = 1 To TotalItems
   If (RowItem < UBound(Items(isCol).Item.Style) + 1) Then
    If (Items(isCol).Item.Style(RowItem) = "O") Then Items(isCol).Item.Values(RowItem) = False
   End If
  Next
  Items(ColItem).Item.Values(RowItem) = True
 ElseIf (iStyle = "Ch") Then
  Items(ColItem).Item.Values(RowItem) = Not (Items(ColItem).Item.Values(RowItem))
 End If
 Call Refresh(True)
 '* Button.
 If (iStyle = "B") Then '* Show the form.
  Call Wait(0.04)
  Items(ColItem).Item.Values(RowItem) = False
  Call Refresh(False)
  Call FormShow(TheForm, TheText)
 ElseIf (iStyle = "C") Then '* Add the Item's.
  Call Wait(0.04)
  Items(ColItem).Item.Values(RowItem) = False
  Call Refresh(False)
 End If
 RaiseEvent Click(RowItem, ColItem, Items(ColItem).Item.Style(RowItem))
On Error GoTo 0
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim LRow As Long, LCol As Long
 
 If (ClickHeader = False) And (MouseX = X) And (MouseY = Y) Then Exit Sub
 LCol = GetRowFromY(Y)  '* Row From Y.
 If (LCol = 0) And (mHeaderHot = True) Then
  LastCol = LCol
  LastRow = LRow
  LRow = GetColFromX(X) '* Col From X.
  If (LRow <= Cols) Then Call DrawHeaders(Value2, LRow, 2)
 Else
  Call DrawHeaders(Value2, -1, 0)
  LastRow = -1
  LastCol = -1
 End If
 Call BorderObject
 MouseX = X
 MouseY = Y
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
 AutoSizeColumn = PropBag.ReadProperty("AutoSizeColumn", True)
 BackColor = PropBag.ReadProperty("BackColor", GetSysColor(COLOR_BTNFACE))
 Set BackgroundPicture = PropBag.ReadProperty("BackgroundPicture", Nothing)
 BorderColor = PropBag.ReadProperty("BorderColor", &HFF9534)
 BorderStyle = PropBag.ReadProperty("BorderStyle", &H0)
 ColumnHeadingHeight = PropBag.ReadProperty("ColumnHeadingHeight", 320)
 Enabled = PropBag.ReadProperty("Enabled", True)
 FlatScrollbars = PropBag.ReadProperty("FlatScrollbars", False)
 FullSelection = PropBag.ReadProperty("FullSelection", False)
 HeadersColor = PropBag.ReadProperty("HeadersColor", GetSysColor(COLOR_BTNFACE))
 HeaderHotTrack = PropBag.ReadProperty("HeaderHotTrack", True)
 ListViewStyle = PropBag.ReadProperty("ListViewStyle", False)
 SelectBackColor = PropBag.ReadProperty("SelectBackColor", ConvertSystemColor(vbHighlight))
 SelectionAlphaBlend = PropBag.ReadProperty("SelectionAlphaBlend", True)
 TextHeaders = PropBag.ReadProperty("TextHeaders", "")
 Set Font = PropBag.ReadProperty("Font", Ambient.Font)
 If (Ambient.UserMode = True) Then
  bTrack = True
  bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
  If Not (bTrackUser32 = True) Then
   If Not (IsFunctionExported("_TrackMouseEvent", "Comctl32") = True) Then
    bTrack = False
   End If
  End If
  If (bTrack = True) Then '* OS supports mouse leave so subclass for it.
   '* Start subclassing the UserControl.
   Call DrawScrollBar(1)
   Call DrawScrollBar(2)
   With UserControl
    Call sc_Subclass(.hWnd)
    Call sc_AddMsg(.hWnd, WM_MOUSEWHEEL)
    Call sc_AddMsg(.hWnd, WM_MOUSEMOVE)
    Call sc_AddMsg(.hWnd, WM_MOUSELEAVE)
    Call sc_AddMsg(.hWnd, WM_KILLFOCUS)
    Call sc_AddMsg(.hWnd, WM_CTLCOLORSCROLLBAR)
    Call sc_AddMsg(.hWnd, WM_VSCROLL)
    Call sc_AddMsg(.hWnd, WM_HSCROLL)
    If (isXp = True) Then Call sc_AddMsg(.hWnd, WM_THEMECHANGED)
   End With
   With UserControl.Parent
    Call sc_Subclass(.hWnd)
    Call sc_AddMsg(.hWnd, WM_WINDOWPOSCHANGING)
    Call sc_AddMsg(.hWnd, WM_WINDOWPOSCHANGED)
    Call sc_AddMsg(.hWnd, WM_GETMINMAXINFO)
    Call sc_AddMsg(.hWnd, WM_LBUTTONDOWN)
    Call sc_AddMsg(.hWnd, WM_SIZE)
   End With
  End If
 End If
End Sub

Private Sub UserControl_Resize()
 If (Ambient.UserMode = True) Then Call Refresh
End Sub

Private Sub UserControl_Terminate()
On Error GoTo Catch
 Call sc_Terminate '* Stop all subclassing.
 TotalItems = 0
 ReDim Items(1)
 Call DestroyWindow(m_hWNdH)
 Call DestroyWindow(m_hWNdV)
 Erase xRow
 Erase yRow
 Erase Items
 If (m_bNoFlatScrollBars = True) Then
  m_hWnd = m_hWNdV
  If (m_hWnd) Then Call UninitializeFlatSB(m_hWnd)
  m_hWnd = m_hWNdH
  If (m_hWnd) Then Call UninitializeFlatSB(m_hWnd)
 End If
 Exit Sub
Catch:
On Error GoTo 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 With PropBag
  Call .WriteProperty("AutoSizeColumn", b_AutoSizeColumn, True)
  Call .WriteProperty("BackColor", m_lBackColor, GetSysColor(COLOR_BTNFACE))
  Call .WriteProperty("BackgroundPicture", m_lBackgroundPic, Nothing)
  Call .WriteProperty("BorderColor", m_lBorderColor, &HFF9534)
  Call .WriteProperty("BorderStyle", m_eBorderStyle, &H0)
  Call .WriteProperty("ColumnHeadingHeight", mColumnHeadingH, 320)
  Call .WriteProperty("Enabled", isEnabled, True)
  Call .WriteProperty("FlatScrollbars", m_bNoFlatScrollBars, False)
  Call .WriteProperty("Font", mFont, Ambient.Font)
  Call .WriteProperty("FullSelection", m_lFullSelect, False)
  Call .WriteProperty("HeadersColor", m_lHeadersColor, GetSysColor(COLOR_BTNFACE))
  Call .WriteProperty("HeaderHotTrack", mHeaderHot, True)
  Call .WriteProperty("ListViewStyle", m_bViewStyle, False)
  Call .WriteProperty("SelectBackColor", m_lSelectBackColor, ConvertSystemColor(vbHighlight))
  Call .WriteProperty("SelectionAlphaBlend", m_bAlphaBlendSel, True)
  Call .WriteProperty("TextHeaders", m_sTextHeaders, "")
 End With
End Sub

'*******************************************************
'* This code is extract of this great control.
'* Name:     vbalScrollButton
'* Author:   Steve McMahon (steve@dogma.demon.co.uk)
'* Date:     28 December 1998
'*******************************************************
Private Property Get SmallChange() As Long
 SmallChange = m_lSmallChange
End Property

Private Property Let SmallChange(ByVal lSmallChange As Long)
 m_lSmallChange = lSmallChange
End Property

Private Sub pGetSI(ByRef tSI As SCROLLINFO, ByVal fMask As Long)
 tSI.fMask = fMask
 tSI.cbSize = LenB(tSI)
 If (m_bNoFlatScrollBars = True) Then
  Call GetScrollInfo(m_hWnd, SB_CTL, tSI)
 Else
  Call FlatSB_GetScrollInfo(m_hWnd, SB_CTL, tSI)
 End If
End Sub

Private Sub pLetSI(ByRef tSI As SCROLLINFO, ByVal fMask As Long)
 tSI.fMask = fMask
 tSI.cbSize = LenB(tSI)
 If (m_bNoFlatScrollBars = True) Then
  Call SetScrollInfo(m_hWnd, SB_CTL, tSI, True)
 Else
  Call FlatSB_SetScrollInfo(m_hWnd, SB_CTL, tSI, True)
 End If
End Sub

Private Property Get Min() As Long
 Dim tSI As SCROLLINFO
 
 Call pGetSI(tSI, SIF_RANGE)
 Min = tSI.nMin
End Property

Private Property Get Max() As Long
 Dim tSI As SCROLLINFO
 
 Call pGetSI(tSI, SIF_RANGE Or SIF_PAGE)
 Max = tSI.nMax - tSI.nPage
End Property

Private Property Get Value() As Long
 Dim tSI As SCROLLINFO
 
 Call pGetSI(tSI, SIF_POS)
 Value = tSI.nPos
End Property

Private Property Get LargeChange() As Long
 Dim tSI As SCROLLINFO
 
 Call pGetSI(tSI, SIF_PAGE)
 LargeChange = tSI.nPage
End Property

Private Property Let Min(ByVal iMin As Long)
 Dim tSI As SCROLLINFO
 
 tSI.nMin = iMin
 tSI.nMax = Max + LargeChange
 Call pLetSI(tSI, SIF_RANGE)
End Property

Private Property Let Max(ByVal iMax As Long)
 Dim tSI As SCROLLINFO
 
 tSI.nMax = iMax + LargeChange
 tSI.nMin = Min
 Call pLetSI(tSI, SIF_RANGE)
 Call pRaiseEvent(False)
End Property

Private Property Let Value(ByVal iValue As Long)
 Dim tSI As SCROLLINFO, lPercent As Long
 
 If (iValue <> Value) Then
  tSI.nPos = iValue
  Call pLetSI(tSI, SIF_POS)
  If (Max > 0) Then lPercent = iValue * 100 \ Max
  Call pRaiseEvent(False)
 End If
End Property

Private Property Let LargeChange(ByVal iLargeChange As Long)
 Dim lCurLargeChange As Long, tSI As SCROLLINFO
 Dim lCurMax         As Long
    
 Call pGetSI(tSI, SIF_ALL)
 tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
 tSI.nPage = iLargeChange
 Call pLetSI(tSI, SIF_PAGE Or SIF_RANGE)
End Property

Private Function pRaiseEvent(ByVal bScroll As Boolean)
 Static s_lLastValue As Long
 
 If (Value <> s_lLastValue) Then
  If (bScroll = True) Then
   RaiseEvent Scroll
  Else
   RaiseEvent Change
  End If
  s_lLastValue = Value
 End If
End Function

' See post: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=58622&lngWId=1
' Thanks MArio Florez.
Private Function RenderIconGrayscale(ByVal Dest_hDC As Long, ByVal hIcon As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Dest_Height As Long, Optional ByVal Dest_Width As Long, Optional ByVal GrayC As Boolean = True) As Boolean
 Dim hBMP_Mask As Long, hBMP_Image As Long
 Dim hBMP_Prev As Long, hIcon_Temp As Long
 Dim hDC_Temp  As Long

 ' Make sure parameters passed are valid
 If (Dest_hDC = 0) Or (hIcon = 0) Then Exit Function
 ' Extract the bitmaps from the icon
 If (GetIconBitmaps(hIcon, hBMP_Mask, hBMP_Image) = False) Then Exit Function
 ' Create a memory DC to work with
 hDC_Temp = CreateCompatibleDC(0)
 If (hDC_Temp = 0) Then GoTo CleanUp
 ' Make the image bitmap gradient
 If (RenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0, , , GrayC) = False) Then GoTo CleanUp
 ' Extract the gradient bitmap out of the DC
 Call SelectObject(hDC_Temp, hBMP_Prev)
 ' Take the newly gradient bitmap and make a gradient icon from it
 hIcon_Temp = CreateIconFromBMP(hBMP_Mask, hBMP_Image)
 If (hIcon_Temp = 0) Then GoTo CleanUp
 ' Draw the newly created gradient icon onto the specified DC
 If (DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon_Temp, Dest_Width, Dest_Height, 0, 0, &H3) <> 0) Then
  RenderIconGrayscale = True
 End If
CleanUp:
 Call DestroyIcon(hIcon_Temp): hIcon_Temp = 0
 Call DeleteDC(hDC_Temp): hDC_Temp = 0
 Call DeleteObject(hBMP_Mask): hBMP_Mask = 0
 Call DeleteObject(hBMP_Image): hBMP_Image = 0
End Function

Private Function GetIconBitmaps(ByVal hIcon As Long, ByRef Return_hBmpMask As Long, ByRef Return_hBmpImage As Long) As Boolean
 Dim TempICONINFO As ICONINFO

 If (GetIconInfo(hIcon, TempICONINFO) = 0) Then Exit Function
 Return_hBmpMask = TempICONINFO.hbmMask
 Return_hBmpImage = TempICONINFO.hbmColor
 GetIconBitmaps = True
End Function

'=============================================================================================================
Private Function RenderBitmapGrayscale(ByVal Dest_hDC As Long, ByVal hBitmap As Long, Optional ByVal Dest_X As Long, Optional ByVal Dest_Y As Long, Optional ByVal Srce_X As Long, Optional ByVal Srce_Y As Long, Optional ByVal GrayC As Boolean = True) As Boolean
 Dim TempBITMAP As BITMAP, hScreen   As Long
 Dim hDC_Temp   As Long, hBMP_Prev   As Long
 Dim MyCounterX As Long, MyCounterY  As Long
 Dim NewColor   As Long, hNewPicture As Long
 Dim DeletePic  As Boolean

 ' Make sure parameters passed are valid
 If (Dest_hDC = 0) Or (hBitmap = 0) Then Exit Function
 ' Get the handle to the screen DC
 hScreen = GetDC(0)
 If (hScreen = 0) Then Exit Function
 ' Create a memory DC to work with the picture
 hDC_Temp = CreateCompatibleDC(hScreen)
 If (hDC_Temp = 0) Then GoTo CleanUp
 ' If the user specifies NOT to alter the original, then make a copy of it to use
 DeletePic = False
 hNewPicture = hBitmap
 ' Select the bitmap into the DC
 hBMP_Prev = SelectObject(hDC_Temp, hNewPicture)
 ' Get the height / width of the bitmap in pixels
 If (GetObjectAPI(hNewPicture, Len(TempBITMAP), TempBITMAP) = 0) Then GoTo CleanUp
 If (TempBITMAP.bmHeight <= 0) Or (TempBITMAP.bmWidth <= 0) Then GoTo CleanUp
 ' Loop through each pixel and conver it to it's grayscale equivelant
 If (GrayC = True) Then
  For MyCounterX = 0 To TempBITMAP.bmWidth - 1
   For MyCounterY = 0 To TempBITMAP.bmHeight - 1
    NewColor = GetPixel(hDC_Temp, MyCounterX, MyCounterY)
    If (NewColor <> -1) Then
     Select Case NewColor
      ' If the color is already a grey shade, no need to convert it
      Case vbBlack, vbWhite, &H101010, &H202020, &H303030, &H404040, &H505050, &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, &HE0E0E0, &HF0F0F0
       NewColor = NewColor
      Case Else
       NewColor = 0.33 * (NewColor Mod 256) + 0.59 * ((NewColor \ 256) Mod 256) + 0.11 * ((NewColor \ 65536) Mod 256)
       NewColor = RGB(NewColor, NewColor, NewColor)
     End Select
     Call SetPixel(hDC_Temp, MyCounterX, MyCounterY, NewColor)
    End If
   Next
  Next
 End If
 ' Display the picture on the specified hDC
 Call BitBlt(Dest_hDC, Dest_X, Dest_Y, TempBITMAP.bmWidth, TempBITMAP.bmHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy)
 RenderBitmapGrayscale = True
CleanUp:
 Call ReleaseDC(0, hScreen): hScreen = 0
 Call SelectObject(hDC_Temp, hBMP_Prev)
 Call DeleteDC(hDC_Temp): hDC_Temp = 0
 If (DeletePic = True) Then
  Call DeleteObject(hNewPicture)
  hNewPicture = 0
 End If
End Function

Private Function CreateIconFromBMP(ByVal hBMP_Mask As Long, ByVal hBMP_Image As Long) As Long
 Dim TempICONINFO As ICONINFO

 If (hBMP_Mask = 0) Or (hBMP_Image = 0) Then Exit Function
 TempICONINFO.fIcon = 1
 TempICONINFO.hbmMask = hBMP_Mask
 TempICONINFO.hbmColor = hBMP_Image
 CreateIconFromBMP = CreateIconIndirect(TempICONINFO)
End Function

'* By: Kristian S.Stangeland
Private Function SafeUBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long
 Dim lAddress&, cElements&, lLBound&, cDims%

 If (Dimension < 1) Then
  SafeUBound = -1
  Exit Function
 End If
 Call CopyMemory(lAddress, ByVal lpArray, 4)
 If (lAddress = 0) Then
  '* The array isn't initilized.
  SafeUBound = -1
  Exit Function
 End If
 '* Calculate the dimensions.
 Call CopyMemory(cDims, ByVal lAddress, 2)
 Dimension = cDims - Dimension + 1
 '* Obtain the needed data.
 Call CopyMemory(cElements, ByVal (lAddress + 16 + ((Dimension - 1) * 8)), 4)
 Call CopyMemory(lLBound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4)
 SafeUBound = cElements + lLBound - 1
End Function

Private Function SafeLBound(ByVal lpArray As Long, Optional Dimension As Long = 1) As Long
 Dim lAddress&, cElements&, lLBound&, cDims%

 If (Dimension < 1) Then
  SafeLBound = -1
  Exit Function
 End If
 Call CopyMemory(lAddress, ByVal lpArray, 4)
 If (lAddress = 0) Then
  '* The array isn't initilized.
  SafeLBound = -1
  Exit Function
 End If
 '* Calculate the dimensions.
 Call CopyMemory(cDims, ByVal lAddress, 2)
 Dimension = cDims - Dimension + 1
 '* Obtain the needed data.
 Call CopyMemory(lLBound, ByVal (lAddress + 20 + ((Dimension - 1) * 8)), 4)
 SafeLBound = lLBound
End Function

'================================================
' Name:          (API's AlphaBlend emulation)
' Class:         cTile.cls
' Author:        Carles P.V.
' Dependencies:
' Last revision: 2003.03.28
'================================================

'========================================================================================
' Methods
'========================================================================================
Private Function CreatePatternFromStdPicture(ByVal Image As StdPicture) As Boolean
 Dim uBI       As BITMAP
 Dim uBIH      As BITMAPINFOHEADER
 Dim aBuffer() As Byte ' Packed DIB
 Dim lhDC      As Long
 Dim lhOldBmp  As Long
    
 If (GetObjectType(Image.Handle) = OBJ_BITMAP) Then
  '-- Get image info
  Call GetObject(Image.Handle, Len(uBI), uBI)
  '-- Prepare DIB header and redim. buffer array
  With uBIH
   .biSize = Len(uBIH)
   .biPlanes = 1
   .biBitCount = 24
   .biWidth = uBI.bmWidth
   .biHeight = uBI.bmHeight
   .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
  End With
  ReDim aBuffer(1 To Len(uBIH) + uBIH.biSizeImage)
  '-- Create DIB brush
  lhDC = CreateCompatibleDC(0)
  If (lhDC <> 0) Then
   lhOldBmp = SelectObject(lhDC, Image.Handle)
   '-- Build packed DIB:
   '-  Merge Header
   Call CopyMemory(aBuffer(1), uBIH, Len(uBIH))
   '-  Get and merge DIB bits
   Call GetDIBits(lhDC, Image.Handle, 0, uBI.bmHeight, aBuffer(Len(uBIH) + 1), uBIH, DIB_RGB_COLORS)
   Call SelectObject(lhDC, lhOldBmp)
   Call DeleteDC(lhDC)
   '-  Create brush from packed DIB
   Call DestroyPattern
   m_hBrush = CreateDIBPatternBrushPt(aBuffer(1), DIB_RGB_COLORS)
  End If
 End If
 '-- Success
 CreatePatternFromStdPicture = (m_hBrush <> 0)
End Function

Private Function CreatePatternFromHatchBrush(ByVal BrushStyle As HatchBrushStyleCts, ByVal Color As OLE_COLOR) As Boolean
 '-- Create brush from system brush
 Call DestroyPattern
 Call OleTranslateColor(Color, 0, Color)
 m_hBrush = CreateHatchBrush(BrushStyle, Color)
 '-- Success
 CreatePatternFromHatchBrush = (m_hBrush <> 0)
End Function

Private Function CreatePatternFromSolidColor(ByVal Color As OLE_COLOR) As Boolean
 '-- Create brush from solid color
 Call DestroyPattern
 Call OleTranslateColor(Color, 0, Color)
 m_hBrush = CreateSolidBrush(Color)
 '-- Success
 CreatePatternFromSolidColor = (m_hBrush <> 0)
End Function

Private Sub Tile(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal ResetBrushOrigin As Boolean = True)
 Dim rTile As RECT, ptOrg As POINTAPI
  
 If (m_hBrush <> 0) Then
  '-- Set brush origin
  If (ResetBrushOrigin = True) Then
   Call SetBrushOrgEx(hDC, X, Y, ptOrg)
  Else
   Call SetBrushOrgEx(hDC, 0, 0, ptOrg)
  End If
  '-- Tile image
  Call SetRect(rTile, X, Y, X + Width, Y + Height)
  Call FillRect(hDC, rTile, m_hBrush)
 End If
End Sub

Private Sub DestroyPattern()
 If (m_hBrush <> 0) Then
  Call DeleteObject(m_hBrush)
  m_hBrush = 0
 End If
End Sub

'* ======================================================================================================
'*  UserControl private routines.
'*  Determine if the passed function is supported.
'* ======================================================================================================
'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    FreeLibrary hMod
  End If
End Function

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      TrackMouseEvent tme
    Else
      TrackMouseEventComCtl tme
    End If
  End If
End Sub

'-uSelfSub code-----------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long) As Boolean             'Subclass the specified window handle
  Dim nAddr As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError "sc_Subclass", "Invalid window handle"
  End If
  
  If z_hWnds Is Nothing Then
    RtlMoveMemory VarPtr(z_DataDataPtr), VarPtrArray(z_Data), 4             'Get the address of z_Data()'s SafeArray header
    z_DataDataPtr = z_DataDataPtr + 12                                      'Bump the address to point to the pvData data pointer
    RtlMoveMemory VarPtr(z_DataOrigData), z_DataDataPtr, 4                  'Get the value of z_Data()'s SafeArray pvData data pointer
  
    nAddr = zGetCallback                                                    'Get the address of this UserControl's zWndProc callback routine
    
    'Initialise the machine-code thunk
    z_Code(6) = -490736517001394.5807@: z_Code(7) = 484417356483292.94@: z_Code(8) = -171798741966746.6996@: z_Code(9) = 843649688964536.7412@: z_Code(10) = -330085705188364.0817@: z_Code(11) = 41621208.9739@: z_Code(12) = -900372920033759.9903@: z_Code(13) = 291516653989344.1016@: z_Code(14) = -621553923181.6984@: z_Code(15) = 291551690021556.6453@: z_Code(16) = 28798458374890.8543@: z_Code(17) = 86444073845629.4399@: z_Code(18) = 636540268579660.4789@: z_Code(19) = 60911183420250.2143@: z_Code(20) = 846934495644380.8767@: z_Code(21) = 14073829823.4668@: z_Code(22) = 501055845239149.5051@: z_Code(23) = 175724720056981.1236@: z_Code(24) = 75457451135513.7931@: z_Code(25) = -576850389355798.3357@: z_Code(26) = 146298060653075.5445@: z_Code(27) = 850256350680294.7583@: z_Code(28) = -4888724176660.092@: z_Code(29) = 21456079546.6867@
    
    zMap VarPtr(z_Code(0))                                                  'Map the address of z_Code()'s first element to the z_Data() array
    z_Data(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    z_Data(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                  'Store CallWindowProc function address in the thunk data
    z_Data(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                   'Store the SetWindowLong function address in the thunk data
    z_Data(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                   'Store the VirtualFree function address in the thunk data
    z_Data(IDX_ME) = ObjPtr(Me)                                             'Store my object address in the thunk data
    z_Data(IDX_CALLBACK) = nAddr                                            'Store the zWndProc address in the thunk data
    zMap z_DataOrigData                                                     'Restore z_Data()'s original data pointer
    
    Set z_hWnds = New Collection                                            'Create the window-handle/thunk-memory-address collection
  End If

  nAddr = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                    'Allocate executable memory
  RtlMoveMemory nAddr, VarPtr(z_Code(0)), CODE_LEN                          'Copy the machine-code to the allocated memory

  On Error GoTo Catch                                                       'Catch double subclassing
    z_hWnds.Add nAddr, "h" & lng_hWnd                                       'Add the hWnd/thunk-address to the collection
  On Error GoTo 0

  zMap nAddr                                                                'Map z_Data() to the subclass thunk machine-code
  z_Data(IDX_EBX) = nAddr                                                   'Patch the data address
  z_Data(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
  z_Data(IDX_BTABLE) = nAddr + CODE_LEN                                     'Store the address of the before table in the thunk data
  z_Data(IDX_ATABLE) = z_Data(IDX_BTABLE) + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
  nAddr = nAddr + WNDPROC_OFF                                               'Execution address of the thunk's WndProc
  z_Data(IDX_WNDPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, nAddr)        'Set the new WndProc and store the original WndProc in the thunk data
  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
  sc_Subclass = True                                                        'Indicate success
  Exit Function                                                             'Exit

Catch:
  zError "sc_Subclass", "Window handle is already subclassed"
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim i     As Long
  Dim nAddr As Long

  If z_hWnds Is Nothing Then                                                'Ensure that subclassing has been started
  Else
    With z_hWnds
      For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        nAddr = .Item(i)                                                    'Map z_Data() to the hWnd thunk address
        If IsBadCodePtr(nAddr) = 0 Then                                     'Ensure that the thunk hasn't already freed itself
          zMap nAddr                                                        'Map the thunk memory to the z_Data() array
          sc_UnSubclass z_Data(IDX_HWND)                                    'UnSubclass
        End If
      Next i                                                                'Next member of the collection
    End With
    
    Set z_hWnds = Nothing                                                   'Destroy the window-handle/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Public Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_hWnds Is Nothing Then                                                'Ensure that subclassing has been started
    zError "UnSubclass", "Subclassing hasn't been started", False
  Else
    zDelMsg lng_hWnd, ALL_MESSAGES, IDX_BTABLE                              'Delete all before messages
    zDelMsg lng_hWnd, ALL_MESSAGES, IDX_ATABLE                              'Delete all after messages
    zMap_hWnd lng_hWnd                                                      'Map the thunk memory to the z_Data() array
    z_Data(IDX_SHUTDOWN) = -1                                               'Set the shutdown indicator
    zMap z_DataOrigData                                                     'Restore z_Data()'s original data pointer
    z_hWnds.Remove "h" & lng_hWnd                                           'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If When And MSG_BEFORE Then                                               'If the message is to be added to the before original WndProc table...
    zAddMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Add the message to the before table
  End If

  If When And MSG_AFTER Then                                                'If message is to be added to the after original WndProc table...
    zAddMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Add the message to the after table
  End If

  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If When And MSG_BEFORE Then                                               'If the message is to be deleted from the before original WndProc table...
    zDelMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Delete the message from the before table
  End If

  If When And MSG_AFTER Then                                                'If the message is to be deleted from the after original WndProc table...
    zDelMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Delete the message from the after table
  End If

  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  sc_CallOrigWndProc = CallWindowProcA(z_Data(IDX_WNDPROC), lng_hWnd, uMsg, _
                                                            wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Function

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim i      As Long                                                        'Loop index

  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  zMap z_Data(nTable)                                                       'Map z_Data() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = z_Data(0)                                                      'Get the current table entry count

    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values", False
      Exit Sub
    End If

    For i = 1 To nCount                                                     'Loop through the table entries
      If z_Data(i) = 0 Then                                                 'If the element is free...
        z_Data(i) = uMsg                                                    'Use this element
        Exit Sub                                                            'Bail
      ElseIf z_Data(i) = uMsg Then                                          'If the message is already in the table...
        Exit Sub                                                            'Bail
      End If
    Next i                                                                  'Next message table entry

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
    z_Data(nCount) = uMsg                                                   'Store the message in the appended table entry
  End If

  z_Data(0) = nCount                                                        'Store the new table entry count
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim i      As Long                                                        'Loop index

  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  zMap z_Data(nTable)                                                       'Map z_Data() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    z_Data(0) = 0                                                           'Zero the table entry count
  Else
    nCount = z_Data(0)                                                      'Get the table entry count
    
    For i = 1 To nCount                                                     'Loop through the table entries
      If z_Data(i) = uMsg Then                                              'If the message is found...
        z_Data(i) = 0                                                       'Null the msg value -- also frees the element for re-use
        Exit Sub                                                            'Exit
      End If
    Next i                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table", False
  End If
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String, Optional ByVal bEnd As Boolean = True)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  
  MsgBox sMsg & ".", IIf(bEnd, vbCritical, vbExclamation) + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
  
  If bEnd Then
    End
  End If
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map z_Data() to the specified address
Private Sub zMap(ByVal nAddr As Long)
  RtlMoveMemory z_DataDataPtr, VarPtr(nAddr), 4                             'Set z_Data()'s SafeArray data pointer to the specified address
End Sub

'Map z_Data() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_hWnds Is Nothing Then                                                'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started", True
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    zMap_hWnd = z_hWnds("h" & lng_hWnd)                                     'Get the thunk address
    zMap zMap_hWnd                                                          'Map z_Data() to the thunk address
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Determine the address of the final private method, zWndProc
Private Function zGetCallback() As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte                                                         'Value pointed at by the vTable entry
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Upper bound of z_Data()
  Dim k     As Long                                                         'vTable entry value
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(Me), 4                                'Get the address of my vTable
  zMap nAddr + &H7A4                                                        'Map z_Data() to the first possible vTable entry for a UserControl

  j = UBound(z_Data())                                                      'Get the upper bound of z_Data()
  
  For i = 0 To j                                                            'Loop through the vTable looking for the first method entry
    k = z_Data(i)                                                           'Get the vTable entry
    
    If k <> 0 Then                                                          'Skip implemented interface entries
      RtlMoveMemory VarPtr(bVal), k, 1                                      'Get the first byte pointed to by this vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'If a method (pcode or native)
        bSub = bVal                                                         'Store which of the method markers was found (pcode or native)
        Exit For                                                            'Method found, quit loop and scan methods
      End If
    End If
  Next i
  
  For i = i To j                                                            'Loop through the remaining vTable entries
    k = z_Data(i)                                                           'Get the vTable entry
    
    If IsBadCodePtr(k) Then                                                 'Is the vTable entry an invalid code address...
      Exit For                                                              'Bad code pointer, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), k, 1                                        'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      Exit For                                                              'Bad method signature, quit loop
    End If
  Next i
  
  If i > j Then                                                             'Loop completed without finding the last method
    zError "zGetCallback", "z_Data() overflow. Increase the number of elements in the z_Data() array"
  End If
  
  'Uncomment the following line to determine the minimum number of elements needed by the z_Data() array
  'Debug.Print "Optimal dimension: z_Data(" & IIf(i > IIf(MSG_ENTRIES > IDX_EBX, MSG_ENTRIES, IDX_EBX), i, IIf(MSG_ENTRIES > IDX_EBX, MSG_ENTRIES, IDX_EBX)) & ")"
 
  zGetCallback = z_Data(i - 1)                                              'Return the last good vTable entry address
End Function

'-Subclass callback: must be private and the last method in the source file-----------------------
Private Sub zWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
 '*************************************************************************************************
 '* bBefore  - Indicates whether the callback is before or after the original WndProc. Usually you
 '*            will know unless the callback for the uMsg value is specified as MSG_BEFORE_AFTER
 '*            (both before and after the original WndProc).
 '* bHandled - In a before original WndProc callback, setting bHandled to True will prevent the
 '*            message being passed to the original WndProc and (if set to do so) the after
 '*            original WndProc callback.
 '* lReturn  - WndProc return value. Set as per the MSDN documentation for the message value,
 '*            and/or, in an after the original WndProc callback, act on the return value as set
 '*            by the original WndProc.
 '* hWnd     - Window handle.
 '* uMsg     - Message value.
 '* wParam   - Message related data.
 '* lParam   - Message related data.
 '*************************************************************************************************
 If (isEnabled = False) Then Exit Sub
 Dim lSC As Long, lScrollcode As Long
 Dim tSI As SCROLLINFO, lV    As Long
 
 Select Case uMsg
  Case WM_THEMECHANGED
   Call Refresh
   RaiseEvent ThemeChanged
  Case WM_MOUSEMOVE
   If Not (bInCtrl = True) Then
    bInCtrl = True
    Call TrackMouseLeave(lng_hWnd)
    RaiseEvent MouseEnter
   End If
  Case WM_KILLFOCUS
   Call Refresh
   RaiseEvent MouseEnter
  Case WM_MOUSELEAVE
   bInCtrl = False
   Call Refresh
   RaiseEvent MouseLeave
  'Case WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED, WM_GETMINMAXINFO, WM_SIZE, WM_LBUTTONDOWN, WM_RBUTTONDOWN
  ' Call Refresh
  Case WM_MOUSEWHEEL
   If (wParam = &H780000) Then
    Value = Value - 1
   ElseIf (wParam = &HFF880000) Then
    Value = Value + 1
   End If
   Call Refresh
  Case WM_CTLCOLORSCROLLBAR
   bHandled = True
  Case WM_VSCROLL, WM_HSCROLL '* Steven
   lScrollcode = (wParam And &HFFFF&)
   If (uMsg = WM_VSCROLL) Then
    m_hWnd = m_hWNdV
   ElseIf (uMsg = WM_HSCROLL) Then
    m_hWnd = m_hWNdH
   End If
   Select Case lScrollcode
    Case SB_THUMBTRACK
     '* Is vertical/horizontal?
     Call pGetSI(tSI, SIF_TRACKPOS)
     Value = tSI.nTrackPos
     Call pRaiseEvent(True)
    Case SB_LEFT, SB_BOTTOM
     Value = Min
     Call pRaiseEvent(False)
    Case SB_RIGHT, SB_TOP
     Value = Max
     Call pRaiseEvent(False)
    Case SB_LINELEFT, SB_LINEUP
     lV = Value
     lSC = m_lSmallChange
     If (lV - lSC <= Min) Then
      Value = Min
     Else
      Value = lV - lSC
     End If
     Call pRaiseEvent(False)
    Case SB_LINERIGHT, SB_LINEDOWN
     lV = Value
     lSC = m_lSmallChange
     If (lV + lSC >= Max) Then
      Value = Max + 1
     Else
      Value = lV + lSC
     End If
     Call pRaiseEvent(False)
    Case SB_PAGELEFT, SB_PAGEUP
     Value = Value - LargeChange
     Call pRaiseEvent(False)
    Case SB_PAGERIGHT, SB_PAGEDOWN
     Value = Value + LargeChange
     Call pRaiseEvent(False)
    Case SB_ENDSCROLL
     Call pRaiseEvent(False)
    If (uMsg = WM_VSCROLL) Then
     Value1 = Value
    ElseIf (uMsg = WM_HSCROLL) Then
     Value2 = Value
    End If
    Call Refresh
   End Select
 End Select
End Sub
