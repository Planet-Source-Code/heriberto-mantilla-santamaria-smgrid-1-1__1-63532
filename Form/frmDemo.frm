VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo SMGrid 1.0b by HACKPRO TM © 2005"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Height          =   420
      Left            =   5055
      TabIndex        =   2
      Top             =   4965
      Width           =   1080
   End
   Begin GridControl.SMGrid SMGrid 
      Height          =   3690
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6509
      AutoSizeColumn  =   0   'False
      BackColor       =   16777215
      BorderStyle     =   1
      ColumnHeadingHeight=   370
      FlatScrollbars  =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextHeaders     =   "^Demo Version|~by HACKPRO TM"
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDemo.frx":058A
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   105
      TabIndex        =   1
      Top             =   3825
      Width           =   6150
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2005       *'
'******************************************************'
'* Comments: This control was necessary to develop it *'
'*           for a program of a thesis of grade of my *'
'*           University, its evolution was stopped by *'
'*           a lot of time, although it is not comple-*'
'*           tely ended, but it's a beginning.        *'
'******************************************************'
'* Now my website is available but alone the version  *'
'* in Spanish.                                        *'
'-----------------------------------------------------*'
'* WebSite:  http://www.geocities.com/hackprotm/      *'
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2005       *'
'******************************************************'
Option Explicit

 Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (Iccex As tagInitCommonControlsEx) As Boolean
 Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
 Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
 Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
  
 Private Type tagInitCommonControlsEx
  lngSize As Long
  lngICC  As Long
 End Type
 
 Private Const ICC_USEREX_CLASSES = &H200
 
 Private m_hMod As Long

 Private i As Integer, j As Integer

Private Sub cmdResume_Click()
 Call frmProperties.Show
End Sub

Private Sub Form_Initialize()
 Dim Iccex As tagInitCommonControlsEx

 Iccex.lngSize = LenB(Iccex)
 Iccex.lngICC = ICC_USEREX_CLASSES
 Call InitCommonControlsEx(Iccex)
 m_hMod = LoadLibrary("shell32.dll")
End Sub

Private Sub Form_Load()
 With SMGrid
  Me.Caption = "Demo " & .GetControlVersion
  .TextHeaders = "^Column Header1|^Column Header2|Column Header3|Column Header4|~Column Header5"
  Call .AddItem("Col1 Row1|Col2 Row1", "&HEED5C4|&HE0E0E0|&HDEEDEF", , "T|B")
  Call .AddItem("Col1 Row2|Col2 Row2", , , "C|B", , , "T|F")
  Call .AddItem("~Right Align|Col2 Row3", , , "C|C", , "", "F")
  Call .AddItem("Col1 Row4|Col2 Row4", , , "C|B|D")
  Call .AddItem("^Center Align|Col2 Row5", , , "T|T")
  Call .ColumnWidth(0, 140)
  Call .ColumnWidth(1, 140)
  Call .ColumnWidth(2, 110)
  Call .ColumnWidth(3, 80)
  Call .ColumnWidth(4, 90)
  Call .ColumnWidth(5, 170)
  For i = 6 To 8
   Call .AddItem("Col1 Row" & i & "|Col2 Row" & i & "|Col3 Row" & i, , , "T|Ch|O|O", , "|T", "||F")
  Next
  For i = 9 To 18
   Call .AddItem("||||^Col5 Row" & i, , "||||&HC56A31", "C")
  Next
  Call .ObjectForm(frmPopUp, frmPopUp)
  'Call .HideCol(0, False)
  'Call .HideCol(1, False)
  'Call .HideCol(2, False)
  'Call .HideCol(3, False)
  .AutoSizeColumn = True
  .HeaderHotTrack = True
  'Set .BackgroundPicture = LoadPicture(App.Path & "\NotePad.bmp", vbResBitmap)
  '.ReDraw
 End With
End Sub

Private Sub Form_Terminate()
On Error Resume Next
 Call FreeLibrary(m_hMod)
On Error GoTo 0
End Sub

Private Sub SMGrid_Click(ByVal Row As Long, ByVal Col As Long, ByVal Style As String)
 lblTitle.Caption = "Row: " & Row & " - Col: " & Col & " (" & SMGrid.ItemText(Col, Row) & ")"
End Sub
