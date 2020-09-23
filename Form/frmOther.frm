VERSION 5.00
Begin VB.Form frmOther 
   Caption         =   "Other SMGrid"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   Icon            =   "frmOther.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin GridControl.SMGrid SMGrid 
      Height          =   7110
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12541
      BackColor       =   16777215
      BorderColor     =   16576
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
      SelectBackColor =   33023
   End
End
Attribute VB_Name = "frmOther"
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

Private Sub Form_Load()
 Dim lngA As Long, i As Long
 Dim iClr As OLE_COLOR

 Call Randomize(Timer)
 '* Start from random character.
 lngA = 32 + ((Rnd * 65281) \ 1)
 With SMGrid
  .Font = "Arial Unicode MS"
  .TextHeaders = " Support Unicode|^Char|^Item"
  .FullSelection = False
  .ListViewStyle = True
  .BackColor = vbWhite
  '* Add 256 items.
  i = 0
  For lngA = lngA To lngA + 45
   i = i + 1
   If (i Mod 2) Then
    iClr = &HC0FFFF
   Else
    iClr = vbWhite
   End If
   Call .AddItem("^" & ChrW$(lngA) & "|^" & CStr(lngA) & "|^" & i, iClr & "|" & iClr & "|" & iClr, "&HC000C0|", "T|T")
   If (.ActualRow > 1) Then Call .RowHeight(.ActualRow - 1, 450)
  Next
  Call .ColumnWidth(0, 1600)
  Call .ColumnWidth(1, 520)
  Call .ColumnWidth(2, 520)
  Call .ChangedEnabled(0, 1, False)
  Call .ChangedEnabled(0, 3, False)
  Call .ChangedEnabled(0, 8, False)
  Call .ChangedEnabled(1, 8, False)
  Call .ChangedEnabled(1, .ActualRow, False)
  Call .ChangedForeColor(0, 1, vbGrayText)
  Call .ChangedForeColor(0, 3, vbGrayText)
  Call .ChangedForeColor(0, 8, vbGrayText)
  Call .ChangedForeColor(1, 8, vbGrayText)
  Call .ChangedForeColor(1, .ActualRow, vbGrayText)
  .AutoSizeColumn = False
  .FlatScrollbars = True
  .ReDraw
 End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
 Call SMGrid.Move(0, 0, Me.ScaleWidth - 1 * Screen.TwipsPerPixelX, Me.ScaleHeight - SMGrid.Top - 1 * Screen.TwipsPerPixelY)
 SMGrid.ReDraw
End Sub
