VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Style Properties"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOther 
      Caption         =   "&Another"
      Height          =   465
      Left            =   3690
      TabIndex        =   1
      Top             =   6735
      Width           =   1365
   End
   Begin GridControl.SMGrid SMGrid 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   11456
      AutoSizeColumn  =   0   'False
      BackColor       =   16777215
      BackgroundPicture=   "frmProperties.frx":038A
      BorderStyle     =   7
      ColumnHeadingHeight=   380
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FullSelection   =   -1  'True
      ListViewStyle   =   -1  'True
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   0
      Picture         =   "frmProperties.frx":1EF1
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmProperties"
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

 Private i As Integer

Private Sub cmdOther_Click()
 Call frmOther.Show
End Sub

Private Sub Form_Load()
 Dim tFont As StdFont
 
 Set tFont = New StdFont
 With SMGrid
  .Font = "Tahoma"
  .Font.Size = 8
  .TextHeaders = "  Propiedad|  Valor"
  tFont.Bold = True
  .FullSelection = True
  Call .AddItem("     Descripción", , , , , , "F|F", tFont)
  Call .AddItem("-", , , , , , "F|F")
  Call .AddItem("Título")
  Call .AddItem("Asunto")
  Call .AddItem("Categoría")
  Call .AddItem("Palabras clave")
  Call .AddItem("Escala|No", , , "|C")
  Call .AddItem("Vínculos obsoletos|0", , , "F|T")
  Call .AddItem("Comentarios")
  Call .AddItem("", , , , , , "F|F")
  Call .AddItem("     Origen", , , , , , "F|F", tFont)
  Call .AddItem("-", , , , , , "F|F")
  Call .AddItem("Autor|Edwin Mantilla Santamaría", , , "F|T")
  Call .AddItem("Guardado por|Heriberto Mantilla Santamaría", , , "F|T")
  Call .AddItem("Número de revisión")
  Call .AddItem("Nombre de aplicación|Microsoft Excel")
  Call .AddItem("Organización|HACKPRO TM", , , "F|T")
  Call .AddItem("Fecha de creación|22/11/2005 08:42 p.m.|")
  Call .AddItem("Fecha en que se guardó por última vez|23/11/2005 02:38 p.m.|", , , "|D")
  Call .ColumnWidth(0, 2450)
  Call .ColumnWidth(1, 2410)
  Call .RowHeight(0, 400)
  Call .RowHeight(2, 200)
  Call .RowHeight(10, 200)
  Call .RowHeight(12, 200)
  For i = 3 To 9
   Call .RowColPicture(i, 0, imgIcon.Picture)
  Next
  For i = 13 To 19
   Call .RowColPicture(i, 0, imgIcon.Picture)
  Next
  .AutoSizeColumn = False
  .ReDraw
 End With
End Sub
