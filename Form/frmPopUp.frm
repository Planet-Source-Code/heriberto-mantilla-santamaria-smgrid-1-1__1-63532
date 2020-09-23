VERSION 5.00
Begin VB.Form frmPopUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PopUp Demo - SMGrid"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3870
   Icon            =   "frmPopUp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin GridControl.SMGrid SMGrid 
      Height          =   2070
      Left            =   105
      TabIndex        =   3
      Top             =   360
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   3651
      BackColor       =   16777215
      BorderStyle     =   6
      ColumnHeadingHeight=   350
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextHeaders     =   "^Choose your Option"
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   165
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   390
      Left            =   2640
      TabIndex        =   1
      Top             =   2520
      Width           =   1110
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   4
      Top             =   2595
      Width           =   60
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Write your opinion for this control."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2910
   End
End
Attribute VB_Name = "frmPopUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2005       *'
'******************************************************'
'* Comments: PopUp form for SMGrid                    *'
'******************************************************'
'* Now my website is available but alone the version  *'
'* in Spanish.                                        *'
'-----------------------------------------------------*'
'* WebSite:  http://www.geocities.com/hackprotm/      *'
'******************************************************'
'*        All rights Reserved © HACKPRO TM 2005       *'
'******************************************************'
Option Explicit

Private Sub cmdSave_Click()
 Call frmDemo.SMGrid.ChangedItem(frmDemo.SMGrid.ColPos, frmDemo.SMGrid.RowPos, SMGrid.ItemText(SMGrid.ColPos, SMGrid.RowPos))
 Call Unload(frmPopUp)
 Set frmPopUp = Nothing
End Sub

Private Sub Form_Load()
 With SMGrid
  Me.Caption = .GetControlVersion
  .TextHeaders = "^Choose your Option"
  Call .AddItem("Excellent", , , "O")
  Call .AddItem("Good", , , "O")
  Call .AddItem("Average", , , "O")
  Call .AddItem("Below Average", , , "O")
  Call .AddItem("Poor", , , "O", , "T")
  Select Case frmDemo.SMGrid.ItemText(frmDemo.SMGrid.ColPos, frmDemo.SMGrid.RowPos)
   Case "Excellent":     Call .ChangedValue(0, 1, True)
   Case "Good":          Call .ChangedValue(0, 2, True)
   Case "Average":       Call .ChangedValue(0, 3, True)
   Case "Below Average": Call .ChangedValue(0, 4, True)
   Case "Poor":          Call .ChangedValue(0, 5, True)
  End Select
  .AutoSizeColumn = True
  .ReDraw
 End With
End Sub

Private Sub SMGrid_Click(ByVal Row As Long, ByVal Col As Long, ByVal Style As String)
 lblValue.Caption = SMGrid.ItemText(Col, Row)
End Sub
