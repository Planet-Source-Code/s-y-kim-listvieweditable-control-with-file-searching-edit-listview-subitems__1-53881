VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmText 
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   9210
   StartUpPosition =   1  '¼ÒÀ¯ÀÚ °¡¿îµ¥
   Begin RichTextLib.RichTextBox txtText 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   5530
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   2.00000e5
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmText.frx":0000
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
   On Error GoTo Bye
   txtText.Move 0, 0, ScaleWidth, ScaleHeight
   Exit Sub
Bye:
End Sub
