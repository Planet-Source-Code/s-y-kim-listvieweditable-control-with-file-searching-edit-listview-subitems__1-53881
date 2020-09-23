VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "ListViewEditable - File Searching Demo"
   ClientHeight    =   6825
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Frame fraItemEdit 
      Caption         =   "Item Edit: (NOTE: Press Return to view the Previous Text after editing an item.)"
      Height          =   1125
      Left            =   90
      TabIndex        =   20
      Top             =   5430
      Width           =   9795
      Begin VB.Label lblItemEdit 
         AutoSize        =   -1  'True
         Caption         =   "lblItemEdit(5)"
         Height          =   180
         Index           =   5
         Left            =   1500
         TabIndex        =   26
         Top             =   810
         Width           =   1110
      End
      Begin VB.Label lblItemEdit 
         AutoSize        =   -1  'True
         Caption         =   "Previous Text:"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   780
         Width           =   1230
      End
      Begin VB.Label lblItemEdit 
         AutoSize        =   -1  'True
         Caption         =   "Text:"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   540
         Width           =   435
      End
      Begin VB.Label lblItemEdit 
         AutoSize        =   -1  'True
         Caption         =   "lblItemEdit(2)"
         Height          =   180
         Index           =   2
         Left            =   1500
         TabIndex        =   23
         Top             =   570
         Width           =   1110
      End
      Begin VB.Label lblItemEdit 
         AutoSize        =   -1  'True
         Caption         =   "Index=0, SubItem Index=0"
         Height          =   180
         Index           =   1
         Left            =   1530
         TabIndex        =   22
         Top             =   330
         Width           =   2565
      End
      Begin VB.Label lblItemEdit 
         AutoSize        =   -1  'True
         Caption         =   "Item:"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   420
      End
   End
   Begin VB.CheckBox chkIncludeFolder 
      Caption         =   "Include folders to the search results if they contain a matched file."
      Height          =   195
      Left            =   1410
      TabIndex        =   19
      Top             =   870
      Width           =   6015
   End
   Begin VB.TextBox txtUpdateFrequency 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      Height          =   270
      Left            =   8820
      TabIndex        =   18
      Text            =   "txtUpdateFrequency"
      Top             =   540
      Width           =   585
   End
   Begin prjListViewEditable.ListViewEditable ListViewEditable1 
      Height          =   2745
      Left            =   100
      TabIndex        =   14
      Top             =   1170
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4842
      HideSelection   =   -1  'True
      Object.Visible         =   0   'False
      ColumnHeadersText=   $"frmMain.frx":0000
      FileSpec        =   "*.mp3"
      IncludeFolder   =   0   'False
      Path            =   "D:\Musics\"
      UpdateFrequency =   25
      HighlightColumn =   0   'False
   End
   Begin VB.CommandButton cmdGetChecked 
      Caption         =   "Get Checked Items"
      Height          =   285
      Left            =   7440
      TabIndex        =   13
      Top             =   5040
      Width           =   2475
   End
   Begin VB.CommandButton cmdTopIndex 
      Caption         =   "Go"
      Height          =   280
      Index           =   0
      Left            =   9240
      TabIndex        =   12
      Top             =   4410
      Width           =   675
   End
   Begin VB.TextBox txtTopIndex 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      Height          =   280
      Left            =   8340
      TabIndex        =   11
      Top             =   4410
      Width           =   855
   End
   Begin VB.CommandButton cmdGetSelected 
      Caption         =   "Get Selected Items"
      Height          =   285
      Left            =   7440
      TabIndex        =   10
      Top             =   4740
      Width           =   2475
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "S&top"
      Height          =   315
      Index           =   3
      Left            =   8670
      TabIndex        =   7
      Top             =   150
      Width           =   1185
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Start Search"
      Height          =   315
      Index           =   1
      Left            =   7020
      TabIndex        =   2
      Top             =   150
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Start &Folder..."
      Height          =   315
      Index           =   0
      Left            =   5070
      TabIndex        =   1
      Top             =   150
      Width           =   1485
   End
   Begin VB.ComboBox cboCategory 
      Height          =   300
      Left            =   1410
      Style           =   2  'µå·Ó´Ù¿î ¸ñ·Ï
      TabIndex        =   0
      Top             =   150
      Width           =   3435
   End
   Begin VB.Label lblTopIndex 
      AutoSize        =   -1  'True
      Caption         =   "TopIndex:"
      Height          =   180
      Left            =   7440
      TabIndex        =   27
      Top             =   4470
      Width           =   855
   End
   Begin VB.Label lblUpdateFrequency2 
      AutoSize        =   -1  'True
      Caption         =   "items"
      Height          =   180
      Left            =   9450
      TabIndex        =   17
      Top             =   600
      Width           =   465
   End
   Begin VB.Label lblUpdateFrequency 
      Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
      AutoSize        =   -1  'True
      Caption         =   "Update display every"
      Height          =   180
      Index           =   2
      Left            =   7020
      TabIndex        =   16
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label lblTargetFiles 
      AutoSize        =   -1  'True
      Caption         =   "File Spec:"
      Height          =   180
      Index           =   1
      Left            =   100
      TabIndex        =   15
      Top             =   210
      Width           =   855
   End
   Begin VB.Label lblVisibleCount 
      AutoSize        =   -1  'True
      Caption         =   "lblVisibleCount:"
      Height          =   180
      Left            =   100
      TabIndex        =   9
      Top             =   4860
      Width           =   1320
   End
   Begin VB.Label lblSearchDir 
      AutoSize        =   -1  'True
      Caption         =   "lblSearchDir"
      Height          =   180
      Left            =   100
      TabIndex        =   8
      Top             =   4170
      Width           =   1020
   End
   Begin VB.Label lblTargetFiles 
      AutoSize        =   -1  'True
      Caption         =   "Target Files:"
      Height          =   180
      Index           =   0
      Left            =   100
      TabIndex        =   6
      Top             =   570
      Width           =   1065
   End
   Begin VB.Label lblItemCount 
      AutoSize        =   -1  'True
      Caption         =   "lblItemCount"
      Height          =   180
      Left            =   100
      TabIndex        =   5
      Top             =   4410
      Width           =   1050
   End
   Begin VB.Label lblIconCount 
      AutoSize        =   -1  'True
      Caption         =   "lblIconCount"
      Height          =   180
      Left            =   100
      TabIndex        =   4
      Top             =   4620
      Width           =   1050
   End
   Begin VB.Label lblDisplayName 
      AutoSize        =   -1  'True
      Caption         =   "lblDisplayName"
      Height          =   180
      Left            =   1440
      TabIndex        =   3
      Top             =   570
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
  
   Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
   
   With ListViewEditable1
      'FileSpec
      If Len(.FileSpec) = 0 Then
         .FileSpec = "*.*"
      End If
      LoadFileSpecs cboCategory, .FileSpec
      
      'DisplayName
      lblDisplayName.Caption = .DisplayName
      
      'SearchDir
      lblSearchDir.Caption = "Search Directory: " & .Path
      
      'Icon Count
      lblIconCount.Caption = "Icon Count: The ImageList for this target file type view contains " & _
                                          .ImageListCount & IIf(.ImageListCount > 1, " different images.", " image.")
                                          
      'UpdateFrequency
      txtUpdateFrequency.Text = .UpdateFrequency
      
      'IncludeFolder
      chkIncludeFolder.Value = Abs(.IncludeFolder)
      
      'Item Count
      lblItemCount = "List Items: " & .FileItemCount & " files - Ready for searching"
      
      'Visible Count
      lblVisibleCount.Caption = "Visible Count: " & ListViewEditable1.VisibleCount & " items are visible on this page"
      
      'Item Text
      lblItemEdit(2) = ""
      lblItemEdit(5) = ""
   End With

End Sub


Private Sub chkIncludeFolder_Click()
   ListViewEditable1.IncludeFolder = Abs(chkIncludeFolder.Value)
End Sub

Private Sub cmdGetChecked_Click()

   Dim i As Long
   Dim Checked() As Long
   With ListViewEditable1
      If .GetCheckedItems(Checked) > 0 Then
         With frmText
            .Caption = "Checked Rows"
            .txtText.Text = ListViewEditable1.RowsArray(Checked)
            .Show vbModeless, Me
            DoEvents
         End With
      End If
   End With
End Sub

Private Sub cmdGetSelected_Click()
   Dim i As Long
   Dim Selected() As Long
   With ListViewEditable1
      If .GetSelectedItems(Selected) > 0 Then
         With frmText
            .Caption = "Selected Rows"
            .txtText.Text = ListViewEditable1.RowsArray(Selected)
            .Show vbModeless, Me
            DoEvents
         End With
      End If
   End With
End Sub

Private Sub cmdTopIndex_Click(Index As Integer)
   With ListViewEditable1
      Select Case Index
      Case 0
         .TopIndex = Val(txtTopIndex.Text)
      End Select
   End With
End Sub


Private Sub cmdSelect_Click(Index As Integer)
  
   With ListViewEditable1
      Select Case Index
      Case 0
         Dim strPath As String
         strPath = .SelectDirectory
         If Len(strPath) Then
            .Path = strPath
            lblDisplayName.Caption = .DisplayName
            .Scan
         End If
      Case 1
         If Len(.Path) Then
            .Scan
         End If
      Case 2
         Unload Me
      Case 3
         .StopScan
      End Select
   End With
End Sub

Private Sub UpdateItemCount()
   With ListViewEditable1
      lblIconCount = "Icon Count: The ImageList for this target file type view contains " & _
                    .ImageListCount & IIf(.ImageListCount > 1, " different images.", " image.")
      If .Stopped Then
         lblItemCount = "List Items: " & .FileItemCount & " files - Search Stopped"
      Else
         lblItemCount = "List Items: " & .FileItemCount & " files - Search Complete"
      End If
      .ResetStopFlag
      txtTopIndex.Text = .TopIndex
   End With
End Sub


Private Sub ListViewEditable1_AfterItemEdit(Cancel As Boolean, Index As Long, Subitem As Long, OldString As String, NewString As String)
   lblItemEdit(1) = "Index=" & Index & ", SubItem Index=" & Subitem
   lblItemEdit(2) = NewString
   lblItemEdit(5) = OldString
End Sub

Private Sub ListViewEditable1_FileDragDropFinish(Count As Long, Files As Long, Folders As Long)
   Call UpdateItemCount
End Sub

Private Sub ListViewEditable1_MouseWheel()
   txtTopIndex.Text = ListViewEditable1.TopIndex
End Sub

Private Sub ListViewEditable1_VScroll()
   txtTopIndex.Text = ListViewEditable1.TopIndex
End Sub

Private Sub ListViewEditable1_Resize()
   lblVisibleCount.Caption = "Visible Count: " & ListViewEditable1.VisibleCount & " items are visible on this page"
End Sub

Private Sub ListViewEditable1_ScanFinish(Stopped As Boolean, Path As String, FileSpec As String, Recursive As Boolean, IncludeFolder As Boolean)
   Call UpdateItemCount
End Sub

Private Sub ListViewEditable1_SearchDirChange(Directory As String)
   lblSearchDir = Directory
   With ListViewEditable1
      lblIconCount = "Icon Count: The ImageList for this target file type view contains " & _
                              .ImageListCount & IIf(.ImageListCount > 1, " different images.", " image.")
      lblItemCount = "ListItems: " & .FileItemCount & " files - Searching ...."
   End With
End Sub

Private Sub cboCategory_Click()
   
   Dim pos As Integer
   Dim lpos As Integer
   Dim Item As String
   
   Item = cboCategory.List(cboCategory.ListIndex)
   pos = InStr(Item, "(") + 1
   lpos = InStr(Item, ")") - pos
   With ListViewEditable1
      .FileSpec = Mid$(Item, pos, lpos)
      lblDisplayName.Caption = .DisplayName
      cmdSelect(1).Enabled = Len(.Path) > 0
   End With
End Sub


Private Sub ListViewEditable1_SubitemClick(Index As Long, Subitem As Long, Button As Integer, Shift As Integer)
   
   lblItemEdit(1) = "Index=" & Index & ", SubItem Index=" & Subitem
   With ListViewEditable1.ListItems(Index)
      If Subitem = 0 Then
         lblItemEdit(2) = .Text
      Else
         lblItemEdit(2) = .SubItems(Subitem)
      End If
      lblItemEdit(5) = ""
   End With
End Sub


Private Sub txtUpdateFrequency_Change()
   ListViewEditable1.UpdateFrequency = Val(txtUpdateFrequency.Text)
End Sub
