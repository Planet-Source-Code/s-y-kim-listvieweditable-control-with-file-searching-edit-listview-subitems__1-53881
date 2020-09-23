Attribute VB_Name = "modMSComctLib"
Option Explicit

Public Enum eImageSizingTypes
   sizeNone = 0
   sizeCheckBox
   sizeIcon
End Enum

Public Enum eLedgerColours
   vbLedgerWhite = &HF9FEFF
   vbLedgerGreen = &HD0FFCC
   vbLedgerYellow = &HE1FAFF
   vbLedgerRed = &HE1E1FF
   vbLedgerGrey = &HE0E0E0
   vbLedgerBeige = &HD9F2F7
   vbLedgerSoftWhite = &HF7F7F7
   vbLedgerPureWhite = &HFFFFFF
End Enum

Public Enum ListColumnAlignmentConstants
   lvwColumnLeft = 0
   lvwColumnRight = 1
   lvwColumnCenter = 2
End Enum 'ListColumnAlignmentConstants

Public Enum eListColumnDataType
   ldtString = 0
   ldtNumber = 1
   ldtDateTime = 2
End Enum

Public Function LoadFileSpecs(ByVal cboProp As ComboBox, Optional SelectedItem As String) As String

   With cboProp
      .Clear
      
      .AddItem "All Files and Folders (*.*)"
      .AddItem "ASF Files (*.asf)"
      .AddItem "Applications (*.exe)"
      .AddItem "Bitmap Files (*.bmp)"
      .AddItem "Dynamic Link Libraries (*.dll)"
      .AddItem "Gif Files (*.gif)"
      .AddItem "Htm Documents (*.htm)"
      .AddItem "Html Documents (*.htm*)"
      .AddItem "Html Documents (*.html)"
      .AddItem "Icons (*.ico)"
      .AddItem "Jpg Files (*.jpg)"
      .AddItem "MP3 Files (*.mp3)"
      .AddItem "MPEG Files (*.mpeg)"
      .AddItem "MPG Files (*.mpg)"
      .AddItem "Microsoft Word Documents (*.doc)"
      .AddItem "Rich Text Format Documents (*.rtf)"
      .AddItem "Text Files (*.txt)"
      .AddItem "True Type Fonts (*.ttf)"
      .AddItem "Visual Basic Forms (*.frm)"
      .AddItem "Visual Basic Modules (*.bas)"
      .AddItem "Visual Basic Projects (*.vbp)"
      .AddItem "Windows Help Files (*.hlp)"
      .AddItem "Windows Html Help Files (*.chm)"
      .AddItem "Windows Shortcuts (*.lnk)"
      
      .Tag = "" 'Clear the tag
   
      
      If Len(SelectedItem) > 0 Then
         
         Dim i As Long
         Dim sFileSpec As String
         Dim pos As Long, lpos As Long
         
         LoadFileSpecs = SelectedItem
         
         For i = 0 To .ListCount - 1
            pos = InStr(1, .List(i), "(") + 1
            lpos = InStr(1, .List(i), ")") - pos
            sFileSpec = Mid$(.List(i), pos, lpos)
            If sFileSpec = SelectedItem Then
               .ListIndex = i
               LoadFileSpecs = sFileSpec
               Exit For
            End If
         Next
         
       End If
   End With
End Function


Public Sub LoadeListColumnDataType(ByVal cboProp As ComboBox, Optional SelectedItem As eListColumnDataType)
   Dim Index As eListColumnDataType
   Dim i As Long
   With cboProp
      .Clear
      Index = ldtString:    .AddItem Index & " - " & "ldtString":    .ItemData(.NewIndex) = Index
      Index = ldtNumber:    .AddItem Index & " - " & "ldtNumber":    .ItemData(.NewIndex) = Index
      Index = ldtDateTime:    .AddItem Index & " - " & "ldtDateTime":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function eListColumnDataTypeDesc( _
   ByVal Index As eListColumnDataType) As String
   Select Case Index
   Case ldtString:      eListColumnDataTypeDesc = "ldtString"
   Case ldtNumber:      eListColumnDataTypeDesc = "ldtNumber"
   Case ldtDateTime:       eListColumnDataTypeDesc = "ldtDateTime"
   Case Else
   End Select
End Function


Public Sub LoadListColumnAlignmentConstants(ByVal cboProp As ComboBox, Optional SelectedItem As ListColumnAlignmentConstants)
   Dim Index As ListColumnAlignmentConstants
   Dim i As Long
   With cboProp
      .Clear
      Index = lvwColumnLeft:    .AddItem Index & " - " & "lvwColumnLeft":    .ItemData(.NewIndex) = Index
      Index = lvwColumnRight:    .AddItem Index & " - " & "lvwColumnRight":    .ItemData(.NewIndex) = Index
      Index = lvwColumnCenter:    .AddItem Index & " - " & "lvwColumnCenter":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub


Public Function ListColumnAlignmentConstantsDesc( _
   ByVal Index As ListColumnAlignmentConstants) As String
   Select Case Index
   Case lvwColumnLeft:        ListColumnAlignmentConstantsDesc = "lvwColumnLeft"
   Case lvwColumnRight:       ListColumnAlignmentConstantsDesc = "lvwColumnRight"
   Case lvwColumnCenter:      ListColumnAlignmentConstantsDesc = "lvwColumnCenter"
   Case Else
   End Select
End Function


Public Sub LoadeImageSizingTypes(ByVal cboProp As ComboBox, Optional SelectedItem As eImageSizingTypes)
   Dim Index As eImageSizingTypes
   Dim i As Long
   With cboProp
      .Clear
      Index = sizeNone:    .AddItem Index & " - " & "sizeNone":    .ItemData(.NewIndex) = Index
      Index = sizeCheckBox:    .AddItem Index & " - " & "sizeCheckBox":    .ItemData(.NewIndex) = Index
      Index = sizeIcon:    .AddItem Index & " - " & "sizeIcon":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function eImageSizingTypesDesc( _
   ByVal Index As eImageSizingTypes) As String
   Select Case Index
   Case sizeNone:       eImageSizingTypesDesc = "sizeNone"
   Case sizeCheckBox:      eImageSizingTypesDesc = "sizeCheckBox"
   Case sizeIcon:       eImageSizingTypesDesc = "sizeIcon"
   Case Else
   End Select
End Function


Public Sub LoadeLedgerColours(ByVal cboProp As ComboBox, Optional SelectedItem As eLedgerColours)
   Dim Index As eLedgerColours
   Dim i As Long
   With cboProp
      .Clear
      Index = vbLedgerWhite:    .AddItem "&H" & Hex$(Index) & " - " & "vbLedgerWhite":    .ItemData(.NewIndex) = Index
      Index = vbLedgerGreen:    .AddItem "&H" & Hex$(Index) & " - " & "vbLedgerGreen":   .ItemData(.NewIndex) = Index
      Index = vbLedgerYellow:    .AddItem "&H" & Hex$(Index) & " - " & "vbLedgerYellow":   .ItemData(.NewIndex) = Index
      Index = vbLedgerRed:    .AddItem "&H" & Hex$(Index) & " - " & "vbLedgerRed":   .ItemData(.NewIndex) = Index
      Index = vbLedgerGrey:    .AddItem "&H" & Hex$(Index) & " - " & "vbLedgerGrey":   .ItemData(.NewIndex) = Index
      Index = vbLedgerBeige:    .AddItem "&H" & Hex$(Index) & " - " & "vbLedgerBeige":   .ItemData(.NewIndex) = Index
      Index = vbLedgerSoftWhite:    .AddItem "&H" & Hex$(Index) & " - " & "vbLedgerSoftWhite":   .ItemData(.NewIndex) = Index
      Index = vbLedgerPureWhite:    .AddItem "&H" & Hex$(Index) & " - " & "vbLedgerPureWhite":   .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function eLedgerColoursDesc( _
   ByVal Index As eLedgerColours) As String
   Select Case Index
   Case vbLedgerWhite:        eLedgerColoursDesc = "vbLedgerWhite"
   Case vbLedgerGreen:        eLedgerColoursDesc = "vbLedgerGreen"
   Case vbLedgerYellow:       eLedgerColoursDesc = "vbLedgerYellow"
   Case vbLedgerRed:       eLedgerColoursDesc = "vbLedgerRed"
   Case vbLedgerGrey:      eLedgerColoursDesc = "vbLedgerGrey"
   Case vbLedgerBeige:        eLedgerColoursDesc = "vbLedgerBeige"
   Case vbLedgerSoftWhite:       eLedgerColoursDesc = "vbLedgerSoftWhite"
   Case vbLedgerPureWhite:       eLedgerColoursDesc = "vbLedgerPureWhite"
   Case Else
   End Select
End Function

Public Sub LoadListSortOrderConstants(ByVal cboProp As ComboBox, Optional SelectedItem As ListSortOrderConstants)
   Dim Index As ListSortOrderConstants
   Dim i As Long
   With cboProp
      .Clear
      Index = lvwAscending:    .AddItem Index & " - " & "lvwAscending":    .ItemData(.NewIndex) = Index
      Index = lvwDescending:    .AddItem Index & " - " & "lvwDescending":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function ListSortOrderConstantsDesc( _
   ByVal Index As ListSortOrderConstants) As String
   Select Case Index
   Case lvwAscending:      ListSortOrderConstantsDesc = "lvwAscending"
   Case lvwDescending:        ListSortOrderConstantsDesc = "lvwDescending"
   Case Else
   End Select
End Function


Public Sub LoadOLEDropConstants(ByVal cboProp As ComboBox, Optional SelectedItem As OLEDropConstants)
   Dim Index As OLEDropConstants
   Dim i As Long
   With cboProp
      .Clear
      Index = ccOLEDropNone:    .AddItem Index & " - " & "ccOLEDropNone":    .ItemData(.NewIndex) = Index
      Index = ccOLEDropManual:    .AddItem Index & " - " & "ccOLEDropManual":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function OLEDropConstantsDesc( _
   ByVal Index As OLEDropConstants) As String
   Select Case Index
   Case ccOLEDropNone:        OLEDropConstantsDesc = "ccOLEDropNone"
   Case ccOLEDropManual:      OLEDropConstantsDesc = "ccOLEDropManual"
   Case Else
   End Select
End Function

Public Sub LoadOLEDragConstants(ByVal cboProp As ComboBox, Optional SelectedItem As OLEDragConstants)
   Dim Index As OLEDragConstants
   Dim i As Long
   With cboProp
      .Clear
      Index = ccOLEDragManual:    .AddItem Index & " - " & "ccOLEDragManual":    .ItemData(.NewIndex) = Index
      Index = ccOLEDragAutomatic:    .AddItem Index & " - " & "ccOLEDragAutomatic":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function OLEDragConstantsDesc( _
   ByVal Index As OLEDragConstants) As String
   Select Case Index
   Case ccOLEDragManual:      OLEDragConstantsDesc = "ccOLEDragManual"
   Case ccOLEDragAutomatic:      OLEDragConstantsDesc = "ccOLEDragAutomatic"
   Case Else
   End Select
End Function


Public Sub LoadAppearanceConstants(ByVal cboProp As ComboBox, Optional SelectedItem As AppearanceConstants)
   Dim Index As AppearanceConstants
   Dim i As Long
   With cboProp
      .Clear
      Index = ccFlat:    .AddItem Index & " - " & "ccFlat":    .ItemData(.NewIndex) = Index
      Index = cc3D:    .AddItem Index & " - " & "cc3D":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function AppearanceConstantsDesc( _
   ByVal Index As AppearanceConstants) As String
   Select Case Index
   Case ccFlat:      AppearanceConstantsDesc = "ccFlat"
   Case cc3D:        AppearanceConstantsDesc = "cc3D"
   Case Else
   End Select
End Function

Public Sub LoadBorderStyleConstants(ByVal cboProp As ComboBox, Optional SelectedItem As BorderStyleConstants)
   Dim Index As BorderStyleConstants
   Dim i As Long
   With cboProp
      .Clear
      Index = ccNone:    .AddItem Index & " - " & "ccNone":    .ItemData(.NewIndex) = Index
      Index = ccFixedSingle:    .AddItem Index & " - " & "ccFixedSingle":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function BorderStyleConstantsDesc( _
   ByVal Index As BorderStyleConstants) As String
   Select Case Index
   Case ccNone:      BorderStyleConstantsDesc = "ccNone"
   Case ccFixedSingle:        BorderStyleConstantsDesc = "ccFixedSingle"
   Case Else
   End Select
End Function

Public Sub LoadListLabelEditConstants(ByVal cboProp As ComboBox, Optional SelectedItem As ListLabelEditConstants)
   Dim Index As ListLabelEditConstants
   Dim i As Long
   With cboProp
      .Clear
      Index = lvwAutomatic:    .AddItem Index & " - " & "lvwAutomatic":    .ItemData(.NewIndex) = Index
      Index = lvwManual:    .AddItem Index & " - " & "lvwManual":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function ListLabelEditConstantsDesc( _
   ByVal Index As ListLabelEditConstants) As String
   Select Case Index
   Case lvwAutomatic:      ListLabelEditConstantsDesc = "lvwAutomatic"
   Case lvwManual:      ListLabelEditConstantsDesc = "lvwManual"
   Case Else
   End Select
End Function

Public Sub LoadListArrangeConstants(ByVal cboProp As ComboBox, Optional SelectedItem As ListArrangeConstants)
   Dim Index As ListArrangeConstants
   Dim i As Long
   With cboProp
      .Clear
      Index = lvwNone:    .AddItem Index & " - " & "lvwNone":    .ItemData(.NewIndex) = Index
      Index = lvwAutoLeft:    .AddItem Index & " - " & "lvwAutoLeft":    .ItemData(.NewIndex) = Index
      Index = lvwAutoTop:    .AddItem Index & " - " & "lvwAutoTop":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function ListArrangeConstantsDesc( _
   ByVal Index As ListArrangeConstants) As String
   Select Case Index
   Case lvwNone:        ListArrangeConstantsDesc = "lvwNone"
   Case lvwAutoLeft:       ListArrangeConstantsDesc = "lvwAutoLeft"
   Case lvwAutoTop:        ListArrangeConstantsDesc = "lvwAutoTop"
   Case Else
   End Select
End Function

Public Sub LoadListViewConstants(ByVal cboProp As ComboBox, Optional SelectedItem As ListViewConstants)
   Dim Index As ListViewConstants
   Dim i As Long
   With cboProp
      .Clear
      Index = lvwIcon:    .AddItem Index & " - " & "lvwIcon":    .ItemData(.NewIndex) = Index
      Index = lvwSmallIcon:    .AddItem Index & " - " & "lvwSmallIcon":    .ItemData(.NewIndex) = Index
      Index = lvwList:    .AddItem Index & " - " & "lvwList":    .ItemData(.NewIndex) = Index
      Index = lvwReport:    .AddItem Index & " - " & "lvwReport":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function ListViewConstantsDesc( _
   ByVal Index As ListViewConstants) As String
   Select Case Index
   Case lvwIcon:        ListViewConstantsDesc = "lvwIcon"
   Case lvwSmallIcon:      ListViewConstantsDesc = "lvwSmallIcon"
   Case lvwList:        ListViewConstantsDesc = "lvwList"
   Case lvwReport:      ListViewConstantsDesc = "lvwReport"
   Case Else
   End Select
End Function

Public Sub LoadMousePointerConstants(ByVal cboProp As ComboBox, Optional SelectedItem As MousePointerConstants)
   Dim Index As MousePointerConstants
   Dim i As Long
   With cboProp
      .Clear
      Index = ccDefault:    .AddItem Index & " - " & "ccDefault":    .ItemData(.NewIndex) = Index
      Index = ccArrow:    .AddItem Index & " - " & "ccArrow":    .ItemData(.NewIndex) = Index
      Index = ccCross:    .AddItem Index & " - " & "ccCross":    .ItemData(.NewIndex) = Index
      Index = ccIBeam:    .AddItem Index & " - " & "ccIBeam":    .ItemData(.NewIndex) = Index
      Index = ccIcon:    .AddItem Index & " - " & "ccIcon":    .ItemData(.NewIndex) = Index
      Index = ccSize:    .AddItem Index & " - " & "ccSize":    .ItemData(.NewIndex) = Index
      Index = ccSizeNESW:    .AddItem Index & " - " & "ccSizeNESW":    .ItemData(.NewIndex) = Index
      Index = ccSizeNS:    .AddItem Index & " - " & "ccSizeNS":    .ItemData(.NewIndex) = Index
      Index = ccSizeNWSE:    .AddItem Index & " - " & "ccSizeNWSE":    .ItemData(.NewIndex) = Index
      Index = ccSizeEW:    .AddItem Index & " - " & "ccSizeEW":    .ItemData(.NewIndex) = Index
      Index = ccUpArrow:    .AddItem Index & " - " & "ccUpArrow":    .ItemData(.NewIndex) = Index
      Index = ccHourglass:    .AddItem Index & " - " & "ccHourglass":    .ItemData(.NewIndex) = Index
      Index = ccNoDrop:    .AddItem Index & " - " & "ccNoDrop":    .ItemData(.NewIndex) = Index
      Index = ccArrowHourglass:    .AddItem Index & " - " & "ccArrowHourglass":    .ItemData(.NewIndex) = Index
      Index = ccArrowQuestion:    .AddItem Index & " - " & "ccArrowQuestion":    .ItemData(.NewIndex) = Index
      Index = ccSizeAll:    .AddItem Index & " - " & "ccSizeAll":    .ItemData(.NewIndex) = Index
      Index = ccCustom:    .AddItem Index & " - " & "ccCustom":    .ItemData(.NewIndex) = Index

      .Tag = "" 'Clear the tag
      For i = 0 To .ListCount - 1
         If .ItemData(i) = SelectedItem Then
            .ListIndex = i
            Exit For
         End If
      Next

   End With
End Sub

Public Function MousePointerConstantsDesc( _
   ByVal Index As MSComctlLib.MousePointerConstants) As String
   Select Case Index
   Case ccDefault:      MousePointerConstantsDesc = "ccDefault"
   Case ccArrow:        MousePointerConstantsDesc = "ccArrow"
   Case ccCross:        MousePointerConstantsDesc = "ccCross"
   Case ccIBeam:        MousePointerConstantsDesc = "ccIBeam"
   Case ccIcon:      MousePointerConstantsDesc = "ccIcon"
   Case ccSize:      MousePointerConstantsDesc = "ccSize"
   Case ccSizeNESW:        MousePointerConstantsDesc = "ccSizeNESW"
   Case ccSizeNS:       MousePointerConstantsDesc = "ccSizeNS"
   Case ccSizeNWSE:        MousePointerConstantsDesc = "ccSizeNWSE"
   Case ccSizeEW:       MousePointerConstantsDesc = "ccSizeEW"
   Case ccUpArrow:      MousePointerConstantsDesc = "ccUpArrow"
   Case ccHourglass:       MousePointerConstantsDesc = "ccHourglass"
   Case ccNoDrop:       MousePointerConstantsDesc = "ccNoDrop"
   Case ccArrowHourglass:        MousePointerConstantsDesc = "ccArrowHourglass"
   Case ccArrowQuestion:      MousePointerConstantsDesc = "ccArrowQuestion"
   Case ccSizeAll:      MousePointerConstantsDesc = "ccSizeAll"
   Case ccCustom:       MousePointerConstantsDesc = "ccCustom"
   Case Else
   End Select
End Function


Public Function GetNewValue(ByVal Ctrl As Object) As String
   Dim strTag As String
   strTag = Ctrl.Tag
   GetNewValue = VBA.Mid$(strTag, VBA.InStr(1, strTag, "|NewValue=") + Len("|NewValue="))
   GetNewValue = VBA.Left$(GetNewValue, VBA.InStr(1, GetNewValue, "|") - 1)
   
End Function

Public Function IsChanged(ByVal Ctrl As Object) As Boolean
   Dim strChanged As String
   Dim strTag As String
   strTag = Ctrl.Tag
   strChanged = VBA.Mid$(strTag, VBA.InStr(1, strTag, "|Changed=") + Len("|Changed="))
   strChanged = VBA.Left$(strChanged, VBA.InStr(1, strChanged, "|") - 1)
   IsChanged = Val(strChanged)
End Function

Public Sub LoadCheckBox(ByVal Ctrl As CheckBox, ByVal InitialValue As CheckBoxConstants)
   Ctrl.Tag = "|InitialValue=" & Abs(InitialValue) & "|Changed=0|NewValue=|"
   Ctrl.Value = InitialValue
End Sub

Public Sub LoadTextBox(ByVal Ctrl As TextBox, ByVal InitialValue As String)
   Ctrl.Tag = "|InitialValue=" & InitialValue & "|Changed=0|NewValue=|"
   Ctrl.Text = InitialValue
End Sub

Public Function ChangeTag(ByVal Ctrl As Object, ByVal NewValue As String) As Boolean

   Dim strTag As String
   Dim strInitialValue As String
   
   strTag = Ctrl.Tag
   If Len(strTag) = 0 Then
      
      Ctrl.Tag = "|InitialValue=" & NewValue & "|Changed=0|NewValue=|"
      
   Else
   
      ChangeTag = True
      
      'strInitialValue = Between(strTag, "|InitialValue=", "|")
      strInitialValue = VBA.Mid$(strTag, VBA.InStr(1, strTag, "|InitialValue=") + Len("|InitialValue="))
      strInitialValue = VBA.Left$(strInitialValue, VBA.InStr(1, strInitialValue, "|") - 1)
      
      If strInitialValue = NewValue Then
         strTag = Replace(strTag, "|Changed=1|", "|Changed=0|")
      Else
         strTag = Replace(strTag, "|Changed=0|", "|Changed=1|")
      End If
      
      'strTag = ToRev(strTag, "=")
      strTag = Left$(strTag, InStrRev(strTag, "="))
      
      '|InitialValue=|Changed=0|NewValue=|
      Ctrl.Tag = strTag & NewValue & "|"
   End If
   
Bye:
End Function



