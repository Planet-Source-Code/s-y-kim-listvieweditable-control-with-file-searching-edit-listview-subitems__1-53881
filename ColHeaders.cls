VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColHeaders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colItems As Collection

Private Sub Class_Initialize()
   Set m_colItems = New Collection
End Sub
Private Sub Class_Terminate()
   Set m_colItems = Nothing
End Sub

Public Property Get Items() As Collection
   Set Items = m_colItems
End Property
Public Property Set Items(ByVal colItems As Collection)
   Set m_colItems = colItems
End Property

Public Property Get Item(Index As Variant) As ColHeader
Attribute Item.VB_UserMemId = 0
   Set Item = m_colItems(Index)
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
   Set NewEnum = m_colItems.[_NewEnum]
End Property

Public Property Get Count() As Long
   Count = m_colItems.Count
End Property

Public Function Remove(Index As Variant) As Boolean
   On Error GoTo Bye
   m_colItems.Remove Index
   Remove = True
   Exit Function
Bye:
End Function

Public Sub Clear()
   Set m_colItems = New Collection
End Sub

Public Function Exists(Index As Variant) As Boolean
   On Error GoTo Bye
   Dim oItem As Object
   Set oItem = m_colItems(Index)
   Exists = True
   Exit Function
Bye:
End Function

Public Function AddEx( _
         Optional ByVal Item As ColHeader, _
         Optional Key As String, _
         Optional Before As Variant, Optional After As Variant) As ColHeader
      
   If ObjPtr(Item) = 0 Then
      Set AddEx = New ColHeader
   Else
      Set AddEx = Item
   End If
   
   If LenB(Key) Then
      m_colItems.Add AddEx, Key, Before, After
   Else
      m_colItems.Add AddEx, , Before, After
   End If
End Function

Public Function Add( _
         Optional ByVal Key As String, _
         Optional ByVal Text As String, _
         Optional ByVal Width As Single, _
         Optional ByVal Alignment As ListColumnAlignmentConstants, _
         Optional ByVal DataType As eListColumnDataType, _
         Optional Before As Variant, Optional After As Variant) As ColHeader
         
   On Error GoTo Bye
   Dim oAdd As ColHeader
   Set oAdd = New ColHeader
   If LenB(Key) Then
      m_colItems.Add oAdd, Key, Before, After
   Else
      m_colItems.Add oAdd, , Before, After
   End If
   With oAdd
      .Text = Text
      .Alignment = Alignment
      .Key = Key
      .Width = Width
      .DataType = DataType
   End With
   Set Add = oAdd
   Exit Function
Bye:
End Function

Public Sub AddFromColumnHeaders(ColumnHeaders As MSComctlLib.ColumnHeaders)
   Dim i As Long
   Me.Clear
   With ColumnHeaders
      For i = 1 To .Count
         With .Item(i)
            Me.Add .Key, .Text, .Width, .Alignment, Val(.Tag)
         End With
      Next i
   End With
End Sub

Public Sub SetColumnHeaders(ColumnHeaders As MSComctlLib.ColumnHeaders)
   Dim i As Long
   
   ColumnHeaders.Clear
   For i = 1 To Me.Count
      With Me.Item(i)
         If Len(.Key) Then
            ColumnHeaders.Add(, , .Text, .Width, .Alignment).Tag = .DataType
         Else
            ColumnHeaders.Add(, .Key, .Text, .Width, .Alignment).Tag = .DataType
         End If
      End With
   Next i
End Sub

Public Function GetTagText(Optional Delimiter As String = vbCrLf) As String
   Dim i As Long
   Dim asText() As String
   If Count > 0 Then
      ReDim asText(1 To Count)
      For i = 1 To Count
         asText(i) = Item(i).GetTagText("|")
      Next i
      GetTagText = Delimiter & Join(asText, Delimiter) & Delimiter
   End If
End Function

Public Sub CreateItemsFromTagText(TagText As String, _
                                                      Optional Delimiter As String = vbCrLf, _
                                                      Optional PropDelimiter As String = "|")
   Dim asText() As String
   Dim i As Long
   asText = Split(TagText, Delimiter)
   For i = 0 To UBound(asText)
      If Len(asText(i)) > 0 Then
         Add.SetPropsFromTagText asText(i), "|"
      End If
   Next
End Sub
