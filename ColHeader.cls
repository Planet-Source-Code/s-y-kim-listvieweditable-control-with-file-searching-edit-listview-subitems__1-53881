VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_strText As String
Private m_strKey As String
Private m_sngWidth As Single
Private m_enAlignment As ListColumnAlignmentConstants
Private m_enDataType As eListColumnDataType

Private Function Between( _
                        ByVal Source As String, _
                        ByVal FirstTarget As String, _
                        ByVal SecondTarget As String, _
                        Optional ByVal Compare As VbCompareMethod = VbCompareMethod.vbTextCompare, _
                        Optional FirstTarget_StartPos As Long, _
                        Optional SecondTarget_EndPos As Long, _
                        Optional ByVal bReturnWholeTextIfNotFound As Boolean) As String

'Returns the text between the two targets.

   On Error GoTo Err_Handler

   FirstTarget_StartPos = InStr(1, Source, FirstTarget, Compare)
   If FirstTarget_StartPos > 0 Then
      SecondTarget_EndPos = VBA.InStr(FirstTarget_StartPos + Len(FirstTarget), _
                            Source, SecondTarget, Compare)
      If SecondTarget_EndPos > 0 Then
         SecondTarget_EndPos = SecondTarget_EndPos + Len(SecondTarget) - 1
         'Source = After(Source, FirstTarget, Compare)
         Source = VBA.Mid$(Source, VBA.InStr(1, Source, FirstTarget, Compare) + Len(FirstTarget))
         'Between = Before(Source, SecondTarget, Compare)
         Between = VBA.Left$(Source, VBA.InStr(1, Source, SecondTarget, Compare) - 1)
      Else 'NOT SECONDTARGET_ENDPOS...
         If bReturnWholeTextIfNotFound Then
            Between = Source
         End If
      End If
   Else 'NOT FIRSTTARGET_STARTPOS...
      If bReturnWholeTextIfNotFound Then
         Between = Source
      End If
   End If
Exit Function

Err_Handler:
End Function

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
   Text = m_strText
End Property
Public Property Let Text(ByVal strText As String)
   m_strText = strText
End Property

Public Property Get Key() As String
   Key = m_strKey
End Property
Public Property Let Key(ByVal strKey As String)
   m_strKey = strKey
End Property

Public Property Get Width() As Single
   Width = m_sngWidth
End Property
Public Property Let Width(ByVal sngWidth As Single)
   m_sngWidth = sngWidth
End Property

Public Property Get Alignment() As ListColumnAlignmentConstants
   Alignment = m_enAlignment
End Property
Public Property Let Alignment(ByVal enAlignment As ListColumnAlignmentConstants)
   m_enAlignment = enAlignment
End Property

Public Property Get DataType() As eListColumnDataType
   DataType = m_enDataType
End Property
Public Property Let DataType(ByVal enDataType As eListColumnDataType)
   m_enDataType = enDataType
End Property

Public Function GetTagText(Optional Delimiter As String = "|") As String
   Dim asText(0 To 4) As String
   asText(0) = "Text=" & m_strText 'String
   asText(1) = "Key=" & m_strKey 'String
   asText(2) = "Width=" & m_sngWidth 'Single
   asText(3) = "Alignment=" & m_enAlignment 'ListColumnAlignmentConstants
   asText(4) = "DataType=" & m_enDataType 'eListColumnDataType
   GetTagText = Delimiter & Join(asText, Delimiter) & Delimiter
End Function

Public Sub SetPropsFromTagText(TagText As String, Optional Delimiter As String = "|")
   '|Text=FullName|Key=|Width=1299.969|Alignment=0|DataType=0|Level=0|
   m_strText = Between(TagText, Delimiter & "Text=", Delimiter)
   m_strKey = Between(TagText, Delimiter & "Key=", Delimiter)
   m_sngWidth = Val(Between(TagText, Delimiter & "Width=", Delimiter))
   m_enAlignment = Val(Between(TagText, Delimiter & "Alignment=", Delimiter))
   m_enDataType = Val(Between(TagText, Delimiter & "DataType=", Delimiter))
End Sub


