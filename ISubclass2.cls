VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISubclass2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub After( _
   lReturn As Long, _
   ByVal hWnd As Long, _
   ByVal uMsg As eMsg, _
   ByVal wParam As Long, _
   ByVal lParam As Long)
'
End Sub

Public Sub Before( _
   bHandled As Boolean, _
   lReturn As Long, _
   hWnd As Long, _
   uMsg As eMsg, _
   wParam As Long, _
   lParam As Long)
'
End Sub


