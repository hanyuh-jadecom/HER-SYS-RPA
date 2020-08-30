VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls検査"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim l採取日 As Date
Dim l検査方法 As String
Dim l検体 As String
Dim l結果 As cls結果

Public Property Let 採取日(value As Date)
    l採取日 = value
End Property
Public Property Get 採取日() As Date
    採取日 = l採取日
End Property

Public Property Let 検査方法(value As String)
    l検査方法 = value
End Property
Public Property Get 検査方法() As String
    検査方法 = l検査方法
End Property

Public Property Let 検体(value As String)
    l検体 = value
End Property
Public Property Get 検体() As String
    検体 = l検体
End Property

Public Property Set 結果(value As cls結果)
    Set l結果 = value
End Property
Public Property Get 結果() As cls結果
    If l結果 Is Nothing Then
        Set 結果 = Nothing
    Else
        Set 結果 = l結果
    End If
End Property

