VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls結果"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim l結果日 As Date
Dim l結果 As String

Public Property Let 結果日(value As Date)
    l結果日 = value
End Property
Public Property Get 結果日() As Date
    結果日 = l結果日
End Property

Public Property Let 結果(value As String)
    l結果 = value
End Property
Public Property Get 結果() As String
    結果 = l結果
End Property

