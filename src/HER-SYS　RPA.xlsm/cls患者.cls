VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls患者"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim l患者姓 As String
Dim l患者名 As String
Dim l患者カナ姓 As String
Dim l患者カナ名 As String
Dim l患者性別 As String
Dim l患者生年月日 As Date
Dim l検査 As cls検査
'Dim l結果 As cls結果

Public Property Let 患者姓(value As String)
    l患者姓 = value
End Property
Public Property Get 患者姓() As String
    患者姓 = l患者姓
End Property

Public Property Let 患者名(value As String)
    l患者名 = value
End Property
Public Property Get 患者名() As String
    患者名 = l患者名
End Property

Public Property Let 患者カナ姓(value As String)
    l患者カナ姓 = StrConv(value, vbWide)
End Property
Public Property Get 患者カナ姓() As String
    患者カナ姓 = l患者カナ姓
End Property

Public Property Let 患者カナ名(value As String)
    l患者カナ名 = StrConv(value, vbWide)
End Property
Public Property Get 患者カナ名() As String
    患者カナ名 = l患者カナ名
End Property

Public Property Let 患者性別(value As String)
    l患者性別 = value
End Property
Public Property Get 患者性別() As String
    患者性別 = l患者性別
End Property

Public Property Let 患者生年月日(value As Date)
    l患者生年月日 = value
End Property
Public Property Get 患者生年月日() As Date
    患者生年月日 = l患者生年月日
End Property

Public Property Set 検査(value As cls検査)
    Set l検査 = value
End Property
Public Property Get 検査() As cls検査
    Set 検査 = l検査
End Property

'Public Property Set 結果(value As cls結果)
'    Set l結果 = value
'End Property
'Public Property Get 結果() As cls結果
'    Set 結果 = l結果
'End Property


