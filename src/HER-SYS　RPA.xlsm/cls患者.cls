VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls³Ò"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim l³Ò© As String
Dim l³Ò¼ As String
Dim l³ÒJi© As String
Dim l³ÒJi¼ As String
Dim l³Ò«Ê As String
Dim l³Ò¶Nú As Date
Dim l¸ As cls¸
'Dim lÊ As clsÊ

Public Property Let ³Ò©(value As String)
    l³Ò© = value
End Property
Public Property Get ³Ò©() As String
    ³Ò© = l³Ò©
End Property

Public Property Let ³Ò¼(value As String)
    l³Ò¼ = value
End Property
Public Property Get ³Ò¼() As String
    ³Ò¼ = l³Ò¼
End Property

Public Property Let ³ÒJi©(value As String)
    l³ÒJi© = StrConv(value, vbWide)
End Property
Public Property Get ³ÒJi©() As String
    ³ÒJi© = l³ÒJi©
End Property

Public Property Let ³ÒJi¼(value As String)
    l³ÒJi¼ = StrConv(value, vbWide)
End Property
Public Property Get ³ÒJi¼() As String
    ³ÒJi¼ = l³ÒJi¼
End Property

Public Property Let ³Ò«Ê(value As String)
    l³Ò«Ê = value
End Property
Public Property Get ³Ò«Ê() As String
    ³Ò«Ê = l³Ò«Ê
End Property

Public Property Let ³Ò¶Nú(value As Date)
    l³Ò¶Nú = value
End Property
Public Property Get ³Ò¶Nú() As Date
    ³Ò¶Nú = l³Ò¶Nú
End Property

Public Property Set ¸(value As cls¸)
    Set l¸ = value
End Property
Public Property Get ¸() As cls¸
    Set ¸ = l¸
End Property

'Public Property Set Ê(value As clsÊ)
'    Set lÊ = value
'End Property
'Public Property Get Ê() As clsÊ
'    Set Ê = lÊ
'End Property


