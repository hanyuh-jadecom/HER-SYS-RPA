VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReadData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const col³Ò¼© As Integer = 1
Const col³Ò¼¼ As Integer = 2
Const col³ÒJi© As Integer = 3
Const col³ÒJi¼ As Integer = 4
Const col³Ò«Ê As Integer = 5
Const col³Ò¶Nú As Integer = 6
Const col¸Ìæú As Integer = 7
Const col¸û@ As Integer = 8
Const col¸Þ¿ As Integer = 9
Const colÊú As Integer = 10
Const colÊ As Integer = 11

Const sheetName As String = ""
Const startRow As Integer = 3

Dim ³Ò()   As cls³Ò

Public Property Let read(value As String)
    Dim wbk As Workbook
    Dim wst As Worksheet
    Dim r As Integer
    
    On Error Resume Next
    Set wbk = Workbooks.Open(value)
    If Err.Number > 0 Then
        MsgBox ("CSVt@CÌÇÝÝÉ¸sµÜµ½" & vbCrLf & Err.Description)
        Exit Property
    End If
    On Error GoTo 0
    If sheetName = "" Then
        Set wst = wbk.Worksheets(1)
    Else
        Set wst = wbk.Worksheets(sheetName)
    End If
    Dim i As Integer
    i = 1
    r = startRow
    While wst.Cells(r, 1) <> ""
        ReDim Preserve ³Ò(i)
        Set ³Ò(i) = New cls³Ò
        ³Ò(i).³ÒJi© = wst.Cells(r, col³ÒJi©)
        ³Ò(i).³ÒJi¼ = wst.Cells(r, col³ÒJi¼)
        ³Ò(i).³Ò© = wst.Cells(r, col³Ò¼©)
        ³Ò(i).³Ò¼ = wst.Cells(r, col³Ò¼¼)
        ³Ò(i).³Ò«Ê = wst.Cells(r, col³Ò«Ê)
        ³Ò(i).³Ò¶Nú = wst.Cells(r, col³Ò¶Nú)
        If wst.Cells(r, col¸Ìæú) <> "" Then
            Set ³Ò(i).¸ = New cls¸
            ³Ò(i).¸.¸û@ = wst.Cells(r, col¸û@)
            ³Ò(i).¸.Ì = wst.Cells(r, col¸Þ¿)
            ³Ò(i).¸.Ìæú = wst.Cells(r, col¸Ìæú)
            If wst.Cells(r, colÊú) <> "" Then
            Set ³Ò(i).¸.Ê = New clsÊ
                ³Ò(i).¸.Ê.Ê = wst.Cells(r, colÊ)
                ³Ò(i).¸.Ê.Êú = wst.Cells(r, colÊú)
            End If
        End If
        i = i + 1
        r = r + 1
    Wend

End Property

Public Property Get ³ÒS() As cls³Ò()
    ³ÒS = ³Ò
End Property

