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

Const col患者名姓 As Integer = 1
Const col患者名名 As Integer = 2
Const col患者カナ姓 As Integer = 3
Const col患者カナ名 As Integer = 4
Const col患者性別 As Integer = 5
Const col患者生年月日 As Integer = 6
Const col検査採取日 As Integer = 7
Const col検査方法 As Integer = 8
Const col検査材料 As Integer = 9
Const col結果日 As Integer = 10
Const col結果 As Integer = 11

Const sheetName As String = ""
Const startRow As Integer = 3

Dim 患者()   As cls患者

Public Property Let read(value As String)
    Dim wbk As Workbook
    Dim wst As Worksheet
    Dim r As Integer
    
    On Error Resume Next
    Set wbk = Workbooks.Open(value)
    If Err.Number > 0 Then
        MsgBox ("CSVファイルの読み込みに失敗しました" & vbCrLf & Err.Description)
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
        ReDim Preserve 患者(i)
        Set 患者(i) = New cls患者
        患者(i).患者カナ姓 = wst.Cells(r, col患者カナ姓)
        患者(i).患者カナ名 = wst.Cells(r, col患者カナ名)
        患者(i).患者姓 = wst.Cells(r, col患者名姓)
        患者(i).患者名 = wst.Cells(r, col患者名名)
        患者(i).患者性別 = wst.Cells(r, col患者性別)
        患者(i).患者生年月日 = wst.Cells(r, col患者生年月日)
        If wst.Cells(r, col検査採取日) <> "" Then
            Set 患者(i).検査 = New cls検査
            患者(i).検査.検査方法 = wst.Cells(r, col検査方法)
            患者(i).検査.検体 = wst.Cells(r, col検査材料)
            患者(i).検査.採取日 = wst.Cells(r, col検査採取日)
            If wst.Cells(r, col結果日) <> "" Then
            Set 患者(i).検査.結果 = New cls結果
                患者(i).検査.結果.結果 = wst.Cells(r, col結果)
                患者(i).検査.結果.結果日 = wst.Cells(r, col結果日)
            End If
        End If
        i = i + 1
        r = r + 1
    Wend

End Property

Public Property Get 患者S() As cls患者()
    患者S = 患者
End Property

