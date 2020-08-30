VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim wst As Worksheet
'Dim r As Integer者

Private Sub btnOpenBrowser_Click()
    OpenBrowser
    login
End Sub

Private Sub CommandButton1_Click()
    Dim Target As String
    Target = Application.GetOpenFilename("CSVファイル(*.csv),*.csv,EXCELファイル(*.xls*),*.xls*")
    If Target = "False" Then Exit Sub
    Me.TextBox1.Text = Target
End Sub

Private Sub CommandButton2_Click()
    If Trim(Me.TextBox1.Text) = "" Then
        MsgBox ("ファイル名を入力してください")
        Exit Sub
    End If
    
    'CSVファイルを開く
    Dim reader As New clsReadData
    reader.read = Me.TextBox1.Text
    
    OpenBrowser
    login
    Sleep (1000)
    
    'Dim 患者 As cls患者
    Dim element As WebElement
    Dim errnumber As Integer
    Dim i As Integer
    Dim i1 As Integer
    For i = 1 To UBound(reader.患者S)
        Set 患者 = reader.患者S()(i)
        '登録されているか？
        'フリガナで検索
        Set element = driver.FindElementByXPath("//input[@data-cy='phoneticText']")
        element.SendKeys (患者.患者カナ姓 & 患者.患者カナ名)
        'Stop
        '生年月日で検索
        Set element = driver.FindElementByXPath("//input[@data-cy='birthdayText']")
        element.SendKeys (患者.患者生年月日)
        'Stop
        '検索ボタンwクリック
        Set element = driver.FindElementByXPath("//button[@data-cy='searchButton']")
        element.Click
        'Stop
        Sleep (1000)
        i1 = 1
        Do While i1 < 10
            If driver.FindElementsByXPath("//span[@data-cy='itemId']").Count > 0 Then Exit Do
            If driver.FindElementsByClass("out-entry").Count > 0 Then Exit Do
            Sleep (1000)
            i1 = i1 + 1
        Loop
        '登録済みなら一覧表が出るはず｡
        errnumber = 0
        On Error Resume Next
        Set element = driver.FindElementByXPath("//span[@data-cy='itemId']")
        errnumber = Err.Number
        On Error GoTo 0
        If errnumber > 0 Then
            '見つからなければ新規作成ボタンを押す
            Set element = driver.FindElementByXPath("//button[@data-cy='basicCreateButton']")
            element.Click
            'Stop
            
            'フリガナ　姓
            Set element = driver.FindElementByXPath("//input[@data-cy='phoneticFamilyName']")
            element.SendKeys (患者.患者カナ姓)
            'フリガナ　名
            Set element = driver.FindElementByXPath("//input[@data-cy='phoneticLastName']")
            element.SendKeys (患者.患者カナ名)
            '漢字　姓
            Set element = driver.FindElementByXPath("//input[@data-cy='nameFamilyName']")
            element.SendKeys (患者.患者姓)
            '漢字　名
            Set element = driver.FindElementByXPath("//input[@data-cy='nameLastName']")
            element.SendKeys (患者.患者名)
            '生年月日
            Set element = driver.FindElementByXPath("//input[@data-cy='birthdayText']")
            element.SendKeys (患者.患者生年月日)
            '
            '保健所
            Dim sel As SelectElement
            Set sel = driver.FindElementByXPath("//select[@data-cy='healthCenterSelect']").AsSelect
            sel.SelectByText (保健所)
            '確認ボタン
            Set element = driver.FindElementByXPath("//button[@data-cy='basicConfirmButton']")
            element.Click
            'Stop
            
            '登録ボタン
            Set element = driver.FindElementByXPath("//button[@data-cy='registButton']")
            element.Click
            'Stop
             
             '完了ボタン
            Set element = driver.FindElementByXPath("//button[@data-cy='completeButton']")
            element.Click
            'Stop
            
            検査
        
        Else
            '見つかればその患者のページに飛ぶ
            'Stop
            Set element = driver.FindElementByXPath("//span[@data-cy='itemName']").FindElementByXPath("a")
            element.Click
            
            検査
            
            'Stop
            
        End If
        'トップに戻る
        Set element = driver.FindElementByClass("nav-button")
        element.Click
        Sleep (1000)
        Set element = driver.FindElementByXPath("//button[@data-cy='registerList']")
        element.Click
    Next
    MsgBox ("終了しました。")
End Sub


