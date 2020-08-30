Attribute VB_Name = "Module1"
Option Explicit


Public Const User_ID As String = ""
Public Const Password  As String = ""

Public Const 外来機関 As String = "石岡第一病院"
Public Const 保健所 As String = "土浦保健所"




Public Const resultCount As Integer = 4

Public 患者 As cls患者
Public driver As Selenium.ChromeDriver

Const HERSYS_URL As String = "https://stop.cov19.mhlw.go.jp/signin/"

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
Public Sub getDriver()
    If driver Is Nothing Then
        Set driver = New Selenium.ChromeDriver
        driver.Get (HERSYS_URL)
    End If
End Sub


Public Sub OpenBrowser()
    getDriver
    If driver Is Nothing Then
        Set driver = New Selenium.ChromeDriver
    End If
    driver.Get (HERSYS_URL)
End Sub

Public Sub login()
    Dim rtn As Integer
    rtn = MsgBox("ログイン時電話による認証があります。準備が良ければOKをクリックしてください。", vbOKCancel)
    If rtn = vbCancel Then Exit Sub
    Dim btn As WebElement
    Set btn = driver.FindElementByXPath("/html/body/div/div/div/header/div[3]/div/button")
    btn.Click
    Dim ipt As WebElement
    Set ipt = driver.FindElementByName("loginfmt")
    ipt.SendKeys (User_ID)
    Set btn = driver.FindElementById("idSIButton9")
    btn.Click
    Set ipt = driver.FindElementByName("passwd")
    ipt.SendKeys (Password)
    Sleep (1000)
    Set btn = driver.FindElementById("idSIButton9")
    btn.Click
    MsgBox "電話での確認が終わったら「OK」をクリックしてください", vbOKOnly
    Set btn = driver.FindElementById("idSIButton9")
    btn.Click


End Sub



Sub 検索登録(検査日 As String, 検査内容 As String, 検体 As String)
    
    'Stop
Dim element As WebElement
Set element = driver.FindElementByClass("search-button-no-transition")
element.Click
'外来機関名を入力
Set element = driver.FindElementByXPath("//input[@data-cy='searchModalOutpatientText']")
element.SendKeys (外来機関)
'検索ボタンをクック
Set element = driver.FindElementByXPath("//button[@data-cy='searchModalSearchButton']")
element.Click
'検索された医療機関名をクリック
Set element = driver.FindElementByXPath("//div[@data-cy='searchModalTableItemSelect']")
element.Click
'確認ボタンをクリック
Set element = driver.FindElementByXPath("//button[@data-cy='searchModalConfirmButton']")
element.Click


    'Stop
Dim select1 As SelectElement
'検査方法
Set select1 = driver.FindElementByXPath("//select[@data-cy='inspection1Method']").AsSelect
select1.SelectByText 検査内容
'検体
Set select1 = driver.FindElementByXPath("//select[@data-cy='inspection1TestMaterial']").AsSelect
select1.SelectByText 検体
'カレンダーを選択
Set element = driver.FindElementByClass("date-time-picker")
'Stop
カレンダー選択 element, CDate(検査日)
'    Stop

'確認ボタンをクリック
Set element = driver.FindElementByXPath("//button[@data-cy='inspectionConfirmButton']")
element.Click

'Stop

'登録ボタンをクリック '
Set element = driver.FindElementByXPath("//button[@data-cy='toViewButton']")
element.Click

'Stop
'登録ボタンをクリック
Set element = driver.FindElementByXPath("//button[@data-cy='modalRegisterButton']")
element.Click

'Stop

End Sub
Sub カレンダー選択(BaseElement As WebElement, date1 As Date)
    '年月を確認する
    Dim element1 As WebElement
    Dim elements1 As WebElements
    'Date　Pickerを開く
    Set element1 = BaseElement.FindElementByXPath("//input[@placeholder = '日付選択']")
    
    
    
    Set element1 = BaseElement
    element1.ScrollIntoView
    element1.Click
    'Stop
    Set elements1 = BaseElement.FindElementsByClass("custom-button-content")
    'Stop
    Do Until Format(date1, "yyyyMM") = elements1(2).Text & Format(Val(elements1(1).Text), "00")
        If Format(date1, "yyyyMM") < elements1(2).Text & Format(Val(elements1(1).Text), "00") Then
            Set elements1 = BaseElement.FindElementsByClass("arrow-month")
            elements1(1).Click
        Else
            Set elements1 = BaseElement.FindElementsByClass("arrow-month")
            elements1(2).Click
        End If
        Set elements1 = BaseElement.FindElementsByClass("custom-button-content")
    Loop
    'Stop
    Set elements1 = BaseElement.FindElementsByClass("datepicker-day-text")
    Dim i As Integer
    Dim i1 As Integer
    i = elements1.Count
    For i1 = 1 To i
        If elements1(i1).Text = Day(date1) Then
            elements1(i1).Click
            Exit For
        End If
    Next
    

End Sub

