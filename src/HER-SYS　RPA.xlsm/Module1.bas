Attribute VB_Name = "Module1"
Option Explicit


Public Const User_ID As String = ""
Public Const Password  As String = ""

Public Const �O���@�� As String = "�Ή����a�@"
Public Const �ی��� As String = "�y�Y�ی���"




Public Const resultCount As Integer = 4

Public ���� As cls����
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
    rtn = MsgBox("���O�C�����d�b�ɂ��F�؂�����܂��B�������ǂ����OK���N���b�N���Ă��������B", vbOKCancel)
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
    MsgBox "�d�b�ł̊m�F���I�������uOK�v���N���b�N���Ă�������", vbOKOnly
    Set btn = driver.FindElementById("idSIButton9")
    btn.Click


End Sub



Sub �����o�^(������ As String, �������e As String, ���� As String)
    
    'Stop
Dim element As WebElement
Set element = driver.FindElementByClass("search-button-no-transition")
element.Click
'�O���@�֖������
Set element = driver.FindElementByXPath("//input[@data-cy='searchModalOutpatientText']")
element.SendKeys (�O���@��)
'�����{�^�����N�b�N
Set element = driver.FindElementByXPath("//button[@data-cy='searchModalSearchButton']")
element.Click
'�������ꂽ��Ë@�֖����N���b�N
Set element = driver.FindElementByXPath("//div[@data-cy='searchModalTableItemSelect']")
element.Click
'�m�F�{�^�����N���b�N
Set element = driver.FindElementByXPath("//button[@data-cy='searchModalConfirmButton']")
element.Click


    'Stop
Dim select1 As SelectElement
'�������@
Set select1 = driver.FindElementByXPath("//select[@data-cy='inspection1Method']").AsSelect
select1.SelectByText �������e
'����
Set select1 = driver.FindElementByXPath("//select[@data-cy='inspection1TestMaterial']").AsSelect
select1.SelectByText ����
'�J�����_�[��I��
Set element = driver.FindElementByClass("date-time-picker")
'Stop
�J�����_�[�I�� element, CDate(������)
'    Stop

'�m�F�{�^�����N���b�N
Set element = driver.FindElementByXPath("//button[@data-cy='inspectionConfirmButton']")
element.Click

'Stop

'�o�^�{�^�����N���b�N '
Set element = driver.FindElementByXPath("//button[@data-cy='toViewButton']")
element.Click

'Stop
'�o�^�{�^�����N���b�N
Set element = driver.FindElementByXPath("//button[@data-cy='modalRegisterButton']")
element.Click

'Stop

End Sub
Sub �J�����_�[�I��(BaseElement As WebElement, date1 As Date)
    '�N�����m�F����
    Dim element1 As WebElement
    Dim elements1 As WebElements
    'Date�@Picker���J��
    Set element1 = BaseElement.FindElementByXPath("//input[@placeholder = '���t�I��']")
    
    
    
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

