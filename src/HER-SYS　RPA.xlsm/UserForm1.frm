VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8910.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim wst As Worksheet
'Dim r As Integer��

Private Sub btnOpenBrowser_Click()
    OpenBrowser
    login
End Sub

Private Sub CommandButton1_Click()
    Dim Target As String
    Target = Application.GetOpenFilename("CSV�t�@�C��(*.csv),*.csv,EXCEL�t�@�C��(*.xls*),*.xls*")
    If Target = "False" Then Exit Sub
    Me.TextBox1.Text = Target
End Sub

Private Sub CommandButton2_Click()
    If Trim(Me.TextBox1.Text) = "" Then
        MsgBox ("�t�@�C��������͂��Ă�������")
        Exit Sub
    End If
    
    'CSV�t�@�C�����J��
    Dim reader As New clsReadData
    reader.read = Me.TextBox1.Text
    
    OpenBrowser
    login
    Sleep (1000)
    
    'Dim ���� As cls����
    Dim element As WebElement
    Dim errnumber As Integer
    Dim i As Integer
    Dim i1 As Integer
    For i = 1 To UBound(reader.����S)
        Set ���� = reader.����S()(i)
        '�o�^����Ă��邩�H
        '�t���K�i�Ō���
        Set element = driver.FindElementByXPath("//input[@data-cy='phoneticText']")
        element.SendKeys (����.���҃J�i�� & ����.���҃J�i��)
        'Stop
        '���N�����Ō���
        Set element = driver.FindElementByXPath("//input[@data-cy='birthdayText']")
        element.SendKeys (����.���Ґ��N����)
        'Stop
        '�����{�^��w�N���b�N
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
        '�o�^�ς݂Ȃ�ꗗ�\���o��͂��
        errnumber = 0
        On Error Resume Next
        Set element = driver.FindElementByXPath("//span[@data-cy='itemId']")
        errnumber = Err.Number
        On Error GoTo 0
        If errnumber > 0 Then
            '������Ȃ���ΐV�K�쐬�{�^��������
            Set element = driver.FindElementByXPath("//button[@data-cy='basicCreateButton']")
            element.Click
            'Stop
            
            '�t���K�i�@��
            Set element = driver.FindElementByXPath("//input[@data-cy='phoneticFamilyName']")
            element.SendKeys (����.���҃J�i��)
            '�t���K�i�@��
            Set element = driver.FindElementByXPath("//input[@data-cy='phoneticLastName']")
            element.SendKeys (����.���҃J�i��)
            '�����@��
            Set element = driver.FindElementByXPath("//input[@data-cy='nameFamilyName']")
            element.SendKeys (����.���Ґ�)
            '�����@��
            Set element = driver.FindElementByXPath("//input[@data-cy='nameLastName']")
            element.SendKeys (����.���Җ�)
            '���N����
            Set element = driver.FindElementByXPath("//input[@data-cy='birthdayText']")
            element.SendKeys (����.���Ґ��N����)
            '
            '�ی���
            Dim sel As SelectElement
            Set sel = driver.FindElementByXPath("//select[@data-cy='healthCenterSelect']").AsSelect
            sel.SelectByText (�ی���)
            '�m�F�{�^��
            Set element = driver.FindElementByXPath("//button[@data-cy='basicConfirmButton']")
            element.Click
            'Stop
            
            '�o�^�{�^��
            Set element = driver.FindElementByXPath("//button[@data-cy='registButton']")
            element.Click
            'Stop
             
             '�����{�^��
            Set element = driver.FindElementByXPath("//button[@data-cy='completeButton']")
            element.Click
            'Stop
            
            ����
        
        Else
            '������΂��̊��҂̃y�[�W�ɔ��
            'Stop
            Set element = driver.FindElementByXPath("//span[@data-cy='itemName']").FindElementByXPath("a")
            element.Click
            
            ����
            
            'Stop
            
        End If
        '�g�b�v�ɖ߂�
        Set element = driver.FindElementByClass("nav-button")
        element.Click
        Sleep (1000)
        Set element = driver.FindElementByXPath("//button[@data-cy='registerList']")
        element.Click
    Next
    MsgBox ("�I�����܂����B")
End Sub


