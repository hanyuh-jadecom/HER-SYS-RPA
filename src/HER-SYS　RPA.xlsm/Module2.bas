Attribute VB_Name = "Module2"
Option Explicit

Sub ����()
    Dim element As WebElement
    '�����^�u���N���b�N
'    Sleep (1000)
'    driver.GoForward
    'Stop
'    Sleep (5000)
    Set element = driver.FindElementByXPath("//div[@data-cy='inspectionTab']")
    Set element = element.FindElementByClass("tab-item-text")
    element.ScrollIntoView
    element.Click
    
    'Stop
                
    Dim errno As Integer
    errno = 0
    Sleep (1000)
    On Error Resume Next
    Set element = driver.FindElementByXPath("//input[@data-cy='inspectionViewDate']")
    errno = Err.Number
    On Error GoTo 0
    If errno > 0 Then
    '���������\������Ă��Ȃ�
        Set element = driver.FindElementByXPath("//button[@data-cy='toCreateButton']")
        element.Click
        If Not (����.���� Is Nothing) Then
            '���������Ă���
            �����o�^ ����.����
            'Stop
        Else
            '���҂����o�^����Ă��āA�������o�^����Ă��Ȃ��͕̂�
            Stop
        End If
    Else
    '�V�K���̓{�^�����Ȃ������B
        '�������m�F
        If element.value = ����.����.�̎�� Then
            ��������
        Else
            MsgBox ("��l�̊��҂ŕ����̌�����������ꍇ�͖��Ή��ł�")
            Exit Sub
        End If
    End If
    'Stop
End Sub

Sub ��������()
    Dim element As WebElement
    Dim find���� As Boolean
    find���� = False
    Dim i As Integer
    Dim elements As WebElements
    For i = 1 To resultCount
        Set elements = driver.FindElementsByXPath("//p[@data-cy='inspection" & i & "Method']")
        If elements.Count > 0 Then
            Set element = elements(1)
            If element.Text = ����.����.�������@ Then
                Set element = driver.FindElementByXPath("//p[@data-cy='inspection" & i & "TestMaterial']")
                If element.Text = ����.����.���� Then
                    find���� = True
                    If Not (����.����.���� Is Nothing) Then
                        Set element = driver.FindElementByXPath("//input[@type='checkbox']")
                       ' Stop
                        If element.IsSelected Then
                        '���ʂ��o�^�ς�
                         'TODO:�o�^����Ă��錋�ʂƓ������ǂ����`�F�b�N����H
                            'Stop
                        Else
                        '���ʂ����o�^
                            ���ʓo�^
                        End If
                    End If
                End If
            End If
        End If
    Next
    If Not find���� Then
        �����o�^ ����.����
    End If
'Stop
End Sub

Sub ���ʓo�^()
    Dim element As WebElement
    Dim elements As WebElements
    Dim find���� As Boolean
    Dim select1 As SelectElement
    
    Set element = driver.FindElementByXPath("//button[@data-cy='toEditButton']")
    element.Click
    'Stop
    Sleep (1000)
    Set elements = driver.FindElementsByClass("date-time-picker")
    'Stop
    �J�����_�[�I�� elements(3), CDate(����.����.����.���ʓ�)
    'Stop
    find���� = False
    Dim i As Integer
    Dim i1 As Integer
    For i = 1 To resultCount
        Set select1 = driver.FindElementByXPath("//select[@data-cy='inspection" & i & "Method']").AsSelect
        If select1.SelectedOption.Text = ����.����.�������@ Then
            Set select1 = driver.FindElementByXPath("//select[@data-cy='inspection" & i & "TestMaterial']").AsSelect
            If select1.SelectedOption.Text = ����.����.���� Then
                find���� = True
                If Not (����.����.���� Is Nothing) Then
                    If ����.����.����.���� = "�A��" Then
                        i1 = 1
                    ElseIf ����.����.����.���� = "�A��" Then
                        i1 = 2
                    Else
                        i1 = 3
                    End If
                    Set element = driver.FindElementsByXPath("//label[@class='control control--radio']")((i - 1) * 3 + i1)
                    element.Click
                    'Stop
                End If
            End If
        End If
    Next
    
    
    'Stop
    Set element = driver.FindElementByXPath("//button[@data-cy='inspectionConfirmButton']")
    element.Click
    Sleep (1000)
    Set element = driver.FindElementByXPath("//button[@data-cy='toViewButton']")
    element.Click
    Sleep (1000)
    Set element = driver.FindElementByXPath("//button[@data-cy='modalRegisterButton']")
    element.Click
    

'Stop

End Sub

Sub �����o�^(���� As cls����)
    
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
select1.SelectByText ����.�������@
'����
Set select1 = driver.FindElementByXPath("//select[@data-cy='inspection1TestMaterial']").AsSelect
select1.SelectByText ����.����
'�J�����_�[��I��
Set element = driver.FindElementByClass("date-time-picker")
'Stop
�J�����_�[�I�� element, CDate(����.�̎��)
    'Stop
Dim i As Integer
i = 1
Dim i1 As Integer
If Not (����.���� Is Nothing) Then
    If ����.����.����.���� = "�A��" Then
        i1 = 1
    ElseIf ����.����.����.���� = "�z��" Then
        i1 = 2
    Else
        i1 = 3
    End If
    Set element = driver.FindElementsByXPath("//label[@class='control control--radio']")((i - 1) * 3 + i1)
    element.Click
    
    Dim elements As WebElements
    Set elements = driver.FindElementsByClass("date-time-picker")
    'Stop
    �J�����_�[�I�� elements(3), CDate(����.����.����.���ʓ�)
    
    'Stop
End If

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
