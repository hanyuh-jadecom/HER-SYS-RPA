Attribute VB_Name = "Module2"
Option Explicit

Sub 検査()
    Dim element As WebElement
    '検査タブをクリック
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
    '検査日が表示されていない
        Set element = driver.FindElementByXPath("//button[@data-cy='toCreateButton']")
        element.Click
        If Not (患者.検査 Is Nothing) Then
            '検査をしている
            検索登録 患者.検査
            'Stop
        Else
            '患者だけ登録されていて、検査が登録されていないのは変
            Stop
        End If
    Else
    '新規入力ボタンがなかった。
        '検査日確認
        If element.value = 患者.検査.採取日 Then
            検査検索
        Else
            MsgBox ("一人の患者で複数の検査日がある場合は未対応です")
            Exit Sub
        End If
    End If
    'Stop
End Sub

Sub 検査検索()
    Dim element As WebElement
    Dim find検査 As Boolean
    find検査 = False
    Dim i As Integer
    Dim elements As WebElements
    For i = 1 To resultCount
        Set elements = driver.FindElementsByXPath("//p[@data-cy='inspection" & i & "Method']")
        If elements.Count > 0 Then
            Set element = elements(1)
            If element.Text = 患者.検査.検査方法 Then
                Set element = driver.FindElementByXPath("//p[@data-cy='inspection" & i & "TestMaterial']")
                If element.Text = 患者.検査.検体 Then
                    find検査 = True
                    If Not (患者.検査.結果 Is Nothing) Then
                        Set element = driver.FindElementByXPath("//input[@type='checkbox']")
                       ' Stop
                        If element.IsSelected Then
                        '結果が登録済み
                         'TODO:登録されている結果と同じかどうかチェックする？
                            'Stop
                        Else
                        '結果が未登録
                            結果登録
                        End If
                    End If
                End If
            End If
        End If
    Next
    If Not find検査 Then
        検索登録 患者.検査
    End If
'Stop
End Sub

Sub 結果登録()
    Dim element As WebElement
    Dim elements As WebElements
    Dim find検査 As Boolean
    Dim select1 As SelectElement
    
    Set element = driver.FindElementByXPath("//button[@data-cy='toEditButton']")
    element.Click
    'Stop
    Sleep (1000)
    Set elements = driver.FindElementsByClass("date-time-picker")
    'Stop
    カレンダー選択 elements(3), CDate(患者.検査.結果.結果日)
    'Stop
    find検査 = False
    Dim i As Integer
    Dim i1 As Integer
    For i = 1 To resultCount
        Set select1 = driver.FindElementByXPath("//select[@data-cy='inspection" & i & "Method']").AsSelect
        If select1.SelectedOption.Text = 患者.検査.検査方法 Then
            Set select1 = driver.FindElementByXPath("//select[@data-cy='inspection" & i & "TestMaterial']").AsSelect
            If select1.SelectedOption.Text = 患者.検査.検体 Then
                find検査 = True
                If Not (患者.検査.結果 Is Nothing) Then
                    If 患者.検査.結果.結果 = "陰性" Then
                        i1 = 1
                    ElseIf 患者.検査.結果.結果 = "陰性" Then
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

Sub 検索登録(検査 As cls検査)
    
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
select1.SelectByText 検査.検査方法
'検体
Set select1 = driver.FindElementByXPath("//select[@data-cy='inspection1TestMaterial']").AsSelect
select1.SelectByText 検査.検体
'カレンダーを選択
Set element = driver.FindElementByClass("date-time-picker")
'Stop
カレンダー選択 element, CDate(検査.採取日)
    'Stop
Dim i As Integer
i = 1
Dim i1 As Integer
If Not (検査.結果 Is Nothing) Then
    If 患者.検査.結果.結果 = "陰性" Then
        i1 = 1
    ElseIf 患者.検査.結果.結果 = "陽性" Then
        i1 = 2
    Else
        i1 = 3
    End If
    Set element = driver.FindElementsByXPath("//label[@class='control control--radio']")((i - 1) * 3 + i1)
    element.Click
    
    Dim elements As WebElements
    Set elements = driver.FindElementsByClass("date-time-picker")
    'Stop
    カレンダー選択 elements(3), CDate(患者.検査.結果.結果日)
    
    'Stop
End If

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
