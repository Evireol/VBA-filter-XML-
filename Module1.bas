Attribute VB_Name = "Module1"
Public Function GetSetFunc(ByVal funcIndex As Integer, ByVal funcEntryPoint As IXMLDOMNode, ByVal nodeIndex As Integer) As Boolean
    Dim myNodes As IXMLDOMNodeList
    Dim myNode As IXMLDOMNode
    Dim myChildNodes As IXMLDOMNodeList
    Dim nChildNode As Integer
    Dim checkPR As Boolean
        checkPR = False
    Dim checkFR As Boolean
        checkFR = False
    Dim i As String
        i = 0
    
    'PersonReply
    Set myNodes = funcEntryPoint.SelectNodes("//PersonReply")
    Set myNode = myNodes(nodeIndex)
    Set myChildNodes = myNode.ChildNodes
    
        Select Case funcIndex
            Case 1
                For nChildNode = 0 To myChildNodes.Length - 1
                    If checkPR = False Then
                        Cells(2, (nChildNode + 1)) = myChildNodes(nChildNode).BaseName  'Наименование столбца
                    End If
                    Cells((nodeIndex + 3), (nChildNode + 1)) = myChildNodes(nChildNode).Text 'Значение столбца
                    i = i + 1
                Next nChildNode
                checkPR = True
            Case 2
                For nChildNode = 0 To myChildNodes.Length - 1
                    UserForm1.Controls("TB" & (nChildNode + 1)).Text = myChildNodes(nChildNode).Text 'Значение столбца
                    i = i + 1
                Next nChildNode
            Case 3
                For nChildNode = 0 To myChildNodes.Length - 1
                    myChildNodes(nChildNode).Text = UserForm1.Controls("TB" & (nChildNode + 1)).Text 'Значение столбца
                    i = i + 1
                Next nChildNode
        End Select
        
    'ficoRisk
    Set myNodes = funcEntryPoint.SelectNodes("//ficoRisk")
    Set myNode = myNodes(nodeIndex)
    Set myChildNodes = myNode.ChildNodes
    
        Select Case funcIndex
            Case 1
                For nChildNode = 0 To myChildNodes.Length - 1
                    If checkFR = False Then
                        Cells(2, (nChildNode + (i + 1))) = myChildNodes(nChildNode).BaseName    'Наименование столбца
                    End If
                    Cells((nodeIndex + 3), (nChildNode + (i + 1))) = myChildNodes(nChildNode).Text   'Значение столбца
                Next nChildNode
                checkFR = True
            Case 2
                For nChildNode = 0 To myChildNodes.Length - 1
                    UserForm1.Controls("TB" & (nChildNode + i + 1)).Text = myChildNodes(nChildNode).Text 'Значение столбца
                Next nChildNode
            Case 3
                For nChildNode = 0 To myChildNodes.Length - 1
                    myChildNodes(nChildNode).Text = UserForm1.Controls("TB" & (nChildNode + i + 1)).Text 'Значение столбца
                Next nChildNode
        End Select

    GetSetFunc = True
    Exit Function
    
End Function
Sub XmlSearch()
    Dim nNode As Integer
    Dim errNum As Integer
    Dim returnVal As Boolean
    
    'Подгрузка файла XML
    Set objXML = New MSXML2.DOMDocument60
    UserForm1.fileSrc = Application.GetOpenFilename
    If UserForm1.fileSrc = "False" Then
        Exit Sub
    End If
    objXML.Load UserForm1.fileSrc

    'Проверка на пустой файл
    Dim entryPoint As IXMLDOMNode
        Set entryPoint = objXML
    Dim myNodes As IXMLDOMNodeList
        Set myNodes = entryPoint.SelectNodes("//PersonReply")
    If myNodes.Length = 0 Then
        errNum = 1
        GoTo ErrZone
    End If
    
    'Вывод результатов
    Application.ScreenUpdating = False
    For nNode = 0 To myNodes.Length - 1
        returnVal = GetSetFunc(1, entryPoint, nNode)
    Next nNode
    
    Exit Sub
    
    
    
    'Зона обработки ошибок (пока что только с выходом из функции)
ErrZone:
    Select Case errNum
        Case 1
            MsgBox "ДАННЫЕ ДЛЯ ВЫВОДА НЕ НАЙДЕНЫ В ВЫБРАННОМ ФАЙЛЕ!", vbCritical, "[ОШИБКА]"
    End Select

End Sub
Public Sub XmlEditSearch()
    Dim myNodes As IXMLDOMNodeList
    Dim strItem As String
        
    'Подгрузка файла XML
    Set objXML = New MSXML2.DOMDocument60
    UserForm1.fileSrc = Application.GetOpenFilename
    If UserForm1.fileSrc = "False" Then
        Exit Sub
    End If
    objXML.Load UserForm1.fileSrc
    Dim entryPoint As IXMLDOMNode
        Set entryPoint = objXML
          
    'Перебор поочередно всех персон
    Set myNodes = entryPoint.SelectNodes("//PersonReply")
    If myNodes.Length = 0 Then
        errNum = 1
        GoTo ErrZone
    End If
       
    'Заполнение ComboBox результатами
    For nNode = 0 To myNodes.Length - 1
        Set myNodes = entryPoint.SelectNodes("//PersonReply")
        Set myNode = myNodes(nNode)
        Set myChildNodes = myNode.ChildNodes
        strItem = ""
        For nChildNode = 2 To 4
            strItem = strItem + myChildNodes(nChildNode).Text + " " 'Значение столбца
        Next nChildNode
        UserForm1.ComboBox1.AddItem strItem
    Next nNode
    
    UserForm1.Show
    
    
    
    'Зона обработки ошибок (пока что только с выходом из функции)
ErrZone:
    Select Case errNum
        Case 1
            MsgBox "ДАННЫЕ ДЛЯ ВЫВОДА НЕ НАЙДЕНЫ В ВЫБРАННОМ ФАЙЛЕ!", vbCritical, "[ОШИБКА]"
    End Select

End Sub
