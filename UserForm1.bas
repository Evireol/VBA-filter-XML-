VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Редактирование данных"
   ClientHeight    =   9885.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15645
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fileSrc As String
Public objXML As MSXML2.DOMDocument60
Sub CommandButton1_Click()
    Dim returnVal As Boolean
    
    'Подгрузка файла XML
    Set UserForm1.objXML = New MSXML2.DOMDocument60
    UserForm1.objXML.Load UserForm1.fileSrc
    
    Dim entryPoint As IXMLDOMNode
        Set entryPoint = UserForm1.objXML
    
    'Проверка на пустой выбор
    If ComboBox1.ListIndex = -1 Then
        errNum = 1
        GoTo ErrZone
    End If
    
    'Вывод результатов из XML
    returnVal = GetSetFunc(2, entryPoint, ComboBox1.ListIndex)
    UserForm1.Frame1.Enabled = True
    UserForm1.Frame2.Enabled = True
    UserForm1.CommandButton2.Enabled = True
    
    
    
    'Зона обработки ошибок (пока что только с выходом из функции)
ErrZone:
    Select Case errNum
        Case 1
            MsgBox "НЕ ВЫБРАН ОБЪЕКТ ДЛЯ ОТОБРАЖЕНИЯ ЕГО ДАННЫХ!", vbCritical, "[ОШИБКА]"
    End Select
End Sub
Sub CommandButton2_Click()
    
    Dim answer As Integer
    Dim returnVal As Boolean
    Dim entryPoint As IXMLDOMNode
        Set entryPoint = UserForm1.objXML
    Dim dateTime As String
    Dim diaFolder As FileDialog
    Dim selected As Boolean
    
    answer = MsgBox("Вы уверены что хотите сохранить изменения?", vbQuestion + vbYesNo + vbDefaultButton2, "Подтверждение действий")
    If answer = vbNo Then
        Exit Sub
    End If
        


'UserForm1.objXML.Save UserForm1.fileSrc

    'Запрос пути сохранения файла и бэкапа
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show

    If selected Then
        'Формирование строки 'дата_время'
        dateTime = DatePart("D", Now) & "-" & DatePart("M", Now) & "-" & DatePart("Yyyy", Now) & _
             "_" & DatePart("H", Now) & "-" & DatePart("N", Now) & "-" & DatePart("S", Now)
             
        'Сохранение бэкапа по указанному пути
        UserForm1.objXML.Save diaFolder.SelectedItems(1) & "\" & dateTime & "_XmlData_BACKUP.xml"
        
        'Запись результатов в XML
        returnVal = GetSetFunc(3, entryPoint, ComboBox1.ListIndex)
        
        'Сохранение файла по указанному пути
        UserForm1.objXML.Save diaFolder.SelectedItems(1) & "\" & dateTime & "_XmlData.xml"
        
        MsgBox "Данные успешно изменены!", vbOKOnly, "Успешно"
    End If
    Set diaFolder = Nothing
    
End Sub
