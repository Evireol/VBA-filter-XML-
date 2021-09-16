VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�������������� ������"
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
    
    '��������� ����� XML
    Set UserForm1.objXML = New MSXML2.DOMDocument60
    UserForm1.objXML.Load UserForm1.fileSrc
    
    Dim entryPoint As IXMLDOMNode
        Set entryPoint = UserForm1.objXML
    
    '�������� �� ������ �����
    If ComboBox1.ListIndex = -1 Then
        errNum = 1
        GoTo ErrZone
    End If
    
    '����� ����������� �� XML
    returnVal = GetSetFunc(2, entryPoint, ComboBox1.ListIndex)
    UserForm1.Frame1.Enabled = True
    UserForm1.Frame2.Enabled = True
    UserForm1.CommandButton2.Enabled = True
    
    
    
    '���� ��������� ������ (���� ��� ������ � ������� �� �������)
ErrZone:
    Select Case errNum
        Case 1
            MsgBox "�� ������ ������ ��� ����������� ��� ������!", vbCritical, "[������]"
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
    
    answer = MsgBox("�� ������� ��� ������ ��������� ���������?", vbQuestion + vbYesNo + vbDefaultButton2, "������������� ��������")
    If answer = vbNo Then
        Exit Sub
    End If
        


'UserForm1.objXML.Save UserForm1.fileSrc

    '������ ���� ���������� ����� � ������
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    selected = diaFolder.Show

    If selected Then
        '������������ ������ '����_�����'
        dateTime = DatePart("D", Now) & "-" & DatePart("M", Now) & "-" & DatePart("Yyyy", Now) & _
             "_" & DatePart("H", Now) & "-" & DatePart("N", Now) & "-" & DatePart("S", Now)
             
        '���������� ������ �� ���������� ����
        UserForm1.objXML.Save diaFolder.SelectedItems(1) & "\" & dateTime & "_XmlData_BACKUP.xml"
        
        '������ ����������� � XML
        returnVal = GetSetFunc(3, entryPoint, ComboBox1.ListIndex)
        
        '���������� ����� �� ���������� ����
        UserForm1.objXML.Save diaFolder.SelectedItems(1) & "\" & dateTime & "_XmlData.xml"
        
        MsgBox "������ ������� ��������!", vbOKOnly, "�������"
    End If
    Set diaFolder = Nothing
    
End Sub
