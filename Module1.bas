Attribute VB_Name = "Module1"
Function ZipCodeToAddress(strZipcode)
Dim objXMLHttp As Object, zipArr
'"-"�n�C�t���������Ă����ꍇ�͎�菜��
strZipcode = Replace(strZipcode, "-", "")
Set objXMLHttp = CreateObject("MSXML2.XMLHTTP")
    objXMLHttp.Open "GET", "http://zip.cgis.biz/csv/zip.php?zn=" & strZipcode, False
    objXMLHttp.send
    
    'API�̌��ʂ�z��ɑ������
    zipArr = Split(Replace(objXMLHttp.responseText, """", ""), ",")
        
        '����Ȓl���Ԃ��Ă����ꍇ�͔z��̗v�f����15�ɂȂ�
        If UBound(zipArr) = 15 Then
            ZipCodeToAddress = zipArr(12) & zipArr(13) & zipArr(14)
        Else
            '�X�֔ԍ����Ԉ���Ă���ꍇ�▢���͂̏ꍇ�́A�󕶎���Ԃ�
            ZipCodeToAddress = ""
        End If
End Function
Sub Form_Click()

UserForm1.Show vbModal

'On Error GoTo ErrExit
    
    '�t�H�[�����J��
    'DoCmd.openForm "UserForm1"
    
    'Exit Sub

'ErrExit:
    'MsgBox "�G���[���������t�H�[�����J�����Ƃ��ł��܂���ł����B" & _
       ' vbCrLf & Err.Description
End Sub
Sub Check()
    Dim i As Long
    Dim j As Long
    j = Worksheets("sheet2").Cells(Rows.Count, 1).End(xlUp).Row + 1
    For i = 1 To j
        If WorksheetFunction.CountIf(Range("A1:A1000"), Cells(i, 1)) > 1 Then
            Cells(i, 1).Font.ColorIndex = 3
        End If
    Next i
End Sub

