Attribute VB_Name = "Module1"
Function ZipCodeToAddress(strZipcode)
Dim objXMLHttp As Object, zipArr
'"-"ハイフンが入っていた場合は取り除く
strZipcode = Replace(strZipcode, "-", "")
Set objXMLHttp = CreateObject("MSXML2.XMLHTTP")
    objXMLHttp.Open "GET", "http://zip.cgis.biz/csv/zip.php?zn=" & strZipcode, False
    objXMLHttp.send
    
    'APIの結果を配列に代入する
    zipArr = Split(Replace(objXMLHttp.responseText, """", ""), ",")
        
        '正常な値が返ってきた場合は配列の要素数が15になる
        If UBound(zipArr) = 15 Then
            ZipCodeToAddress = zipArr(12) & zipArr(13) & zipArr(14)
        Else
            '郵便番号が間違っている場合や未入力の場合は、空文字を返す
            ZipCodeToAddress = ""
        End If
End Function
Sub Form_Click()

UserForm1.Show vbModal

'On Error GoTo ErrExit
    
    'フォームを開く
    'DoCmd.openForm "UserForm1"
    
    'Exit Sub

'ErrExit:
    'MsgBox "エラーが発生しフォームを開くことができませんでした。" & _
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

