VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "受付フォーム"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ComboBox1_Enter()
 Me!ComboBox1.DropDown
End Sub
Private Sub ComboBox2_Enter()
 Me!ComboBox2.DropDown
End Sub
Private Sub ComboBox3_Enter()
 Me!ComboBox3.DropDown
End Sub
Private Sub ComboBox4_Enter()
 Me!ComboBox4.DropDown
End Sub

Private Sub CommandButton1_Click()
Dim endCell As Integer
'年齢計算
'Dim 生年月日 As Variant
'Dim 年齢 As Variant
'back:
'生年月日 = InputBox("生年月日を入力してください。yyyy/mm/dd形式")
'If IsDate(生年月日) Then
'年齢 = DateDiff("yyyy", 生年月日, Date) + (Format(生年月日, "mmdd") > Format(Date, "mmdd"))
'MsgBox 年齢 & "歳です。"
'Else
'MsgBox "入力形式がちがいます。"
'GoTo back
'End If


endCell = Worksheets("sheet2").Cells(Rows.Count, 1).End(xlUp).Row + 1
Debug.Print endCell
    With Worksheets("sheet2")
        .Cells(endCell, 1).Value = TextBox1.Value
        .Cells(endCell, 2).Value = TextBox2.Text
        .Cells(endCell, 3).Value = ComboBox2.Text & "/" & ComboBox3.Text & "/" & ComboBox4.Text
        .Cells(endCell, 4).Value = ComboBox1.Value
        .Cells(endCell, 5).Value = TextBox4.Value
        .Cells(endCell, 6).Value = TextBox5.Text
        .Cells(endCell, 7).Value = TextBox6.Text
        .Cells(endCell, 8).Value = TextBox7.Text
    End With
    SetUp
End Sub

Private Sub CommandButton2_Click()
'変数宣言
Dim Postalnum As Variant


'郵便番号を変数に格納
Postalnum = TextBox4.Text


' 郵便番号のテキストボックスに3桁以上の郵便番号があり、住所がブランクの場合のみ住所を変換させる
If ((Len(Postalnum) >= 3) And (TextBox5.Value = "")) Then

'変数を渡して、結果をテキストボックスに入れる
TextBox5.Text = ZipCodeToAddress(Postalnum)


End If

End Sub

Private Sub CommandButton3_Click()
    Unload UserForm1
End Sub

Private Sub TextBox1_Change()
    TextBox1.Value = Format(TextBox1.Value, "00000")
    If Len(TextBox1.Text) > 5 Then
        ''すでに5文字目が入力されている
        TextBox1.Value = Left(TextBox1.Value, 5)
    End If
End Sub
Private Sub TextBox3_Click()
    .IMEMode = fmIMEModeAlpha
End Sub
Private Sub TextBox4_Change()
    If Len(TextBox4.Text) > 7 Then
        ''すでに7文字目が入力されている
        TextBox4.Value = Left(TextBox4.Value, 7)
    End If
End Sub

Private Sub TextBox5_Enter()
  With Me!TextBox5
    '選択長さの位置にカーソル移動
    .SelStart = .SelLength
  End With
End Sub

Private Sub UserForm_Initialize()
    With ComboBox1
        .AddItem "男"
        .AddItem "女"
        .AddItem "その他"
        .AddItem "無回答"
    End With
    ComboBox1.ListIndex = 3
    
    Dim i As Integer
    '年のコンボボックス　去年から10年間
    For i = Year(Date) - 100 To Year(Date) Step 1
        ComboBox2.AddItem i
    Next
    '初期値は現在の年
    ComboBox2.ListIndex = 100
    
    '月のコンボボックス　１２ヶ月
    For i = 1 To 12
    If i < 10 Then
        ComboBox3.AddItem 0 & i
    Else
        ComboBox3.AddItem i
    End If
    Next
    '初期値は現在の月
    ComboBox3.ListIndex = Month(Date) - 1
    
    '日のコンボボックス　31日
    For i = 1 To 31
        ComboBox4.AddItem i
    Next
    '初期値は現在の年
    ComboBox4.ListIndex = Day(Date) - 1
    
    Dim endCell As Integer
    endCell = Worksheets("sheet2").Cells(Rows.Count, 1).End(xlUp).Row
    Debug.Print endCell
    TextBox1.Value = Format(Cells(endCell, 1).Value + 1, "00000")
    
    ComboBox1.Style = fmStyleDropDownList
    ComboBox2.Style = fmStyleDropDownList
    ComboBox3.Style = fmStyleDropDownList
    ComboBox4.Style = fmStyleDropDownList
    'TextBox1.Value
    'TextBox2.Text = ""
    'TextBox4.Value = ""
    'TextBox5.Text = ""
    'TextBox6.Text = ""
    'TextBox7.Text = ""
        
    TextBox2.SetFocus

    
End Sub
Public Sub SetUp()
    ComboBox1.ListIndex = 3
    '初期値は現在の年
    ComboBox2.ListIndex = 100

    '初期値は現在の月
    ComboBox3.ListIndex = Month(Date) - 1
    
    '初期値は現在の年
    ComboBox4.ListIndex = Day(Date) - 1
    
    Dim endCell As Integer
    endCell = Worksheets("sheet2").Cells(Rows.Count, 1).End(xlUp).Row
    'Debug.Print endCell
    TextBox1.Value = Format(Cells(endCell, 1).Value + 1, "00000")

    TextBox2.Text = ""
    TextBox4.Text = ""
    TextBox5.Text = ""
    TextBox6.Text = ""
    TextBox7.Text = ""
    
    TextBox2.SetFocus

End Sub
