VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "��t�t�H�[��"
   ClientHeight    =   7020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
'�N��v�Z
'Dim ���N���� As Variant
'Dim �N�� As Variant
'back:
'���N���� = InputBox("���N��������͂��Ă��������Byyyy/mm/dd�`��")
'If IsDate(���N����) Then
'�N�� = DateDiff("yyyy", ���N����, Date) + (Format(���N����, "mmdd") > Format(Date, "mmdd"))
'MsgBox �N�� & "�΂ł��B"
'Else
'MsgBox "���͌`�����������܂��B"
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
'�ϐ��錾
Dim Postalnum As Variant


'�X�֔ԍ���ϐ��Ɋi�[
Postalnum = TextBox4.Text


' �X�֔ԍ��̃e�L�X�g�{�b�N�X��3���ȏ�̗X�֔ԍ�������A�Z�����u�����N�̏ꍇ�̂ݏZ����ϊ�������
If ((Len(Postalnum) >= 3) And (TextBox5.Value = "")) Then

'�ϐ���n���āA���ʂ��e�L�X�g�{�b�N�X�ɓ����
TextBox5.Text = ZipCodeToAddress(Postalnum)


End If

End Sub

Private Sub CommandButton3_Click()
    Unload UserForm1
End Sub

Private Sub TextBox1_Change()
    TextBox1.Value = Format(TextBox1.Value, "00000")
    If Len(TextBox1.Text) > 5 Then
        ''���ł�5�����ڂ����͂���Ă���
        TextBox1.Value = Left(TextBox1.Value, 5)
    End If
End Sub
Private Sub TextBox3_Click()
    .IMEMode = fmIMEModeAlpha
End Sub
Private Sub TextBox4_Change()
    If Len(TextBox4.Text) > 7 Then
        ''���ł�7�����ڂ����͂���Ă���
        TextBox4.Value = Left(TextBox4.Value, 7)
    End If
End Sub

Private Sub TextBox5_Enter()
  With Me!TextBox5
    '�I�𒷂��̈ʒu�ɃJ�[�\���ړ�
    .SelStart = .SelLength
  End With
End Sub

Private Sub UserForm_Initialize()
    With ComboBox1
        .AddItem "�j"
        .AddItem "��"
        .AddItem "���̑�"
        .AddItem "����"
    End With
    ComboBox1.ListIndex = 3
    
    Dim i As Integer
    '�N�̃R���{�{�b�N�X�@���N����10�N��
    For i = Year(Date) - 100 To Year(Date) Step 1
        ComboBox2.AddItem i
    Next
    '�����l�͌��݂̔N
    ComboBox2.ListIndex = 100
    
    '���̃R���{�{�b�N�X�@�P�Q����
    For i = 1 To 12
    If i < 10 Then
        ComboBox3.AddItem 0 & i
    Else
        ComboBox3.AddItem i
    End If
    Next
    '�����l�͌��݂̌�
    ComboBox3.ListIndex = Month(Date) - 1
    
    '���̃R���{�{�b�N�X�@31��
    For i = 1 To 31
        ComboBox4.AddItem i
    Next
    '�����l�͌��݂̔N
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
    '�����l�͌��݂̔N
    ComboBox2.ListIndex = 100

    '�����l�͌��݂̌�
    ComboBox3.ListIndex = Month(Date) - 1
    
    '�����l�͌��݂̔N
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
