Attribute VB_Name = "Game_code"
Option Explicit

Public username2 As String

Sub Check_username()

    Cells(2, 2) = InputBox("이름을 입력해주세요")

End Sub


Function Initial_word() As String

Dim DB As Object
Dim r_end As Long
Dim random_index As Long
Dim word_DB() As Variant

Set DB = ThisWorkbook.Worksheets("Word_DB")
r_end = DB.Cells(Rows.Count, 1).End(xlUp).Row

ReDim word_DB(1 To (r_end - 6))

word_DB = Range(DB.Cells(7, 1), DB.Cells(r_end, 1))
Initial_word = word_DB(Application.RandBetween(1, (r_end - 6)), 1)

End Function


Sub register_word_DB(answer As String)

Dim DB As Object
Dim r_end As Long

Set DB = ThisWorkbook.Worksheets("Word_DB")
r_end = DB.Cells(Rows.Count, 3).End(xlUp).Row

DB.Cells(r_end + 1, 3) = answer

End Sub




Function find_next_problem(initial_c As String) As String

Dim DB As Object
Dim r_end As Long
Dim word_DB() As Variant
Dim word As Variant
Dim cand_no As Long, cand() As String

Set DB = ThisWorkbook.Worksheets("Word_DB")
r_end = DB.Cells(Rows.Count, 1).End(xlUp).Row

ReDim word_DB(1 To (r_end - 6))
word_DB = Range(DB.Cells(7, 1), DB.Cells(r_end, 1))

For Each word In word_DB
    If Left(word, 1) = initial_c Then
        cand_no = cand_no + 1
        ReDim Preserve cand(1 To cand_no)
        cand(cand_no) = word
    End If
Next word

If cand_no < 1 Then
    MsgBox "승리하셨습니다!"
    Call Update_word_DB
Else
    find_next_problem = cand(1)
End If

End Function


Sub Update_word_DB()

Dim DB As Object
Dim word_DB As New Collection
Dim new_words As Range, old_words As Range, word As Range
Dim word_s As Variant, i As Long

Set DB = ThisWorkbook.Worksheets("Word_DB")
If DB.Cells(Rows.Count, 3).End(xlUp).Row < 2 Then: Exit Sub

Set old_words = Range(DB.Cells(2, 1), DB.Cells(DB.Cells(Rows.Count, 1).End(xlUp).Row, 3))
Set new_words = Range(DB.Cells(2, 3), DB.Cells(DB.Cells(Rows.Count, 3).End(xlUp).Row, 3))

'지금은 old-new 순서대로 collection에 등록시키지만, 이후엔 사용자가 입력한 단어들만 골라서 중복체크 후 입력할 수 있게 할 것
On Error Resume Next
    For Each word In old_words
        word_DB.Add Item:=word.Value, Key:=word.Value
    Next
    
    For Each word In new_words
        word_DB.Add Item:=word.Value, Key:=word.Value
    Next
On Error GoTo 0

'word_db의 단어들을 다시 배운 단어로 넣어놓음
i = 7
For Each word_s In word_DB
    DB.Cells(i, 1) = word_s
    i = i + 1
Next

new_words.Value = vbNullString

End Sub

