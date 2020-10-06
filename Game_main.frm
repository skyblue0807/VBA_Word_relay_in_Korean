VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Game_main 
   Caption         =   "게임: "
   ClientHeight    =   3000
   ClientLeft      =   110
   ClientTop       =   460
   ClientWidth     =   4880
   OleObjectBlob   =   "Game_main.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "Game_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public username As String


Private Sub UserForm_Initialize()
'유저폼이 열릴 때 / 제목에 사용자 이름 표기, 첫 단어 랜덤으로 제공

Dim MAIN As Object

Set MAIN = ThisWorkbook.Worksheets("MAIN")

username = MAIN.Cells(2, 2)
Game_main.Caption = "게임: " & username & "님"
Game_main.Problem_txt.Value = Initial_word

End Sub


Private Sub CB_Answer_Click()
'사용자가 답변 입력 시

Dim user_answer As String, Problem As String, next_problem As String
Dim DB As Object

Problem = Problem_txt.Value

username = ActiveSheet.Cells(2, 2)
Set DB = ThisWorkbook.Worksheets("Word_DB")
user_answer = Txt_answer.Value

'정답 확인 -> 프로시저로 독립 예정
If InStr(user_answer, Right(Problem, 1)) <> 1 Then
    MsgBox ("정답이 틀립니다!")
    Exit Sub
End If

'정답을 시트에 등록하기
Call register_word_DB(user_answer)

'사용자가 준 정답에 대한 다음 답안 찾아내기
next_problem = find_next_problem(Right(user_answer, 1))

'Game_main에 다음 문제 표기하기
'Call Update_Game_main -> 필요시 독립
Call register_word_DB(next_problem)
Game_main.Problem_txt.Value = next_problem

End Sub

Private Sub CommandButton2_Click()

Unload Me

End Sub


