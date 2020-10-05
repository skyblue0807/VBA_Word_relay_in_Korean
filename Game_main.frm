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

Public username1 As String


Private Sub CommandButton1_Click()
'comment test for git
'add one line

Dim user_answer As String, Problem As String, next_problem As String
Dim DB As Object

Problem = Problem_txt.Value

username1 = Cells(2, 2)
Set DB = ThisWorkbook.Worksheets(username1)
user_answer = TextBox1.Value

'정답 확인 -> 프로시저로 독립
If InStr(user_answer, Right(Problem, 1)) <> 1 Then
    MsgBox ("정답이 틀립니다!")
    Exit Sub
End If

'정답을 시트에 등록하기
Call register_word_DB(user_answer, username1)

'사용자가 준 정답에 대한 다음 답안 찾아내기
next_problem = find_next_problem(Right(user_answer, 1), username1)

'Game_main에 다음 문제 표기하기
'Call Update_Game_main -> 필요시 독립
Call register_word_DB(next_problem, username1)
Game_main.Problem_txt.Value = next_problem

End Sub

Private Sub CommandButton2_Click()

Unload Me

End Sub


Private Sub UserForm_Initialize()

Dim MAIN As Object

Set MAIN = ThisWorkbook.Worksheets("MAIN")

username1 = MAIN.Cells(2, 2)
Game_main.Caption = "게임: " & username1 & "님"
Game_main.Problem_txt.Value = Initial_word(username1)

End Sub
