Attribute VB_Name = "Module1"
Option Explicit

Public Sub Q1()

    ThisWorkbook.Worksheets("Q1").Name = "問題1"

End Sub

Public Sub Q2()

    With ThisWorkbook.Worksheets("Q2").Range("A2")

        .Value = "田村 純"
        .Font.Bold = True

    End With

End Sub

Public Sub Q3()

    With ThisWorkbook.Worksheets("Q3")
    
        .Range("B3:D7").Borders.LineStyle = True
        .Range("B3:D3").Bold = True
        .Range("B3:D3").Interior.Color = 15189684
    
    End With

End Sub

'BMI取得
Public Function getBmi(ByVal height As Double, ByVal weight As Double) As Double
    getBmi = weight / (height * height)
End Function

'BMI判定
Public Function getBmiJudge(ByVal bmi As Double) As String

    Const COL_IDX_BMI As Long = 1
    Const COL_IDX_JUDGE As Long = 2

    Dim rowIdx As Long
    For rowIdx = 2 To 3
    
        '判定値が基準のBMI値未満である場合
        If shBmi.Cells(rowIdx, COL_IDX_BMI).Value > bmi Then
            getBmiJudge = shBmi.Cells(rowIdx, COL_IDX_JUDGE).Value
            Exit Function
        End If
    
    Next
    
    getBmiJudge = shBmi.Cells(4, COL_IDX_JUDGE).Value

End Function

Public Sub Q5()
    
    '今日の日付を取得
    Dim today As String
    today = Format(Date, "DD")

    Call MsgBox(Replace("今日は{1}日です", "{1}", today))

End Sub

Public Sub Q9()

    Dim rowIdx As Long
    
    With ThisWorkbook.Worksheets("Q9")
    
        While .Cells()



End Sub





