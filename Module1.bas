Attribute VB_Name = "Module1"
Option Explicit

Public Sub Q1()

    ThisWorkbook.Worksheets("Q1").Name = "���1"

End Sub

Public Sub Q2()

    With ThisWorkbook.Worksheets("Q2").Range("A2")

        .Value = "�c�� ��"
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

'BMI�擾
Public Function getBmi(ByVal height As Double, ByVal weight As Double) As Double
    getBmi = weight / (height * height)
End Function

'BMI����
Public Function getBmiJudge(ByVal bmi As Double) As String

    Const COL_IDX_BMI As Long = 1
    Const COL_IDX_JUDGE As Long = 2

    Dim rowIdx As Long
    For rowIdx = 2 To 3
    
        '����l�����BMI�l�����ł���ꍇ
        If shBmi.Cells(rowIdx, COL_IDX_BMI).Value > bmi Then
            getBmiJudge = shBmi.Cells(rowIdx, COL_IDX_JUDGE).Value
            Exit Function
        End If
    
    Next
    
    getBmiJudge = shBmi.Cells(4, COL_IDX_JUDGE).Value

End Function

Public Sub Q5()
    
    '�����̓��t���擾
    Dim today As String
    today = Format(Date, "DD")

    Call MsgBox(Replace("������{1}���ł�", "{1}", today))

End Sub

Public Sub Q9()

    Dim rowIdx As Long
    
    With ThisWorkbook.Worksheets("Q9")
    
        While .Cells()



End Sub





