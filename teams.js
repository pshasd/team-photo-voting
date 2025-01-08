'Sheet1 (메인 투표 시트)
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Column = 4 Then  ' D열(투표 열)
        Application.EnableEvents = False
        
        ' 현재 선택된 수 계산
        Dim checkedCount As Integer
        checkedCount = WorksheetFunction.CountIf(Range("D2:D25"), True)
        
        ' 3개 초과 선택 방지
        If checkedCount > 3 Then
            MsgBox "최대 3팀까지만 선택 가능합니다!", vbExclamation
            Target.Value = False
            checkedCount = checkedCount - 1
        End If
        
        ' 선택 수 표시 업데이트
        Range("F2").Value = checkedCount & " / 3팀 선택됨"
        Range("F3").Value = String(checkedCount, "■") & String(3 - checkedCount, "□")
        
        ' 투표 완료 가능 여부 표시
        If checkedCount = 3 Then
            Range("F4").Value = "✓ 투표 가능"
            Range("F4").Interior.Color = RGB(198, 239, 206)
        Else
            Range("F4").Value = "3팀 선택 필요"
            Range("F4").Interior.Color = RGB(255, 235, 238)
        End If
        
        Application.EnableEvents = True
    End If
End Sub

'Sheet2 (집계 시트)
Private Sub Worksheet_Calculate()
    ' 자동 집계 업데이트
    Application.ScreenUpdating = False
    
    With Me.ChartObjects("VoteChart")
        .Activate
        .Chart.SetSourceData Source:=Range("A2:B25")
        .Chart.Axes(xlValue).MaximumScale = WorksheetFunction.Max(Range("B2:B25")) + 1
    End With
    
    Application.ScreenUpdating = True
End Sub

'ThisWorkbook
Private Sub Workbook_Open()
    ' 시트 보호 설정
    Sheets("투표").Protect Password:="", UserInterfaceOnly:=True
    Sheets("집계").Protect Password:="", UserInterfaceOnly:=True
    
    ' 날짜 체크
    If Date > DateValue("2025-01-09") Then
        MsgBox "투표가 마감되었습니다. (마감일: 2025년 1월 9일)", vbInformation
        ThisWorkbook.ReadOnly = True
    End If
End Sub
