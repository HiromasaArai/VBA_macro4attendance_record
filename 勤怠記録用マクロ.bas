Option Explicit

' Read Me
' シート名: active
' B3: No
' C3: 日付
' D3: 曜日
' E3: 時刻
' F3: 標題
' G3: 内容
' H3: チケット
' I3: 補足
' J3: 時間


Private Function getFormat化日付(arg As Variant) As String
    getFormat化日付 = Format(arg, "YYYYMMDD")
End Function

Private Function isTerm4DateKbn(dateKbn As String) As Boolean
    Dim isTerm1, isTerm2, isTerm3 As Boolean

    isTerm1 = dateKbn = "休日"
    isTerm2 = dateKbn = "祝日"
    isTerm3 = dateKbn = "有休"

    isTerm4DateKbn = isTerm1 Or isTerm2 Or isTerm3
End Function

Private Function kintaiCreateFunc(mugic_num As String) As Double
    Dim ws       As Worksheet: Set ws = Worksheets("active")
    Dim start_rg As Range: Set start_rg = ws.Range("C3")

    Dim to_month As String: to_month = Format(Date, "YYYYMM")
    Dim my_date  As String: my_date = to_month + mugic_num
    Dim my_sum   As Double: my_sum = 0

    Do While start_rg.Value <> ""
        If getFormat化日付(start_rg.Value) = my_date Then
            my_sum = my_sum + start_rg.Offset(0, 7).Value
        End If
        Set start_rg = start_rg.Offset(1, 0)
    Loop
    kintaiCreateFunc = my_sum
End Function

Private Function getDateStatus(mugic_num As String) As String
    Dim ws       As Worksheet: Set ws = Worksheets("active")
    Dim start_rg As Range: Set start_rg = ws.Range("C3")

    Dim to_month As String: to_month = Format(Date, "YYYYMM")
    Dim my_date  As String: my_date = to_month + mugic_num
    Dim ans_str  As String: ans_str = ""
    Dim Format化日付 As String
    Dim date_kbn As String
    Dim isFormat化日付の判定 As Boolean

    Do While start_rg.Value <> ""
        Format化日付 = getFormat化日付(start_rg.Value)
        isFormat化日付の判定 = Format化日付 = my_date
        date_kbn = start_rg.Offset(0, 3).Value
        If isFormat化日付の判定 And isTerm4DateKbn(date_kbn) Then
            ans_str = date_kbn
            Exit Do
        End If
        Set start_rg = start_rg.Offset(1, 0)
    Loop
    getDateStatus = ans_str
End Function

Sub Mcr001_重複する文字の色を変えて数字を合わせる()
    Call lastDateCreate
    Call toDayCreate
    Call loopNumChange
    Call loopColorChange(Range("C3"))
    Call loopColorChange(Range("D3"))
    Call loopColorChange(Range("H3"))
    Call loopColorChange(Range("I3"))
End Sub

Private Sub toDayCreate()
    Dim todayRg          As Range: Set todayRg = Nothing
    Dim myRg             As Range: Set myRg = Range("C3")

    Dim isNotDoubleToday As Boolean: isNotDoubleToday = True
    Dim today            As Date: today = Date
    Dim i As Long
    Dim myTime As Double
    Dim grayOutColor As Variant: grayOutColor = RGB(220, 220, 220)
    
    Do While myRg.Value <> ""
        Set myRg = myRg.Offset(1, 0)
        If myRg.Value = today Then
            If todayRg Is Nothing Then Set todayRg = myRg

            If myRg.Value = myRg.Offset(-1, 0).Value Then isNotDoubleToday = False
        End If
    Loop

    If isNotDoubleToday Then
        With todayRg
            .Offset(0, 2).Value = 9
            .Offset(0, 2).NumberFormatLocal = "0.00"
            .Offset(0, 7).Value = 0.5
            .Offset(0, 7).NumberFormatLocal = "0.00"
            Range(.Offset(0, 3), .Offset(0, 6)).ShrinkToFit = True
        End With

        myTime = 17
        For i = 0 To 16
            With todayRg
                .Offset(1, 0).EntireRow.Insert
                .Offset(1, 0).Value = today
                .Offset(1, 1).Value = .Offset(0, 1).Value
                .Offset(1, 2).Value = myTime
                .Offset(1, 2).NumberFormatLocal = "0.00"
                Range(.Offset(1, 3), .Offset(1, 6)).ShrinkToFit = True

                If myTime = 12 Or myTime = 12.5 Then
                    .Offset(1, 3).Value = "休憩"
                    .Offset(1, 4).Value = "休憩"
                    Range(.Offset(1, 2), .Offset(1, 7)).Interior.Color = grayOutColor
                Else
                    .Offset(1, 7).Value = 0.5
                    .Offset(1, 7).NumberFormatLocal = "0.00"
                End If
            End With
            myTime = myTime - 0.5
        Next i
    End If
End Sub

Private Sub loopColorChange(myRg As Range)
    Dim beforeRg As Range: Set beforeRg = Nothing
    Dim chkRg    As Range: Set chkRg = Range(Cells(myRg.Row, 2).Address)
    Dim isTerm1, isTerm2 As Boolean

    Do While chkRg.Value <> ""
        If Not (beforeRg Is Nothing) Then
            isTerm1 = myRg.Value = beforeRg.Value
            isTerm2 = myRg.Value <> ""
            If isTerm1 And isTerm2 Then
                myRg.Font.ThemeColor = xlThemeColorDark2
            Else
                myRg.Font.ThemeColor = 2
            End If
        End If
        Set beforeRg = myRg
        Set myRg = myRg.Offset(1, 0)
        Set chkRg = Range(Cells(myRg.Row, 2).Address)
    Loop

    Set beforeRg = Nothing
    Set chkRg = Nothing
End Sub

Private Function isTerm4loopNumChange(myRg As Range) As Boolean
    Dim isTerm1, isTerm2, isTerm3, isTerm4 As Boolean

    isTerm1 = myRg.Offset(0, 4).Value <> ""
    isTerm2 = myRg.Offset(0, 1).Value <> ""
    isTerm3 = myRg.Offset(1, 0).Value <> ""
    isTerm4 = myRg.Offset(0, 3).Value <> ""

    isTerm4loopNumChange = isTerm1 Or isTerm2 Or isTerm3 Or isTerm4
End Function

Private Sub loopNumChange()
    Dim myRg     As Range: Set myRg = Range("B3")
    Dim beforeRg As Range: Set beforeRg = Nothing
    
    Do While isTerm4loopNumChange(myRg)
        If Not (beforeRg Is Nothing) Then
            myRg.Value = beforeRg.Value + 1
        End If
        Set beforeRg = myRg
        Set myRg = myRg.Offset(1, 0)
    Loop
    Set beforeRg = Nothing
    Set myRg = Nothing
End Sub

Private Function getDayOfWeek(workDay As Date) As String
    getDayOfWeek = Format(workDay, "aaa")
End Function

Private Function isTerm4Holiday(workDay As Date) As Boolean
    Dim dayOfWeek As String: dayOfWeek = getDayOfWeek(workDay)
    Dim isTerm1, isTerm2 As Boolean

    isTerm1 = dayOfWeek = "土"
    isTerm2 = dayOfWeek = "日"
    isTerm4Holiday = isTerm1 Or isTerm2
End Function

Private Sub lastDateCreate()
    Dim lastDayRg As Range: Set lastDayRg = Range("C2").End(xlDown)
    Dim myRg      As Range: Set myRg = Nothing

    Dim lastDayFromToday As Date: lastDayFromToday = Date + 30
    Dim workDay          As Date: workDay = lastDayRg.Value
    Dim i                As Integer
    Dim offsetRow        As Integer
    Dim dateDiff         As Integer: dateDiff = lastDayFromToday - workDay
    Dim grayOutColor     As Variant: grayOutColor = RGB(220, 220, 220)

    For i = 0 To dateDiff
        workDay = workDay + 1
        offsetRow = i + 1
        With lastDayRg
            .Offset(offsetRow, 0).Value = workDay
            .Offset(offsetRow, 1).Value = getDayOfWeek(workDay)
            Set myRg = Range(.Offset(offsetRow, -1), .Offset(offsetRow, 7))
            If isTerm4Holiday(workDay) Then
                .Offset(offsetRow, 3).Value = "休日"
                myRg.Interior.Color = grayOutColor
            End If
        End With
        myRg.Borders.LineStyle = xlContinuous
        
    Next i
End Sub

