Attribute VB_Name = "omDateFunctions"
Option Compare Database
Option Explicit
'@lib 2026-10-11 sha=4a654d8c

' Leeg (Null/Empty) = niet voorbij, geeft False. Vergelijking op dagdatum (Date), geen tijdstip.
' Let op: VBA's And is niet kortsluitend zoals C#'s && — daarom de Null-check apart, niet in
' één samengestelde And-uitdrukking (zie CM-guideline implementation/vba).
Public Function IsDateInPast(dt As Variant) As Boolean
    If omStringFunctions.NotIsNullOrEmpty(dt) Then
        IsDateInPast = (Date > dt)
    Else
        IsDateInPast = False
    End If
End Function

Public Function GetMonthYearEndDate(ByVal Month As Long, ByVal year As Long) As Date
    GetMonthYearEndDate = DateAdd("m", 1, GetMonthYearStartDate(Month, year)) - 1
End Function
Public Function GetPeriodEndDate(ByVal Period As Long) As Date
    GetPeriodEndDate = GetMonthYearEndDate(Period Mod 100, Int(Period / 100))
End Function
Public Function GetMonthYearStartDate(ByVal Month As Long, ByVal year As Long) As Date
    GetMonthYearStartDate = DateSerial(year, Month, 1)
End Function
Public Function GetPeriodStartDate(ByVal Period As Long) As Date
    GetPeriodStartDate = GetMonthYearStartDate(Period Mod 100, Int(Period / 100))
End Function
Public Function GetVarPeriodStartDate(ByVal Period As Variant) As Date
    GetVarPeriodStartDate = GetPeriodStartDate(Nz(Period, 0))
End Function
Public Function GetVarPeriodEndDate(ByVal Period As Variant) As Date
    GetVarPeriodEndDate = GetPeriodEndDate(Nz(Period, 0))
End Function

Public Function ConvertDateToYYYYMMDD(ByVal dt As Date) As Long
    ConvertDateToYYYYMMDD = (year(dt) * 100 + Month(dt)) * 100 + Day(dt)
End Function
Public Function ConvertVarDateToYYYYMMDD(ByVal dt As Variant) As Long
    If IsDate(dt) Then
        ConvertVarDateToYYYYMMDD = ConvertDateToYYYYMMDD(CDate(dt))
    Else
        ConvertVarDateToYYYYMMDD = 0
    End If
End Function
Public Function ConvertDateToPeriod(ByVal dt As Date) As Long
    ConvertDateToPeriod = year(dt) * 100 + Month(dt)
End Function
Public Function ConvertVarDateToPeriod(ByVal dt As Variant, Optional NullValue As Long = 0) As Long
    If IsDate(dt) Then
        ConvertVarDateToPeriod = ConvertDateToPeriod(CDate(dt))
    Else
        ConvertVarDateToPeriod = NullValue
    End If
End Function
Public Function GetDateMonthStartDate(ByVal dt As Date)
    GetDateMonthStartDate = GetMonthYearStartDate(Month(dt), year(dt))
End Function
Public Function GetDateMonthEndDate(ByVal dt As Date)
    GetDateMonthEndDate = GetMonthYearEndDate(Month(dt), year(dt))
End Function
Public Function ConvertVarDateToDate(ByVal dt As Variant) As Date
    If IsDate(dt) Then
        ConvertVarDateToDate = CDate(dt)
    Else
        ConvertVarDateToDate = 0
    End If
End Function
Public Function NextBirthDate(ByVal dt As Variant) As Variant

    If IsDate(dt) Then
        NextBirthDate = DateSerial(year(Now), Month(dt), Day(dt))
        If NextBirthDate < Now Then
            NextBirthDate = DateSerial(year(Now) + 1, Month(dt), Day(dt))
        End If
    Else
        NextBirthDate = Null
    End If
End Function
Public Function GetTimeStamp(Optional ts As Date = 0) As String

    If ts = 0 Then ts = Now
    GetTimeStamp = format(Now, "yyyyMMdd_hhmmss")
End Function

Public Function YYYYMM_Offset(lMonthTemp As Long, lMonthOffset As Long) As Long
Dim lYear As Long
Dim lMonth As Long
Dim lCount As Long

    lYear = Int(lMonthTemp / 100)
    lMonth = lMonthTemp Mod 100
    If lMonthOffset > 0 Then
        For lCount = 1 To lMonthOffset
            lMonth = lMonth + 1
            If lMonth > 12 Then
                lMonth = 1
                lYear = lYear + 1
            End If
        Next lCount
    ElseIf lMonthOffset < 0 Then
        For lCount = 1 To Abs(lMonthOffset)
            lMonth = lMonth - 1
            If lMonth <= 0 Then
                lMonth = 12
                lYear = lYear - 1
            End If
        Next lCount
    End If
    YYYYMM_Offset = lYear * 100 + lMonth

End Function
Public Function YYYYWW(ByVal dateTemp As Variant) As Long

    If IsDate(dateTemp) Then
        YYYYWW = year(dateTemp) * 100 + format(dateTemp, "ww", vbMonday)
    Else
        YYYYWW = 0
    End If

End Function
Public Function YYYYWW_Date(ByVal lYYYYWW As Long) As Date
Dim lYear As Long
Dim lWeek As Long
Dim dateStart As Date
Dim lWeekCount As Long

    lYear = Int(lYYYYWW / 100)
    lWeek = lYYYYWW Mod 100
    If lWeek = 1 Then
        YYYYWW_Date = "01/01/" & lYear
    Else
        dateStart = ("01/01/" & lYear)
        dateStart = dateStart - Weekday(dateStart, vbMonday) + 1
        lWeekCount = 1
        While lWeekCount < lWeek
            dateStart = dateStart + 7
            lWeekCount = lWeekCount + 1
        Wend
        YYYYWW_Date = dateStart
    End If

End Function

Public Function WorkingDays(ByVal lYearMonth As Long) As Long
Dim dateWork As Date

    dateWork = "01/" & str(lYearMonth Mod 100) & "/" & str(Int(lYearMonth / 100))
    If Month(dateWork) <> (lYearMonth Mod 100) Then
        dateWork = str(lYearMonth Mod 100) & "/01/" & str(Int(lYearMonth / 100))
    End If
    While lYearMonth Mod 100 = Month(dateWork)
        If Weekday(dateWork, vbMonday) > 0 And Weekday(dateWork, vbMonday) < 6 Then
            WorkingDays = WorkingDays + 1
        End If
        dateWork = dateWork + 1
    Wend
End Function
Public Function WorkingDaysFrom(ByVal lYearMonth As Long, ByVal dateFrom As Variant) As Long
Dim dateWork As Date

    If IsDate(dateFrom) Then
        dateWork = dateFrom
        If Month(dateWork) <> (lYearMonth Mod 100) Then
            dateWork = str(lYearMonth Mod 100) & "/" & Month(dateWork) & "/" & str(Int(lYearMonth / 100))
        End If
        While lYearMonth Mod 100 = Month(dateWork)
            If Weekday(dateWork, vbMonday) > 0 And Weekday(dateWork, vbMonday) < 6 Then
                WorkingDaysFrom = WorkingDaysFrom + 1
            End If
            dateWork = dateWork + 1
        Wend
    Else
        WorkingDaysFrom = 0
    End If

End Function
Public Function WorkingDaysTill(ByVal lYearMonth As Long, ByVal dateFrom As Variant) As Long
Dim dateWork As Date

    If IsDate(dateFrom) Then
        dateWork = dateFrom
        If Month(dateWork) <> (lYearMonth Mod 100) Then
            dateWork = str(lYearMonth Mod 100) & "/" & Month(dateWork) & "/" & str(Int(lYearMonth / 100))
        End If
        While lYearMonth Mod 100 = Month(dateWork)
            If Weekday(dateWork, vbMonday) > 0 And Weekday(dateWork, vbMonday) < 6 Then
                WorkingDaysTill = WorkingDaysTill + 1
            End If
            dateWork = dateWork - 1
        Wend
    Else
        WorkingDaysTill = 0
    End If

End Function
Public Function WorkingDaysBetween(ByVal dateStart As Variant, ByVal dateEnd As Variant) As Long
Dim dateWork As Date

    If IsDate(dateStart) And IsDate(dateEnd) Then
        dateWork = dateStart
        While dateWork <= dateEnd
            If Weekday(dateWork, vbMonday) > 0 And Weekday(dateWork, vbMonday) < 6 Then
                WorkingDaysBetween = WorkingDaysBetween + 1
            End If
            dateWork = dateWork + 1
        Wend
    Else
        WorkingDaysBetween = 0
    End If
End Function
Public Function MonthsBetween(ByVal dateStart As Variant, ByVal dateEnd As Variant) As Double
    If IsDate(dateStart) And IsDate(dateEnd) Then
        MonthsBetween = (DateDiff("m", dateStart, dateEnd) + 1) * (DateDiff("d", dateStart, dateEnd) + 1) / (DateDiff("d", DateSerial(year(dateStart), Month(dateStart), 1), DateSerial(year(dateEnd), Month(dateEnd) + 1, 1) - 1) + 1)
    Else
        MonthsBetween = 0
    End If
End Function

Rem  functions to be replaced
Public Function Period(ByVal dt As Variant) As Long
    Period = ConvertVarDateToPeriod(dt)
End Function
Public Function PeriodAddMonths(ByVal dt As Variant, months As Double) As Long
    dt = DateAdd("m", months, dt)
    PeriodAddMonths = ConvertVarDateToPeriod(dt)
End Function
Public Function MonthAddMonths(ByVal dt As Variant, months As Double) As Long
    dt = DateAdd("m", months, dt)
    MonthAddMonths = Month(dt)
End Function
Public Function YearAddMonths(ByVal dt As Variant, months As Double) As Long
    dt = DateAdd("m", months, dt)
    YearAddMonths = year(dt)
End Function
Public Function DateYYYYMMDD(ByVal dt As Variant) As Long
    DateYYYYMMDD = ConvertVarDateToYYYYMMDD(CDate(dt))
End Function
Public Function StartDate(ByVal Period As Long) As Date
    StartDate = GetPeriodStartDate(Period)
End Function
Public Function UniDate(dt As Date) As Long
    UniDate = ConvertDateToYYYYMMDD(dt)
End Function
Public Function DateOnly(dt As Variant) As Date
    DateOnly = Int(ConvertVarDateToDate(dt))
End Function
Public Function YYYYMM(dt As Variant) As Long
    YYYYMM = ConvertVarDateToPeriod(dt)
End Function
Public Function IsPeriod(Period As Variant, Optional yearDelta = 100) As Boolean
    Period = Nz(Period, 0)
    IsPeriod = ((Period / 100) > year(Date) - yearDelta) And ((Period / 100) < year(Date) + yearDelta)
    IsPeriod = IsPeriod And ((Period Mod 100) > 0) And ((Period Mod 100) < 13)
End Function

Public Function AddToDate(dt As Date, Optional addYear As Long = 0, Optional addMonth As Long = 0, Optional addDay As Long = 0, Optional EndOfMonth As Boolean = False) As Date
    AddToDate = dt
    AddToDate = DateAdd("yyyy", addYear, AddToDate)
    AddToDate = DateAdd("m", addMonth, AddToDate)
    AddToDate = DateAdd("d", addDay, AddToDate)
    If EndOfMonth Then
        AddToDate = GetDateMonthEndDate(AddToDate)
    End If
End Function

Public Function IsTimeOnly(dt As Variant) As Boolean
    If Not IsDate(dt) Then
        IsTimeOnly = False
        Exit Function
    End If
    IsTimeOnly = (DateValue(dt) = 0)
End Function

Public Function ConcatenateCurrentDateIfTimeOnly(dt As Variant, Optional defaultDate As Date = 0) As Variant

    If Not IsDate(dt) Then
        ConcatenateCurrentDateIfTimeOnly = dt
        Exit Function
    End If
    ConcatenateCurrentDateIfTimeOnly = dt
    If IsTimeOnly(dt) Then
        ConcatenateCurrentDateIfTimeOnly = IIf(defaultDate = 0, Date, defaultDate) & " " & dt
    End If

End Function

Public Function NumberToDHMS(t As Double) As String

    NumberToDHMS = Int(t) & "d " & format(t - Int(t), "hh:nn:ss")
End Function

Public Function GetTimeBetween(StartDate As Date, EndDate As Date, excludeWeekends As Boolean)
Dim currentdate As Date
Dim calculatedTime As Date

    calculatedTime = EndDate - StartDate
    If excludeWeekends Then
        currentdate = DateOnly(StartDate)
        While currentdate < EndDate
            If Weekday(currentdate, vbMonday) > 5 Then
                If DateOnly(StartDate) = currentdate Then
                    calculatedTime = calculatedTime - (currentdate + 1 - StartDate)
                ElseIf DateOnly(EndDate) = currentdate Then
                    calculatedTime = calculatedTime - (EndDate - currentdate)
                Else
                    calculatedTime = calculatedTime - 1
                End If
            End If
            currentdate = currentdate + 1
        Wend
    End If
    GetTimeBetween = calculatedTime
End Function
Public Function GetTimeBetweenTimeRangesFlat(StartDate As Date, EndDate As Date, excludeWeekends As Boolean, range1StartDate As Date, range1EndDate As Date, Optional range2StartDate As Date = "0:0", Optional range2EndDate As Date = "0:0", Optional range3StartDate As Date = "0:0", Optional range3EndDate As Date = "0:0", Optional range4StartDate As Date = "0:0", Optional range4EndDate As Date = "0:0")
Dim arr() As Object
Dim tr As omTimeRange

    omDateFunctions.AddTimeRangeToArray arr, range1StartDate, range1EndDate, True, True
    omDateFunctions.AddTimeRangeToArray arr, range2StartDate, range2EndDate
    omDateFunctions.AddTimeRangeToArray arr, range3StartDate, range3EndDate
    omDateFunctions.AddTimeRangeToArray arr, range4StartDate, range4EndDate
    GetTimeBetweenTimeRangesFlat = GetTimeBetweenTimeRanges(StartDate, EndDate, excludeWeekends, arr)

End Function
Public Sub AddTimeRangeToArray(arr() As Object, StartDate As Date, EndDate As Date, Optional AllowZeros As Boolean = False, Optional Clear As Boolean = False)
Dim tr As omTimeRange
    If (StartDate <> 0 And EndDate <> 0) Or AllowZeros Then
        Set tr = New omTimeRange
        tr.StartTime = TimeValue(StartDate)
        tr.EndTime = TimeValue(EndDate)
        omArrayFunctions.ObjectArrayAdd arr, tr
    End If
End Sub
Public Sub AddTimeRangeToArray_Test()
Dim arr() As Object
    AddTimeRangeToArray arr, "10:15", "10:20"
End Sub
Public Function GetTimeBetweenTimeRanges(StartDate As Date, EndDate As Date, excludeWeekends As Boolean, timeRanges() As Object)
Dim currentdate As Date
Dim calculatedTime As Long
Dim tr As Variant
Dim calcStartSeconds As Long
Dim calcEndSeconds As Long
Dim trStartInSeconds As Long
Dim trEndInSeconds As Long

    currentdate = DateOnly(StartDate)
    While currentdate <= DateOnly(EndDate)
        If Weekday(currentdate, vbMonday) <= 5 Or Not excludeWeekends Then
            calcStartSeconds = GetTimeInSeconds("0:0")
            calcEndSeconds = GetTimeInSeconds("0:0", True)
            If DateOnly(StartDate) = currentdate Then
                calcStartSeconds = GetTimeInSeconds(TimeValue(StartDate))
            End If
            If DateOnly(EndDate) = currentdate Then
                calcEndSeconds = GetTimeInSeconds(TimeValue(EndDate), IIf(TimeValue(EndDate) <> 0, True, False))
            End If
            For Each tr In timeRanges
                If TypeName(tr) = "omTimeRange" Then
                    trStartInSeconds = tr.GetStartTimeInSeconds()
                    trEndInSeconds = tr.GetEndTimeInSeconds()
                    If calcStartSeconds < trEndInSeconds And calcEndSeconds > trStartInSeconds Then
                        calculatedTime = calculatedTime + (IIf(calcEndSeconds > trEndInSeconds, trEndInSeconds, calcEndSeconds) - IIf(calcStartSeconds > trStartInSeconds, calcStartSeconds, trStartInSeconds))
                    End If
                End If
            Next
        End If
        currentdate = currentdate + 1
    Wend
    GetTimeBetweenTimeRanges = DateAdd("s", calculatedTime, "0:0")

End Function

Public Sub GetTimeBetweenTimeRanges_Test()
Dim arr(1) As Object 'omTimeRange
Dim tr As omTimeRange
    Set tr = New omTimeRange
    tr.StartTime = "9:00"
    tr.EndTime = "12:00"

    Set arr(0) = tr
    Debug.Print GetTimeBetweenTimeRanges("2018-10-16 9:00", "2018-10-17 11:45", False, arr)
End Sub

Public Sub GetTimeBetweenTimeRangesFlat_Test()
    Debug.Print GetTimeBetweenTimeRangesFlat("2018-10-16 9:00", "2018-10-17 11:45", True, "9:0", "12:0")
End Sub

Public Function GetTimeInSeconds(t As Date, Optional ZeroIsDay As Boolean = False) As Long
    If t = "0:0" And ZeroIsDay Then
        GetTimeInSeconds = CLng(1) * 24 * 60 * 60
    Else
        GetTimeInSeconds = DateDiff("s", "0:0", t)
    End If
End Function

Public Function IsWeekDay(dt As Date, Optional FirstDayOfWeek As VbDayOfWeek = VbDayOfWeek.vbMonday) As Boolean
    IsWeekDay = (Weekday(dt, FirstDayOfWeek) <= 5)
End Function

Public Function GetQuarter(dt As Date, Optional FirstDayOfWeek As VbDayOfWeek = VbDayOfWeek.vbMonday, Optional FirstWeekOfYear As VbFirstWeekOfYear = VbFirstWeekOfYear.vbFirstFourDays) As Long
    GetQuarter = DatePart("q", dt, FirstDayOfWeek, FirstWeekOfYear)
End Function
Public Function GetWeekNumber(dt As Date, Optional FirstDayOfWeek As VbDayOfWeek = VbDayOfWeek.vbMonday, Optional FirstWeekOfYear As VbFirstWeekOfYear = VbFirstWeekOfYear.vbFirstFourDays) As Long
    GetWeekNumber = DatePart("ww", dt, FirstDayOfWeek, FirstWeekOfYear)
End Function
Public Function GetWeekDay(dt As Date, Optional FirstDayOfWeek As VbDayOfWeek = VbDayOfWeek.vbMonday, Optional FirstWeekOfYear As VbFirstWeekOfYear = VbFirstWeekOfYear.vbFirstFourDays) As Long
    GetWeekDay = DatePart("w", dt, FirstDayOfWeek, FirstWeekOfYear)
End Function

' The table Calendar (one row per day) belongs to PopulateCalendar: this creates it when there is no table with that name.
' A table that exists already, also a linked one (for example to SQL Server), is checked: it must have exactly the fields
' Date, Year, Month, Day, Period, WeekDay, WeekNumber, Quarter, IsWeekend, IsHoliday, Description, each of a type that can hold the value
' (a whole number field may be wider than needed, for example Long for Year). When it does not match, error vbObjectError + 1 says what is wrong.
Public Sub CreateCalendar()
Dim db As DAO.Database
Dim td As DAO.TableDef
Dim problems As String

    Set db = CurrentDb
    On Error Resume Next
    Set td = db.TableDefs("Calendar")
    On Error GoTo 0
    If td Is Nothing Then
        db.Execute "CREATE TABLE [Calendar] ([Date] DATETIME CONSTRAINT PK_Calendar PRIMARY KEY, [Year] SHORT, [Month] BYTE, [Day] BYTE, [Period] LONG, " & _
                   "[WeekDay] BYTE, [WeekNumber] BYTE, [Quarter] BYTE, [IsWeekend] YESNO, [IsHoliday] YESNO, [Description] TEXT(255))"
        Exit Sub
    End If
    problems = CalendarLayoutProblems(td)
    If Len(problems) > 0 Then Err.Raise vbObjectError + 1, "CreateCalendar", "The table Calendar does not have the layout that PopulateCalendar needs: " & problems
End Sub

' What is different between the table and the layout of the calendar, as text; "" when it matches.
Private Function CalendarLayoutProblems(td As DAO.TableDef) As String
Dim names As Variant
Dim kinds As Variant
Dim i As Long
Dim fld As DAO.Field
Dim found As Boolean
Dim rank As Long
Dim needed As Long
Dim result As String

    names = Array("Date", "Year", "Month", "Day", "Period", "WeekDay", "WeekNumber", "Quarter", "IsWeekend", "IsHoliday", "Description")
    kinds = Array("D", "2", "1", "1", "3", "1", "1", "1", "B", "B", "T") 'D = date, 1/2/3 = whole number of at least byte/integer/long, B = yes/no, T = text
    For i = 0 To UBound(names)
        found = False
        For Each fld In td.Fields
            If StrComp(fld.Name, names(i), vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next fld
        If Not found Then
            result = result & "field " & names(i) & " is missing; "
        Else
            Select Case kinds(i)
                Case "D"
                    If fld.Type <> dbDate Then result = result & names(i) & " must be a date; "
                Case "B"
                    If fld.Type <> dbBoolean Then result = result & names(i) & " must be yes/no; "
                Case "T"
                    If fld.Type <> dbText And fld.Type <> dbMemo Then result = result & names(i) & " must be text; "
                Case Else
                    Select Case fld.Type
                        Case dbByte: rank = 1
                        Case dbInteger: rank = 2
                        Case dbLong: rank = 3
                        Case dbBigInt: rank = 4
                        Case Else: rank = 0
                    End Select
                    needed = CLng(kinds(i))
                    If rank < needed Then result = result & names(i) & " must be a whole number" & IIf(needed > 1, " of at least " & IIf(needed = 2, "integer", "long") & " size", "") & "; "
            End Select
        End If
    Next i
    For Each fld In td.Fields
        found = False
        For i = 0 To UBound(names)
            If StrComp(fld.Name, names(i), vbTextCompare) = 0 Then found = True
        Next i
        If Not found Then result = result & "field " & fld.Name & " is not part of the calendar; "
    Next fld
    CalendarLayoutProblems = result
End Function

' Adds the days that are missing in the table Calendar: from the day after the last one (or 1 January 2017) up to the end of next year.
' Creates the table first when it does not exist (see CreateCalendar).
Public Sub PopulateCalendar()
Dim rs As New ADODB.Recordset
Dim minDate As Date
Dim maxDate As Date

    CreateCalendar
    minDate = Nz(DMax("Date", "Calendar"), DateSerial(2017, 1, 1) - 1) + 1
    maxDate = DateSerial(year(Now) + 2, 1, 1) - 1
    rs.Open "Calendar", CurrentProject.connection, adOpenForwardOnly, adLockOptimistic
    While minDate <= maxDate
        rs.AddNew
        rs("Date") = minDate
        rs("Year") = year(minDate)
        rs("Month") = Month(minDate)
        rs("Day") = Day(minDate)
        rs("Period") = year(minDate) * 100 + Month(minDate)
        rs("WeekDay") = GetWeekDay(minDate)
        rs("WeekNumber") = GetWeekNumber(minDate)
        rs("Quarter") = GetQuarter(minDate)
        rs("IsWeekend") = Not IsWeekDay(minDate)
        rs("IsHoliday") = False
        rs.Update
        minDate = minDate + 1
    Wend
    rs.Close
    Set rs = Nothing

End Sub

Public Function DaysBetween(StartDate As Date, EndDate As Date) As Long
    DaysBetween = DateDiff("d", StartDate, EndDate)
End Function

' Number of whole years between two dates (the age on EndDate of something that started on StartDate).
Public Function YearsBetween(StartDate As Date, EndDate As Date) As Long
    YearsBetween = DateDiff("yyyy", StartDate, EndDate)
    If DateSerial(Year(EndDate), Month(StartDate), Day(StartDate)) > EndDate Then
        YearsBetween = YearsBetween - 1
    End If
End Function

Public Function EOMonth(dt As Date) As Date
    EOMonth = DateSerial(year(dt), Month(dt) + 1, 0)
End Function

Public Function CalculatePayDate(dt As Date, Days As Integer, EndOfMonth As Boolean, ExtraDays As Integer) As Date

    dt = DateAdd("d", Days, dt)
    If EndOfMonth Then
        dt = DateSerial(year(dt), Month(dt), 1)
        dt = DateAdd("m", 1, dt)
        dt = DateAdd("d", -1, dt)
    End If
    dt = DateAdd("d", ExtraDays, dt)
    CalculatePayDate = dt
End Function

Public Function GetFiscalYear(ByVal d As Date, _
        Optional ByVal startMonth As Integer = 1, _
        Optional ByVal offset As Integer = 0) As Integer
    Dim baseYear As Integer
    baseYear = year(d)
    If Month(d) < startMonth Then baseYear = baseYear - 1
    GetFiscalYear = baseYear + offset
End Function
