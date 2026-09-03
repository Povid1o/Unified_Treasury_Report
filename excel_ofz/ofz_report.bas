Attribute VB_Name = "modOFZ"
Option Explicit

'=============================================================================
' Отчёт «Ставки ОФЗ» целиком внутри Excel.
' Порт reports/ofz_rates/etl.py на VBA: CBonds JSON API -> формат BI.
'
' Запросы уходят из процесса Excel (WinINET, как у обычного веб-запроса Excel),
' поэтому проходят там, где корпоративный файрвол гасит python/requests.
'
' ТОЛЬКО ДЛЯ EXCEL ПОД WINDOWS: в Excel для macOS нет COM, а на нём здесь
' держится и HTTP, и разбор ответа. Собирать книгу можно где угодно.
'
' Точки входа (назначаются на кнопки листа «Параметры»):
'   RunOfzReport   — сформировать отчёт
'   ExportOfzXlsx  — выгрузить лист «Отчёт» отдельным .xlsx для BI
'   ExportOfzCsv   — то же в CSV (UTF-8 BOM), как у Python-версии
'   TestConnection — проверить, что API отвечает (диагностика файрвола)
'=============================================================================

Private Const BASE_URL As String = "https://ws.cbonds.info/services/json"
Private Const TIMEOUT_SEC As Long = 30

' Глубина архива get_index_value_new у демо-подписки CBonds — 100 календарных дней
' (см. CBonds_API/README.md, раздел «Ограничения»).
Private Const MAX_SPAN_DAYS As Long = 100

Private Const SH_PARAMS As String = "Параметры"
Private Const SH_REPORT As String = "Отчёт"
Private Const SH_LOG As String = "Лог"
Private Const SH_INDEX As String = "Индексы"
Private Const SH_GAPS As String = "Пропуски"

Private Const GROUP_YIELD As String = "Доходность ОФЗ"
Private Const TYPE_VAL As String = "Значение"

Private mLogRow As Long
Private mTransport As String   ' какой HTTP-объект реально сработал

' Позиция разбора JSON (см. раздел 8). Все переменные уровня модуля обязаны
' стоять здесь, в блоке объявлений: объявление после первой процедуры даёт
' ошибку компиляции «Only comments may appear after End Sub».
Private mJson As String
Private mPos As Long

' Запрос к API и разбор ответа держатся на COM-объектах Windows
' (MSXML2/WinHttp, ADODB.Stream, Scripting.Dictionary). В Excel для macOS COM
' нет вообще, поэтому там книга может только храниться и открываться.
Private Sub RequireWindows()
#If Mac Then
    Err.Raise vbObjectError + 590, , _
        "Эта книга работает только в Excel для Windows. Запрос к CBonds и разбор ответа " & _
        "используют компоненты Windows (MSXML2/WinHttp, ADODB, Scripting), которых в Excel " & _
        "для macOS нет. На Mac запускайте Python-версию: python console.py ofz-rates"
#End If
End Sub

'=============================================================================
' 1. ТОЧКИ ВХОДА
'=============================================================================

Public Sub RunOfzReport()
    Dim t0 As Single: t0 = Timer
    Dim rowsDict As Object
    Dim tenors() As String, i As Long
    Dim rangeFrom As Date, rangeTo As Date
    Dim wanted As Object            ' Nothing либо словарь запрошенных дат (режим 3)
    Dim login As String, pwd As String
    Dim limitN As Long, total As Long, spanDays As Long
    Dim idxMap As Object, indexKey As String, typeId As String
    Dim got As Long

    On Error GoTo Fail
    RequireWindows
    Application.ScreenUpdating = False
    Application.StatusBar = "Отчёт «Ставки ОФЗ»: подготовка..."

    LogReset
    SetStatusCell "Выполняется...", False

    ' -- учётные данные ------------------------------------------------------
    login = ParamStr("p_login")
    pwd = ParamStr("p_password")
    If Len(login) = 0 Then Err.Raise vbObjectError + 601, , "Не заполнен логин CBonds (лист «" & SH_PARAMS & "»)."
    If Len(pwd) = 0 Then
        pwd = InputBox("Пароль CBonds (не сохраняется в книге):", "CBonds")
        If Len(pwd) = 0 Then Err.Raise vbObjectError + 602, , "Пароль не введён — запуск отменён."
    End If

    ' -- период --------------------------------------------------------------
    ResolvePeriod rangeFrom, rangeTo, wanted

    spanDays = DateDiff("d", rangeFrom, rangeTo)
    If spanDays > MAX_SPAN_DAYS Then
        LogWrite "WARN", "Запрошенный период (" & spanDays & " дн.) превышает документированную " & _
            "глубину архива CBonds для get_index_value_new (" & MAX_SPAN_DAYS & " календарных дней) — " & _
            "часть дат может не вернуться."
    End If
    LogWrite "INFO", "Формирование отчёта «Ставки ОФЗ» за период " & Iso(rangeFrom) & ".." & Iso(rangeTo)

    ' -- показатели ----------------------------------------------------------
    tenors = SplitList(ParamStr("p_tenors"))
    If UBound(tenors) < 0 Then Err.Raise vbObjectError + 603, , "Не задан ни один срок кривой доходности."
    limitN = ParamLong("p_limit", 200)
    Set idxMap = LoadIndexMap()

    Set rowsDict = CreateObject("Scripting.Dictionary")

    For i = 0 To UBound(tenors)
        Application.StatusBar = "Ставки ОФЗ: запрос " & tenors(i) & " (" & (i + 1) & " из " & (UBound(tenors) + 1) & ")..."
        indexKey = "RUB_Yield_Curve_" & tenors(i)

        If Not idxMap.Exists(indexKey) Then
            LogWrite "WARN", "Пропускаю срок " & tenors(i) & ": индекс " & indexKey & " отсутствует на листе «" & _
                SH_INDEX & "» и в списке индексов, доступных аккаунту (get_index_types). " & _
                "Проверенный эндпоинт: get_index_value_new."
        Else
            typeId = idxMap(indexKey)
            got = FetchIndexValues(login, pwd, typeId, tenors(i), Iso(rangeFrom), Iso(rangeTo), limitN, rowsDict)
            If got = 0 Then
                LogWrite "WARN", "Нет данных по сроку " & tenors(i) & " за период " & Iso(rangeFrom) & ".." & Iso(rangeTo)
            Else
                LogWrite "INFO", "Доходность ОФЗ " & tenors(i) & ": получено " & got & " значений"
            End If
        End If
    Next i

    LogKnownGaps

    If rowsDict.Count = 0 Then
        Err.Raise vbObjectError + 604, , "Итоговый набор данных пуст — ни одного показателя не удалось получить. " & _
            "Подробности на листе «" & SH_LOG & "»."
    End If

    ' -- режим 3: оставляем только запрошенные даты ---------------------------
    If Not wanted Is Nothing Then FilterByWantedDates rowsDict, wanted

    ' -- запись результата ----------------------------------------------------
    Application.StatusBar = "Ставки ОФЗ: запись результата..."
    total = WriteReport(rowsDict)

    LogWrite "INFO", "Итоговая таблица: " & total & " строк, " & CountDistinctTenors(rowsDict) & " показателей"
    LogWrite "INFO", "Готово за " & Format$(Timer - t0, "0.0") & " с (транспорт: " & mTransport & ")"

    SetStatusCell "Готово: " & total & " строк, " & Format$(Now, "dd.mm.yyyy hh:nn:ss"), False
    Application.StatusBar = False
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets(SH_REPORT).Activate
    MsgBox "Отчёт сформирован: " & total & " строк на листе «" & SH_REPORT & "».", vbInformation, "Ставки ОФЗ"
    Exit Sub

Fail:
    ' Текст ошибки забираем первым же действием: любой оператор On Error
    ' (а он есть внутри LogWrite и SetStatusCell) обнуляет объект Err.
    Dim errMsg As String
    errMsg = Err.Description

    Application.StatusBar = False
    Application.ScreenUpdating = True
    LogWrite "ERROR", errMsg
    SetStatusCell "Ошибка: " & errMsg, True
    MsgBox "Не удалось сформировать отчёт." & vbCrLf & vbCrLf & errMsg & vbCrLf & vbCrLf & _
        "Подробности — на листе «" & SH_LOG & "».", vbCritical, "Ставки ОФЗ"
End Sub

Public Sub ExportOfzCsv()
    Dim ws As Worksheet, lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim st As Object, rowText As String, outPath As String

    On Error GoTo Fail
    RequireWindows
    Set ws = ThisWorkbook.Worksheets(SH_REPORT)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 6
    If lastRow < 2 Then Err.Raise vbObjectError + 610, , "Лист «" & SH_REPORT & "» пуст — сначала сформируйте отчёт."

    outPath = ParamStr("p_csv_path")
    If Len(outPath) = 0 Then outPath = ThisWorkbook.Path & Application.PathSeparator & "ofz_report.csv"

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                     ' adTypeText
    st.Charset = "utf-8"            ' ADODB пишет UTF-8 с BOM — как to_csv(encoding="utf-8-sig")
    st.Open

    For r = 1 To lastRow
        rowText = ""
        For c = 1 To lastCol
            If c > 1 Then rowText = rowText & ","
            rowText = rowText & CsvField(ws.Cells(r, c))
        Next c
        st.WriteText rowText & vbCrLf
    Next r

    st.SaveToFile outPath, 2        ' adSaveCreateOverWrite
    st.Close

    LogWrite "INFO", "Отчёт сохранён: " & outPath
    MsgBox "Сохранено " & (lastRow - 1) & " строк:" & vbCrLf & outPath, vbInformation, "Ставки ОФЗ"
    Exit Sub

Fail:
    MsgBox "Не удалось выгрузить CSV." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Ставки ОФЗ"
End Sub

' Отдельный .xlsx с единственным листом-выгрузкой: то, что скармливается BI.
' Рабочую книгу отдавать в BI не стоит — в ней лежат логин и служебные листы.
Public Sub ExportOfzXlsx()
    Dim ws As Worksheet, wbOut As Workbook, lastRow As Long, outPath As String
    Dim prevAlerts As Boolean

    On Error GoTo Fail
    Set ws = ThisWorkbook.Worksheets(SH_REPORT)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 615, , "Лист «" & SH_REPORT & "» пуст — сначала сформируйте отчёт."

    outPath = ParamStr("p_xlsx_path")
    If Len(outPath) = 0 Then outPath = ThisWorkbook.Path & Application.PathSeparator & "ofz_report.xlsx"

    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ws.Copy                              ' Copy без аргументов -> новая книга из одного листа
    Set wbOut = ActiveWorkbook
    wbOut.Worksheets(1).Name = "ofz_report"
    ' 51 = xlOpenXMLWorkbook (.xlsx, без макросов)
    wbOut.SaveAs outPath, 51
    wbOut.Close False

    Application.DisplayAlerts = prevAlerts
    Application.ScreenUpdating = True

    LogWrite "INFO", "Выгрузка для BI сохранена: " & outPath
    MsgBox "Сохранено " & (lastRow - 1) & " строк:" & vbCrLf & outPath, vbInformation, "Ставки ОФЗ"
    Exit Sub

Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Не удалось выгрузить XLSX." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Ставки ОФЗ"
End Sub

Public Sub TestConnection()
    Dim login As String, pwd As String, body As String, resp As String
    Dim status As Long, root As Object, items As Collection

    On Error GoTo Fail
    RequireWindows
    LogReset

    login = ParamStr("p_login")
    pwd = ParamStr("p_password")
    If Len(login) = 0 Then Err.Raise vbObjectError + 620, , "Не заполнен логин CBonds."
    If Len(pwd) = 0 Then pwd = InputBox("Пароль CBonds (не сохраняется в книге):", "CBonds")
    If Len(pwd) = 0 Then Err.Raise vbObjectError + 621, , "Пароль не введён."

    ' Минимальный запрос: одно значение ключевой ставки ЦБ.
    body = BuildPayload(login, pwd, "21755", "", "", 1)
    resp = HttpPostJson(BASE_URL & "/get_index_value_new/?lang=rus", body, status)
    Set root = JsonParse(resp)
    RaiseIfApiError root
    Set items = GetItems(root)

    LogWrite "INFO", "Соединение с CBonds API установлено. HTTP " & status & ", транспорт: " & mTransport & _
        ", записей в ответе: " & items.Count
    MsgBox "Связь с CBonds API есть." & vbCrLf & vbCrLf & _
        "HTTP-статус: " & status & vbCrLf & _
        "Транспорт: " & mTransport & vbCrLf & _
        "Записей в тестовом ответе: " & items.Count, vbInformation, "Проверка подключения"
    Exit Sub

Fail:
    Dim errMsg As String
    errMsg = Err.Description                 ' LogWrite ниже обнулит Err (см. RunOfzReport)
    LogWrite "ERROR", errMsg
    MsgBox "Связи с CBonds API нет." & vbCrLf & vbCrLf & errMsg & vbCrLf & vbCrLf & _
        "Если ошибка транспортная (таймаут, отказ соединения, ошибка сертификата) — " & _
        "запрос гасит прокси/файрвол. Проверьте, что для узла ws.cbonds.info разрешён HTTPS.", _
        vbCritical, "Проверка подключения"
End Sub

'=============================================================================
' 2. ПАРАМЕТРЫ И ПЕРИОД
'=============================================================================

Private Function ParamStr(ByVal nm As String) As String
    Dim v As Variant
    On Error GoTo Fail
    v = ThisWorkbook.Names(nm).RefersToRange.Value2
    If IsError(v) Then ParamStr = "": Exit Function
    If IsEmpty(v) Then ParamStr = "": Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbSingle Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        ParamStr = Trim$(Str$(v))
    Else
        ParamStr = Trim$(CStr(v))
    End If
    Exit Function
Fail:
    ParamStr = ""
End Function

Private Function ParamLong(ByVal nm As String, ByVal dflt As Long) As Long
    Dim s As String
    s = ParamStr(nm)
    If Len(s) = 0 Then ParamLong = dflt Else ParamLong = CLng(Val(s))
    If ParamLong <= 0 Then ParamLong = dflt
End Function

Private Sub SetStatusCell(ByVal msg As String, ByVal isError As Boolean)
    On Error Resume Next
    With ThisWorkbook.Names("p_status").RefersToRange
        .Value = msg
        .Font.Color = IIf(isError, RGB(176, 0, 32), RGB(0, 110, 60))
    End With
End Sub

' Три взаимоисключающих режима периода — приоритет как в build_report() (Python):
'   1) одна дата + глубина истории назад   2) интервал дат   3) список конкретных дат
Private Sub ResolvePeriod(ByRef rangeFrom As Date, ByRef rangeTo As Date, ByRef wanted As Object)
    Dim mode As String, s As String
    Dim parts() As String, i As Long, d As Date
    Dim minD As Date, maxD As Date
    Dim asOf As Date, lookback As Long

    Set wanted = Nothing
    mode = Left$(ParamStr("p_mode"), 1)

    Select Case mode
        Case "3"
            s = ParamStr("p_dates")
            If Len(s) = 0 Then Err.Raise vbObjectError + 630, , "Режим 3: не заполнен список дат (p_dates)."
            parts = SplitList(s)
            Set wanted = CreateObject("Scripting.Dictionary")
            For i = 0 To UBound(parts)
                d = ParseIsoDate(parts(i))
                If Not wanted.Exists(Iso(d)) Then wanted.Add Iso(d), True
                If i = 0 Then
                    minD = d: maxD = d
                Else
                    If d < minD Then minD = d
                    If d > maxD Then maxD = d
                End If
            Next i
            If wanted.Count = 0 Then Err.Raise vbObjectError + 631, , "Режим 3: список дат пуст."
            ' У CBonds нет метода «дай данные только за эти даты» — запрашиваем
            ' диапазон [min; max] и фильтруем результат.
            rangeFrom = minD: rangeTo = maxD

        Case "2"
            Dim sFrom As String, sTo As String
            sFrom = ParamStr("p_date_from"): sTo = ParamStr("p_date_to")
            If Len(sFrom) = 0 Or Len(sTo) = 0 Then
                Err.Raise vbObjectError + 632, , "Режим 2: нужно указать обе границы интервала (p_date_from и p_date_to)."
            End If
            rangeFrom = ParseIsoDate(sFrom)
            rangeTo = ParseIsoDate(sTo)
            If rangeFrom > rangeTo Then
                Err.Raise vbObjectError + 633, , "Начало интервала (" & Iso(rangeFrom) & ") позже конца (" & Iso(rangeTo) & ")."
            End If

        Case Else
            s = ParamStr("p_date")
            If Len(s) = 0 Then asOf = Date Else asOf = ParseIsoDate(s)
            lookback = ParamLong("p_lookback", 90)
            rangeFrom = asOf - lookback
            rangeTo = asOf
    End Select
End Sub

Private Sub FilterByWantedDates(ByRef rowsDict As Object, ByVal wanted As Object)
    Dim k As Variant, row As Variant
    Dim keep As Object, seen As Object, missing As String
    Set keep = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")

    For Each k In rowsDict.Keys
        row = rowsDict(k)
        If wanted.Exists(row(0)) Then
            keep.Add k, row
            If Not seen.Exists(row(0)) Then seen.Add row(0), True
        End If
    Next k

    For Each k In wanted.Keys
        If Not seen.Exists(k) Then missing = missing & IIf(Len(missing) > 0, ", ", "") & k
    Next k
    If Len(missing) > 0 Then
        LogWrite "WARN", "Не найдено данных на даты: " & missing & " (выходной/нет торгов?)"
    End If

    If keep.Count = 0 Then
        Err.Raise vbObjectError + 634, , "Ни на одну из запрошенных дат данных не найдено."
    End If
    Set rowsDict = keep
End Sub

'=============================================================================
' 3. ЗАПРОС К API
'=============================================================================

' Забирает значения одного индекса и складывает валидные строки в rowsDict.
' Ключ словаря — "группа|показатель|дата": дубли по (date_, name_group, name_st)
' схлопываются с сохранением последнего значения, как drop_duplicates(keep="last").
Private Function FetchIndexValues(ByVal login As String, ByVal pwd As String, _
                                  ByVal typeId As String, ByVal tenor As String, _
                                  ByVal dateFrom As String, ByVal dateTo As String, _
                                  ByVal limitN As Long, ByRef rowsDict As Object) As Long
    Dim body As String, resp As String, status As Long
    Dim root As Object, items As Collection, it As Variant
    Dim d As String, valRaw As Variant, v As Double
    Dim key As String, n As Long

    On Error GoTo Fail

    body = BuildPayload(login, pwd, typeId, dateFrom, dateTo, limitN)
    resp = HttpPostJson(BASE_URL & "/get_index_value_new/?lang=rus", body, status)
    Set root = JsonParse(resp)
    RaiseIfApiError root
    Set items = GetItems(root)

    For Each it In items
        d = NormalizeDate(DictGet(it, "date"))
        valRaw = DictGet(it, "value")

        If Len(d) = 0 Then
            LogWrite "WARN", "Срок " & tenor & ": запись без даты пропущена."
        ElseIf Not IsIsoDate(d) Then
            Err.Raise vbObjectError + 640, , "Срок " & tenor & ": дата не в формате YYYY-MM-DD: '" & d & "'"
        ElseIf IsEmpty(valRaw) Or IsNull(valRaw) Or (VarType(valRaw) = vbString And Len(Trim$(CStr(valRaw))) = 0) Then
            LogWrite "WARN", "Срок " & tenor & ", дата " & d & ": пустое значение пропущено."
        Else
            If VarType(valRaw) = vbString Then
                If Not IsNumericStr(CStr(valRaw)) Then
                    Err.Raise vbObjectError + 641, , "Срок " & tenor & ", дата " & d & _
                        ": значение '" & CStr(valRaw) & "' не приводится к числу."
                End If
                v = Val(Replace$(CStr(valRaw), ",", "."))
            Else
                v = CDbl(valRaw)
            End If

            key = GROUP_YIELD & "|" & tenor & "|" & d
            rowsDict(key) = Array(d, GROUP_YIELD, tenor, "%", TYPE_VAL, Round(v, 2))
            n = n + 1
        End If
    Next it

    FetchIndexValues = n
    Exit Function

Fail:
    ' Как в Python: ошибка по одному сроку не валит весь отчёт.
    LogWrite "ERROR", "Ошибка запроса кривой доходности ОФЗ (" & tenor & "): " & Err.Description
    FetchIndexValues = 0
End Function

Private Function BuildPayload(ByVal login As String, ByVal pwd As String, ByVal typeId As String, _
                              ByVal dateFrom As String, ByVal dateTo As String, ByVal limitN As Long) As String
    Dim f As String
    f = "{""field"":""type_id"",""operator"":""eq"",""value"":""" & JsonEsc(typeId) & """}"
    If Len(dateFrom) > 0 Then f = f & ",{""field"":""date"",""operator"":""ge"",""value"":""" & dateFrom & """}"
    If Len(dateTo) > 0 Then f = f & ",{""field"":""date"",""operator"":""le"",""value"":""" & dateTo & """}"

    BuildPayload = "{""auth"":{""login"":""" & JsonEsc(login) & """,""password"":""" & JsonEsc(pwd) & """}," & _
                   """filters"":[" & f & "]," & _
                   """quantity"":{""limit"":" & limitN & ",""offset"":0}," & _
                   """sorting"":[{""field"":""date"",""order"":""desc""}]," & _
                   """fields"":[]}"
End Function

Private Sub RaiseIfApiError(ByVal root As Object)
    Dim e As Variant, msg As String
    If root Is Nothing Then Exit Sub
    If Not HasKey(root, "error") Then Exit Sub

    If IsObject(root("error")) Then
        Set e = root("error")
        If HasKey(e, "err_str") Then msg = CStr(e("err_str")) Else msg = "неизвестная ошибка API"
    Else
        msg = CStr(root("error"))
    End If
    Err.Raise vbObjectError + 650, , "API error: " & msg
End Sub

Private Function GetItems(ByVal root As Object) As Collection
    If HasKey(root, "items") Then
        If IsObject(root("items")) Then
            Set GetItems = root("items")
            Exit Function
        End If
    End If
    Set GetItems = New Collection
End Function

' root — Scripting.Dictionary у нормального ответа, но при неожиданной форме
' ответа (массив, скаляр) метода Exists может не быть: не падаем.
Private Function HasKey(ByVal obj As Object, ByVal key As String) As Boolean
    On Error GoTo Fail
    If obj Is Nothing Then Exit Function
    HasKey = obj.Exists(key)
    Exit Function
Fail:
    HasKey = False
End Function

'=============================================================================
' 4. HTTP
'=============================================================================

' Пробует транспорты по очереди. Первым — MSXML2.XMLHTTP: он ходит через
' WinINET и подхватывает системные настройки прокси/аутентификации, то есть
' ведёт себя так же, как встроенные веб-запросы Excel. ServerXMLHTTP и
' WinHttpRequest используют стек WinHTTP (отдельная конфигурация прокси) —
' оставлены как запасные.
Private Function HttpPostJson(ByVal url As String, ByVal body As String, ByRef statusOut As Long) As String
    Dim progIds As Variant, i As Long
    Dim http As Object, lastErr As String
    Dim payload() As Byte, respBytes() As Byte

    progIds = Array("MSXML2.XMLHTTP.6.0", "MSXML2.ServerXMLHTTP.6.0", "WinHttp.WinHttpRequest.5.1", "MSXML2.XMLHTTP")
    payload = StringToUtf8Bytes(body)

    For i = LBound(progIds) To UBound(progIds)
        Set http = Nothing
        On Error Resume Next
        Set http = CreateObject(CStr(progIds(i)))
        On Error GoTo 0
        If Not http Is Nothing Then
            On Error Resume Next
            Err.Clear
            If InStr(1, CStr(progIds(i)), "ServerXMLHTTP") > 0 Then
                http.setTimeouts TIMEOUT_SEC * 1000, TIMEOUT_SEC * 1000, TIMEOUT_SEC * 1000, TIMEOUT_SEC * 1000
            ElseIf InStr(1, CStr(progIds(i)), "WinHttpRequest") > 0 Then
                http.SetTimeouts TIMEOUT_SEC * 1000, TIMEOUT_SEC * 1000, TIMEOUT_SEC * 1000, TIMEOUT_SEC * 1000
            End If
            http.Open "POST", url, False
            http.setRequestHeader "Content-Type", "application/json; charset=utf-8"
            http.setRequestHeader "Accept", "application/json"
            http.send payload

            If Err.Number = 0 Then
                statusOut = CLng(http.Status)
                respBytes = http.responseBody
                On Error GoTo 0
                mTransport = CStr(progIds(i))
                RaiseIfHttpError statusOut
                HttpPostJson = Utf8BytesToString(respBytes)
                Exit Function
            Else
                lastErr = CStr(progIds(i)) & ": " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        Else
            lastErr = CStr(progIds(i)) & ": объект недоступен на этой машине"
        End If
    Next i

    Err.Raise vbObjectError + 660, , "Не удалось выполнить HTTP-запрос ни одним из транспортов. " & _
        "Последняя ошибка — " & lastErr & ". Похоже на блокировку прокси/файрволом узла ws.cbonds.info."
End Function

Private Sub RaiseIfHttpError(ByVal status As Long)
    Select Case status
        Case 200
            Exit Sub
        Case 301
            Err.Raise vbObjectError + 671, , "HTTP 301: используется HTTP вместо HTTPS"
        Case 403
            Err.Raise vbObjectError + 672, , "HTTP 403: неверный логин/пароль или нет доступа к операции"
        Case 500
            Err.Raise vbObjectError + 673, , "HTTP 500: синтаксическая ошибка в запросе или нет доступа к операции"
        Case 504
            Err.Raise vbObjectError + 674, , "HTTP 504: таймаут сервера"
        Case Else
            Err.Raise vbObjectError + 675, , "HTTP " & status & ": неожиданный ответ сервера"
    End Select
End Sub

Private Function StringToUtf8Bytes(ByVal s As String) As Byte()
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1                      ' adTypeBinary
    st.Position = 3                  ' пропускаем BOM, который ADODB дописывает в начало
    StringToUtf8Bytes = st.Read
    st.Close
End Function

Private Function Utf8BytesToString(ByRef b() As Byte) As String
    Dim st As Object
    On Error GoTo EmptyResult
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open
    st.Write b
    st.Position = 0
    st.Type = 2: st.Charset = "utf-8"
    Utf8BytesToString = st.ReadText
    st.Close
    Exit Function
EmptyResult:
    Utf8BytesToString = ""
End Function

'=============================================================================
' 5. ЗАПИСЬ РЕЗУЛЬТАТА
'=============================================================================

Private Function WriteReport(ByVal rowsDict As Object) As Long
    Dim ws As Worksheet
    Dim keys() As String, i As Long, n As Long
    Dim outArr() As Variant, row As Variant
    Dim k As Variant, lastRow As Long

    n = rowsDict.Count
    ReDim keys(0 To n - 1)
    i = 0
    For Each k In rowsDict.Keys
        keys(i) = CStr(k): i = i + 1
    Next k

    ' Ключ собран как "группа|показатель|дата", поэтому обычная двоичная
    ' сортировка строк даёт порядок sort_values(["name_group","name_st","date_"]).
    SortStrings keys, 0, n - 1

    ReDim outArr(1 To n, 1 To 6)
    For i = 0 To n - 1
        row = rowsDict(keys(i))
        outArr(i + 1, 1) = row(0)
        outArr(i + 1, 2) = row(1)
        outArr(i + 1, 3) = row(2)
        outArr(i + 1, 4) = row(3)
        outArr(i + 1, 5) = row(4)
        outArr(i + 1, 6) = row(5)
    Next i

    Set ws = ThisWorkbook.Worksheets(SH_REPORT)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).ClearContents

    ' date_ обязан остаться текстом "YYYY-MM-DD": с General Excel превратит его
    ' в дату и в BI уедет локальный формат вместо ISO.
    ws.Range("A2").Resize(n, 1).NumberFormat = "@"
    ws.Range("A2").Resize(n, 6).Value = outArr
    WriteReport = n
End Function

Private Sub SortStrings(ByRef a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long
    Dim pivot As String, tmp As String
    If lo >= hi Then Exit Sub
    i = lo: j = hi
    pivot = a((lo + hi) \ 2)
    Do While i <= j
        Do While StrComp(a(i), pivot, vbBinaryCompare) < 0: i = i + 1: Loop
        Do While StrComp(a(j), pivot, vbBinaryCompare) > 0: j = j - 1: Loop
        If i <= j Then
            tmp = a(i): a(i) = a(j): a(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortStrings a, lo, j
    If i < hi Then SortStrings a, i, hi
End Sub

Private Function CountDistinctTenors(ByVal rowsDict As Object) As Long
    Dim k As Variant, row As Variant, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    For Each k In rowsDict.Keys
        row = rowsDict(k)
        If Not d.Exists(row(2)) Then d.Add row(2), True
    Next k
    CountDistinctTenors = d.Count
End Function

Private Function CsvField(ByVal cell As Range) As String
    Dim v As Variant, s As String
    v = cell.Value2
    If IsEmpty(v) Then
        CsvField = ""
        Exit Function
    End If
    If VarType(v) = vbDouble Or VarType(v) = vbSingle Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        s = Trim$(Str$(v))           ' Str$ не зависит от локали: всегда точка
    Else
        s = CStr(v)
    End If
    If InStr(s, """") > 0 Or InStr(s, ",") > 0 Or InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then
        s = """" & Replace$(s, """", """""") & """"
    End If
    CsvField = s
End Function

'=============================================================================
' 6. СПРАВОЧНИКИ И ЛОГ
'=============================================================================

Private Function LoadIndexMap() As Object
    Dim ws As Worksheet, r As Long, lastRow As Long
    Dim d As Object, k As String, v As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Worksheets(SH_INDEX)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        k = Trim$(CStr(ws.Cells(r, 1).Value))
        v = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(k) > 0 And Len(v) > 0 Then
            If Not d.Exists(k) Then d.Add k, v
        End If
    Next r
    Set LoadIndexMap = d
End Function

' Показатели, которые не отдаёт ни один доступный аккаунту эндпоинт
' (лист «Пропуски»). Аналог log_known_gaps() в Python.
Private Sub LogKnownGaps()
    Dim ws As Worksheet, r As Long, lastRow As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_GAPS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 Then
            LogWrite "WARN", "Показатель не получен: группа='" & ws.Cells(r, 1).Value & "', показатель='" & _
                ws.Cells(r, 2).Value & "'. Причина: " & ws.Cells(r, 3).Value
        End If
    Next r
End Sub

Private Sub LogReset()
    Dim ws As Worksheet, lastRow As Long
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Clear
    mLogRow = 1
    mTransport = "—"
End Sub

Private Sub LogWrite(ByVal level As String, ByVal msg As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG)
    If ws Is Nothing Then Exit Sub
    If mLogRow < 1 Then mLogRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    mLogRow = mLogRow + 1
    ws.Cells(mLogRow, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(mLogRow, 2).Value = level
    ws.Cells(mLogRow, 3).Value = msg
    Select Case level
        Case "ERROR": ws.Cells(mLogRow, 2).Font.Color = RGB(176, 0, 32)
        Case "WARN": ws.Cells(mLogRow, 2).Font.Color = RGB(176, 110, 0)
        Case Else: ws.Cells(mLogRow, 2).Font.Color = RGB(90, 90, 90)
    End Select
End Sub

'=============================================================================
' 7. ДАТЫ И СТРОКИ
'=============================================================================

Private Function Iso(ByVal d As Date) As String
    Iso = Format$(d, "yyyy-mm-dd")
End Function

' Принимает как текст "YYYY-MM-DD", так и настоящую дату Excel (если ячейка
' отформатирована как дата и пришёл серийный номер).
Private Function ParseIsoDate(ByVal s As String) As Date
    Dim parts() As String
    s = Trim$(s)
    If IsNumericStr(s) And InStr(s, "-") = 0 Then
        If Val(s) > 20000 Then
            ParseIsoDate = CDate(CDbl(Val(s)))
            Exit Function
        End If
    End If
    parts = Split(s, "-")
    If UBound(parts) <> 2 Then
        Err.Raise vbObjectError + 680, , "Некорректный формат даты '" & s & "', ожидается YYYY-MM-DD"
    End If
    On Error GoTo Bad
    ParseIsoDate = DateSerial(CInt(parts(0)), CInt(parts(1)), CInt(parts(2)))
    Exit Function
Bad:
    Err.Raise vbObjectError + 681, , "Некорректный формат даты '" & s & "', ожидается YYYY-MM-DD"
End Function

' Значения date у CBonds приходят как "YYYY-MM-DD" либо "YYYY-MM-DD HH:MM:SS".
Private Function NormalizeDate(ByVal v As Variant) As String
    Dim s As String
    If IsEmpty(v) Or IsNull(v) Then NormalizeDate = "": Exit Function
    s = Trim$(CStr(v))
    If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
    If InStr(s, "T") > 0 Then s = Left$(s, InStr(s, "T") - 1)
    NormalizeDate = s
End Function

Private Function IsIsoDate(ByVal s As String) As Boolean
    Dim i As Long, c As String
    If Len(s) <> 10 Then Exit Function
    For i = 1 To 10
        c = Mid$(s, i, 1)
        If i = 5 Or i = 8 Then
            If c <> "-" Then Exit Function
        ElseIf c < "0" Or c > "9" Then
            Exit Function
        End If
    Next i
    IsIsoDate = True
End Function

Private Function IsNumericStr(ByVal s As String) As Boolean
    Dim i As Long, c As String, seenDigit As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c >= "0" And c <= "9" Then
            seenDigit = True
        ElseIf c = "." Or c = "," Or c = "-" Or c = "+" Or c = "e" Or c = "E" Then
            ' допустимые символы числа
        Else
            Exit Function
        End If
    Next i
    IsNumericStr = seenDigit
End Function

Private Function SplitList(ByVal s As String) As String()
    Dim parts() As String, out() As String, i As Long, n As Long, t As String
    s = Replace$(Replace$(Replace$(s, ";", ","), vbLf, ","), vbCr, ",")
    parts = Split(s, ",")
    ReDim out(0 To UBound(parts))
    n = -1
    For i = 0 To UBound(parts)
        t = Trim$(parts(i))
        If Len(t) > 0 Then
            n = n + 1
            out(n) = t
        End If
    Next i
    If n < 0 Then
        SplitList = Split("", ",")      ' Split("") даёт пустой массив: LBound=0, UBound=-1
    Else
        ReDim Preserve out(0 To n)
        SplitList = out
    End If
End Function

'=============================================================================
' 8. МИНИМАЛЬНЫЙ JSON-ПАРСЕР (рекурсивный спуск)
'   объект -> Scripting.Dictionary, массив -> Collection
'=============================================================================

' Состояние парсера (mJson, mPos) объявлено в шапке модуля.

Private Function JsonParse(ByVal s As String) As Variant
    Dim v As Variant
    If Len(Trim$(s)) = 0 Then Err.Raise vbObjectError + 690, , "Пустой ответ сервера (тело ответа не получено)."
    mJson = s
    mPos = 1
    JsonSkipWs
    If IsObject(JsonReadValue(v)) Then Set JsonParse = v Else JsonParse = v
End Function

' Обёртка: кладёт разобранное значение в out и возвращает его же — нужна,
' чтобы одинаково обработать объекты (требуют Set) и скаляры.
Private Function JsonReadValue(ByRef out As Variant) As Variant
    Dim v As Variant
    Dim c As String

    JsonSkipWs
    If mPos > Len(mJson) Then Err.Raise vbObjectError + 691, , "JSON: неожиданный конец данных"
    c = Mid$(mJson, mPos, 1)

    Select Case c
        Case "{"
            Set v = JsonObject()
        Case "["
            Set v = JsonArray()
        Case """"
            v = JsonString()
        Case "t"
            JsonExpect "true": v = True
        Case "f"
            JsonExpect "false": v = False
        Case "n"
            JsonExpect "null": v = Null
        Case Else
            v = JsonNumber()
    End Select

    If IsObject(v) Then
        Set out = v
        Set JsonReadValue = v
    Else
        out = v
        JsonReadValue = v
    End If
End Function

Private Function JsonObject() As Object
    Dim d As Object, k As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    mPos = mPos + 1                                  ' пропускаем "{"
    JsonSkipWs
    If Mid$(mJson, mPos, 1) = "}" Then
        mPos = mPos + 1
        Set JsonObject = d
        Exit Function
    End If

    Do
        JsonSkipWs
        If Mid$(mJson, mPos, 1) <> """" Then Err.Raise vbObjectError + 692, , "JSON: ожидался ключ объекта в позиции " & mPos
        k = JsonString()
        JsonSkipWs
        If Mid$(mJson, mPos, 1) <> ":" Then Err.Raise vbObjectError + 693, , "JSON: ожидалось ':' в позиции " & mPos
        mPos = mPos + 1
        JsonReadValue v
        If IsObject(v) Then Set d(k) = v Else d(k) = v
        JsonSkipWs
        Select Case Mid$(mJson, mPos, 1)
            Case ","
                mPos = mPos + 1
            Case "}"
                mPos = mPos + 1
                Set JsonObject = d
                Exit Function
            Case Else
                Err.Raise vbObjectError + 694, , "JSON: ожидалось ',' или '}' в позиции " & mPos
        End Select
    Loop
End Function

Private Function JsonArray() As Collection
    Dim col As New Collection, v As Variant
    mPos = mPos + 1                                  ' пропускаем "["
    JsonSkipWs
    If Mid$(mJson, mPos, 1) = "]" Then
        mPos = mPos + 1
        Set JsonArray = col
        Exit Function
    End If

    Do
        JsonReadValue v
        col.Add v
        JsonSkipWs
        Select Case Mid$(mJson, mPos, 1)
            Case ","
                mPos = mPos + 1
            Case "]"
                mPos = mPos + 1
                Set JsonArray = col
                Exit Function
            Case Else
                Err.Raise vbObjectError + 695, , "JSON: ожидалось ',' или ']' в позиции " & mPos
        End Select
    Loop
End Function

Private Function JsonString() As String
    Dim sb As String, c As String, code As String
    mPos = mPos + 1                                  ' пропускаем открывающую кавычку
    Do
        If mPos > Len(mJson) Then Err.Raise vbObjectError + 696, , "JSON: незакрытая строка"
        c = Mid$(mJson, mPos, 1)
        If c = """" Then
            mPos = mPos + 1
            JsonString = sb
            Exit Function
        ElseIf c = "\" Then
            mPos = mPos + 1
            c = Mid$(mJson, mPos, 1)
            Select Case c
                Case """": sb = sb & """"
                Case "\": sb = sb & "\"
                Case "/": sb = sb & "/"
                Case "b": sb = sb & Chr$(8)
                Case "f": sb = sb & Chr$(12)
                Case "n": sb = sb & vbLf
                Case "r": sb = sb & vbCr
                Case "t": sb = sb & vbTab
                Case "u"
                    code = Mid$(mJson, mPos + 1, 4)
                    sb = sb & ChrW$(CLng("&H" & code))
                    mPos = mPos + 4
                Case Else
                    sb = sb & c
            End Select
            mPos = mPos + 1
        Else
            sb = sb & c
            mPos = mPos + 1
        End If
    Loop
End Function

Private Function JsonNumber() As Variant
    Dim startPos As Long, c As String
    startPos = mPos
    Do While mPos <= Len(mJson)
        c = Mid$(mJson, mPos, 1)
        If (c >= "0" And c <= "9") Or c = "-" Or c = "+" Or c = "." Or c = "e" Or c = "E" Then
            mPos = mPos + 1
        Else
            Exit Do
        End If
    Loop
    If mPos = startPos Then Err.Raise vbObjectError + 697, , "JSON: не число в позиции " & mPos
    ' Val() всегда трактует точку как десятичный разделитель, независимо от локали.
    JsonNumber = Val(Mid$(mJson, startPos, mPos - startPos))
End Function

Private Sub JsonSkipWs()
    Dim c As String
    Do While mPos <= Len(mJson)
        c = Mid$(mJson, mPos, 1)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Then mPos = mPos + 1 Else Exit Do
    Loop
End Sub

Private Sub JsonExpect(ByVal literal As String)
    If StrComp(Mid$(mJson, mPos, Len(literal)), literal, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 698, , "JSON: ожидалось '" & literal & "' в позиции " & mPos
    End If
    mPos = mPos + Len(literal)
End Sub

' Безопасное чтение поля словаря: нет ключа -> Empty.
Private Function DictGet(ByVal obj As Variant, ByVal key As String) As Variant
    If Not IsObject(obj) Then DictGet = Empty: Exit Function
    If obj Is Nothing Then DictGet = Empty: Exit Function
    On Error GoTo Fail
    If obj.Exists(key) Then
        If IsObject(obj(key)) Then Set DictGet = obj(key) Else DictGet = obj(key)
    Else
        DictGet = Empty
    End If
    Exit Function
Fail:
    DictGet = Empty
End Function
