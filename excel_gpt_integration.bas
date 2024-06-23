Function GPT(sPrompt As String, sCell As Range) As String
    ' Константа для интервала между запросами в секундах, для лимита 1 запрос в секунду
    Const RequestInterval As Double = 1.0
    ' Объявление переменных для HTTP-запроса и ответа
    Dim objHTTP As Object
    Dim URL As String
    Dim payload As String
    Dim jsonText As String
    Dim Parsed As Dictionary
    Dim sContext As String

    ' Подготовка контекста для запроса, объединение запроса пользователя и значения ячейки
    sContext = sPrompt & " " & sCell.Value

    ' Создаем объект для HTTP-запросов
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")

    ' Адрес API
    URL = "https://api.openai.com/v1/chat/completions"

    ' Формируем JSON тело запроса
    payload = "{""model"": ""gpt-4o"", ""temperature"": 0, ""messages"": [{""role"": ""user"", ""content"": """ & sContext & """}]}"

    ' Открытие запроса
    objHTTP.Open "POST", URL, False

    ' Установка необходимых заголовков
    objHTTP.setRequestHeader "Authorization", "Bearer $OPENAI_API_KEY"  ' Замените $OPENAI_API_KEY на ваш ключ API
    objHTTP.setRequestHeader "Content-Type", "application/json"

    ' Отправка запроса с JSON телом
    objHTTP.Send payload

    ' Проверяем статус ответа
    If objHTTP.Status = 200 Then
        jsonText = objHTTP.ResponseText
        ' Парсинг JSON ответа
        Set Parsed = JsonConverter.ParseJson(jsonText)

        ' Извлечение необходимой информации из JSON
        GPT = Parsed("choices")(1)("message")("content")
    Else
        ' Обработка ошибок
        GPT = "Error: " & objHTTP.Status & " " & objHTTP.statusText
    End If

    ' Очистка объекта
    Set objHTTP = Nothing
    
    ' Задержка перед следующим запросом
    WaitSeconds RequestInterval
End Function

' Процедура ожидания заданного количества секунд
Public Sub WaitSeconds(Seconds As Double)
    Dim endtime As Double
    endtime = Timer + Seconds
    Do While Timer < endtime And Timer >= 0
        DoEvents ' Передача управления для обработки других задач в очереди сообщений
    Loop
End Sub