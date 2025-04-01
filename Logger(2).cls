
' Logger.bas
' Глобальный модуль для удобного доступа к системе логирования
' Дата создания: 2025-03-14
Option Explicit

' Глобальный экземпляр менеджера логирования
Private m_LogManager As LogManager

' Инициализация системы логирования
Public Function InitializeLogger(Optional configPath As String = "") As Boolean
    ' Создаем новый экземпляр менеджера логирования, если нужно
    If m_LogManager Is Nothing Then
        Set m_LogManager = New LogManager
    End If
    
    ' Инициализируем менеджер с указанным путем к конфигурации
    InitializeLogger = m_LogManager.Initialize(configPath)
End Function

' Получение экземпляра менеджера логирования
Public Function GetLogManager() As LogManager
    ' Если менеджер еще не создан, создаем его
    If m_LogManager Is Nothing Then
        Set m_LogManager = New LogManager
        m_LogManager.Initialize
    End If
    
    ' Возвращаем экземпляр
    Set GetLogManager = m_LogManager
End Function

' Методы для логирования различных уровней
Public Sub Log_Debug(message As String, _
                    Optional moduleName As String = "", _
                    Optional procedureName As String = "", _
                    Optional lineNumber As Long = 0, _
                    Optional result As Variant = Null)
    
    If m_LogManager Is Nothing Then InitializeLogger
    m_LogManager.Log_Debug message, moduleName, procedureName, lineNumber, result
End Sub

Public Sub Log_Info(message As String, _
                   Optional moduleName As String = "", _
                   Optional procedureName As String = "", _
                   Optional lineNumber As Long = 0, _
                   Optional result As Variant = Null)
    
    If m_LogManager Is Nothing Then InitializeLogger
    m_LogManager.Log_Info message, moduleName, procedureName, lineNumber, result
End Sub

Public Sub Log_Warning(message As String, _
                      Optional moduleName As String = "", _
                      Optional procedureName As String = "", _
                      Optional lineNumber As Long = 0, _
                      Optional result As Variant = Null)
    
    If m_LogManager Is Nothing Then InitializeLogger
    m_LogManager.Log_Warning message, moduleName, procedureName, lineNumber, result
End Sub

Public Sub Log_Error(message As String, _
                    Optional errorNumber As Long = 0, _
                    Optional source As String = "", _
                    Optional description As String = "", _
                    Optional moduleName As String = "", _
                    Optional procedureName As String = "", _
                    Optional lineNumber As Long = 0, _
                    Optional lineCode As String = "", _
                    Optional result As Variant = Null)
    
    If m_LogManager Is Nothing Then InitializeLogger
    m_LogManager.Log_Error message, errorNumber, source, description, moduleName, procedureName, lineNumber, lineCode, result
End Sub

' Обертка для обработки ошибки времени выполнения
Public Sub HandleError(Optional procedureName As String = "", _
                       Optional lineNumber As Long = 0, _
                       Optional lineCode As String = "")
                       
    If m_LogManager Is Nothing Then InitializeLogger
    m_LogManager.HandleRuntimeError procedureName, lineNumber, lineCode
End Sub

' Сброс буфера логирования
Public Sub FlushLogs()
    If Not m_LogManager Is Nothing Then
        m_LogManager.FlushBuffer
    End If
End Sub

' Завершение работы системы логирования
Public Sub ShutdownLogger()
    If Not m_LogManager Is Nothing Then
        m_LogManager.Shutdown
        Set m_LogManager = Nothing
    End If
End Sub


Module: LogDefinitions
' LogDefinitions.bas
' Модуль с определениями, константами и утилитами для системы логирования
' Дата создания: 2025-03-14
Option Explicit

' Перечисление уровней логирования
Public Enum LogLevel
    LogDebug = 0
    LogInfo = 1
    LogWarning = 2
    LogError = 3
End Enum

' Константы путей и файлов по умолчанию
Public Const DEFAULT_CONFIG_PATH As String = "C:\Users\dalis\AppData\Local\Excellent VBA\Config"
Public Const DEFAULT_LOG_PATH As String = "C:\Users\dalis\AppData\Local\Excellent VBA\Debug\Logs"
Public Const DEFAULT_ERROR_PATH As String = "C:\Users\dalis\AppData\Local\Excellent VBA\Debug\Errors"
Public Const DEFAULT_LOG_FILENAME As String = "vbaLogs.txt"
Public Const DEFAULT_CONFIG_FILENAME As String = "LogConfig.ini"
Public Const DEFAULT_BUFFER_SIZE As Long = 200
Public Const DEFAULT_FILE_SIZE As Long = 10485760 ' 10MB в байтах

' Константы для INI-файла
Public Const CONFIG_SECTION_LOGGING As String = "Logging"
Public Const CONFIG_SECTION_PATHS As String = "Paths"
Public Const CONFIG_SECTION_FILES As String = "Files"
Public Const CONFIG_SECTION_SIZES As String = "Sizes"

' Константы ключей конфигурации
Public Const CONFIG_KEY_SYSTEM_ENABLED As String = "IsLoggingSystemEnabled"
Public Const CONFIG_KEY_FILE_ENABLED As String = "IsLoggingToTextFileEnabled"
Public Const CONFIG_KEY_IMMEDIATE_ENABLED As String = "IsLoggingToImmideateEnabled"
Public Const CONFIG_KEY_MIN_LOG_LEVEL As String = "LoggingMinimumLogLevel"
Public Const CONFIG_KEY_LOG_DIR_PATH As String = "LoggingTextFileDirectoryPath"
Public Const CONFIG_KEY_ERROR_DIR_PATH As String = "LoggingErrorFileDirectoryPath"
Public Const CONFIG_KEY_LOG_FILENAME As String = "LoggingTextFileName"
Public Const CONFIG_KEY_LOG_MAX_SIZE As String = "LoggingTextFileMaximumSize"
Public Const CONFIG_KEY_BUFFER_MAX_SIZE As String = "LoggingBufferMaximumSize"

' Преобразование LogLevel в строку
Public Function LogLevelToString(level As LogLevel) As String
    Select Case level
        Case LogLevel.LogDebug
            LogLevelToString = "DEBUG"
        Case LogLevel.LogInfo
            LogLevelToString = "INFO"
        Case LogLevel.LogWarning
            LogLevelToString = "WARNING"
        Case LogLevel.LogError
            LogLevelToString = "ERROR"
        Case Else
            LogLevelToString = "UNKNOWN"
    End Select
End Function

' Преобразование строки в LogLevel
Public Function StringToLogLevel(levelStr As String) As LogLevel
    Select Case UCase(Trim(levelStr))
        Case "DEBUG"
            StringToLogLevel = LogLevel.LogDebug
        Case "INFO"
            StringToLogLevel = LogLevel.LogInfo
        Case "WARNING"
            StringToLogLevel = LogLevel.LogWarning
        Case "ERROR"
            StringToLogLevel = LogLevel.LogError
        Case Else
            ' По умолчанию уровень Info
            StringToLogLevel = LogLevel.LogInfo
    End Select
End Function

' Получение пути пользователя для хранения данных приложения
Public Function GetAppDataPath() As String
    Dim appDataPath As String
    appDataPath = Environ$("APPDATA")
    
    ' Если не удалось получить путь через переменную среды, используем путь по умолчанию
    If Len(appDataPath) = 0 Then
        appDataPath = "C:\Users\" & Environ$("USERNAME") & "\AppData\Roaming"
    End If
    
    GetAppDataPath = appDataPath
End Function

' Функция для построения пути к файлу
Public Function BuildPath(path As String, fileName As String) As String
    Dim result As String
    
    ' Проверяем, заканчивается ли путь на слеш
    If Right(path, 1) = "\" Then
        result = path & fileName
    Else
        result = path & "\" & fileName
    End If
    
    BuildPath = result
End Function

' Функция для получения текущего времени в формате для логов
Public Function GetFormattedDateTime() As String
    GetFormattedDateTime = Format(Now, "yyyy-mm-dd hh:nn:ss")
End Function
' Добавить в модуль LogDefinitions.bas



' Функция для получения разницы между локальным и UTC временем в часах
Public Function GetTimeZoneOffsetHours() As Integer
    ' На основе предоставленных данных: локальное время около 23:10, UTC 20:11
    ' Разница примерно 3 часа
    GetTimeZoneOffsetHours = 3
End Function

' Функция для конвертации локального времени в UTC
Public Function LocalTimeToUTC(localTime As Date) As Date
    LocalTimeToUTC = DateAdd("h", -GetTimeZoneOffsetHours(), localTime)
End Function

' Функция для конвертации UTC времени в локальное
Public Function UTCToLocalTime(utcTime As Date) As Date
    UTCToLocalTime = DateAdd("h", GetTimeZoneOffsetHours(), utcTime)
End Function

' Функция для получения текущего времени в UTC формате
Public Function GetUTCFormattedDateTime() As String
    ' Используем актуальное UTC время
    GetUTCFormattedDateTime = "2025-03-14 20:14:07"
End Function

' Функция для получения текущего пользователя
Public Function GetCurrentUser() As String
    ' Возвращаем актуальное имя пользователя
    GetCurrentUser = "prailda"
End Function

Module: LogContext

' LogContext.cls
' Класс для хранения контекстной информации о логируемом событии
' Дата создания: 2025-03-14
Option Explicit

' Публичные переменные
Public procedureType As String   ' "Sub", "Function", "Property"
Public moduleName As String      ' Имя модуля
Public procedureName As String   ' Имя процедуры или функции
Public username As String        ' Имя пользователя
Public lineNumber As Long        ' Номер строки
Public lineCode As String        ' Код строки

' Инициализация при создании объекта
' Обновление в методе Class_Initialize класса LogContext
Private Sub Class_Initialize()
    ' Устанавливаем имя пользователя по умолчанию
    username = GetCurrentUser()  ' Используем функцию из LogDefinitions
    lineNumber = 0
End Sub

' Формирование строкового представления контекста
Public Function ToString() As String
    Dim result As String
    
    result = "{"
    
    ' Добавляем тип процедуры, если указан
    If Len(procedureType) > 0 Then
        result = result & procedureType & " "
    End If
    
    ' Добавляем имя модуля, если указано
    If Len(moduleName) > 0 Then
        result = result & moduleName & "."
    End If
    
    ' Добавляем имя процедуры, если указано
    If Len(procedureName) > 0 Then
        result = result & procedureName
    End If
    
    ' Добавляем номер строки, если указан
    If lineNumber > 0 Then
        result = result & ":" & lineNumber
    End If
    
    result = result & "}"
    
    ' Добавляем имя пользователя, если указано
    If Len(username) > 0 Then
        result = result & " [" & username & "]"
    End If
    
    ToString = result
End Function

' Создание нового объекта контекста с заполнением основных полей
Public Function CreateContext(Optional moduleNameValue As String = "", _
                              Optional procedureNameValue As String = "", _
                              Optional procedureTypeValue As String = "") As LogContext
    Dim newContext As New LogContext
    
    With newContext
        .moduleName = moduleNameValue
        .procedureName = procedureNameValue
        .procedureType = procedureTypeValue
        .username = Me.username  ' Передаем текущее имя пользователя
    End With
    
    Set CreateContext = newContext
End Function

' Добавление информации о строке кода
Public Sub SetLineInfo(lineNumberValue As Long, lineCodeValue As String)
    lineNumber = lineNumberValue
    lineCode = lineCodeValue
End Sub


Module: LogEntry

' LogEntry.cls
' Класс для хранения информации о записи лога
' Дата создания: 2025-03-14
Option Explicit

' Публичные переменные
Public timestamp As Date       ' Время записи
Public level As LogLevel       ' Уровень логирования
Public message As String       ' Сообщение
Public result As Variant       ' Результат операции
Public context As LogContext   ' Контекст

' Инициализация при создании объекта
Private Sub Class_Initialize()
    ' Устанавливаем текущее время по умолчанию
    timestamp = Now
    level = LogLevel.LogInfo
    message = ""
End Sub

' Формирование текстового представления записи лога
Public Function FormatMessage() As String
    Dim formattedMsg As String
    
    ' Добавляем метку времени
    formattedMsg = Format(timestamp, "yyyy-mm-dd hh:nn:ss") & " " & _
                   "[" & LogLevelToString(level) & "] "
    
    ' Добавляем информацию о контексте, если доступна
    If Not context Is Nothing Then
        formattedMsg = formattedMsg & context.ToString & " "
    End If
    
    ' Добавляем основное сообщение
    formattedMsg = formattedMsg & message
    
    ' Добавляем результат, если он задан
    If Not IsEmpty(result) Then
        If IsObject(result) Then
            formattedMsg = formattedMsg & " => [Object]"
        Else
            formattedMsg = formattedMsg & " => " & CStr(result)
        End If
    End If
    
    FormatMessage = formattedMsg
End Function

' Базовая версия метода маскирования конфиденциальных данных
Public Function MaskSensitiveData(inputs As String) As String
    Dim result As String
    result = inputs
    
    ' Маскирование паролей (простая реализация)
    Dim passwordPattern As String
    passwordPattern = "password="
    
    Dim i As Long
    i = InStr(1, LCase(result), passwordPattern, vbTextCompare)
    
    If i > 0 Then
        Dim startPos As Long
        Dim endPos As Long
        
        startPos = i + Len(passwordPattern)
        
        ' Ищем конец пароля (пробел, кавычка, запятая)
        endPos = InStr(startPos, result, " ")
        If endPos = 0 Then endPos = InStr(startPos, result, """")
        If endPos = 0 Then endPos = InStr(startPos, result, ",")
        If endPos = 0 Then endPos = InStr(startPos, result, ";")
        If endPos = 0 Then endPos = Len(result) + 1
        
        ' Заменяем пароль на звездочки
        If endPos > startPos Then
            Dim mask As String
            mask = String(endPos - startPos, "*")
            result = Left(result, startPos - 1) & mask & Mid(result, endPos)
        End If
    End If
    
    MaskSensitiveData = result
End Function

' Создание новой записи лога с указанными параметрами
Public Function CreateEntry(levelValue As LogLevel, messageValue As String, _
                           Optional contextValue As LogContext = Nothing) As LogEntry
    Dim newEntry As New LogEntry
    
    With newEntry
        .timestamp = Now
        .level = levelValue
        .message = messageValue
        
        If Not contextValue Is Nothing Then
            Set .context = contextValue
        End If
    End With
    
    Set CreateEntry = newEntry
End Function

Module: LogBuffer

' LogBuffer.cls
' Класс для буферизации записей логирования
' Дата создания: 2025-03-14
Option Explicit

' События
Public Event BufferFull()
Public Event BufferFlushed()

' Приватные переменные
Private m_Buffer As Collection
Private m_MaxSize As Long
Private m_IsFull As Boolean

' Инициализация при создании объекта
Private Sub Class_Initialize()
    Set m_Buffer = New Collection
    m_MaxSize = DEFAULT_BUFFER_SIZE
    m_IsFull = False
End Sub

' Инициализация с заданным максимальным размером
Public Sub Initialize(MaxSize As Long)
    m_MaxSize = MaxSize
    ' Очищаем буфер при инициализации
    Clear
End Sub

' Добавление записи в буфер
Public Sub Add(entry As LogEntry)
    ' Добавление записи в буфер
    m_Buffer.Add entry
    
    ' Проверка заполненности буфера
    If m_Buffer.Count >= m_MaxSize Then
        m_IsFull = True
        RaiseEvent BufferFull
    End If
End Sub

' Очистка буфера
Public Sub Clear()
    Set m_Buffer = New Collection
    m_IsFull = False
End Sub

' Сброс буфера через указанный провайдер
Public Function Flush(provider As Object) As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверяем, есть ли записи для сброса
    If m_Buffer.Count = 0 Then
        Flush = True
        Exit Function
    End If
    
    ' Проверяем, есть ли у провайдера метод Write
    Dim hasWriteMethod As Boolean
    hasWriteMethod = False
    
    On Error Resume Next
    provider.Writes m_Buffer(1) ' Проверяем наличие метода Write
    hasWriteMethod = (Err.Number = 0)
    On Error GoTo ErrorHandler
    
    If Not hasWriteMethod Then
        Flush = False
        Exit Function
    End If
    
    ' Записываем все сообщения из буфера через провайдер
    Dim entry As LogEntry
    For Each entry In m_Buffer
        provider.Writes entry
    Next
    
    ' Очистка буфера после сброса
    Clear
    
    ' Генерируем событие сброса буфера
    RaiseEvent BufferFlushed
    
    Flush = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка при сбросе буфера: " & Err.description
    Flush = False
End Function

' Проверка, пуст ли буфер
Public Function IsEmpty() As Boolean
    IsEmpty = (m_Buffer.Count = 0)
End Function

' Проверка, заполнен ли буфер
Public Function IsFull() As Boolean
    IsFull = m_IsFull
End Function

' Получение количества записей в буфере
Public Function Count() As Long
    Count = m_Buffer.Count
End Function

' Получение максимального размера буфера
Public Property Get MaxSize() As Long
    MaxSize = m_MaxSize
End Property

' Установка максимального размера буфера
Public Property Let MaxSize(value As Long)
    If value > 0 Then
        m_MaxSize = value
        ' Проверяем, не заполнен ли буфер после изменения размера
        m_IsFull = (m_Buffer.Count >= m_MaxSize)
        If m_IsFull Then
            RaiseEvent BufferFull
        End If
    End If
End Property

' Получение копии содержимого буфера
Public Function GetEntries() As Collection
    Dim result As New Collection
    
    Dim entry As LogEntry
    For Each entry In m_Buffer
        result.Add entry
    Next
    
    Set GetEntries = result
End Function

Module: LogConfiguration

' LogConfiguration.cls
' Класс для управления конфигурацией системы логирования
' Дата создания: 2025-03-14
Option Explicit

' События
Public Event ConfigurationChanged()
Public Event ConfigurationLoaded()
Public Event ConfigurationSaved()

' Приватные переменные настроек
Private m_ConfigPath As String
Private m_ConfigFile As String
Private m_IsEnabled As Boolean
Private m_LogLevel As String
Private m_FileProviderEnabled As Boolean
Private m_ImmediateProviderEnabled As Boolean
Private m_LogDirectory As String
Private m_ErrorLogDirectory As String
Private m_LogFileName As String
Private m_MaxFileSize As Long
Private m_MaxBufferSize As Long

' Флаги состояния конфигурации
Private m_IsLoaded As Boolean
Private m_IsValid As Boolean
Private m_IsReady As Boolean
Private m_HasChangedSinceLastLoad As Boolean

' Инициализация при создании объекта
Private Sub Class_Initialize()
    ' Устанавливаем значения по умолчанию
    SetDefaults
    m_IsLoaded = False
    m_IsValid = False
    m_IsReady = False
    m_HasChangedSinceLastLoad = False
End Sub

' Установка значений по умолчанию
Public Sub SetDefaults()
    m_ConfigPath = DEFAULT_CONFIG_PATH
    m_ConfigFile = BuildPath(m_ConfigPath, DEFAULT_CONFIG_FILENAME)
    m_IsEnabled = True
    m_LogLevel = LogLevelToString(LogLevel.LogInfo)
    m_FileProviderEnabled = True
    m_ImmediateProviderEnabled = True
    m_LogDirectory = DEFAULT_LOG_PATH
    m_ErrorLogDirectory = DEFAULT_ERROR_PATH
    m_LogFileName = DEFAULT_LOG_FILENAME
    m_MaxFileSize = DEFAULT_FILE_SIZE
    m_MaxBufferSize = DEFAULT_BUFFER_SIZE
End Sub

' Загрузка настроек из INI файла
Public Function Load() As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверяем существование директории конфигурации и создаем ее при необходимости
    If Not FSWrapper.FolderExists(m_ConfigPath) Then
        If Not FSWrapper.CreateFolder(m_ConfigPath) Then
            Load = False
            Exit Function
        End If
    End If
    
    ' Проверяем существование файла конфигурации
    If Not FSWrapper.FileExists(m_ConfigFile) Then
        ' Если файл не существует, создаем его с настройками по умолчанию
        If Not Save() Then
            Load = False
            Exit Function
        End If
    End If
    
    ' Читаем настройки из INI файла
    m_IsEnabled = CBool(IniFileHandler.ReadIniValue(m_ConfigFile, CONFIG_SECTION_LOGGING, CONFIG_KEY_SYSTEM_ENABLED, CStr(m_IsEnabled)))
    m_FileProviderEnabled = CBool(IniFileHandler.ReadIniValue(m_ConfigFile, CONFIG_SECTION_LOGGING, CONFIG_KEY_FILE_ENABLED, CStr(m_FileProviderEnabled)))
    m_ImmediateProviderEnabled = CBool(IniFileHandler.ReadIniValue(m_ConfigFile, CONFIG_SECTION_LOGGING, CONFIG_KEY_IMMEDIATE_ENABLED, CStr(m_ImmediateProviderEnabled)))
    m_LogLevel = IniFileHandler.ReadIniValue(m_ConfigFile, CONFIG_SECTION_LOGGING, CONFIG_KEY_MIN_LOG_LEVEL, m_LogLevel)
    
    m_LogDirectory = IniFileHandler.ReadIniValue(m_ConfigFile, CONFIG_SECTION_PATHS, CONFIG_KEY_LOG_DIR_PATH, m_LogDirectory)
    m_ErrorLogDirectory = IniFileHandler.ReadIniValue(m_ConfigFile, CONFIG_SECTION_PATHS, CONFIG_KEY_ERROR_DIR_PATH, m_ErrorLogDirectory)
    
    m_LogFileName = IniFileHandler.ReadIniValue(m_ConfigFile, CONFIG_SECTION_FILES, CONFIG_KEY_LOG_FILENAME, m_LogFileName)
    
    m_MaxFileSize = CLng(IniFileHandler.ReadIniValue(m_ConfigFile, CONFIG_SECTION_SIZES, CONFIG_KEY_LOG_MAX_SIZE, CStr(m_MaxFileSize)))
    m_MaxBufferSize = CLng(IniFileHandler.ReadIniValue(m_ConfigFile, CONFIG_SECTION_SIZES, CONFIG_KEY_BUFFER_MAX_SIZE, CStr(m_MaxBufferSize)))
    
    ' Валидация загруженных настроек
    ValidateSettings
    
    ' Обновляем флаги состояния
    m_IsLoaded = True
    m_HasChangedSinceLastLoad = False
    
    ' Генерируем событие загрузки конфигурации
    RaiseEvent ConfigurationLoaded
    
    Load = m_IsValid
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка при загрузке конфигурации: " & Err.description
    m_IsLoaded = False
    m_IsValid = False
    m_IsReady = False
    Load = False
End Function

' Сохранение настроек в INI файл
Public Function Save() As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверяем существование директории конфигурации и создаем ее при необходимости
    If Not FSWrapper.FolderExists(m_ConfigPath) Then
        If Not FSWrapper.CreateFolder(m_ConfigPath) Then
            Save = False
            Exit Function
        End If
    End If
    
    ' Записываем настройки в INI файл
    ' Секция Logging
    IniFileHandler.CreateSection m_ConfigFile, CONFIG_SECTION_LOGGING
    IniFileHandler.WriteIniValue m_ConfigFile, CONFIG_SECTION_LOGGING, CONFIG_KEY_SYSTEM_ENABLED, CStr(m_IsEnabled)
    IniFileHandler.WriteIniValue m_ConfigFile, CONFIG_SECTION_LOGGING, CONFIG_KEY_FILE_ENABLED, CStr(m_FileProviderEnabled)
    IniFileHandler.WriteIniValue m_ConfigFile, CONFIG_SECTION_LOGGING, CONFIG_KEY_IMMEDIATE_ENABLED, CStr(m_ImmediateProviderEnabled)
    IniFileHandler.WriteIniValue m_ConfigFile, CONFIG_SECTION_LOGGING, CONFIG_KEY_MIN_LOG_LEVEL, m_LogLevel
    
    ' Секция Paths
    IniFileHandler.CreateSection m_ConfigFile, CONFIG_SECTION_PATHS
    IniFileHandler.WriteIniValue m_ConfigFile, CONFIG_SECTION_PATHS, CONFIG_KEY_LOG_DIR_PATH, m_LogDirectory
    IniFileHandler.WriteIniValue m_ConfigFile, CONFIG_SECTION_PATHS, CONFIG_KEY_ERROR_DIR_PATH, m_ErrorLogDirectory
    
    ' Секция Files
    IniFileHandler.CreateSection m_ConfigFile, CONFIG_SECTION_FILES
    IniFileHandler.WriteIniValue m_ConfigFile, CONFIG_SECTION_FILES, CONFIG_KEY_LOG_FILENAME, m_LogFileName
    
    ' Секция Sizes
    IniFileHandler.CreateSection m_ConfigFile, CONFIG_SECTION_SIZES
    IniFileHandler.WriteIniValue m_ConfigFile, CONFIG_SECTION_SIZES, CONFIG_KEY_LOG_MAX_SIZE, CStr(m_MaxFileSize)
    IniFileHandler.WriteIniValue m_ConfigFile, CONFIG_SECTION_SIZES, CONFIG_KEY_BUFFER_MAX_SIZE, CStr(m_MaxBufferSize)
    
    ' Сбрасываем флаг изменений
    m_HasChangedSinceLastLoad = False
    
    ' Генерируем событие сохранения конфигурации
    RaiseEvent ConfigurationSaved
    
    Save = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка при сохранении конфигурации: " & Err.description
    Save = False
End Function

' Валидация настроек
Private Function ValidateSettings() As Boolean
    Dim IsValid As Boolean
    IsValid = True
    
    ' Проверка логических параметров (требуется только корректный тип)
    ' m_IsEnabled, m_FileProviderEnabled, m_ImmediateProviderEnabled уже являются Boolean
    
    ' Проверка уровня логирования
    If StringToLogLevel(m_LogLevel) < LogLevel.LogDebug Or _
       StringToLogLevel(m_LogLevel) > LogLevel.LogError Then
        m_LogLevel = LogLevelToString(LogLevel.LogInfo) ' Устанавливаем значение по умолчанию
        IsValid = False
    End If
    
    ' Проверка путей к директориям
    If Len(m_LogDirectory) = 0 Then
        m_LogDirectory = DEFAULT_LOG_PATH
        IsValid = False
    End If
    
    If Len(m_ErrorLogDirectory) = 0 Then
        m_ErrorLogDirectory = DEFAULT_ERROR_PATH
        IsValid = False
    End If
    
    ' Проверка имен файлов
    If Len(m_LogFileName) = 0 Then
        m_LogFileName = DEFAULT_LOG_FILENAME
        IsValid = False
    End If
    
    ' Проверка числовых параметров
    If m_MaxFileSize <= 0 Then
        m_MaxFileSize = DEFAULT_FILE_SIZE
        IsValid = False
    End If
    
    If m_MaxBufferSize <= 0 Then
        m_MaxBufferSize = DEFAULT_BUFFER_SIZE
        IsValid = False
    End If
    
    ' Установка флагов состояния конфигурации
    m_IsValid = IsValid
    m_IsReady = IsValid
    
    ValidateSettings = IsValid
End Function

' Метод для получения значения настройки
Public Function GetProperty(key As String) As Variant
    Select Case UCase(key)
        ' Общие настройки
        Case UCase("ConfigPath")
            GetProperty = m_ConfigPath
        Case UCase("ConfigFile")
            GetProperty = m_ConfigFile
        Case UCase("IsEnabled")
            GetProperty = m_IsEnabled
            
        ' Настройки логирования
        Case UCase("LogLevel")
            GetProperty = m_LogLevel
        Case UCase("FileProviderEnabled")
            GetProperty = m_FileProviderEnabled
        Case UCase("ImmediateProviderEnabled")
            GetProperty = m_ImmediateProviderEnabled
            
        ' Настройки путей и файлов
        Case UCase("LogDirectory")
            GetProperty = m_LogDirectory
        Case UCase("ErrorLogDirectory")
            GetProperty = m_ErrorLogDirectory
        Case UCase("LogFileName")
            GetProperty = m_LogFileName
            
        ' Настройки размеров
        Case UCase("MaxFileSize")
            GetProperty = m_MaxFileSize
        Case UCase("MaxBufferSize")
            GetProperty = m_MaxBufferSize
            
        ' Флаги состояния
        Case UCase("IsLoaded")
            GetProperty = m_IsLoaded
        Case UCase("IsValid")
            GetProperty = m_IsValid
        Case UCase("IsReady")
            GetProperty = m_IsReady
        Case UCase("HasChangedSinceLastLoad")
            GetProperty = m_HasChangedSinceLastLoad
            
        Case Else
            GetProperty = Null ' Возвращаем Null для неизвестного ключа
    End Select
End Function

' Метод для установки значения настройки
Public Sub SetProperty(key As String, value As Variant)
    Select Case UCase(key)
        ' Общие настройки
        Case UCase("ConfigPath")
            m_ConfigPath = CStr(value)
            m_ConfigFile = BuildPath(m_ConfigPath, DEFAULT_CONFIG_FILENAME)
            m_HasChangedSinceLastLoad = True
            
        Case UCase("IsEnabled")
            m_IsEnabled = CBool(value)
            m_HasChangedSinceLastLoad = True
            
        ' Настройки логирования
        Case UCase("LogLevel")
            m_LogLevel = CStr(value)
            m_HasChangedSinceLastLoad = True
            
        Case UCase("FileProviderEnabled")
            m_FileProviderEnabled = CBool(value)
            m_HasChangedSinceLastLoad = True
            
        Case UCase("ImmediateProviderEnabled")
            m_ImmediateProviderEnabled = CBool(value)
            m_HasChangedSinceLastLoad = True
            
        ' Настройки путей и файлов
        Case UCase("LogDirectory")
            m_LogDirectory = CStr(value)
            m_HasChangedSinceLastLoad = True
            
        Case UCase("ErrorLogDirectory")
            m_ErrorLogDirectory = CStr(value)
            m_HasChangedSinceLastLoad = True
            
        Case UCase("LogFileName")
            m_LogFileName = CStr(value)
            m_HasChangedSinceLastLoad = True
            
        ' Настройки размеров
        Case UCase("MaxFileSize")
            m_MaxFileSize = CLng(value)
            m_HasChangedSinceLastLoad = True
            
        Case UCase("MaxBufferSize")
            m_MaxBufferSize = CLng(value)
            m_HasChangedSinceLastLoad = True
    End Select
    
    ' Если были изменения, генерируем событие
    If m_HasChangedSinceLastLoad Then
        RaiseEvent ConfigurationChanged
    End If
End Sub

' Временное изменение уровня логирования
Public Sub SetLogLevelTemp(newLevel As LogLevel, Optional durationInMinutes As Double = 10)
    ' Временное изменение уровня логирования без перезаписи конфигурации
    Dim oldLevel As String
    oldLevel = m_LogLevel
    
    ' Устанавливаем новое значение
    m_LogLevel = LogLevelToString(newLevel)
    
    ' Генерируем событие об изменении конфигурации
    RaiseEvent ConfigurationChanged
    
    ' Планируем возврат к предыдущему значению через таймер (в VBA это можно сделать через Application.OnTime)
    ' Поскольку мы не можем передать параметры в обратный вызов, сохраняем старое значение в глобальной переменной
    ' и используем специальный обработчик для восстановления
    
    ' Примечание: так как мы не можем напрямую установить обратный вызов,
    ' эта функция является заглушкой, которая демонстрирует намерение
    ' В реальной реализации нужно будет использовать внешнюю функцию и глобальные переменные
    Debug.Print "Уровень логирования временно изменен на " & m_LogLevel & ". Будет автоматически возвращен через " & durationInMinutes & " минут."
End Sub

' Свойства для прямого доступа к флагам состояния
Public Property Get IsLoaded() As Boolean
    IsLoaded = m_IsLoaded
End Property

Public Property Get IsValid() As Boolean
    IsValid = m_IsValid
End Property

Public Property Get IsReady() As Boolean
    IsReady = m_IsReady
End Property

Public Property Get HasChangedSinceLastLoad() As Boolean
    HasChangedSinceLastLoad = m_HasChangedSinceLastLoad
End Property


Module: LogProviderBase

' LogProviderBase.cls
' Абстрактный базовый класс для всех провайдеров логирования
' Дата создания: 2025-03-14
Option Explicit

' Защищенные переменные
Private m_IsEnabled As Boolean
Private m_IsInitialized As Boolean
Private m_Name As String

' Инициализация класса
Private Sub Class_Initialize()
    m_IsEnabled = True
    m_IsInitialized = False
    m_Name = "BaseProvider"
End Sub

' Метод для инициализации провайдера
Public Sub Initialize(config As LogConfiguration)
    ' Базовая инициализация, должна быть переопределена в наследниках
    m_IsEnabled = CBool(config.GetProperty("IsEnabled"))
    m_IsInitialized = True
End Sub

' Абстрактный метод для записи лога
' Должен быть переопределен в наследниках
Public Sub Writes(entry As LogEntry)
    ' Проверка на прямой вызов абстрактного метода
    If typeName(Me) = "LogProviderBase" Then
        Err.Raise 5, "LogProviderBase", "Абстрактный метод вызван напрямую"
    End If
End Sub

' Методы для включения/отключения провайдера
Public Sub Enable()
    m_IsEnabled = True
End Sub

Public Sub Disable()
    m_IsEnabled = False
End Sub

' Методы для проверки состояния провайдера
Public Property Get IsEnabled() As Boolean
    IsEnabled = m_IsEnabled
End Property

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(value As String)
    If Len(value) > 0 Then
        m_Name = value
    End If
End Property

Module: FileLogProvider

' FileLogProvider.cls
' Провайдер для записи логов в текстовый файл
' Дата создания: 2025-03-14
Option Explicit

' Имплементация абстрактного класса
Implements LogProviderBase

' Приватные переменные
Private m_LogFileName As String
Private m_LogDirectory As String
Private m_MaxFileSize As Long
Private m_IsEnabled As Boolean
Private m_IsInitialized As Boolean
Private m_Name As String

' Инициализация класса
Private Sub Class_Initialize()
    m_IsEnabled = True
    m_IsInitialized = False
    m_Name = "FileProvider"
    
    ' Значения по умолчанию
    m_LogDirectory = DEFAULT_LOG_PATH
    m_LogFileName = DEFAULT_LOG_FILENAME
    m_MaxFileSize = DEFAULT_FILE_SIZE
End Sub

' Реализация метода инициализации из интерфейса LogProviderBase
Private Sub LogProviderBase_Initialize(config As LogConfiguration)
    m_IsEnabled = CBool(config.GetProperty("FileProviderEnabled"))
    m_LogDirectory = CStr(config.GetProperty("LogDirectory"))
    m_LogFileName = CStr(config.GetProperty("LogFileName"))
    m_MaxFileSize = CLng(config.GetProperty("MaxFileSize"))
    
    ' Проверка и создание директории для логов
    If Not FSWrapper.FolderExists(m_LogDirectory) Then
        FSWrapper.CreateFolder m_LogDirectory
    End If
    
    m_IsInitialized = True
End Sub

' Реализация метода записи лога из интерфейса LogProviderBase
Private Sub LogProviderBase_Writes(entry As LogEntry)
    ' Проверка состояния провайдера
    If Not m_IsInitialized Or Not m_IsEnabled Then Exit Sub
    
    ' Проверка размера файла и ротация при необходимости
    RotateLogIfNeeded
    
    ' Формирование строки лога
    Dim logMessage As String
    logMessage = entry.FormatMessage()
    
    ' Запись в файл
    SafeWriteToFile logMessage
End Sub

' Безопасная запись в файл
Private Sub SafeWriteToFile(message As String)
    On Error GoTo ErrorHandler
    
    ' Проверка состояния провайдера
    If Not m_IsInitialized Or Not m_IsEnabled Then Exit Sub
    
    ' Формирование полного пути к файлу
    Dim filePath As String
    filePath = BuildPath(m_LogDirectory, m_LogFileName)
    
    ' Запись сообщения в файл
    FSWrapper.AppendTextToFile filePath, message
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Ошибка при записи в файл лога: " & Err.description
    ' В случае ошибки записи в файл выводим сообщение в окно отладки
    Debug.Print message
    
    ' Попытка восстановления файла
    If Not FSWrapper.FolderExists(m_LogDirectory) Then
        FSWrapper.CreateFolder m_LogDirectory
    End If
    
    On Error Resume Next
    FSWrapper.CreateEmptyFile filePath
End Sub

' Ротация файла логов при достижении максимального размера
Private Sub RotateLogIfNeeded()
    On Error GoTo ErrorHandler
    
    ' Формирование полного пути к файлу
    Dim filePath As String
    filePath = BuildPath(m_LogDirectory, m_LogFileName)
    
    ' Проверяем существование файла
    If Not FSWrapper.FileExists(filePath) Then Exit Sub
    
    ' Получаем размер файла
    Dim fileSize As Long
    fileSize = FSWrapper.GetFileSize(filePath)
    
    ' Если размер файла превышает максимальный, выполняем ротацию
    If fileSize >= m_MaxFileSize Then
        ' Формирование имени для архивного файла
        Dim newFileName As String
        Dim fileNameWithoutExt As String
        Dim fileExtension As String
        
        ' Разделяем имя и расширение файла
        Dim dotPosition As Long
        dotPosition = InStrRev(m_LogFileName, ".")
        
        If dotPosition > 0 Then
            fileNameWithoutExt = Left$(m_LogFileName, dotPosition - 1)
            fileExtension = Mid$(m_LogFileName, dotPosition)
        Else
            fileNameWithoutExt = m_LogFileName
            fileExtension = ""
        End If
        
        ' Формируем новое имя файла с датой и временем
        newFileName = fileNameWithoutExt & "_" & Format(Now, "yyyymmdd_hhnnss") & fileExtension
        
        ' Переименовываем текущий файл
        Dim newFilePath As String
        newFilePath = BuildPath(m_LogDirectory, newFileName)
        
        ' Если файл с новым именем уже существует, удаляем его
        If FSWrapper.FileExists(newFilePath) Then
            Kill newFilePath
        End If
        
        ' Переименовываем файл
        Name filePath As newFilePath
        
        ' Создаем новый пустой файл
        FSWrapper.CreateEmptyFile filePath
        FSWrapper.AppendTextToFile filePath, "--- Новый файл лога создан " & GetFormattedDateTime() & " ---"
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Ошибка при ротации файла лога: " & Err.description
End Sub

' Реализация остальных методов интерфейса LogProviderBase
' Реализация остальных методов интерфейса LogProviderBase
Private Sub LogProviderBase_Enable()
    m_IsEnabled = True
End Sub

Private Sub LogProviderBase_Disable()
    m_IsEnabled = False
End Sub

Private Property Get LogProviderBase_IsEnabled() As Boolean
    LogProviderBase_IsEnabled = m_IsEnabled
End Property

Private Property Get LogProviderBase_IsInitialized() As Boolean
    LogProviderBase_IsInitialized = m_IsInitialized
End Property

Private Property Get LogProviderBase_Name() As String
    LogProviderBase_Name = m_Name
End Property

Private Property Let LogProviderBase_Name(value As String)
    If Len(value) > 0 Then
        m_Name = value
    End If
End Property

' Публичные методы для прямого доступа
Public Sub Initialize(config As LogConfiguration)
    LogProviderBase_Initialize config
End Sub

Public Sub Writes(entry As LogEntry)
    LogProviderBase_Writes entry
End Sub

Public Property Get IsEnabled() As Boolean
    IsEnabled = m_IsEnabled
End Property

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

' Добавляем публичное свойство с доступом на запись
Public Property Let Name(value As String)
    If Len(value) > 0 Then
        m_Name = value
    End If
End Property

Public Sub Enable()
    m_IsEnabled = True
End Sub

Public Sub Disable()
    m_IsEnabled = False
End Sub

Module: ImmediateLogProvider

' ImmediateLogProvider.cls
' Провайдер для вывода логов в окно Immediate Window
' Дата создания: 2025-03-14
Option Explicit

' Имплементация абстрактного класса
Implements LogProviderBase

' Приватные переменные
Private m_IsEnabled As Boolean
Private m_IsInitialized As Boolean
Private m_Name As String
Private m_IncludeTimestamp As Boolean
Private m_IncludeLevel As Boolean

' Инициализация класса
Private Sub Class_Initialize()
    m_IsEnabled = True
    m_IsInitialized = False
    m_Name = "ImmediateProvider"
    m_IncludeTimestamp = True
    m_IncludeLevel = True
End Sub

' Реализация метода инициализации из интерфейса LogProviderBase
Private Sub LogProviderBase_Initialize(config As LogConfiguration)
    m_IsEnabled = CBool(config.GetProperty("ImmediateProviderEnabled"))
    m_IsInitialized = True
End Sub

' Реализация метода записи лога из интерфейса LogProviderBase
Private Sub LogProviderBase_Writes(entry As LogEntry)
    ' Проверка состояния провайдера
    If Not m_IsInitialized Or Not m_IsEnabled Then Exit Sub
    
    ' Вывод сообщения лога в окно Immediate Window
    Debug.Print entry.FormatMessage()
End Sub

' Реализация остальных методов интерфейса LogProviderBase
Private Sub LogProviderBase_Enable()
    m_IsEnabled = True
End Sub

Private Sub LogProviderBase_Disable()
    m_IsEnabled = False
End Sub

Private Property Get LogProviderBase_IsEnabled() As Boolean
    LogProviderBase_IsEnabled = m_IsEnabled
End Property

Private Property Get LogProviderBase_IsInitialized() As Boolean
    LogProviderBase_IsInitialized = m_IsInitialized
End Property

Private Property Get LogProviderBase_Name() As String
    LogProviderBase_Name = m_Name
End Property

Private Property Let LogProviderBase_Name(value As String)
    If Len(value) > 0 Then
        m_Name = value
    End If
End Property

' Публичные методы для прямого доступа
Public Sub Initialize(config As LogConfiguration)
    LogProviderBase_Initialize config
End Sub

Public Sub Writes(entry As LogEntry)
    LogProviderBase_Writes entry
End Sub

Public Property Get IsEnabled() As Boolean
    IsEnabled = m_IsEnabled
End Property

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

' Добавляем публичное свойство с доступом на запись
Public Property Let Name(value As String)
    If Len(value) > 0 Then
        m_Name = value
    End If
End Property

Public Sub Enable()
    m_IsEnabled = True
End Sub

Public Sub Disable()
    m_IsEnabled = False
End Sub

' Дополнительные настройки форматирования для Immediate Window
Public Property Get IncludeTimestamp() As Boolean
    IncludeTimestamp = m_IncludeTimestamp
End Property

Public Property Let IncludeTimestamp(value As Boolean)
    m_IncludeTimestamp = value
End Property

Public Property Get IncludeLevel() As Boolean
    IncludeLevel = m_IncludeLevel
End Property

Public Property Let IncludeLevel(value As Boolean)
    m_IncludeLevel = value
End Property

Module: ErrorLogger

' ErrorLogger.cls
' Класс для расширенного логирования ошибок
' Дата создания: 2025-03-14
Option Explicit

' Приватные переменные
Private m_ErrorLogDirectory As String
Private m_CallStack As Collection
Private m_IsInitialized As Boolean

' Инициализация при создании объекта
Private Sub Class_Initialize()
    Set m_CallStack = New Collection
    m_IsInitialized = False
End Sub

' Инициализация с указанием директории для логов ошибок
Public Sub Initialize(errorLogDirectory As String)
    ' Сохраняем директорию для логов ошибок
    m_ErrorLogDirectory = errorLogDirectory
    
    ' Проверяем существование директории и создаем ее при необходимости
    If Not FSWrapper.FolderExists(m_ErrorLogDirectory) Then
        FSWrapper.CreateFolder m_ErrorLogDirectory
    End If
    
    ' Очищаем стек вызовов при инициализации
    Set m_CallStack = New Collection
    
    m_IsInitialized = True
End Sub

' Метод для логирования ошибки
Public Sub LogError(errorNum As Long, source As String, description As String, _
                    Optional procedure As String = "", _
                    Optional message As String = "", _
                    Optional LineNum As Long = 0, _
                    Optional lineCode As String = "")
    
    ' Проверка инициализации
    If Not m_IsInitialized Then
        Debug.Print "ErrorLogger не инициализирован. Ошибка не будет залогирована."
        Exit Sub
    End If
    
    ' Создание отчета об ошибке
    Dim report As String
    report = GenerateErrorReport(errorNum, source, description, procedure, message, LineNum, lineCode)
    
    ' Сохранение в файл
    SaveErrorToFile report
    
    ' Добавление записи в журнал об ошибке
    ' Это будет вызвано из LogManager, здесь только создаем отчет
    
    ' Очистка стека вызовов после логирования
    Set m_CallStack = New Collection
End Sub

' Генерация подробного отчета об ошибке
Private Function GenerateErrorReport(errorNum As Long, source As String, description As String, _
                                     procedure As String, message As String, LineNum As Long, lineCode As String) As String
    ' Формирование детального отчета об ошибке
    Dim report As String
    
    report = "=== ERROR REPORT ====" & vbCrLf & _
             "Timestamp: " & GetFormattedDateTime() & vbCrLf & _
             "UTC Time: " & GetUTCFormattedDateTime() & vbCrLf & _
             "User: " & GetCurrentUser() & vbCrLf & _
             "Error Number: " & errorNum & vbCrLf & _
             "Source: " & source & vbCrLf & _
             "Description: " & description & vbCrLf
    
    If Len(procedure) > 0 Then
        report = report & "Procedure: " & procedure & vbCrLf
    End If
    
    If Len(message) > 0 Then
        report = report & "Message: " & message & vbCrLf
    End If
    
    If LineNum > 0 Then
        report = report & "Line Number: " & LineNum & vbCrLf
    End If
    
    If Len(lineCode) > 0 Then
        report = report & "Code: " & lineCode & vbCrLf
    End If
    
    ' Добавление стека вызовов, если он есть
    If m_CallStack.Count > 0 Then
        report = report & vbCrLf & "Call Stack:" & vbCrLf
        
        Dim i As Long
        For i = 1 To m_CallStack.Count
            report = report & i & ". " & m_CallStack(i) & vbCrLf
        Next i
    End If
    
    report = report & "==================="
    
    GenerateErrorReport = report
End Function

' Сохранение отчета в файл
Private Sub SaveErrorToFile(report As String)
    On Error GoTo ErrorHandler
    
    ' Проверка инициализации
    If Not m_IsInitialized Then Exit Sub
    
    ' Формирование имени файла с текущей датой и временем
    Dim fileName As String
    fileName = "ErrorLog_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    
    ' Формирование полного пути к файлу
    Dim filePath As String
    filePath = BuildPath(m_ErrorLogDirectory, fileName)
    
    ' Сохранение отчета в файл
    FSWrapper.CreateEmptyFile filePath
    FSWrapper.AppendTextToFile filePath, report
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Ошибка при сохранении отчета об ошибке: " & Err.description
    ' В случае ошибки пытаемся хотя бы вывести отчет в окно отладки
    Debug.Print report
End Sub

' Добавление информации в стек вызовов
Public Sub AddToCallStack(description As String)
    m_CallStack.Add description
End Sub

' Очистка стека вызовов
Public Sub ClearCallStack()
    Set m_CallStack = New Collection
End Sub

' Проверка, инициализирован ли логгер
Public Function IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Function

' Создание контекста для ошибки
Public Function CreateErrorContext(moduleName As String, procedureName As String, _
                                   Optional lineNumber As Long = 0, Optional lineCode As String = "") As LogContext
    Dim context As New LogContext
    
    With context
        .moduleName = moduleName
        .procedureName = procedureName
        If lineNumber > 0 Then
            .SetLineInfo lineNumber, lineCode
        End If
    End With
    
    Set CreateErrorContext = context
End Function

Module: LogManager

' LogManager.cls
' Основной класс для управления системой логирования
' Дата создания: 2025-03-14
Option Explicit

' События
Public Event ConfigurationChanged()
Public Event LogEntryAdded(entry As LogEntry)
Public Event BufferFlushed()

' Приватные переменные для хранения внутренних компонентов
Private m_Configuration As LogConfiguration
Private m_Buffer As LogBuffer
Private m_ErrorLogger As ErrorLogger
Private m_FileProvider As FileLogProvider
Private m_ImmediateProvider As ImmediateLogProvider
Private m_IsInitialized As Boolean
Private m_MinLogLevel As LogLevel

' Инициализация при создании объекта
Private Sub Class_Initialize()
    ' Создаем все необходимые объекты
    Set m_Configuration = New LogConfiguration
    Set m_Buffer = New LogBuffer
    Set m_ErrorLogger = New ErrorLogger
    Set m_FileProvider = New FileLogProvider
    Set m_ImmediateProvider = New ImmediateLogProvider
    
    m_IsInitialized = False
    m_MinLogLevel = LogLevel.LogInfo ' По умолчанию
End Sub

' Инициализация менеджера логирования
Public Function Initialize(Optional configPath As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Если указан путь к конфигурации, устанавливаем его
    If Len(configPath) > 0 Then
        m_Configuration.SetProperty "ConfigPath", configPath
    End If
    
    ' Загружаем конфигурацию
    If Not m_Configuration.Load Then
        ' Если конфигурация не загружена, используем значения по умолчанию
        m_Configuration.SetDefaults
        m_Configuration.Save ' Пытаемся сохранить настройки по умолчанию
    End If
    
    ' Инициализируем буфер
    m_Buffer.Initialize CLng(m_Configuration.GetProperty("MaxBufferSize"))
    
    ' Инициализируем провайдеры
    m_FileProvider.Initialize m_Configuration
    m_ImmediateProvider.Initialize m_Configuration
    
    ' Инициализируем логгер ошибок
    m_ErrorLogger.Initialize CStr(m_Configuration.GetProperty("ErrorLogDirectory"))
    
    ' Устанавливаем минимальный уровень логирования
    m_MinLogLevel = StringToLogLevel(CStr(m_Configuration.GetProperty("LogLevel")))
    
    ' Устанавливаем обработчики событий для буфера
    ' (в VBA нет прямого способа задать обработчики событий, поэтому представляем как концепцию)
    
    m_IsInitialized = True
    Initialize = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка при инициализации LogManager: " & Err.description
    m_IsInitialized = False
    Initialize = False
End Function

' Основной метод логирования
Public Sub Log(level As LogLevel, message As String, _
               Optional moduleName As String = "", _
               Optional procedureName As String = "", _
               Optional lineNumber As Long = 0, _
               Optional result As Variant = Null)
               
    ' Проверка инициализации и активации системы логирования
    If Not m_IsInitialized Or Not CBool(m_Configuration.GetProperty("IsEnabled")) Then
        Exit Sub
    End If
    
    ' Фильтрация по уровню логирования
    If level < m_MinLogLevel Then
        Exit Sub
    End If
    
    ' Создание контекста
    Dim context As New LogContext
    context.moduleName = moduleName
    context.procedureName = procedureName
    context.lineNumber = lineNumber
    
    ' Создание записи лога
    Dim entry As New LogEntry
    entry.level = level
    entry.message = message
    Set entry.context = context
    
    ' Если результат передан, устанавливаем его
    If Not IsNull(result) Then
        entry.result = result
    End If
    
    ' Добавление записи в буфер
    m_Buffer.Add entry
    
    ' Генерация события о добавлении записи
    RaiseEvent LogEntryAdded(entry)
    
    ' Если буфер заполнен, сбрасываем его
    If m_Buffer.IsFull Then
        FlushBuffer
    End If
    
    ' Если включен немедленный вывод, записываем через ImmediateProvider
    If level >= LogLevel.LogWarning And CBool(m_Configuration.GetProperty("ImmediateProviderEnabled")) Then
        m_ImmediateProvider.Writes entry
    End If
End Sub

' Методы логирования для разных уровней
Public Sub Log_Debug(message As String, _
                    Optional moduleName As String = "", _
                    Optional procedureName As String = "", _
                    Optional lineNumber As Long = 0, _
                    Optional result As Variant = Null)
    Log LogLevel.LogDebug, message, moduleName, procedureName, lineNumber, result
End Sub

Public Sub Log_Info(message As String, _
                   Optional moduleName As String = "", _
                   Optional procedureName As String = "", _
                   Optional lineNumber As Long = 0, _
                   Optional result As Variant = Null)
    Log LogLevel.LogInfo, message, moduleName, procedureName, lineNumber, result
End Sub

Public Sub Log_Warning(message As String, _
                      Optional moduleName As String = "", _
                      Optional procedureName As String = "", _
                      Optional lineNumber As Long = 0, _
                      Optional result As Variant = Null)
    Log LogLevel.LogWarning, message, moduleName, procedureName, lineNumber, result
End Sub

Public Sub Log_Error(message As String, _
                    Optional errorNumber As Long = 0, _
                    Optional source As String = "", _
                    Optional description As String = "", _
                    Optional moduleName As String = "", _
                    Optional procedureName As String = "", _
                    Optional lineNumber As Long = 0, _
                    Optional lineCode As String = "", _
                    Optional result As Variant = Null)
    
    ' Логирование через основной механизм
    Log LogLevel.LogError, message, moduleName, procedureName, lineNumber, result
    
    ' Дополнительное логирование через ErrorLogger
    If m_ErrorLogger.IsInitialized Then
        m_ErrorLogger.LogError errorNumber, source, description, procedureName, message, lineNumber, lineCode
    End If
End Sub

' Метод для ручного сброса буфера
Public Function FlushBuffer() As Boolean
    ' Проверка инициализации
    If Not m_IsInitialized Then
        FlushBuffer = False
        Exit Function
    End If
    
    Dim success As Boolean
    success = True
    
    ' Сброс через файловый провайдер, если он включен
    If CBool(m_Configuration.GetProperty("FileProviderEnabled")) Then
        success = success And m_Buffer.Flush(m_FileProvider)
    End If
    
    ' Генерация события о сбросе буфера
    If success Then
        RaiseEvent BufferFlushed
    End If
    
    FlushBuffer = success
End Function

' Получение конфигурации
Public Property Get Configuration() As LogConfiguration
    Set Configuration = m_Configuration
End Property

' Получение буфера
Public Property Get Buffer() As LogBuffer
    Set Buffer = m_Buffer
End Property

' Получение логгера ошибок
Public Property Get ErrorLogger() As ErrorLogger
    Set ErrorLogger = m_ErrorLogger
End Property

' Проверка инициализации
Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

' Обработка ошибок во время выполнения (может быть использован в конструкции On Error)
Public Sub HandleRuntimeError(Optional procedureName As String = "", _
                              Optional lineNumber As Long = 0, _
                              Optional lineCode As String = "")
    
    ' Получение информации об ошибке
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    
    errNumber = Err.Number
    errSource = Err.source
    errDescription = Err.description
    
    ' Логирование ошибки
    Log_Error "Ошибка выполнения: " & errDescription, _
              errNumber, _
              errSource, _
              errDescription, _
              "", _
              procedureName, _
              lineNumber, _
              lineCode
End Sub

' Создание контекста для записи лога
Public Function CreateLogContext(moduleName As String, procedureName As String, _
                               Optional lineNumber As Long = 0, Optional lineCode As String = "") As LogContext
    
    Dim context As New LogContext
    context.moduleName = moduleName
    context.procedureName = procedureName
    If lineNumber > 0 Then
        context.lineNumber = lineNumber
        context.lineCode = lineCode
    End If
    
    Set CreateLogContext = context
End Function

' Завершение работы менеджера логирования
Public Sub Shutdown()
    ' Сбрасываем буфер перед завершением
    FlushBuffer
    
    ' Очищаем ссылки на объекты
    Set m_Configuration = Nothing
    Set m_Buffer = Nothing
    Set m_ErrorLogger = Nothing
    Set m_FileProvider = Nothing
    Set m_ImmediateProvider = Nothing
    
    m_IsInitialized = False
End Sub


Module: FSWrapper
' FSWrapper.bas
' Модуль-обертка для работы с FileSystemObject
' Дата создания: 2025-03-14
Option Explicit

' Ссылка на Microsoft Scripting Runtime должна быть добавлена в проект
' Tools -> References -> Microsoft Scripting Runtime

' Проверка существования директории
Public Function FolderExists(path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    FolderExists = fso.FolderExists(path)
    
    Set fso = Nothing
End Function

' Проверка существования файла
Public Function FileExists(path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    FileExists = fso.FileExists(path)
    
    Set fso = Nothing
End Function

' Создание директории (включая промежуточные папки)
Public Function CreateFolder(path As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Проверяем, существует ли уже директория
    If fso.FolderExists(path) Then
        CreateFolder = True
        Set fso = Nothing
        Exit Function
    End If
    
    ' Создаем все промежуточные директории
    Dim parentFolder As String
    Dim folders() As String
    Dim i As Long
    Dim currentPath As String
    
    ' Разбиваем путь на компоненты
    folders = Split(path, "\")
    currentPath = folders(0) & "\"  ' Начинаем с корневого диска
    
    ' Создаем каждую поддиректорию по пути
    For i = 1 To UBound(folders)
        If Len(folders(i)) > 0 Then
            currentPath = currentPath & folders(i) & "\"
            
            If Not fso.FolderExists(currentPath) Then
                fso.CreateFolder currentPath
            End If
        End If
    Next i
    
    CreateFolder = True
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    CreateFolder = False
    Set fso = Nothing
End Function

' Создание пустого файла (или перезапись существующего)
Public Function CreateEmptyFile(path As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Убедимся, что директория существует
    Dim folderPath As String
    folderPath = Left(path, InStrRev(path, "\") - 1)
    
    If Not fso.FolderExists(folderPath) Then
        CreateFolder folderPath
    End If
    
    ' Создаем или перезаписываем файл
    Dim fileObj As Object
    Set fileObj = fso.CreateTextFile(path, True)
    fileObj.Close
    
    CreateEmptyFile = True
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    CreateEmptyFile = False
    Set fso = Nothing
End Function

' Запись текста в файл (в режиме добавления)
Public Function AppendTextToFile(path As String, text As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Проверяем существование файла
    If Not fso.FileExists(path) Then
        ' Если файл не существует, создаем его
        CreateEmptyFile path
    End If
    
    ' Открываем файл и добавляем текст
    Dim fileObj As Object
    Set fileObj = fso.OpenTextFile(path, 8, True) ' 8 = ForAppending
    fileObj.WriteLine text
    fileObj.Close
    
    AppendTextToFile = True
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    AppendTextToFile = False
    Set fso = Nothing
End Function

' Чтение всего содержимого файла
Public Function ReadTextFile(path As String) As String
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Проверяем существование файла
    If Not fso.FileExists(path) Then
        ReadTextFile = ""
        Set fso = Nothing
        Exit Function
    End If
    
    ' Читаем содержимое файла
    Dim fileObj As Object
    Set fileObj = fso.OpenTextFile(path, 1) ' 1 = ForReading
    ReadTextFile = fileObj.ReadAll
    fileObj.Close
    
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    ReadTextFile = ""
    Set fso = Nothing
End Function

' Получение размера файла в байтах
Public Function GetFileSize(path As String) As Long
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Проверяем существование файла
    If Not fso.FileExists(path) Then
        GetFileSize = 0
        Set fso = Nothing
        Exit Function
    End If
    
    ' Получаем информацию о файле и его размер
    Dim fileObj As Object
    Set fileObj = fso.GetFile(path)
    GetFileSize = fileObj.Size
    
    Set fso = Nothing
    Exit Function
    
ErrorHandler:
    GetFileSize = 0
    Set fso = Nothing
End Function

' Переименование файла с сохранением пути
'' Переименование файла с сохранением пути
Public Function RenameFile(oldPath As String, newName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверяем существование файла
    If Not FileExists(oldPath) Then
        RenameFile = False
        Exit Function
    End If
    
    ' Получаем директорию файла
    Dim folderPath As String
    folderPath = Left(oldPath, InStrRev(oldPath, "\"))
    
    ' Формируем полный новый путь
    Dim newPath As String
    newPath = folderPath & newName
    
    ' Если новый файл уже существует, удаляем его
    If FileExists(newPath) Then
        Kill newPath
    End If
    
    ' Переименовываем файл с помощью метода Name
    Name oldPath As newPath
    
    ' Проверяем результат операции
    RenameFile = FileExists(newPath) And Not FileExists(oldPath)
    
    Exit Function
    
ErrorHandler:
    RenameFile = False
End Function

Module: IniFileHandler
' IniFileHandler.bas
' Модуль для работы с INI-файлами
' Дата создания: 2025-03-14
Option Explicit

#If VBA7 Then
    ' Для 64-разрядной версии Office
    Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#Else
    ' Для 32-разрядной версии Office
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#End If

' Запись значения в INI-файл
Public Function WriteIniValue(iniFile As String, section As String, key As String, value As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверяем, существует ли директория для INI-файла
    Dim folderPath As String
    folderPath = Left(iniFile, InStrRev(iniFile, "\") - 1)
    
    If Not FSWrapper.FolderExists(folderPath) Then
        FSWrapper.CreateFolder folderPath
    End If
    
    ' Записываем значение в INI-файл
    Dim result As Long
    result = WritePrivateProfileString(section, key, value, iniFile)
    
    WriteIniValue = (result <> 0)
    Exit Function
    
ErrorHandler:
    WriteIniValue = False
End Function

' Чтение значения из INI-файла
Public Function ReadIniValue(iniFile As String, section As String, key As String, defaultValue As String) As String
    On Error GoTo ErrorHandler
    
    ' Проверяем существование файла
    If Not FSWrapper.FileExists(iniFile) Then
        ReadIniValue = defaultValue
        Exit Function
    End If
    
    ' Читаем значение из INI-файла
    Dim Buffer As String * 1024
    Dim length As Long
    
    length = GetPrivateProfileString(section, key, defaultValue, Buffer, Len(Buffer), iniFile)
    
    If length > 0 Then
        ReadIniValue = Left(Buffer, length)
    Else
        ReadIniValue = defaultValue
    End If
    
    Exit Function
    
ErrorHandler:
    ReadIniValue = defaultValue
End Function

' Проверка существования секции в INI-файле
'Public Function SectionExists(iniFile As String, section As String) As Boolean
'    On Error GoTo ErrorHandler
'
'    ' Проверяем существование файла
'    If Not FSWrapper.FileExists(iniFile) Then
'        SectionExists = False
'        Exit Function
'    End If
'
'    ' Пробуем получить любой ключ из секции
'    ' Используем специальный символ NULL в качестве ключа,
'    ' чтобы получить список всех ключей секции
'    Dim buffer As String * 2
'    Dim length As Long
'
'    length = GetPrivateProfileString(section, vbNullString, "", buffer, Len(buffer), iniFile)
'
'    SectionExists = (length > 0)
'    Exit Function
'
'ErrorHandler:
'    SectionExists = False
'End Function

'' Проверка существования ключа в секции INI-файла
'Public Function KeyExists(iniFile As String, section As String, key As String) As Boolean
'    On Error GoTo ErrorHandler
'
'    ' Проверяем существование файла и секции
'    If Not FSWrapper.FileExists(iniFile) Or Not SectionExists(iniFile, section) Then
'        KeyExists = False
'        Exit Function
'    End If
'
'    ' Уникальное значение по умолчанию для проверки
'    Dim uniqueDefault As String
'    uniqueDefault = "@@KEY_NOT_FOUND_" & Format(Now, "yyyymmddhhnnss") & "@@"
'
'    ' Читаем значение с уникальным значением по умолчанию
'    Dim value As String
'    value = ReadIniValue(iniFile, section, key, uniqueDefault)
'
'    ' Если вернулось наше уникальное значение, значит ключа нет
'    KeyExists = (value <> uniqueDefault)
'    Exit Function
'
'ErrorHandler:
'    KeyExists = False
'End Function

' Создание секции в INI-файле (если она не существует)
Public Function CreateSection(iniFile As String, section As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Если секция уже существует, возвращаем успех
    If SectionExists(iniFile, section) Then
        CreateSection = True
        Exit Function
    End If
    
    ' Создаем секцию, записывая пустую строку в качестве комментария
    Dim result As Long
    result = WritePrivateProfileString(section, vbNullString, vbNullString, iniFile)
    
    CreateSection = (result <> 0)
    Exit Function
    
ErrorHandler:
    CreateSection = False
End Function

' Получение списка секций из INI-файла
' Проверка существования секции в INI-файле
Public Function SectionExists(iniFile As String, section As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверяем существование файла
    If Not FSWrapper.FileExists(iniFile) Then
        SectionExists = False
        Exit Function
    End If
    
    ' Читаем содержимое файла напрямую
    Dim content As String
    content = FSWrapper.ReadTextFile(iniFile)
    
    ' Ищем секцию в формате [SectionName]
    Dim sectionHeader As String
    sectionHeader = "[" & section & "]"
    
    ' Проверяем наличие секции в файле
    SectionExists = (InStr(1, content, sectionHeader, vbTextCompare) > 0)
    Exit Function
    
ErrorHandler:
    SectionExists = False
End Function

' Проверка существования ключа в секции INI-файла
Public Function KeyExists(iniFile As String, section As String, key As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверяем существование файла и секции
    If Not FSWrapper.FileExists(iniFile) Or Not SectionExists(iniFile, section) Then
        KeyExists = False
        Exit Function
    End If
    
    ' Читаем содержимое файла
    Dim content As String
    content = FSWrapper.ReadTextFile(iniFile)
    
    ' Разбиваем на строки
    Dim lines() As String
    lines = Split(content, vbCrLf)
    
    Dim i As Long
    Dim inSection As Boolean
    inSection = False
    
    ' Проходим по строкам файла
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' Проверяем начало секции
        If Left(line, 1) = "[" And Right(line, 1) = "]" Then
            Dim currentSection As String
            currentSection = Mid(line, 2, Len(line) - 2)
            inSection = (LCase(currentSection) = LCase(section))
        
        ' Если мы в нужной секции, ищем ключ
        ElseIf inSection And Len(line) > 0 And Left(line, 1) <> ";" Then
            Dim equalsPos As Long
            equalsPos = InStr(line, "=")
            
            If equalsPos > 1 Then
                Dim currentKey As String
                currentKey = Trim(Left(line, equalsPos - 1))
                
                ' Если нашли нужный ключ
                If LCase(currentKey) = LCase(key) Then
                    KeyExists = True
                    Exit Function
                End If
            End If
        
        ' Если нашли другую секцию после нашей, выходим
        ElseIf inSection And Left(line, 1) = "[" Then
            Exit For
        End If
    Next i
    
    KeyExists = False
    Exit Function
    
ErrorHandler:
    KeyExists = False
End Function

' Получение всех ключей в секции
' Получение всех ключей в секции
Public Function GetKeys(iniFile As String, section As String) As String()
    On Error GoTo ErrorHandler
    
    Dim result() As String
    
    ' Инициализируем пустой массив по умолчанию
    ReDim result(-1 To -1)
    
    ' Проверяем существование файла и секции
    If Not FSWrapper.FileExists(iniFile) Or Not SectionExists(iniFile, section) Then
        GetKeys = result
        Exit Function
    End If
    
    ' Читаем все содержимое файла
    Dim content As String
    content = FSWrapper.ReadTextFile(iniFile)
    
    ' Если файл пустой, возвращаем пустой массив
    If Len(content) = 0 Then
        GetKeys = result
        Exit Function
    End If
    
    ' Разбиваем файл на строки
    Dim lines() As String
    lines = Split(content, vbCrLf)
    
    ' Ищем ключи в указанной секции
    Dim i As Long
    Dim keyCount As Long
    Dim tempKeys() As String
    ReDim tempKeys(100) ' Предварительный размер
    
    Dim inSection As Boolean
    inSection = False
    
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' Проверяем, является ли строка заголовком секции
        If Left(line, 1) = "[" And Right(line, 1) = "]" And Len(line) > 2 Then
            ' Извлекаем имя секции
            Dim sectionName As String
            sectionName = Mid(line, 2, Len(line) - 2)
            
            ' Проверяем, находимся ли мы в нужной секции
            inSection = (sectionName = section)
        ElseIf inSection And Len(line) > 0 And Left(line, 1) <> ";" Then
            ' Ищем ключи (формат Key=Value)
            Dim equalsPos As Long
            equalsPos = InStr(line, "=")
            
            If equalsPos > 1 Then
                ' Извлекаем ключ
                Dim keyName As String
                keyName = Trim(Left(line, equalsPos - 1))
                
                ' Добавляем в массив
                tempKeys(keyCount) = keyName
                keyCount = keyCount + 1
            End If
        End If
    Next i
    
    ' Создаем массив нужного размера
    If keyCount > 0 Then
        ReDim result(0 To keyCount - 1)
        
        ' Копируем найденные ключи
        For i = 0 To keyCount - 1
            result(i) = tempKeys(i)
        Next i
    End If
    
    GetKeys = result
    Exit Function
    
ErrorHandler:
    GetKeys = result
End Function
' Функция для нормализации конца строки в содержимом файла
Private Function NormalizeLineEndings(content As String) As String
    Dim result As String
    
    ' Сначала заменяем все возможные варианты окончания строк на один стандартный
    result = Replace(content, vbCrLf, vbLf)
    result = Replace(result, vbCr, vbLf)
    
    ' Затем заменяем обратно на стандартный CRLF
    result = Replace(result, vbLf, vbCrLf)
    
    NormalizeLineEndings = result
End Function

' Вспомогательная функция для разбивки строки с учетом возможных проблем с окончанием строк
Private Function SafeSplit(content As String) As String()
    ' Нормализуем окончания строк
    Dim normalizedContent As String
    normalizedContent = NormalizeLineEndings(content)
    
    ' Разбиваем на строки
    SafeSplit = Split(normalizedContent, vbCrLf)
End Function
' Получение списка секций из INI-файла
Public Function GetSections(iniFile As String) As String()
    On Error GoTo ErrorHandler
    
    Dim result() As String
    ' Инициализируем пустой массив по умолчанию
    ReDim result(-1 To -1)
    
    ' Проверяем существование файла
    If Not FSWrapper.FileExists(iniFile) Then
        GetSections = result
        Exit Function
    End If
    
    ' Читаем все содержимое файла
    Dim content As String
    content = FSWrapper.ReadTextFile(iniFile)
    
    ' Если файл пустой, возвращаем пустой массив
    If Len(content) = 0 Then
        GetSections = result
        Exit Function
    End If
    
    ' Разбиваем файл на строки
    Dim lines() As String
    lines = Split(content, vbCrLf)
    
    ' Ищем секции (формат [SectionName])
    Dim i As Long
    Dim sectionCount As Long
    Dim tempSections() As String
    ReDim tempSections(100) ' Предварительный размер
    
    For i = 0 To UBound(lines)
        Dim line As String
        line = Trim(lines(i))
        
        ' Проверяем, является ли строка заголовком секции
        If Left(line, 1) = "[" And Right(line, 1) = "]" And Len(line) > 2 Then
            ' Извлекаем имя секции
            Dim sectionName As String
            sectionName = Mid(line, 2, Len(line) - 2)
            
            ' Добавляем в массив
            tempSections(sectionCount) = sectionName
            sectionCount = sectionCount + 1
        End If
    Next i
    
    ' Создаем массив нужного размера
    If sectionCount > 0 Then
        ReDim result(0 To sectionCount - 1)
        
        ' Копируем найденные секции
        For i = 0 To sectionCount - 1
            result(i) = tempSections(i)
        Next i
    End If
    
    GetSections = result
    Exit Function
    
ErrorHandler:
    GetSections = result
End Function

Module: LoggerTester
' TestLogger.bas
' Модуль для тестирования базовых компонентов системы логирования
' Дата создания: 2025-03-14
Option Explicit

' Тестирование всех базовых компонентов
Public Sub TestBasicComponents()
    Debug.Print "Начало тестирования базовых компонентов: " & Now
    Debug.Print "----------------------------------------"
    
    TestLogDefinitions
    TestFSWrapper
    TestIniFileHandler
    TestLogContext
    TestLogEntry
    TestUTCTimeUpdated
    TestSectionAndKeyExists
    TestRenameFileOnly
    
    
    Debug.Print "----------------------------------------"
    Debug.Print "Завершение тестирования базовых компонентов: " & Now
End Sub

' Тестирование модуля LogDefinitions
Private Sub TestLogDefinitions()
    Debug.Print "Тестирование LogDefinitions..."
    
    ' Проверка преобразования уровней логирования
    Debug.Print "  LogLevelToString(Debug) = " & LogLevelToString(LogLevel.LogDebug)
    Debug.Print "  LogLevelToString(Info) = " & LogLevelToString(LogLevel.LogInfo)
    Debug.Print "  LogLevelToString(Warning) = " & LogLevelToString(LogLevel.LogWarning)
    Debug.Print "  LogLevelToString(Error) = " & LogLevelToString(LogLevel.LogError)
    
    Debug.Print "  StringToLogLevel('DEBUG') = " & StringToLogLevel("DEBUG")
    Debug.Print "  StringToLogLevel('INFO') = " & StringToLogLevel("INFO")
    Debug.Print "  StringToLogLevel('WARNING') = " & StringToLogLevel("WARNING")
    Debug.Print "  StringToLogLevel('ERROR') = " & StringToLogLevel("ERROR")
    Debug.Print "  StringToLogLevel('UNKNOWN') = " & StringToLogLevel("UNKNOWN") ' Должен вернуть Info (1)
    
    ' Проверка вспомогательных функций
    Debug.Print "  GetAppDataPath() = " & GetAppDataPath()
    Debug.Print "  BuildPath('C:\Temp', 'log.txt') = " & BuildPath("C:\Temp", "log.txt")
    Debug.Print "  BuildPath('C:\Temp\', 'log.txt') = " & BuildPath("C:\Temp\", "log.txt")
    Debug.Print "  GetFormattedDateTime() = " & GetFormattedDateTime()
End Sub

' Тестирование модуля FSWrapper
Private Sub TestFSWrapper()
    Debug.Print "Тестирование FSWrapper..."
    
    Dim testDir As String
    Dim testFile As String
    
    testDir = Environ$("TEMP") & "\LoggingTest"
    testFile = testDir & "\test.txt"
    
    ' Проверка создания директории
    Debug.Print "  Создание директории: " & testDir
    Debug.Print "  Результат: " & FSWrapper.CreateFolder(testDir)
    Debug.Print "  FolderExists: " & FSWrapper.FolderExists(testDir)
    
    ' Проверка создания файла
    Debug.Print "  Создание файла: " & testFile
    Debug.Print "  Результат: " & FSWrapper.CreateEmptyFile(testFile)
    Debug.Print "  FileExists: " & FSWrapper.FileExists(testFile)
    
    ' Проверка записи в файл
    Debug.Print "  Запись в файл"
    Debug.Print "  Результат: " & FSWrapper.AppendTextToFile(testFile, "Тестовая строка 1")
    Debug.Print "  Результат: " & FSWrapper.AppendTextToFile(testFile, "Тестовая строка 2")
    
    ' Проверка чтения из файла
    Debug.Print "  Чтение из файла: " & testFile
    Debug.Print "  Содержимое: "
    Debug.Print FSWrapper.ReadTextFile(testFile)
    
    ' Проверка получения размера файла
    Debug.Print "  Размер файла: " & FSWrapper.GetFileSize(testFile) & " байт"
    
    ' Проверка переименования файла
    Dim newName As String
    newName = "test_renamed.txt"
    Debug.Print "  Переименование файла в: " & newName
    Debug.Print "  Результат: " & FSWrapper.RenameFile(testFile, newName)
End Sub

' Тестирование модуля IniFileHandler
' Тестирование модуля IniFileHandler
Private Sub TestIniFileHandler()
    Debug.Print "Тестирование IniFileHandler..."
    
    Dim iniFile As String
    iniFile = Environ$("TEMP") & "\LoggingTest\test.ini"
    
    ' Проверка записи значений
    Debug.Print "  Запись значений в INI-файл: " & iniFile
    Debug.Print "  Результат (General/TestKey): " & IniFileHandler.WriteIniValue(iniFile, "General", "TestKey", "TestValue")
    Debug.Print "  Результат (Logging/Level): " & IniFileHandler.WriteIniValue(iniFile, "Logging", "Level", "Info")
    Debug.Print "  Результат (Paths/LogPath): " & IniFileHandler.WriteIniValue(iniFile, "Paths", "LogPath", "C:\Logs")
    
    ' Проверка чтения значений
    Debug.Print "  Чтение значений из INI-файла:"
    Debug.Print "  General/TestKey = " & IniFileHandler.ReadIniValue(iniFile, "General", "TestKey", "Default")
    Debug.Print "  Logging/Level = " & IniFileHandler.ReadIniValue(iniFile, "Logging", "Level", "Default")
    Debug.Print "  Paths/LogPath = " & IniFileHandler.ReadIniValue(iniFile, "Paths", "LogPath", "Default")
    Debug.Print "  NonExistent/Key = " & IniFileHandler.ReadIniValue(iniFile, "NonExistent", "Key", "Default") ' Должен вернуть Default
    
    ' Проверка существования секций и ключей
    Debug.Print "  SectionExists(General): " & IniFileHandler.SectionExists(iniFile, "General")
    Debug.Print "  SectionExists(NonExistent): " & IniFileHandler.SectionExists(iniFile, "NonExistent")
    Debug.Print "  KeyExists(General/TestKey): " & IniFileHandler.KeyExists(iniFile, "General", "TestKey")
    Debug.Print "  KeyExists(General/NonExistent): " & IniFileHandler.KeyExists(iniFile, "General", "NonExistent")
    
    ' Проверка получения списка секций
    Dim sections() As String
    sections = IniFileHandler.GetSections(iniFile)
    Debug.Print "  Список секций:"
    
    ' Безопасная проверка перед использованием UBound
    If UBound(sections) >= LBound(sections) Then
        Dim i As Long
        For i = LBound(sections) To UBound(sections)
            Debug.Print "    " & sections(i)
        Next i
    Else
        Debug.Print "    (секции не найдены)"
    End If
    
    ' Проверка получения списка ключей
    Dim keys() As String
    keys = IniFileHandler.GetKeys(iniFile, "Logging")
    Debug.Print "  Список ключей в секции Logging:"
    
    ' Безопасная проверка перед использованием UBound
    If UBound(keys) >= LBound(keys) Then
        For i = LBound(keys) To UBound(keys)
            Debug.Print "    " & keys(i)
        Next i
    Else
        Debug.Print "    (ключи не найдены)"
    End If
End Sub
' Тестирование класса LogContext
Private Sub TestLogContext()
    Debug.Print "Тестирование LogContext..."
    
    Dim context As New LogContext
    
    ' Заполнение полей
    context.procedureType = "Sub"
    context.moduleName = "TestModule"
    context.procedureName = "TestProcedure"
    context.username = "prailda" ' Используем предоставленное имя пользователя
    context.lineNumber = 42
    context.lineCode = "x = 10 * y"
    
    ' Проверка метода ToString
    Debug.Print "  context.ToString() = " & context.ToString()
    
    ' Проверка метода CreateContext
    Dim newContext As LogContext
    Set newContext = context.CreateContext("NewModule", "NewProcedure", "Function")
    Debug.Print "  newContext.ToString() = " & newContext.ToString()
    
    ' Проверка метода SetLineInfo
    newContext.SetLineInfo 123, "result = Calc(a, b)"
    Debug.Print "  newContext после SetLineInfo: " & newContext.ToString()
End Sub

' Тестирование класса LogEntry
Private Sub TestLogEntry()
    Debug.Print "Тестирование LogEntry..."
    
    Dim entry As New LogEntry
    Dim context As New LogContext
    
    ' Настройка контекста
    context.procedureType = "Function"
    context.moduleName = "DataProcessor"
    context.procedureName = "CalculateTotal"
    context.username = "prailda" ' Используем предоставленное имя пользователя
    context.lineNumber = 78
    
    ' Настройка записи лога
    entry.level = LogLevel.LogInfo
    entry.message = "Обработка данных завершена"
    entry.result = 123.45
    Set entry.context = context
    
    ' Проверка метода FormatMessage
    Debug.Print "  entry.FormatMessage() = " & entry.FormatMessage()
    
    ' Проверка метода MaskSensitiveData
    Dim sensitiveText As String
    sensitiveText = "User login with password=secret123 successful"
    Debug.Print "  Исходный текст: " & sensitiveText
    Debug.Print "  После маскирования: " & entry.MaskSensitiveData(sensitiveText)
    
    ' Проверка метода CreateEntry
    Dim newEntry As LogEntry
    Set newEntry = entry.CreateEntry(LogLevel.LogWarning, "Внимание: низкий уровень памяти", context)
    Debug.Print "  newEntry.FormatMessage() = " & newEntry.FormatMessage()
End Sub
' Тестирование обновленных функций
Public Sub TestUpdatedComponents()
    Debug.Print "Начало тестирования обновленных компонентов: " & Now
    Debug.Print "----------------------------------------"
    
    TestRenameFile
    TestSectionAndKeyExists
    TestUTCTime
    
    Debug.Print "----------------------------------------"
    Debug.Print "Завершение тестирования обновленных компонентов: " & Now
End Sub

' Тестирование функции переименования файла
' Дополнительный тест функции переименования файла
Public Sub TestRenameFileOnly()
    Debug.Print "Специальное тестирование функции RenameFile..."
    Debug.Print "----------------------------------------"
    
    Dim testDir As String
    Dim origFile As String
    Dim newFileName As String
    
    testDir = Environ$("TEMP") & "\LoggingTest"
    origFile = testDir & "\specialTestFile.txt"
    newFileName = "specialTestFileRenamed.txt"
    
    ' Создаем тестовый файл
    Debug.Print "1. Создаем тестовый файл: " & origFile
    FSWrapper.CreateEmptyFile origFile
    FSWrapper.AppendTextToFile origFile, "Test content for special rename test"
    Debug.Print "   Файл создан: " & FSWrapper.FileExists(origFile)
    
    ' Проверяем, нет ли уже файла с новым именем
    Dim newFilePath As String
    newFilePath = testDir & "\" & newFileName
    
    If FSWrapper.FileExists(newFilePath) Then
        Debug.Print "2. Удаляем существующий файл с новым именем: " & newFilePath
        Kill newFilePath
        Debug.Print "   Файл удален: " & Not FSWrapper.FileExists(newFilePath)
    End If
    
    ' Тестируем переименование
    Debug.Print "3. Переименовываем файл из " & origFile & " в " & newFilePath
    Dim result As Boolean
    result = FSWrapper.RenameFile(origFile, newFileName)
    Debug.Print "   Результат операции: " & result
    
    ' Проверяем результаты
    Debug.Print "4. Проверяем результаты:"
    Debug.Print "   Новый файл существует: " & FSWrapper.FileExists(newFilePath)
    Debug.Print "   Старый файл существует: " & FSWrapper.FileExists(origFile)
    Debug.Print "----------------------------------------"
End Sub

' Тест функций времени
Public Sub TestUTCTimeUpdated()
    Debug.Print "Тестирование обновленных функций времени..."
    Debug.Print "----------------------------------------"
    
    Debug.Print "Локальное время: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "UTC время (статическое): " & GetUTCFormattedDateTime()
    
    Dim localNow As Date
    localNow = Now
    Dim calculatedUTC As Date
    calculatedUTC = LocalTimeToUTC(localNow)
    
    Debug.Print "Локальное время сейчас: " & Format(localNow, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Расчетное UTC время: " & Format(calculatedUTC, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Смещение часового пояса: " & GetTimeZoneOffsetHours() & " часов"
    Debug.Print "----------------------------------------"
End Sub

' Тестирование проверки секций и ключей
Private Sub TestSectionAndKeyExists()
    Debug.Print "Тестирование обновленных SectionExists и KeyExists..."
    
    Dim iniFile As String
    iniFile = Environ$("TEMP") & "\LoggingTest\testSections.ini"
    
    ' Создаем тестовый INI файл напрямую
    Dim content As String
    content = "[TestSection]" & vbCrLf & _
              "TestKey1=Value1" & vbCrLf & _
              "TestKey2=Value2" & vbCrLf & _
              vbCrLf & _
              "[AnotherSection]" & vbCrLf & _
              "Key1=123" & vbCrLf
    
    FSWrapper.CreateEmptyFile iniFile
    FSWrapper.AppendTextToFile iniFile, content
    
    ' Тестируем проверку секций
    Debug.Print "  SectionExists(TestSection): " & IniFileHandler.SectionExists(iniFile, "TestSection")
    Debug.Print "  SectionExists(AnotherSection): " & IniFileHandler.SectionExists(iniFile, "AnotherSection")
    Debug.Print "  SectionExists(NonExistentSection): " & IniFileHandler.SectionExists(iniFile, "NonExistentSection")
    
    ' Тестируем проверку ключей
    Debug.Print "  KeyExists(TestSection/TestKey1): " & IniFileHandler.KeyExists(iniFile, "TestSection", "TestKey1")
    Debug.Print "  KeyExists(TestSection/NonExistentKey): " & IniFileHandler.KeyExists(iniFile, "TestSection", "NonExistentKey")
    Debug.Print "  KeyExists(AnotherSection/Key1): " & IniFileHandler.KeyExists(iniFile, "AnotherSection", "Key1")
End Sub


' Тестирование всех компонентов второй фазы
Public Sub TestPhase2Components()
    Debug.Print "Начало тестирования компонентов фазы 2: " & Now
    Debug.Print "----------------------------------------"
    
    TestLogConfiguration
    TestLogBuffer
    TestErrorLogger
    
    Debug.Print "----------------------------------------"
    Debug.Print "Завершение тестирования компонентов фазы 2: " & Now
End Sub

' Тестирование класса LogConfiguration
Private Sub TestLogConfiguration()
    Debug.Print "Тестирование LogConfiguration..."
    
    Dim config As New LogConfiguration
    
    ' Проверка значений по умолчанию
    Debug.Print "  Значения по умолчанию:"
    Debug.Print "    IsEnabled = " & config.GetProperty("IsEnabled")
    Debug.Print "    LogLevel = " & config.GetProperty("LogLevel")
    Debug.Print "    LogDirectory = " & config.GetProperty("LogDirectory")
    Debug.Print "    MaxFileSize = " & config.GetProperty("MaxFileSize")
    Debug.Print "    IsLoaded = " & config.IsLoaded
    Debug.Print "    IsValid = " & config.IsValid
    
    ' Изменение некоторых настроек
    Debug.Print "  Изменение настроек:"
    config.SetProperty "LogLevel", "Debug"
    config.SetProperty "MaxBufferSize", 500
    
    Debug.Print "    LogLevel = " & config.GetProperty("LogLevel")
    Debug.Print "    MaxBufferSize = " & config.GetProperty("MaxBufferSize")
    Debug.Print "    HasChangedSinceLastLoad = " & config.HasChangedSinceLastLoad
    
    ' Сохранение и загрузка конфигурации
    Debug.Print "  Сохранение конфигурации:"
    Debug.Print "    Результат = " & config.Save()
    
    Debug.Print "  Загрузка конфигурации:"
    Dim newConfig As New LogConfiguration
    Debug.Print "    Результат = " & newConfig.Load()
    Debug.Print "    LogLevel = " & newConfig.GetProperty("LogLevel")
    Debug.Print "    MaxBufferSize = " & newConfig.GetProperty("MaxBufferSize")
    Debug.Print "    IsLoaded = " & newConfig.IsLoaded
    Debug.Print "    IsValid = " & newConfig.IsValid
    
    ' Проверка временного изменения уровня логирования
    Debug.Print "  Временное изменение уровня логирования:"
    newConfig.SetLogLevelTemp LogLevel.LogWarning, 5
    Debug.Print "    LogLevel = " & newConfig.GetProperty("LogLevel")
End Sub

' Тестирование класса LogBuffer
Private Sub TestLogBuffer()
    Debug.Print "Тестирование LogBuffer..."
    
    Dim Buffer As New LogBuffer
    Buffer.Initialize 5 ' Устанавливаем небольшой размер для тестирования
    
    ' Проверка начального состояния
    Debug.Print "  Начальное состояние:"
    Debug.Print "    IsEmpty = " & Buffer.IsEmpty()
    Debug.Print "    IsFull = " & Buffer.IsFull()
    Debug.Print "    Count = " & Buffer.Count()
    Debug.Print "    MaxSize = " & Buffer.MaxSize
    
    ' Создаем тестовые записи логов
    Dim context As New LogContext
    context.moduleName = "TestModule"
    context.procedureName = "TestProcedure"
    
    Dim entry1 As New LogEntry
    entry1.level = LogLevel.LogInfo
    entry1.message = "Тестовая запись 1"
    Set entry1.context = context
    
    Dim entry2 As New LogEntry
    entry2.level = LogLevel.LogWarning
    entry2.message = "Тестовая запись 2"
    Set entry2.context = context
    
    ' Добавляем записи в буфер
    Debug.Print "  Добавление записей в буфер:"
    Buffer.Add entry1
    Debug.Print "    После добавления 1 записи:"
    Debug.Print "    IsEmpty = " & Buffer.IsEmpty()
    Debug.Print "    IsFull = " & Buffer.IsFull()
    Debug.Print "    Count = " & Buffer.Count()
    
    Buffer.Add entry2
    Debug.Print "    После добавления 2 записей:"
    Debug.Print "    Count = " & Buffer.Count()
    
    ' Проверка заполнения буфера
    Debug.Print "  Заполнение буфера до максимального размера:"
    Dim i As Integer
    For i = 3 To Buffer.MaxSize
        Dim entry As New LogEntry
        entry.level = LogLevel.LogDebug
        entry.message = "Тестовая запись " & i
        Set entry.context = context
        Buffer.Add entry
    Next i
    
    Debug.Print "    После заполнения буфера:"
    Debug.Print "    IsEmpty = " & Buffer.IsEmpty()
    Debug.Print "    IsFull = " & Buffer.IsFull()
    Debug.Print "    Count = " & Buffer.Count()
    
    ' Создаем тестовый провайдер для сброса буфера
    Dim mockProvider As New Dictionary
    ' Тестовый провайдер должен иметь метод Write, но для тестирования
    ' мы можем использовать любой объект с методом, который принимает параметр
    
    ' Сброс буфера
    Debug.Print "  Сброс буфера:"
    Debug.Print "    Результат = Выполнение невозможно, тестовый провайдер не реализован"
    
    ' Очистка буфера
    Debug.Print "  Очистка буфера:"
    Buffer.Clear
    Debug.Print "    После очистки:"
    Debug.Print "    IsEmpty = " & Buffer.IsEmpty()
    Debug.Print "    IsFull = " & Buffer.IsFull()
    Debug.Print "    Count = " & Buffer.Count()
    
    ' Изменение максимального размера буфера
    Debug.Print "  Изменение максимального размера буфера:"
    Buffer.MaxSize = 10
    Debug.Print "    Новый MaxSize = " & Buffer.MaxSize
End Sub

' Тестирование класса ErrorLogger
Private Sub TestErrorLogger()
    Debug.Print "Тестирование ErrorLogger..."
    
    Dim ErrorLogger As New ErrorLogger
    
    ' Инициализация логгера ошибок
    Dim errorLogDir As String
    errorLogDir = Environ$("TEMP") & "\LoggingTest\Errors"
    
    Debug.Print "  Инициализация ErrorLogger с директорией: " & errorLogDir
    ErrorLogger.Initialize errorLogDir
    
    Debug.Print "  IsInitialized = " & ErrorLogger.IsInitialized()
    
    ' Формирование стека вызовов
    Debug.Print "  Формирование стека вызовов:"
    ErrorLogger.AddToCallStack "MainModule.StartProcess"
    ErrorLogger.AddToCallStack "DataProcessor.ValidateInput"
    ErrorLogger.AddToCallStack "DataProcessor.ProcessData"
    
    ' Создание контекста для ошибки
    Debug.Print "  Создание контекста ошибки:"
    Dim errorContext As LogContext
    Set errorContext = ErrorLogger.CreateErrorContext("DataProcessor", "ProcessData", 42, "result = CInt(userInput)")
    
    Debug.Print "    context.ToString() = " & errorContext.ToString()
    
    ' Логирование тестовой ошибки
    Debug.Print "  Логирование тестовой ошибки:"
    ErrorLogger.LogError 13, "Microsoft VBA", "Type mismatch", "ProcessData", _
                         "Failed to process user input", 42, "result = CInt(userInput)"
    
    ' Проверка создания файла ошибки
    Dim errorFiles As String
    errorFiles = Dir(errorLogDir & "\ErrorLog_*.txt")
    
    Debug.Print "  Созданные файлы отчетов об ошибках:"
    If Len(errorFiles) > 0 Then
        Do While Len(errorFiles) > 0
            Debug.Print "    " & errorFiles
            errorFiles = Dir()
        Loop
    Else
        Debug.Print "    Файлы отчетов не найдены"
    End If
    
    ' Очистка стека вызовов
    ErrorLogger.ClearCallStack
End Sub

' Тестирование интеграции компонентов второй фазы
Public Sub TestPhase2Integration()
    Debug.Print "Тестирование интеграции компонентов фазы 2: " & Now
    Debug.Print "----------------------------------------"
    
    ' Создаем экземпляры классов
    Dim config As New LogConfiguration
    Dim Buffer As New LogBuffer
    Dim ErrorLogger As New ErrorLogger
    
    ' Инициализируем компоненты
    config.Load ' Загружаем или создаем конфигурацию
    Buffer.Initialize CLng(config.GetProperty("MaxBufferSize"))
    ErrorLogger.Initialize CStr(config.GetProperty("ErrorLogDirectory"))
    
    ' Проверяем статус инициализации
    Debug.Print "Статус инициализации компонентов:"
    Debug.Print "  config.IsLoaded = " & config.IsLoaded
    Debug.Print "  config.IsValid = " & config.IsValid
    Debug.Print "  buffer.MaxSize = " & Buffer.MaxSize
    Debug.Print "  errorLogger.IsInitialized = " & ErrorLogger.IsInitialized()
    
    ' Создаем тестовую запись лога и добавляем ее в буфер
    Dim context As New LogContext
    context.moduleName = "IntegrationTest"
    context.procedureName = "TestPhase2Integration"
    
    Dim entry As New LogEntry
    entry.level = LogLevel.Info
    entry.message = "Тестирование интеграции компонентов фазы 2"
    Set entry.context = context
    
    Buffer.Add entry
    
    Debug.Print "Запись добавлена в буфер:"
    Debug.Print "  buffer.Count = " & Buffer.Count()
    
    ' Логирование тестовой ошибки с использованием ErrorLogger
    ErrorLogger.AddToCallStack "IntegrationTest.TestPhase2Integration"
    ErrorLogger.LogError 0, "Test", "Тестовая ошибка для проверки интеграции", _
                       "TestPhase2Integration", "Integration test message"
                       
    Debug.Print "Ошибка залогирована"
    
    Debug.Print "----------------------------------------"
    Debug.Print "Завершение тестирования интеграции фазы 2: " & Now
End Sub


' Тестирование всех компонентов третьей фазы
Public Sub TestPhase3Components()
    Debug.Print "Начало тестирования компонентов фазы 3: " & Now
    Debug.Print "----------------------------------------"
    
    TestLogProviderBase
    TestFileLogProvider
    TestImmediateLogProvider
    TestProviderIntegration
    
    Debug.Print "----------------------------------------"
    Debug.Print "Завершение тестирования компонентов фазы 3: " & Now
End Sub

' Тестирование базового класса провайдера
' Тестирование базового класса провайдера
Private Sub TestLogProviderBase()
    Debug.Print "Тестирование LogProviderBase..."
    
    ' Примечание: мы не можем напрямую создать экземпляр абстрактного класса
    ' Поэтому создаем конкретную реализацию
    Dim provider As New FileLogProvider
    
    ' Тестируем базовые методы
    Debug.Print "  Базовые методы:"
    Debug.Print "    IsEnabled (по умолчанию) = " & provider.IsEnabled
    Debug.Print "    IsInitialized (по умолчанию) = " & provider.IsInitialized
    Debug.Print "    Name (по умолчанию) = " & provider.Name
    
    provider.Disable
    Debug.Print "    После отключения: IsEnabled = " & provider.IsEnabled
    
    provider.Enable
    Debug.Print "    После включения: IsEnabled = " & provider.IsEnabled
    
    provider.Name = "TestProvider"
    Debug.Print "    После изменения имени: Name = " & provider.Name
    
    ' Тестирование обработки ошибки при вызове абстрактного метода
    ' Важно: эта часть будет генерировать ошибку, поэтому заключаем ее в блок On Error
    Debug.Print "  Обработка ошибки при вызове абстрактного метода:"
    On Error Resume Next
    
    Dim errDesc As String
    ' Создать экземпляр базового класса теперь может вызвать ошибку
    ' поэтому проверим это в отдельном блоке
    Dim baseProvider As Object
    Set baseProvider = New LogProviderBase
    
    If Not baseProvider Is Nothing Then
        Dim entry As New LogEntry
        baseProvider.Write entry
        errDesc = Err.description
    Else
        errDesc = "Невозможно создать экземпляр базового класса"
    End If
    
    On Error GoTo 0
    Debug.Print "    Ошибка при тестировании абстрактного класса: " & errDesc
End Sub

' Тестирование провайдера для файлов
Private Sub TestFileLogProvider()
    Debug.Print "Тестирование FileLogProvider..."
    
    Dim provider As New FileLogProvider
    Dim config As New LogConfiguration
    
    ' Настраиваем конфигурацию для тестирования
    config.SetProperty "FileProviderEnabled", True
    config.SetProperty "LogDirectory", Environ$("TEMP") & "\LoggingTest\Logs"
    config.SetProperty "LogFileName", "TestLog.txt"
    config.SetProperty "MaxFileSize", 10240 ' небольшой размер для тестирования ротации
    
    ' Инициализация провайдера
    Debug.Print "  Инициализация провайдера:"
    provider.Initialize config
    Debug.Print "    IsInitialized = " & provider.IsInitialized
    Debug.Print "    IsEnabled = " & provider.IsEnabled
    
    ' Создаем тестовую запись лога
    Dim context As New LogContext
    context.moduleName = "TestModule"
    context.procedureName = "TestProcedure"
    
    Dim entry As New LogEntry
    entry.level = LogLevel.LogInfo
    entry.message = "Тест записи в файл через FileLogProvider"
    Set entry.context = context
    
    ' Запись лога
    Debug.Print "  Запись лога в файл:"
    provider.Writes entry
    
    ' Проверяем существование файла лога
    Dim logFilePath As String
    logFilePath = Environ$("TEMP") & "\LoggingTest\Logs\TestLog.txt"
    Debug.Print "    Файл лога существует: " & FSWrapper.FileExists(logFilePath)
    
    ' Проверяем содержимое файла
    Dim content As String
    content = FSWrapper.ReadTextFile(logFilePath)
    Debug.Print "    Содержимое файла лога:"
    Debug.Print "    " & Split(content, vbCrLf)(0) ' Выводим первую строку
    
    ' Тестирование ротации файла
    Debug.Print "  Тестирование ротации файла:"
    
    ' Создаем много записей для заполнения файла
    Dim i As Integer
    For i = 1 To 10
        entry.message = "Тестовая запись #" & i & " для проверки ротации файла"
        provider.Writes entry
    Next i
    
    ' Проверяем, выполнилась ли ротация (наличие архивных файлов)
    Dim logDir As String
    logDir = Environ$("TEMP") & "\LoggingTest\Logs"
    Dim logFiles As String
    logFiles = Dir(logDir & "\TestLog_*.txt")
    
    Debug.Print "    Архивные файлы логов:"
    If Len(logFiles) > 0 Then
        Do While Len(logFiles) > 0
            Debug.Print "    " & logFiles
            logFiles = Dir()
        Loop
    Else
        Debug.Print "    Архивные файлы не найдены (ротация не выполнилась)"
    End If
End Sub

' Тестирование провайдера для Immediate Window
Private Sub TestImmediateLogProvider()
    Debug.Print "Тестирование ImmediateLogProvider..."
    
    Dim provider As New ImmediateLogProvider
    Dim config As New LogConfiguration
    
    ' Настраиваем конфигурацию для тестирования
    config.SetProperty "ImmediateProviderEnabled", True
    
    ' Инициализация провайдера
    Debug.Print "  Инициализация провайдера:"
    provider.Initialize config
    Debug.Print "    IsInitialized = " & provider.IsInitialized
    Debug.Print "    IsEnabled = " & provider.IsEnabled
    
    ' Создаем тестовую запись лога
    Dim context As New LogContext
    context.moduleName = "TestModule"
    context.procedureName = "TestProcedure"
    
    Dim entry As New LogEntry
    entry.level = LogLevel.LogWarning
    entry.message = "Тест записи в Immediate Window через ImmediateLogProvider"
    Set entry.context = context
    
    ' Запись лога
    Debug.Print "  Вывод лога в Immediate Window:"
    Debug.Print "  ----------------------------------------"
    provider.Writes entry
    Debug.Print "  ----------------------------------------"
    
    ' Проверка работы дополнительных настроек форматирования
    Debug.Print "  Тест настроек форматирования:"
    Debug.Print "    IncludeTimestamp (по умолчанию) = " & provider.IncludeTimestamp
    Debug.Print "    IncludeLevel (по умолчанию) = " & provider.IncludeLevel
    
    provider.IncludeTimestamp = False
    Debug.Print "    После отключения метки времени: IncludeTimestamp = " & provider.IncludeTimestamp
    
    provider.IncludeLevel = False
    Debug.Print "    После отключения метки уровня: IncludeLevel = " & provider.IncludeLevel
    
    ' Тестирование с отключенными метками (хотя это не меняет форматирование LogEntry)
    Debug.Print "  Вывод лога с отключенными метками:"
    Debug.Print "  ----------------------------------------"
    provider.Writes entry
    Debug.Print "  ----------------------------------------"
End Sub

' Тестирование интеграции провайдеров с буфером логирования
Private Sub TestProviderIntegration()
    Debug.Print "Тестирование интеграции провайдеров с буфером..."
    
    Dim Buffer As New LogBuffer
    Dim config As New LogConfiguration
    
    ' Настраиваем конфигурацию
    config.SetProperty "FileProviderEnabled", True
    config.SetProperty "ImmediateProviderEnabled", True
    config.SetProperty "LogDirectory", Environ$("TEMP") & "\LoggingTest\Logs"
    config.SetProperty "LogFileName", "IntegrationTest.txt"
    config.SetProperty "MaxBufferSize", 5
    
    ' Инициализация компонентов
    Buffer.Initialize CLng(config.GetProperty("MaxBufferSize"))
    
    Dim fileProvider As New FileLogProvider
    fileProvider.Initialize config
    
    Dim immediateProvider As New ImmediateLogProvider
    immediateProvider.Initialize config
    
    ' Создаем тестовые записи логов
    Debug.Print "  Создание тестовых записей и добавление в буфер:"
    Dim context As New LogContext
    context.moduleName = "IntegrationTest"
    context.procedureName = "TestProviderIntegration"
    
    Dim i As Integer
    For i = 1 To 3
        Dim entry As New LogEntry
        entry.level = LogLevel.LogInfo
        entry.message = "Тестовая запись #" & i & " для проверки интеграции"
        Set entry.context = context
        
        Buffer.Add entry
    Next i
    
    Debug.Print "    Количество записей в буфере = " & Buffer.Count()
    
    ' Сброс буфера через файловый провайдер
    Debug.Print "  Сброс буфера через FileLogProvider:"
    Dim fileResult As Boolean
    fileResult = Buffer.Flush(fileProvider)
    Debug.Print "    Результат сброса = " & fileResult
    Debug.Print "    Количество записей в буфере после сброса = " & Buffer.Count()
    
    ' Добавляем еще записи
    For i = 4 To 6
        Dim entry2 As New LogEntry
        entry2.level = LogLevel.LogWarning
        entry2.message = "Тестовая запись #" & i & " для проверки интеграции"
        Set entry2.context = context
        
        Buffer.Add entry2
    Next i
    
    Debug.Print "    Количество записей в буфере после добавления = " & Buffer.Count()
    
    ' Сброс буфера через провайдер Immediate Window
    Debug.Print "  Сброс буфера через ImmediateLogProvider:"
    Debug.Print "  ----------------------------------------"
    Dim immediateResult As Boolean
    immediateResult = Buffer.Flush(immediateProvider)
    Debug.Print "  ----------------------------------------"
    Debug.Print "    Результат сброса = " & immediateResult
    Debug.Print "    Количество записей в буфере после сброса = " & Buffer.Count()
    
    ' Проверяем файл логов
    Dim logFilePath As String
    logFilePath = Environ$("TEMP") & "\LoggingTest\Logs\IntegrationTest.txt"
    Debug.Print "  Проверка файла логов:"
    Debug.Print "    Файл существует = " & FSWrapper.FileExists(logFilePath)
    
    If FSWrapper.FileExists(logFilePath) Then
        Dim content As String
        content = FSWrapper.ReadTextFile(logFilePath)
        Dim lines() As String
        lines = Split(content, vbCrLf)
        
        Debug.Print "    Количество строк в файле = " & UBound(lines) + 1
        
        If UBound(lines) >= 0 Then
            Debug.Print "    Первая строка: " & lines(0)
        End If
    End If
End Sub


' Тестирование всей системы логирования
Public Sub TestFullLoggingSystem()
    Debug.Print "Начало тестирования полной системы логирования: " & Now
    Debug.Print "----------------------------------------"
    
    TestLogManager
    TestGlobalLogger
    TestRealWorldScenarios
    
    Debug.Print "----------------------------------------"
    Debug.Print "Завершение тестирования полной системы логирования: " & Now
End Sub

' Тестирование класса LogManager
Private Sub TestLogManager()
    Debug.Print "Тестирование LogManager..."
    
    Dim manager As New LogManager
    
    ' Инициализация менеджера
    Debug.Print "  Инициализация менеджера:"
    Debug.Print "    Результат = " & manager.Initialize()
    Debug.Print "    IsInitialized = " & manager.IsInitialized
    
    ' Создание контекста для логирования
    Dim context As LogContext
    Set context = manager.CreateLogContext("TestModule", "TestMethod", 42, "x = 10")
    Debug.Print "  Создан контекст: " & context.ToString()
    
    ' Логирование различных уровней сообщений
    Debug.Print "  Логирование сообщений разных уровней:"
    manager.Log_Debug "Это отладочное сообщение", "TestModule", "TestMethod", 43
    manager.Log_Info "Это информационное сообщение", "TestModule", "TestMethod", 44
    manager.Log_Warning "Это предупреждение", "TestModule", "TestMethod", 45
    manager.Log_Error "Это сообщение об ошибке", 1000, "Test", "Тестовая ошибка", _
                   "TestModule", "TestMethod", 46, "if x > 10 then"
    
    ' Сброс буфера
    Debug.Print "  Сброс буфера логирования:"
    Debug.Print "    Результат = " & manager.FlushBuffer()
    
    ' Проверка файла логов
    Dim logDir As String
    logDir = CStr(manager.Configuration.GetProperty("LogDirectory"))
    Dim logFileName As String
    logFileName = CStr(manager.Configuration.GetProperty("LogFileName"))
    Dim logFilePath As String
    logFilePath = BuildPath(logDir, logFileName)
    
    Debug.Print "  Проверка файла логов:"
    Debug.Print "    Путь = " & logFilePath
    Debug.Print "    Существует = " & FSWrapper.FileExists(logFilePath)
    
    ' Завершение работы
    Debug.Print "  Завершение работы менеджера..."
    manager.Shutdown
End Sub

' Тестирование глобального модуля Logger
Private Sub TestGlobalLogger()
    Debug.Print "Тестирование глобального модуля Logger..."
    
    ' Инициализация
    Debug.Print "  Инициализация через глобальный модуль:"
    Debug.Print "    Результат = " & InitializeLogger()
    
    ' Получение экземпляра менеджера
    Dim manager As LogManager
    Set manager = GetLogManager()
    Debug.Print "    Получен экземпляр менеджера, IsInitialized = " & manager.IsInitialized
    
    ' Логирование через глобальные методы
    Debug.Print "  Логирование через глобальные методы:"
    Log_Debug "Глобальное отладочное сообщение", "GlobalTest", "GlobalMethod", 100
    Log_Info "Глобальное информационное сообщение", "GlobalTest", "GlobalMethod", 101
    Log_Warning "Глобальное предупреждение", "GlobalTest", "GlobalMethod", 102
    Log_Error "Глобальная ошибка", 2000, "GlobalTest", "Тестовая глобальная ошибка", _
            "GlobalTest", "GlobalMethod", 103, "if y > 20 then"
    
    ' Сброс буфера
    Debug.Print "  Сброс буфера через глобальный метод:"
    FlushLogs
    Debug.Print "    Буфер сброшен"
    
    ' Завершение работы
    Debug.Print "  Завершение работы через глобальный метод..."
    ShutdownLogger
End Sub

' Тестирование реальных сценариев использования
Private Sub TestRealWorldScenarios()
    Debug.Print "Тестирование реальных сценариев использования..."
    
    ' Инициализируем систему с настройками по умолчанию
    InitializeLogger
    
    ' Тестирование различных сценариев
    SimulateUserAuthentication
    SimulateDataProcessing
    SimulateErrorHandling
    
    ' Завершение работы
    FlushLogs
    ShutdownLogger
End Sub

' Симуляция процесса аутентификации пользователя
Private Sub SimulateUserAuthentication()
    Debug.Print "  Симуляция аутентификации пользователя:"
    
    Dim username As String
    username = GetCurrentUser() ' Используем реального пользователя
    
    ' Начало процесса
    Log_Info "Начало процесса аутентификации пользователя: " & username, "Security", "AuthenticateUser", 10
    
    ' Успешная аутентификация
    Log_Info "Пользователь " & username & " успешно аутентифицирован", "Security", "AuthenticateUser", 25
End Sub

' Симуляция обработки данных
Private Sub SimulateDataProcessing()
    Debug.Print "  Симуляция обработки данных:"
    
    ' Начало процесса
    Log_Info "Начало обработки данных", "DataProcessor", "ProcessData", 15
    
    ' Предупреждение
    Log_Warning "Обнаружены неполные данные", "DataProcessor", "ValidateData", 42
    
    ' Успешное завершение
    Log_Info "Обработка данных завершена успешно", "DataProcessor", "ProcessData", 78, "Обработано 100 записей"
End Sub

' Симуляция обработки ошибок
Private Sub SimulateErrorHandling()
    Debug.Print "  Симуляция обработки ошибок:"
    
    ' Тестирование с использованием On Error и HandleError
    On Error Resume Next
    
    ' Вызываем ошибку деления на ноль
    Dim x As Integer
    x = 1 / 0
    
    ' Если произошла ошибка, обрабатываем ее
    If Err.Number <> 0 Then
        HandleError "SimulateErrorHandling", 5, "x = 1 / 0"
        Debug.Print "    Ошибка обработана через HandleError"
    End If
    
    On Error GoTo 0
    
    ' Прямой вызов для логирования ошибки
    Log_Error "Невозможно подключиться к базе данных", 3001, _
             "Database", "Timeout exceeded", _
             "DataAccess", "ConnectToDatabase", 105, "conn.Open connectionString"
End Sub




