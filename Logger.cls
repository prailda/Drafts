' ====================================================
' Модуль определений LoggerDefinitions
' ====================================================
Option Explicit

' Категории сообщений логов
Public Enum LogCategory
    LogTrace = 1    ' Категория регистрации логов трассировки
    LogDebug = 2    ' Категория регистрации отладочных сообщений
    LogStep = 3     ' Категория регистрации этапов выполнения программы
    LogInfo = 4     ' Категория регистрации информационных сообщений
    LogWarning = 5  ' Категория регистрации предупреждающих сообщений
    LogError = 6    ' Категория регистрации сообщений об ошибках
    LogSuccess = 7  ' Категория регистрации сообщений об успехе
End Enum

' Уровни сообщений логов
Public Enum LogLevel
    Minimum = 1    ' Запись только [LogSuccess, LogError]
    Normal = 2     ' Запись только [LogInfo, LogWarning, LogError, LogSuccess]
    DebugMode = 3  ' Запись всех сообщений
End Enum

' Правила ротации файлов логов
Public Enum RotationRule
    Session = 1  ' Новый файл для каждой новой сессии
    Day = 2      ' Новый файл для каждого нового дня
    Week = 3     ' Новый файл для каждого новой недели
    Month = 4    ' Новый файл для каждого нового месяца
End Enum

' Состояния логера
Public Enum LoggerState
    LoggerNotAvailable = 1    ' Система логирования повреждена/не работает
    LoggerDisabled = 2        ' Система логирования выключена
    LoggerInitialized = 3     ' Система логирования инициализирована
    LoggerNotInitialized = 4  ' Система логирования не инициализирована
End Enum

' Константы логера
Public Const DEFAULT_BUFFER_SIZE As Integer = 500
Public Const DEFAULT_LOG_LEVEL As LogLevel = LogLevel.Normal
Public Const DEFAULT_CONFIG_FOLDER As String = "C:\Users\%USERNAME%\AppData\Local\Excellent App\Config"
Public Const DEFAULT_CONFIG_FILE As String = "logConfiguration.ini"
Public Const DEFAULT_LOG_FOLDER As String = "C:\Users\%USERNAME%\AppData\Local\Excellent App\Debug"
Public Const DEFAULT_LOG_FILE As String = "logs.txt"
Public Const DEFAULT_FILE_SIZE_ROTATION As Boolean = True
Public Const DEFAULT_FILE_SIZE_ROTATION_LIMIT_MB As Long = 50
Public Const DEFAULT_FILE_DATE_ROTATION As Boolean = False
Public Const DEFAULT_FILE_DATE_ROTATION_RULE As RotationRule = RotationRule.Week


' ====================================================
' Класс LogMessage
' ====================================================
Option Explicit

' Свойства сообщения лога
Public timestamp As Date
Public category As LogCategory
Public message As String
Public additionalInfo As String
Public callStack As String

' Преобразование сообщения в строку
Public Function ToString() As String
    Dim result As String
    
    ' Форматирование даты и времени
    result = Format(timestamp, "yyyy-mm-dd hh:nn:ss") & " "
    
    ' Добавление категории
    Select Case category
        Case LogCategory.LogTrace
            result = result & "[TRACE] "
        Case LogCategory.LogDebug
            result = result & "[DEBUG] "
        Case LogCategory.LogStep
            result = result & "[STEP]  "
        Case LogCategory.LogInfo
            result = result & "[INFO]  "
        Case LogCategory.LogWarning
            result = result & "[WARN]  "
        Case LogCategory.LogError
            result = result & "[ERROR] "
        Case LogCategory.LogSuccess
            result = result & "[OK]    "
    End Select
    
    ' Добавление сообщения
    result = result & message
    
    ' Добавление дополнительной информации при её наличии
    If Len(additionalInfo) > 0 Then
        result = result & " | " & additionalInfo
    End If
    
    ' Добавление стека вызовов при его наличии и для определенных категорий
    If Len(callStack) > 0 And (category = LogCategory.LogError Or category = LogCategory.LogTrace) Then
        result = result & vbNewLine & "CallStack: " & callStack
    End If
    
    ToString = result
End Function


' ====================================================
' Класс LogMessageBuilder
' ====================================================
Option Explicit

Private m_callStack As CallStack

Private Sub Class_Initialize()
    Set m_callStack = New CallStack
End Sub

' Создание обычного сообщения лога
Public Function BuildLogMessage(ByVal category As LogCategory, _
                               ByVal message As String, _
                               Optional ByVal additionalInfo As String = "") As LogMessage
    On Error GoTo ErrorHandler
    
    Dim logMsg As New LogMessage
    
    logMsg.timestamp = Now()
    logMsg.category = category
    logMsg.message = message
    logMsg.additionalInfo = additionalInfo
    
    ' Получение стека вызовов для категорий, требующих этой информации
    If category = LogCategory.LogTrace Then
        logMsg.callStack = m_callStack.GetCallStack()
    End If
    
    Set BuildLogMessage = logMsg
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка в BuildLogMessage: " & Err.Description
    Set BuildLogMessage = Nothing
End Function

' Создание сообщения об ошибке
Public Function BuildErrorLogMessage(ByVal errorObject As ErrObject, _
                                    Optional ByVal additionalInfo As String = "") As LogMessage
    On Error GoTo ErrorHandler
    
    Dim logMsg As New LogMessage
    
    logMsg.timestamp = Now()
    logMsg.category = LogCategory.LogError
    logMsg.message = "Ошибка #" & errorObject.Number & ": " & errorObject.Description
    logMsg.additionalInfo = additionalInfo
    logMsg.callStack = m_callStack.GetCallStack()
    
    Set BuildErrorLogMessage = logMsg
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка в BuildErrorLogMessage: " & Err.Description
    Set BuildErrorLogMessage = Nothing
End Function


' ====================================================
' Класс LogBuffer
' ====================================================
Option Explicit

Private m_buffer As Dictionary
Private m_max_size As Long

Private Sub Class_Initialize()
    Set m_buffer = New Dictionary
    m_max_size = DEFAULT_BUFFER_SIZE
End Sub

' Установка максимального размера буфера
Public Sub SetMaxSize(ByVal size As Long)
    If size > 0 Then
        m_max_size = size
    End If
End Sub

' Добавление сообщения в буфер
Public Function Add(ByVal logMessage As LogMessage) As Boolean
    On Error GoTo ErrorHandler
    
    ' Генерация уникального ключа для сообщения
    Dim key As String
    key = Format(Now(), "yyyymmddhhnnss") & "_" & Format(Timer * 1000, "000000")
    
    ' Добавление сообщения в буфер
    m_buffer.Add key, logMessage
    
    Add = True
    Exit Function
    
ErrorHandler:
    Add = False
End Function

' Получение всех сообщений
Public Function GetAll() As Collection
    On Error GoTo ErrorHandler
    
    Dim result As New Collection
    Dim key As Variant
    
    For Each key In m_buffer.Keys()
        result.Add m_buffer.Item(key)
    Next key
    
    Set GetAll = result
    Exit Function
    
ErrorHandler:
    Set GetAll = New Collection
End Function

' Очистка буфера
Public Function Clear() As Boolean
    On Error GoTo ErrorHandler
    
    Set m_buffer = New Dictionary
    
    Clear = True
    Exit Function
    
ErrorHandler:
    Clear = False
End Function

' Получение количества сообщений в буфере
Public Function Count() As Long
    Count = m_buffer.Count
End Function

' Проверка заполненности буфера
Public Function IsBufferFull() As Boolean
    IsBufferFull = (m_buffer.Count >= m_max_size)
End Function


' ====================================================
' Класс ImmediateLoggerProvider
' ====================================================
Option Explicit

Implements ILoggerProvider

Private m_is_enabled As Boolean

Private Sub Class_Initialize()
    m_is_enabled = True
End Sub

' Установка состояния провайдера
Public Sub SetEnabled(ByVal state As Boolean)
    m_is_enabled = state
End Sub

' Инициализация провайдера
Private Function ILoggerProvider_Initialize() As Boolean
    ILoggerProvider_Initialize = True
End Function

' Вывод сообщения
Private Function ILoggerProvider_Log(ByVal logMessage As LogMessage) As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_is_enabled Then
        ILoggerProvider_Log = False
        Exit Function
    End If
    
    Debug.Print logMessage.ToString()
    
    ILoggerProvider_Log = True
    Exit Function
    
ErrorHandler:
    ILoggerProvider_Log = False
End Function

' Сброс буфера
Private Function ILoggerProvider_Flush() As Boolean
    ILoggerProvider_Flush = True ' Для Immediate Provider ничего делать не нужно
End Function


' ====================================================
' Класс FileLoggerProvider
' ====================================================
Option Explicit

Implements ILoggerProvider

Private m_is_enabled As Boolean
Private m_file_path As String
Private m_size_rotation_enabled As Boolean
Private m_size_rotation_limit As Long
Private m_date_rotation_enabled As Boolean
Private m_date_rotation_rule As RotationRule
Private m_fso As FileSystemObject
Private m_last_rotation_check As Date

Private Sub Class_Initialize()
    m_is_enabled = False
    m_file_path = ""
    m_size_rotation_enabled = DEFAULT_FILE_SIZE_ROTATION
    m_size_rotation_limit = DEFAULT_FILE_SIZE_ROTATION_LIMIT_MB
    m_date_rotation_enabled = DEFAULT_FILE_DATE_ROTATION
    m_date_rotation_rule = DEFAULT_FILE_DATE_ROTATION_RULE
    m_last_rotation_check = Now()
    
    Set m_fso = New FileSystemObject
End Sub

' Установка параметров провайдера
Public Sub Configure(ByVal filePath As String, _
                    ByVal sizeRotationEnabled As Boolean, _
                    ByVal sizeRotationLimit As Long, _
                    ByVal dateRotationEnabled As Boolean, _
                    ByVal dateRotationRule As RotationRule)
                    
    m_file_path = filePath
    m_size_rotation_enabled = sizeRotationEnabled
    m_size_rotation_limit = sizeRotationLimit
    m_date_rotation_enabled = dateRotationEnabled
    m_date_rotation_rule = dateRotationRule
End Sub

' Установка состояния провайдера
Public Sub SetEnabled(ByVal state As Boolean)
    m_is_enabled = state
End Sub

' Инициализация провайдера
Private Function ILoggerProvider_Initialize() As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка наличия пути к файлу
    If m_file_path = "" Then
        ILoggerProvider_Initialize = False
        Exit Function
    End If
    
    ' Проверка существования директории
    Dim folderPath As String
    folderPath = Left(m_file_path, InStrRev(m_file_path, "\") - 1)
    
    If Not m_fso.FolderExists(folderPath) Then
        m_fso.CreateFolder folderPath
    End If
    
    ' Если файл не существует, создаем его с заголовком
    If Not m_fso.FileExists(m_file_path) Then
        Dim ts As TextStream
        Set ts = m_fso.CreateTextFile(m_file_path, True)
        ts.WriteLine "=== Log File Created: " & Now() & " ==="
        ts.WriteLine String(50, "=")
        ts.Close
    End If
    
    ILoggerProvider_Initialize = True
    Exit Function
    
ErrorHandler:
    ILoggerProvider_Initialize = False
End Function

' Вывод сообщения
Private Function ILoggerProvider_Log(ByVal logMessage As LogMessage) As Boolean
    On Error GoTo ErrorHandler
    
    If Not m_is_enabled Then
        ILoggerProvider_Log = False
        Exit Function
    End If
    
    ' Проверка необходимости ротации
    CheckRotation
    
    ' Запись в файл
    Dim ts As TextStream
    Set ts = m_fso.OpenTextFile(m_file_path, ForAppending, True)
    ts.WriteLine logMessage.ToString()
    ts.Close
    
    ILoggerProvider_Log = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка в FileLoggerProvider.Log: " & Err.Description
    ILoggerProvider_Log = False
End Function

' Сброс буфера
Private Function ILoggerProvider_Flush() As Boolean
    ILoggerProvider_Flush = m_fso.FileExists(m_file_path)
End Function

' Проверка ротации файла
Private Sub CheckRotation()
    On Error Resume Next
    
    ' Проверка ротации по размеру
    If m_size_rotation_enabled Then
        Dim fileSize As Long
        If m_fso.FileExists(m_file_path) Then
            fileSize = m_fso.GetFile(m_file_path).Size / (1024 * 1024) ' В мегабайтах
            
            If fileSize >= m_size_rotation_limit Then
                RotateFile "size"
            End If
        End If
    End If
    
    ' Проверка ротации по дате (только раз в час для оптимизации)
    If m_date_rotation_enabled And DateDiff("h", m_last_rotation_check, Now()) >= 1 Then
        m_last_rotation_check = Now()
        
        Dim needRotation As Boolean
        needRotation = False
        
        Select Case m_date_rotation_rule
            Case RotationRule.Session
                ' Для сессии ротация уже выполнена при инициализации
                
            Case RotationRule.Day
                ' Новый файл каждый день
                Static lastDay As Integer
                If lastDay = 0 Then
                    lastDay = Day(Now())
                ElseIf lastDay <> Day(Now()) Then
                    needRotation = True
                    lastDay = Day(Now())
                End If
                
            Case RotationRule.Week
                ' Новый файл каждую неделю (понедельник)
                Static lastWeek As Integer
                Dim currentWeek As Integer
                currentWeek = DatePart("ww", Now(), vbMonday)
                
                If lastWeek = 0 Then
                    lastWeek = currentWeek
                ElseIf lastWeek <> currentWeek Then
                    needRotation = True
                    lastWeek = currentWeek
                End If
                
            Case RotationRule.Month
                ' Новый файл каждый месяц
                Static lastMonth As Integer
                If lastMonth = 0 Then
                    lastMonth = Month(Now())
                ElseIf lastMonth <> Month(Now()) Then
                    needRotation = True
                    lastMonth = Month(Now())
                End If
        End Select
        
        If needRotation Then
            RotateFile "date"
        End If
    End If
End Sub

' Ротация файла
Private Sub RotateFile(ByVal reason As String)
    On Error Resume Next
    
    If Not m_fso.FileExists(m_file_path) Then Exit Sub
    
    ' Формирование имени архивного файла
    Dim basePath As String
    Dim extension As String
    Dim newFileName As String
    
    basePath = Left(m_file_path, InStrRev(m_file_path, ".") - 1)
    extension = Mid(m_file_path, InStrRev(m_file_path, "."))
    newFileName = basePath & "_" & Format(Now(), "yyyymmdd_hhnnss") & "_" & reason & extension
    
    ' Перемещение текущего файла в архив
    m_fso.MoveFile m_file_path, newFileName
    
    ' Создание нового файла с заголовком
    Dim ts As TextStream
    Set ts = m_fso.CreateTextFile(m_file_path, True)
    ts.WriteLine "=== Log File Created: " & Now() & " (Rotation reason: " & reason & ") ==="
    ts.WriteLine String(50, "=")
    ts.Close
End Sub


' ====================================================
' Класс LoggerConfiguration
' ====================================================
Option Explicit

Private m_config As Dictionary
Private m_config_file_path As String
Private m_fso As FileSystemObject

Private Sub Class_Initialize()
    Set m_config = New Dictionary
    Set m_fso = New FileSystemObject
    
    ' Получение пути к файлу конфигурации
    Dim configPath As String
    configPath = Replace(DEFAULT_CONFIG_FOLDER, "%USERNAME%", Environ$("USERNAME"))
    
    ' Проверка существования директории
    If Not m_fso.FolderExists(configPath) Then
        On Error Resume Next
        m_fso.CreateFolder configPath
        On Error GoTo 0
    End If
    
    m_config_file_path = configPath & "\" & DEFAULT_CONFIG_FILE
End Sub

' Получение конфигурации
Public Function GetConfiguration() As Dictionary
    Set GetConfiguration = m_config
End Function

' Установка конфигурации
Public Function SetConfiguration(ByVal config As Dictionary) As Boolean
    On Error GoTo ErrorHandler
    
    Set m_config = config
    SetConfiguration = True
    Exit Function
    
ErrorHandler:
    SetConfiguration = False
End Function

' Получение конфигурации по умолчанию
Public Function GetDefaultConfiguration() As Dictionary
    On Error GoTo ErrorHandler
    
    Dim config As New Dictionary
    
    ' Основные параметры
    config.Add "IsLoggingEnable", True
    config.Add "LoggerConfigurationFilePath", Replace(DEFAULT_CONFIG_FOLDER, "%USERNAME%", Environ$("USERNAME"))
    config.Add "LoggerConfigurationFile", DEFAULT_CONFIG_FILE
    config.Add "LoggingMinimumLevel", DEFAULT_LOG_LEVEL
    config.Add "IsImmediateLoggingEnable", True
    config.Add "IsTextFileLoggingEnable", False
    
    ' Параметры текстового логера
    config.Add "LoggerTextFilePath", Replace(DEFAULT_LOG_FOLDER, "%USERNAME%", Environ$("USERNAME"))
    config.Add "LoggerTextFile", DEFAULT_LOG_FILE
    config.Add "IsLoggerTextFileSizeRotationEnable", DEFAULT_FILE_SIZE_ROTATION
    config.Add "LoggerTextFileSizeRotationRule", DEFAULT_FILE_SIZE_ROTATION_LIMIT_MB
    config.Add "IsLoggerTextFileDateRotationEnable", DEFAULT_FILE_DATE_ROTATION
    config.Add "LoggerTextFileDateRotationRule", DEFAULT_FILE_DATE_ROTATION_RULE
    
    ' Параметры буфера
    config.Add "LoggerBufferMaximumSize", DEFAULT_BUFFER_SIZE
    
    Set GetDefaultConfiguration = config
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка в GetDefaultConfiguration: " & Err.Description
    Set GetDefaultConfiguration = New Dictionary
End Function

' Сохранение конфигурации
Public Function SaveConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка наличия конфигурации
    If m_config Is Nothing Or m_config.Count = 0 Then
        SaveConfiguration = False
        Exit Function
    End If
    
    ' Запись файла конфигурации
    Dim ts As TextStream
    Set ts = m_fso.CreateTextFile(m_config_file_path, True)
    
    Dim key As Variant
    For Each key In m_config.Keys()
        Dim value As Variant
        value = m_config.Item(key)
        
        ' Преобразование значения в строку
        Dim valueStr As String
        If IsObject(value) Then
            valueStr = "[Object]"
        ElseIf VarType(value) = vbBoolean Then
            valueStr = IIf(value, "True", "False")
        Else
            valueStr = CStr(value)
        End If
        
        ts.WriteLine key & "=" & valueStr
    Next key
    
    ts.Close
    
    SaveConfiguration = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка в SaveConfiguration: " & Err.Description
    SaveConfiguration = False
End Function

' Загрузка конфигурации
Public Function LoadConfiguration() As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка существования файла конфигурации
    If Not m_fso.FileExists(m_config_file_path) Then
        ' Если файл не существует, создаем конфигурацию по умолчанию
        Set m_config = GetDefaultConfiguration()
        SaveConfiguration
        LoadConfiguration = True
        Exit Function
    End If
    
    ' Чтение файла конфигурации
    Dim ts As TextStream
    Set ts = m_fso.OpenTextFile(m_config_file_path, ForReading)
    
    Dim config As New Dictionary
    
    ' Чтение строк файла
    Do Until ts.AtEndOfStream
        Dim line As String
        line = ts.ReadLine
        
        ' Пропуск пустых строк и комментариев
        If Len(line) > 0 And Left(line, 1) <> "'" And Left(line, 1) <> "#" Then
            ' Разделение на ключ и значение
            Dim parts As Variant
            parts = Split(line, "=", 2)
            
            If UBound(parts) = 1 Then
                Dim key As String
                Dim value As String
                
                key = Trim(parts(0))
                value = Trim(parts(1))
                
                ' Преобразование значения в соответствующий тип
                Dim typedValue As Variant
                
                If value = "True" Or value = "False" Then
                    typedValue = (value = "True")
                ElseIf IsNumeric(value) Then
                    If InStr(value, ".") > 0 Then
                        typedValue = CDbl(value)
                    Else
                        typedValue = CLng(value)
                    End If
                Else
                    typedValue = value
                End If
                
                ' Добавление в словарь
                config.Add key, typedValue
            End If
        End If
    Loop
    
    ts.Close
    
    ' Проверка валидности конфигурации
    If config.Count = 0 Then
        ' Если конфигурация пуста, создаем конфигурацию по умолчанию
        Set m_config = GetDefaultConfiguration()
        SaveConfiguration
    Else
        Set m_config = config
    End If
    
    LoadConfiguration = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Ошибка в LoadConfiguration: " & Err.Description
    
    ' В случае ошибки создаем конфигурацию по умолчанию
    Set m_config = GetDefaultConfiguration()
    SaveConfiguration
    
    LoadConfiguration = False
End Function

' Проверка валидности конфигурации
Public Function IsConfigValid() As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка наличия конфигурации
    If m_config Is Nothing Or m_config.Count = 0 Then
        IsConfigValid = False
        Exit Function
    End If
    
    ' Проверка наличия всех необходимых параметров
    Dim requiredKeys As Variant
    requiredKeys = Array("IsLoggingEnable", "LoggingMinimumLevel", "IsImmediateLoggingEnable", _
                          "IsTextFileLoggingEnable", "LoggerBufferMaximumSize")
    
    Dim key As Variant
    For Each key In requiredKeys
        If Not m_config.Exists(key) Then
            IsConfigValid = False
            Exit Function
        End If
    Next key
    
    ' Проверка параметров текстового логера, если он включен
    If m_config.Item("IsTextFileLoggingEnable") Then
        Dim textLoggerKeys As Variant
        textLoggerKeys = Array("LoggerTextFilePath", "LoggerTextFile", _
                               "IsLoggerTextFileSizeRotationEnable", "LoggerTextFileSizeRotationRule", _
                               "IsLoggerTextFileDateRotationEnable", "LoggerTextFileDateRotationRule")
        
        For Each key In textLoggerKeys
            If Not m_config.Exists(key) Then
                IsConfigValid = False
                Exit Function
            End If
        Next key
    End If
    
    IsConfigValid = True
    Exit Function
    
ErrorHandler:
    IsConfigValid = False
End Function


' ====================================================
' Класс Logger
' ====================================================
Option Explicit

Private m_logger_state As LoggerState
Private m_configuration As LoggerConfiguration
Private m_buffer As LogBuffer
Private m_providers As Collection
Private m_min_log_level As LogLevel
Private m_message_builder As LogMessageBuilder

' Инициализация компонентов логера
Public Function Initialize() As Boolean
    On Error GoTo ErrorHandler
    
    ' Установка начального состояния
    m_logger_state = LoggerState.LoggerNotInitialized
    
    ' Создание объектов
    Set m_configuration = New LoggerConfiguration
    Set m_buffer = New LogBuffer
    Set m_providers = New Collection
    Set m_message_builder = New LogMessageBuilder
    
    ' Загрузка конфигурации
    If Not m_configuration.LoadConfiguration() Then
        m_logger_state = LoggerState.LoggerNotAvailable
        Initialize = False
        Exit Function
    End If
    
    ' Получение конфигурации
    Dim config As Dictionary
    Set config = m_configuration.GetConfiguration()
    
    ' Проверка активации логирования
    If Not config.Item("IsLoggingEnable") Then
        m_logger_state = LoggerState.LoggerDisabled
        Initialize = True
        Exit Function
    End If
    
    ' Установка минимального уровня логирования
    m_min_log_level = config.Item("LoggingMinimumLevel")
    
    ' Установка размера буфера
    m_buffer.SetMaxSize config.Item("LoggerBufferMaximumSize")
    
    ' Инициализация провайдеров
    If config.Item("IsImmediateLoggingEnable") Then
        Dim immediateProvider As New ImmediateLoggerProvider
        immediateProvider.SetEnabled True
        If immediateProvider.Initialize() Then
            m_providers.Add immediateProvider
        End If
    End If
    
    If config.Item("IsTextFileLoggingEnable") Then
        Dim fileProvider As New FileLoggerProvider
        fileProvider.SetEnabled True
        
        ' Формирование пути к файлу логов
        Dim logFilePath As String
        logFilePath = config.Item("LoggerTextFilePath") & "\" & config.Item("LoggerTextFile")
        
        ' Настройка параметров ротации
        fileProvider.Configure logFilePath, _
                              config.Item("IsLoggerTextFileSizeRotationEnable"), _
                              config.Item("LoggerTextFileSizeRotationRule"), _
                              config.Item("IsLoggerTextFileDateRotationEnable"), _
                              config.Item("LoggerTextFileDateRotationRule")
        
        If fileProvider.Initialize() Then
            m_providers.Add fileProvider
        End If
    End If
    
    ' Проверка наличия хотя бы одного провайдера
    If m_providers.Count = 0 Then
        m_logger_state = LoggerState.LoggerNotAvailable
        Initialize = False
        Exit Function
    End If
    
    ' Установка состояния инициализации
    m_logger_state = LoggerState.LoggerInitialized
    
    Initialize = True
    Exit Function
    
ErrorHandler:
    m_logger_state = LoggerState.LoggerNotAvailable
    Initialize = False
End Function

' Получение текущего состояния логера
Public Function GetState() As LoggerState
    GetState = m_logger_state
End Function

' Проверка соответствия категории минимальному уровню логирования
Private Function IsCategoryAllowed(ByVal category As LogCategory) As Boolean
    Select Case m_min_log_level
        Case LogLevel.Minimum
            ' Только ошибки и успех
            IsCategoryAllowed = (category = LogCategory.LogError Or category = LogCategory.LogSuccess)
            
        Case LogLevel.Normal
            ' Информация, предупреждения, ошибки и успех
            IsCategoryAllowed = (category = LogCategory.LogInfo Or _
                                category = LogCategory.LogWarning Or _
                                category = LogCategory.LogError Or _
                                category = LogCategory.LogSuccess)
            
        Case LogLevel.DebugMode
            ' Все категории
            IsCategoryAllowed = True
            
        Case Else
            IsCategoryAllowed = False
    End Select
End Function

' Логирование трассировки
Public Function LogTrace(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка состояния логера
    If m_logger_state <> LoggerState.LoggerInitialized Then
        LogTrace = False
        Exit Function
    End If
    
    ' Проверка соответствия уровню логирования
    If Not IsCategoryAllowed(LogCategory.LogTrace) Then
        LogTrace = True
        Exit Function
    End If
    
    ' Создание сообщения
    Dim logMsg As LogMessage
    Set logMsg = m_message_builder.BuildLogMessage(LogCategory.LogTrace, message, additionalInfo)
    
    ' Добавление сообщения в буфер
    m_buffer.Add logMsg
    
    ' Проверка заполненности буфера
    If m_buffer.IsBufferFull() Then
        Flush
    End If
    
    LogTrace = True
    Exit Function
    
ErrorHandler:
    LogTrace = False
End Function

' Логирование отладочной информации
Public Function LogDebug(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка состояния логера
    If m_logger_state <> LoggerState.LoggerInitialized Then
        LogDebug = False
        Exit Function
    End If
    
    ' Проверка соответствия уровню логирования
    If Not IsCategoryAllowed(LogCategory.LogDebug) Then
        LogDebug = True
        Exit Function
    End If
    
    ' Создание сообщения
    Dim logMsg As LogMessage
    Set logMsg = m_message_builder.BuildLogMessage(LogCategory.LogDebug, message, additionalInfo)
    
    ' Добавление сообщения в буфер
    m_buffer.Add logMsg
    
    ' Проверка заполненности буфера
    If m_buffer.IsBufferFull() Then
        Flush
    End If
    
    LogDebug = True
    Exit Function
    
ErrorHandler:
    LogDebug = False
End Function

' Логирование шага выполнения
Public Function LogStep(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка состояния логера
    If m_logger_state <> LoggerState.LoggerInitialized Then
        LogStep = False
        Exit Function
    End If
    
    ' Проверка соответствия уровню логирования
    If Not IsCategoryAllowed(LogCategory.LogStep) Then
        LogStep = True
        Exit Function
    End If
    
    ' Создание сообщения
    Dim logMsg As LogMessage
    Set logMsg = m_message_builder.BuildLogMessage(LogCategory.LogStep, message, additionalInfo)
    
    ' Добавление сообщения в буфер
    m_buffer.Add logMsg
    
    ' Проверка заполненности буфера
    If m_buffer.IsBufferFull() Then
        Flush
    End If
    
    LogStep = True
    Exit Function
    
ErrorHandler:
    LogStep = False
End Function

' Логирование информационного сообщения
Public Function LogInfo(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка состояния логера
    If m_logger_state <> LoggerState.LoggerInitialized Then
        LogInfo = False
        Exit Function
    End If
    
    ' Проверка соответствия уровню логирования
    If Not IsCategoryAllowed(LogCategory.LogInfo) Then
        LogInfo = True
        Exit Function
    End If
    
    ' Создание сообщения
    Dim logMsg As LogMessage
    Set logMsg = m_message_builder.BuildLogMessage(LogCategory.LogInfo, message, additionalInfo)
    
    ' Добавление сообщения в буфер
    m_buffer.Add logMsg
    
    ' Проверка заполненности буфера
    If m_buffer.IsBufferFull() Then
        Flush
    End If
    
    LogInfo = True
    Exit Function
    
ErrorHandler:
    LogInfo = False
End Function

' Логирование предупреждения
Public Function LogWarning(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка состояния логера
    If m_logger_state <> LoggerState.LoggerInitialized Then
        LogWarning = False
        Exit Function
    End If
    
    ' Проверка соответствия уровню логирования
    If Not IsCategoryAllowed(LogCategory.LogWarning) Then
        LogWarning = True
        Exit Function
    End If
    
    ' Создание сообщения
    Dim logMsg As LogMessage
    Set logMsg = m_message_builder.BuildLogMessage(LogCategory.LogWarning, message, additionalInfo)
    
    ' Добавление сообщения в буфер
    m_buffer.Add logMsg
    
    ' Проверка заполненности буфера
    If m_buffer.IsBufferFull() Then
        Flush
    End If
    
    LogWarning = True
    Exit Function
    
ErrorHandler:
    LogWarning = False
End Function

' Логирование ошибки
Public Function LogError(ByVal errorObject As ErrObject, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка состояния логера
    If m_logger_state <> LoggerState.LoggerInitialized Then
        LogError = False
        Exit Function
    End If
    
    ' Проверка соответствия уровню логирования
    If Not IsCategoryAllowed(LogCategory.LogError) Then
        LogError = True
        Exit Function
    End If
    
    ' Создание сообщения
    Dim logMsg As LogMessage
    Set logMsg = m_message_builder.BuildErrorLogMessage(errorObject, additionalInfo)
    
    ' Добавление сообщения в буфер
    m_buffer.Add logMsg
    
    ' Принудительный сброс буфера при ошибке
    Flush
    
    LogError = True
    Exit Function
    
ErrorHandler:
    LogError = False
End Function

' Логирование успешного выполнения
Public Function LogSuccess(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка состояния логера
    If m_logger_state <> LoggerState.LoggerInitialized Then
        LogSuccess = False
        Exit Function
    End If
    
    ' Проверка соответствия уровню логирования
    If Not IsCategoryAllowed(LogCategory.LogSuccess) Then
        LogSuccess = True
        Exit Function
    End If
    
    ' Создание сообщения
    Dim logMsg As LogMessage
    Set logMsg = m_message_builder.BuildLogMessage(LogCategory.LogSuccess, message, additionalInfo)
    
    ' Добавление сообщения в буфер
    m_buffer.Add logMsg
    
    ' Проверка заполненности буфера
    If m_buffer.IsBufferFull() Then
        Flush
    End If
    
    LogSuccess = True
    Exit Function
    
ErrorHandler:
    LogSuccess = False
End Function

' Принудительный сброс буфера
Public Function ForceFlush() As Boolean
    ForceFlush = Flush()
End Function

' Сброс буфера
Private Function Flush() As Boolean
    On Error GoTo ErrorHandler
    
    ' Проверка состояния логера
    If m_logger_state <> LoggerState.LoggerInitialized Then
        Flush = False
        Exit Function
    End If
    
    ' Получение всех сообщений из буфера
    Dim messages As Collection
    Set messages = m_buffer.GetAll()
    
    ' Выход, если буфер пуст
    If messages.Count = 0 Then
        Flush = True
        Exit Function
    End If
    
    ' Отправка сообщений во все провайдеры
    Dim provider As Variant
    Dim logMsg As Variant
    
    For Each provider In m_providers
        For Each logMsg In messages
            provider.Log logMsg
        Next logMsg
        
        provider.Flush
    Next provider
    
    ' Очистка буфера
    m_buffer.Clear
    
    Flush = True
    Exit Function
    
ErrorHandler:
    Flush = False
End Function


' ====================================================
' Класс LoggerFactory
' ====================================================
Option Explicit

Private m_instance As Logger

' Получение экземпляра логера
Public Function GetLogger() As Logger
    On Error GoTo ErrorHandler
    
    If m_instance Is Nothing Then
        Set m_instance = New Logger
        m_instance.Initialize
    End If
    
    Set GetLogger = m_instance
    Exit Function
    
ErrorHandler:
    Set GetLogger = Nothing
End Function


' ====================================================
' Модуль Log (Фасад для системы логирования)
' ====================================================
Option Explicit

' Логирование трассировки
Public Function LogTrace(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim logger As Logger
    Set logger = LoggerFactory.GetLogger()
    
    If logger Is Nothing Then
        LogTrace = False
        Exit Function
    End If
    
    LogTrace = logger.LogTrace(message, additionalInfo)
    Exit Function
    
ErrorHandler:
    LogTrace = False
End Function

' Логирование отладочной информации
Public Function LogDebug(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim logger As Logger
    Set logger = LoggerFactory.GetLogger()
    
    If logger Is Nothing Then
        LogDebug = False
        Exit Function
    End If
    
    LogDebug = logger.LogDebug(message, additionalInfo)
    Exit Function
    
ErrorHandler:
    LogDebug = False
End Function

' Логирование шага выполнения
Public Function LogStep(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim logger As Logger
    Set logger = LoggerFactory.GetLogger()
    
    If logger Is Nothing Then
        LogStep = False
        Exit Function
    End If
    
    LogStep = logger.LogStep(message, additionalInfo)
    Exit Function
    
ErrorHandler:
    LogStep = False
End Function

' Логирование информационного сообщения
Public Function LogInfo(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim logger As Logger
    Set logger = LoggerFactory.GetLogger()
    
    If logger Is Nothing Then
        LogInfo = False
        Exit Function
    End If
    
    LogInfo = logger.LogInfo(message, additionalInfo)
    Exit Function
    
ErrorHandler:
    LogInfo = False
End Function

' Логирование предупреждения
Public Function LogWarning(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim logger As Logger
    Set logger = LoggerFactory.GetLogger()
    
    If logger Is Nothing Then
        LogWarning = False
        Exit Function
    End If
    
    LogWarning = logger.LogWarning(message, additionalInfo)
    Exit Function
    
ErrorHandler:
    LogWarning = False
End Function

' Логирование ошибки
Public Function LogError(ByVal errorObject As ErrObject, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim logger As Logger
    Set logger = LoggerFactory.GetLogger()
    
    If logger Is Nothing Then
        LogError = False
        Exit Function
    End If
    
    LogError = logger.LogError(errorObject, additionalInfo)
    Exit Function
    
ErrorHandler:
    LogError = False
End Function

' Логирование успешного выполнения
Public Function LogSuccess(ByVal message As String, Optional ByVal additionalInfo As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    Dim logger As Logger
    Set logger = LoggerFactory.GetLogger()
    
    If logger Is Nothing Then
        LogSuccess = False
        Exit Function
    End If
    
    LogSuccess = logger.LogSuccess(message, additionalInfo)
    Exit Function
    
ErrorHandler:
    LogSuccess = False
End Function

' Принудительный сброс буфера
Public Function ForceFlush() As Boolean
    On Error GoTo ErrorHandler
    
    Dim logger As Logger
    Set logger = LoggerFactory.GetLogger()
    
    If logger Is Nothing Then
        ForceFlush = False
        Exit Function
    End If
    
    ForceFlush = logger.ForceFlush()
    Exit Function
    
ErrorHandler:
    ForceFlush = False
End Function


' ====================================================
' Модуль DebugSystem
' ====================================================
Option Explicit

' Инициализация системы отладки
Public Function InitializeDebugSystem() As Boolean
    On Error GoTo ErrorHandler
    
    ' Получение экземпляра логера через фабрику
    Dim logger As Logger
    Set logger = LoggerFactory.GetLogger()
    
    ' Проверка состояния логера
    If logger.GetState() <> LoggerState.LoggerInitialized Then
        ' Логирование ошибки инициализации через аварийный логер
        EmergencyLog "Ошибка инициализации системы логирования"
        InitializeDebugSystem = False
        Exit Function
    End If
    
    ' Инициализация успешна
    InitializeDebugSystem = True
    Exit Function
    
ErrorHandler:
    EmergencyLog "Ошибка при инициализации системы отладки: " & Err.Description
    InitializeDebugSystem = False
End Function

' Аварийное логирование
Private Sub EmergencyLog(ByVal message As String)
    On Error Resume Next
    Debug.Print "[EMERGENCY] " & Now() & " - " & message
End Sub


' ====================================================
' Пример использования системы логирования в коде приложения
' ====================================================

' В модуле ThisWorkbook
Private Sub Workbook_Open()
    On Error Resume Next
    
    ' Инициализация системы отладки при запуске Excel
    DebugSystem.InitializeDebugSystem
    
    If Err.Number <> 0 Then
        MsgBox "Ошибка при инициализации системы отладки: " & Err.Description, vbExclamation
    Else
        ' Логирование успешного запуска приложения
        Log.LogInfo "Приложение успешно запущено", "Excel версия: " & Application.Version
    End If
End Sub

' В пользовательском коде
Public Sub ProcessData()
    On Error GoTo ErrorHandler
    
    ' Логирование начала выполнения
    Log.LogTrace "Вход в ProcessData", "Начало обработки данных"
    
    ' Логирование этапов выполнения
    Log.LogStep "Этап 1: Получение данных из источника"
    ' Код этапа 1...
    
    ' Отладочная информация
    Log.LogDebug "Количество записей: 100", "Время выполнения: 1.2 сек"
    
    Log.LogStep "Этап 2: Фильтрация данных"
    ' Код этапа 2...
    
    ' Предупреждение
    Log.LogWarning "Найдены пустые значения", "Количество: 5"
    
    Log.LogStep "Этап 3: Сохранение результатов"
    ' Код этапа 3...
    
    ' Информационное сообщение
    Log.LogInfo "Обработка завершена", "Время выполнения: 3.5 сек"
    
    ' Сообщение об успехе
    Log.LogSuccess "Данные успешно обработаны", "Обработано записей: 95"
    
    ' Логирование завершения выполнения
    Log.LogTrace "Выход из ProcessData", "Обработка завершена"
    Exit Sub
    
ErrorHandler:
    ' Логирование ошибки
    Log.LogError Err, "Ошибка при обработке данных"
    
    ' Показ сообщения пользователю
    MsgBox "Произошла ошибка при обработке данных: " & Err.Description, vbExclamation
End Sub