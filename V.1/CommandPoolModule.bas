Attribute VB_Name = "CommandPoolModule"
'' Модуль: CommandPoolModule.bas
'' Описание: Общий модуль для реализации паттерна Singleton для пулов команд
'Option Explicit
'
'' Экземпляры пулов команд
'Private m_LogCommandPool As CommandPool
'Private m_DataCommandPool As CommandPool
'Private m_UICommandPool As CommandPool
'
'' Экземпляр инвокера команд
'Private m_CommandInvoker As CommandInvoker
'
'' Получение экземпляра пула команд для логирования
'Public Function GetLogCommandPool() As CommandPool
'    If m_LogCommandPool Is Nothing Then
'        Set m_LogCommandPool = New CommandPool
'        m_LogCommandPool.Initialize "LogCommands", 10, 30 ' Макс. 10 команд, 30 минут жизни
'    End If
'    Set GetLogCommandPool = m_LogCommandPool
'End Function
'
'' Получение экземпляра пула команд для работы с данными
'Public Function GetDataCommandPool() As CommandPool
'    If m_DataCommandPool Is Nothing Then
'        Set m_DataCommandPool = New CommandPool
'        m_DataCommandPool.Initialize "DataCommands", 15, 30 ' Макс. 15 команд, 30 минут жизни
'    End If
'    Set GetDataCommandPool = m_DataCommandPool
'End Function
'
'' Получение экземпляра пула команд для интерфейса
'Public Function GetUICommandPool() As CommandPool
'    If m_UICommandPool Is Nothing Then
'        Set m_UICommandPool = New CommandPool
'        m_UICommandPool.Initialize "UICommands", 8, 30 ' Макс. 8 команд, 30 минут жизни
'    End If
'    Set GetUICommandPool = m_UICommandPool
'End Function
'
'' Получение экземпляра инвокера команд
'Public Function GetCommandInvoker() As CommandInvoker
'    If m_CommandInvoker Is Nothing Then
'        Set m_CommandInvoker = New CommandInvoker
'    End If
'    Set GetCommandInvoker = m_CommandInvoker
'End Function
'
'' Получение статистики всех пулов команд
'Public Function GetPoolsStatistics() As String
'    Dim stats As String
'    stats = "======= Command Pool Statistics =======" & vbCrLf & vbCrLf
'
'    If Not m_LogCommandPool Is Nothing Then
'        stats = stats & m_LogCommandPool.GetStatistics() & vbCrLf & vbCrLf
'    End If
'
'    If Not m_DataCommandPool Is Nothing Then
'        stats = stats & m_DataCommandPool.GetStatistics() & vbCrLf & vbCrLf
'    End If
'
'    If Not m_UICommandPool Is Nothing Then
'        stats = stats & m_UICommandPool.GetStatistics() & vbCrLf & vbCrLf
'    End If
'
'    stats = stats & "======================================"
'
'    GetPoolsStatistics = stats
'End Function
'' Добавьте этот метод в модуль CommandPoolModule.bas
'Public Function CreateRequestInputCommand(ByVal prompt As String, _
'                                         Optional ByVal title As String = "Ввод данных", _
'                                         Optional ByVal defaultValue As String = "") As RequestInputCommand
'    Dim cmd As New RequestInputCommand
'    cmd.Initialize prompt, title, defaultValue
'    Set CreateRequestInputCommand = cmd
'End Function
'
'Public Sub ReleaseAllPools()
'    On Error Resume Next ' Добавляем обработку ошибок
'
'    If Not m_LogCommandPool Is Nothing Then
'        ' Проверяем, что объект инициализирован корректно
'        If m_LogCommandPool.IsInitialized Then
'            m_LogCommandPool.ClearAllObjects
'        End If
'        Set m_LogCommandPool = Nothing
'    End If
'
'    If Not m_DataCommandPool Is Nothing Then
'        If m_DataCommandPool.IsInitialized Then
'            m_DataCommandPool.ClearAllObjects
'        End If
'        Set m_DataCommandPool = Nothing
'    End If
'
'    If Not m_UICommandPool Is Nothing Then
'        If m_UICommandPool.IsInitialized Then
'            m_UICommandPool.ClearAllObjects
'        End If
'        Set m_UICommandPool = Nothing
'    End If
'
'    Set m_CommandInvoker = Nothing
'
'    ' Проверяем, были ли ошибки
'    If Err.Number <> 0 Then
'        Debug.Print "Error in ReleaseAllPools: " & Err.Description
'        Err.Clear
'    End If
'
'    On Error GoTo 0
'End Sub
'


' Модуль: CommandPoolModule.bas
' Описание: Общий модуль для реализации паттерна Singleton для пулов команд
Option Explicit

' Экземпляры пулов команд
Private m_LogCommandPool As CommandPool
Private m_DataCommandPool As CommandPool
Private m_UICommandPool As CommandPool
Private m_CommandFactory As CommandFactory


' Экземпляр инвокера команд
Private m_CommandInvoker As CommandInvoker

' Флаг инициализации модуля
Private m_ModuleInitialized As Boolean

' Инициализация всех пулов
Public Sub InitializeAllPools()
    If m_ModuleInitialized Then Exit Sub
    
    ' Инициализируем пулы команд
    Set m_LogCommandPool = New CommandPool
    m_LogCommandPool.Initialize "LogCommands", 10, 30
    
    Set m_DataCommandPool = New CommandPool
    m_DataCommandPool.Initialize "DataCommands", 15, 30
    
    Set m_UICommandPool = New CommandPool
    m_UICommandPool.Initialize "UICommands", 8, 30
    
    ' Инициализируем инвокер
    Set m_CommandInvoker = New CommandInvoker
    
    m_ModuleInitialized = True
    
    ' Выводим информацию об инициализации
    Debug.Print "=== Все пулы команд инициализированы ==="
End Sub

' Получение экземпляра пула команд для логирования
Public Function GetLogCommandPool() As CommandPool
    ' Убедимся, что все пулы инициализированы
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    Set GetLogCommandPool = m_LogCommandPool
End Function

' Получение экземпляра пула команд для работы с данными
Public Function GetDataCommandPool() As CommandPool
    ' Убедимся, что все пулы инициализированы
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    Set GetDataCommandPool = m_DataCommandPool
End Function

' Получение экземпляра пула команд для интерфейса
Public Function GetUICommandPool() As CommandPool
    ' Убедимся, что все пулы инициализированы
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    Set GetUICommandPool = m_UICommandPool
End Function

' Получение экземпляра инвокера команд
Public Function GetCommandInvoker() As CommandInvoker
    ' Убедимся, что все пулы инициализированы
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    Set GetCommandInvoker = m_CommandInvoker
End Function

' Получение экземпляра фабрики команд
Public Function GetCommandFactory() As CommandFactory
    ' Убедимся, что все компоненты инициализированы
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    If m_CommandFactory Is Nothing Then
        Set m_CommandFactory = New CommandFactory
    End If
    
    Set GetCommandFactory = m_CommandFactory
End Function




' Получение статистики всех пулов команд
Public Function GetPoolsStatistics() As String
    ' Убедимся, что все пулы инициализированы
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    ' Собираем статистику о пулах и командах
    Dim stats As String
    stats = "======= Command Pool Statistics =======" & vbCrLf & vbCrLf
    
    ' Добавляем статистику о пуле команд логирования
    stats = stats & "--- Log Commands Pool ---" & vbCrLf
    stats = stats & "Available commands: " & m_LogCommandPool.AvailableObjectCount & vbCrLf
    stats = stats & "In-use commands: " & m_LogCommandPool.InUseObjectCount & vbCrLf
    stats = stats & "Max pool size: " & m_LogCommandPool.MaxPoolSize & vbCrLf & vbCrLf
    
    ' Добавляем статистику о пуле команд для работы с данными
    stats = stats & "--- Data Commands Pool ---" & vbCrLf
    stats = stats & "Available commands: " & m_DataCommandPool.AvailableObjectCount & vbCrLf
    stats = stats & "In-use commands: " & m_DataCommandPool.InUseObjectCount & vbCrLf
    stats = stats & "Max pool size: " & m_DataCommandPool.MaxPoolSize & vbCrLf & vbCrLf
    
    ' Добавляем статистику о пуле команд для интерфейса
    stats = stats & "--- UI Commands Pool ---" & vbCrLf
    stats = stats & "Available commands: " & m_UICommandPool.AvailableObjectCount & vbCrLf
    stats = stats & "In-use commands: " & m_UICommandPool.InUseObjectCount & vbCrLf
    stats = stats & "Max pool size: " & m_UICommandPool.MaxPoolSize & vbCrLf & vbCrLf
    
    ' Добавляем информацию о командах в истории
    stats = stats & "--- Command History ---" & vbCrLf
    stats = stats & "Commands in history: " & m_CommandInvoker.CommandHistoryCount & vbCrLf
    stats = stats & "Can undo operations: " & IIf(m_CommandInvoker.CanUndo, "Yes", "No") & vbCrLf & vbCrLf
    
    stats = stats & "======================================"
    
    GetPoolsStatistics = stats
End Function

' Очистка всех пулов команд
Public Sub ReleaseAllPools()
    On Error Resume Next
    
    If Not m_ModuleInitialized Then Exit Sub
    
    ' Сохраняем статистику перед освобождением ресурсов
    Debug.Print "=== Статистика перед освобождением ресурсов ==="
    Debug.Print GetPoolsStatistics()
    
    If Not m_LogCommandPool Is Nothing Then
        m_LogCommandPool.ClearAllObjects
        Set m_LogCommandPool = Nothing
    End If
    
    If Not m_DataCommandPool Is Nothing Then
        m_DataCommandPool.ClearAllObjects
        Set m_DataCommandPool = Nothing
    End If
    
    If Not m_UICommandPool Is Nothing Then
        m_UICommandPool.ClearAllObjects
        Set m_UICommandPool = Nothing
    End If
    
    Set m_CommandInvoker = Nothing
    
    m_ModuleInitialized = False
    
    Debug.Print "=== Все ресурсы пулов освобождены ==="
    
    If Err.Number <> 0 Then
        Debug.Print "Ошибка при освобождении ресурсов: " & Err.Description
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub


' Добавить в любой удобный модуль
Public Sub InitBeforeDemo()
    On Error Resume Next
    ' Освобождаем все ресурсы перед демонстрацией
    ReleaseAllPools
    
    ' Принудительно инициализируем пулы
    InitializeAllPools
    
    ' Выводим информацию
    Debug.Print "=== Пулы инициализированы и готовы к демонстрации ==="
    Debug.Print GetPoolsStatistics()
    
    MsgBox "Все пулы команд инициализированы и готовы к демонстрации!", _
           vbInformation, "Подготовка"
End Sub

