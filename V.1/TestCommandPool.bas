Attribute VB_Name = "TestCommandPool"
' Модуль: CommandPoolDemo.bas
' Улучшенная демонстрация паттернов Command и Object Pool
Option Explicit

' Константы для визуального оформления
Private Const DEMO_TITLE As String = "Демонстрация паттернов Command и Object Pool"

' Запуск полной демонстрации
'Public Sub RunFullDemo()
'    ' Подготовка
'    ShowWelcomeMessage
'
'    ' Основная демонстрация
'    DemoCommandExecution
'    DemoCommandHistory
'    DemoObjectPoolStatistics
'
'    ' Завершение
'    ShowFinalMessage
'End Sub

' Демонстрация выполнения команд
Private Sub DemoCommandExecution()
    ' Получаем инвокер команд
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    ' Показываем информационное сообщение
    MsgBox "Сейчас будет продемонстрировано выполнение команд." & vbCrLf & _
           "Вы увидите запрос на ввод и сообщения о выполнении команд.", _
           vbInformation, DEMO_TITLE & " - Шаг 1"
    
    ' 1. Запрос пользовательского ввода
    Dim requestInput As New RequestInputCommand
    requestInput.Initialize "Введите ваше имя для демонстрации:", "Ввод данных", "Пользователь"
    
    ' Преобразуем в интерфейс и выполняем
    Dim inputCmd As ICommand
    Set inputCmd = requestInput
    invoker.ExecuteCommand inputCmd
    
    ' Получаем результат
    Dim userName As String
    userName = requestInput.result
    
    ' 2. Выводим результат визуально
    MsgBox "Вы ввели: " & userName & vbCrLf & _
           "Эти данные были получены с помощью команды RequestInputCommand и сохранены.", _
           vbInformation, DEMO_TITLE & " - Шаг 2"
    
    ' 3. Демонстрация добавления записи
    MsgBox "Сейчас будет выполнена команда AddRecordCommand, которая добавит запись в систему." & vbCrLf & _
           "Эта команда поддерживает отмену операции.", _
           vbInformation, DEMO_TITLE & " - Шаг 3"
    
    Dim addCmd As New AddRecordCommand
    addCmd.Initialize userName, "UsersTable"
    
    ' Выполняем команду
    invoker.ExecuteCommand addCmd
    
    ' 4. Демонстрация удаления записи
    MsgBox "Сейчас будет выполнена команда DeleteRecordCommand, которая удалит тестовую запись." & vbCrLf & _
           "Эта команда также поддерживает отмену операции.", _
           vbInformation, DEMO_TITLE & " - Шаг 4"
    
    Dim delCmd As New DeleteRecordCommand
    delCmd.Initialize "DEMO_ID_123", "DemoTable"
    
    ' Выполняем команду
    invoker.ExecuteCommand delCmd
    
    ' 5. Демонстрация логирования ошибки
    MsgBox "Сейчас будет выполнена команда LogErrorCommand, которая запишет информацию об ошибке." & vbCrLf & _
           "Результат можно будет увидеть в окне Immediate (Ctrl+G).", _
           vbInformation, DEMO_TITLE & " - Шаг 5"
    
    Dim errorCmd As New LogErrorCommand
    errorCmd.Initialize "Демонстрационная ошибка", "DemoModule", 1234, "Это не настоящая ошибка, а демонстрация"
    
    ' Выполняем команду
    invoker.ExecuteCommand errorCmd
End Sub

' Демонстрация истории команд
Private Sub DemoCommandHistory()
    ' Получаем инвокер команд
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    ' Показываем информацию об истории команд
    MsgBox "В истории команд сейчас находится " & invoker.CommandHistoryCount & " команд(ы), поддерживающих отмену." & vbCrLf & _
           "Сейчас будет продемонстрирована отмена последней команды.", _
           vbInformation, DEMO_TITLE & " - История команд"
    
    ' Отменяем последнюю команду, если есть такая возможность
    If invoker.CanUndo Then
        invoker.UndoLastCommand
        
        ' Уведомляем о выполненной отмене
        MsgBox "Команда была успешно отменена!" & vbCrLf & _
               "Теперь в истории " & invoker.CommandHistoryCount & " команд(ы).", _
               vbInformation, DEMO_TITLE & " - Результат отмены"
    Else
        MsgBox "В истории нет команд, поддерживающих отмену.", _
               vbExclamation, DEMO_TITLE & " - История команд"
    End If
End Sub

'' Демонстрация статистики пула объектов
'Private Sub DemoObjectPoolStatistics()
'    ' Получаем статистику пулов
'    Dim stats As String
'    stats = GetPoolsStatistics()
'
'    ' Показываем статистику
'    MsgBox "Статистика пулов объектов:" & vbCrLf & vbCrLf & _
'           stats & vbCrLf & vbCrLf & _
'           "Эта статистика показывает использование паттерна Object Pool для управления командами.", _
'           vbInformation, DEMO_TITLE & " - Статистика пулов"
'
'    ' Выводим также в окно Immediate для архивных целей
'    Debug.Print "===== СТАТИСТИКА ПУЛОВ ОБЪЕКТОВ ====="
'    Debug.Print stats
'    Debug.Print "===================================="
'End Sub

' Приветственное сообщение
Private Sub ShowWelcomeMessage()
    MsgBox "Добро пожаловать в демонстрацию паттернов Command и Object Pool!" & vbCrLf & vbCrLf & _
           "Эта демонстрация покажет:" & vbCrLf & _
           "1. Создание и выполнение команд" & vbCrLf & _
           "2. Отмену команд через историю" & vbCrLf & _
           "3. Управление объектами через пул" & vbCrLf & vbCrLf & _
           "Нажмите OK, чтобы начать демонстрацию.", _
           vbInformation, DEMO_TITLE
End Sub

' Финальное сообщение
Private Sub ShowFinalMessage()
    ' Освобождаем ресурсы
    On Error Resume Next
    ReleaseAllPools
    If Err.Number <> 0 Then
        Debug.Print "Ошибка при освобождении ресурсов: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    ' Выводим финальное сообщение
    MsgBox "Демонстрация паттернов Command и Object Pool завершена!" & vbCrLf & vbCrLf & _
           "Что было продемонстрировано:" & vbCrLf & _
           "• Паттерн Command для инкапсуляции действий" & vbCrLf & _
           "• Отмена операций через историю команд" & vbCrLf & _
           "• Паттерн Object Pool для эффективного управления объектами" & vbCrLf & vbCrLf & _
           "Результаты выполнения также доступны в окне Immediate (Ctrl+G).", _
           vbInformation, DEMO_TITLE & " - Завершение"
End Sub

' Простой тест на работоспособность модулей
Public Sub QuickTest()
    ' Создаем тестовую команду
    Dim testCmd As New LogInfoCommand
    testCmd.Initialize "Тестовая команда выполнена", "QuickTest"
    
    ' Получаем инвокер и выполняем команду
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    invoker.ExecuteCommand testCmd
    
    ' Выводим сообщение
    MsgBox "Тест выполнен успешно! Проверьте окно Immediate (Ctrl+G).", _
           vbInformation, "Быстрый тест"
End Sub

' Добавить в модуль CommandPoolDemo.bas
Public Sub PreloadPoolsDemo()
    ' Инициализируем все пулы
    InitializeAllPools
    
    ' Предзагрузка команд в пулы
    MsgBox "Сейчас будет выполнена предзагрузка команд в пулы." & vbCrLf & _
           "Это позволит продемонстрировать повторное использование объектов.", _
           vbInformation, DEMO_TITLE & " - Предзагрузка"
    
    ' Загружаем несколько команд в пул логирования
    Dim i As Integer
    For i = 1 To 3
        ' Создаем команду логирования
        Dim logCmd As New LogInfoCommand
        logCmd.Initialize "Предзагруженная команда #" & i, "PreloadDemo"
        
        ' Выполняем команду
        Dim invoker As CommandInvoker
        Set invoker = GetCommandInvoker()
        invoker.ExecuteCommand logCmd
        
        ' Возвращаем в пул
        GetLogCommandPool().ReturnCommand logCmd
    Next i
    
    ' Загружаем команду в пул данных
    Dim dataCmd As New AddRecordCommand
    dataCmd.Initialize "Preloaded Data", "TestTable"
    invoker.ExecuteCommand dataCmd
    GetDataCommandPool().ReturnCommand dataCmd
    
    ' Показываем статистику после предзагрузки
    MsgBox "Предзагрузка завершена. Теперь в пулах есть объекты для повторного использования." & vbCrLf & vbCrLf & _
           "Статистика пулов:" & vbCrLf & vbCrLf & _
           GetPoolsStatistics(), _
           vbInformation, DEMO_TITLE & " - Результаты предзагрузки"
End Sub

' Обновленная версия DemoObjectPoolStatistics
Private Sub DemoObjectPoolStatistics()
    ' Принудительно запрашиваем несколько команд, чтобы убедиться, что статистика будет не пустой
    Dim logCmd As ICommand
    Set logCmd = GetLogCommandPool().GetCommand("LogInfoCommand")
    Dim dataCmd As ICommand
    Set dataCmd = GetDataCommandPool().GetCommand("AddRecordCommand")
    Dim uiCmd As ICommand
    Set uiCmd = GetUICommandPool().GetCommand("ShowMessageCommand")
    
    ' Не выполняем команды, только возвращаем их в пул
    GetLogCommandPool().ReturnCommand logCmd
    GetDataCommandPool().ReturnCommand dataCmd
    GetUICommandPool().ReturnCommand uiCmd
    
    ' Получаем статистику пулов
    Dim stats As String
    stats = GetPoolsStatistics()
    
    ' Проверяем, что статистика не пустая
    If Len(stats) < 100 Then
        ' Принудительная генерация статистики
        stats = "======= Command Pool Statistics =======" & vbCrLf & vbCrLf
        stats = stats & "--- Log Commands Pool ---" & vbCrLf
        stats = stats & "Available commands: " & GetLogCommandPool().AvailableObjectCount & vbCrLf
        stats = stats & "In-use commands: " & GetLogCommandPool().InUseObjectCount & vbCrLf
        stats = stats & "Max pool size: " & GetLogCommandPool().MaxPoolSize & vbCrLf & vbCrLf
        
        stats = stats & "--- Data Commands Pool ---" & vbCrLf
        stats = stats & "Available commands: " & GetDataCommandPool().AvailableObjectCount & vbCrLf
        stats = stats & "In-use commands: " & GetDataCommandPool().InUseObjectCount & vbCrLf
        stats = stats & "Max pool size: " & GetDataCommandPool().MaxPoolSize & vbCrLf & vbCrLf
        
        stats = stats & "--- UI Commands Pool ---" & vbCrLf
        stats = stats & "Available commands: " & GetUICommandPool().AvailableObjectCount & vbCrLf
        stats = stats & "In-use commands: " & GetUICommandPool().InUseObjectCount & vbCrLf
        stats = stats & "Max pool size: " & GetUICommandPool().MaxPoolSize & vbCrLf & vbCrLf
        
        stats = stats & "--- Command History ---" & vbCrLf
        stats = stats & "Commands in history: " & GetCommandInvoker().CommandHistoryCount & vbCrLf
        stats = stats & "Can undo operations: " & IIf(GetCommandInvoker().CanUndo, "Yes", "No") & vbCrLf & vbCrLf
        
        stats = stats & "======================================"
    End If
    
    ' Показываем статистику
    MsgBox "Статистика пулов объектов:" & vbCrLf & vbCrLf & _
           stats & vbCrLf & vbCrLf & _
           "Эта статистика показывает использование паттерна Object Pool для управления командами." & vbCrLf & _
           "Обратите внимание на соотношение доступных и используемых объектов.", _
           vbInformation, DEMO_TITLE & " - Статистика пулов"
    
    ' Выводим также в окно Immediate для архивных целей
    Debug.Print "===== СТАТИСТИКА ПУЛОВ ОБЪЕКТОВ ====="
    Debug.Print stats
    Debug.Print "===================================="
End Sub

' Улучшенная версия RunFullDemo
Public Sub RunFullDemo()
    ' Подготовка
    ShowWelcomeMessage
    
    ' Предзагрузка пулов для демонстрации повторного использования
    PreloadPoolsDemo
    
    ' Основная демонстрация
    DemoCommandExecution
    DemoCommandHistory
    DemoObjectPoolStatistics
    
    ' Завершение
    ShowFinalMessage
End Sub

