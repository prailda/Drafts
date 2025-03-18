VBA
'-------------------------------------------
' Component: SetValueCommand
'-------------------------------------------
' Класс SetCellValueCommand
' Команда для установки значения ячейки с использованием новой архитектуры
Option Explicit

Implements ICommandWithValidation

' Внутренняя структура для хранения данных команды
Private Type TSetCellValueCommand
    sheetName As String          ' Имя листа
    cellAddress As String        ' Адрес ячейки
    newValue As Variant          ' Новое значение
    OldValue As Variant          ' Старое значение для отмены
    ValidationManager As ValidationManager ' Менеджер валидации
    logger As logger             ' Логгер
    commandName As String        ' Имя команды
    ExecutionTimestamp As Date   ' Время выполнения
    UndoTimestamp As Date        ' Время отмены
    ExecutedSuccessfully As Boolean ' Успешно ли выполнена
    UndoneSuccessfully As Boolean ' Успешно ли отменена
    ValidationErrors As String   ' Ошибки валидации
    IsValidated As Boolean       ' Была ли выполнена валидация
    IsValid As Boolean           ' Валидна ли команда
End Type

Private this As TSetCellValueCommand

' Инициализация
Private Sub Class_Initialize()
    Set this.ValidationManager = ValidationManager.GetInstance
    Set this.logger = logger.GetInstance
    this.commandName = "SetCellValueCommand"
    this.ExecutedSuccessfully = False
    this.UndoneSuccessfully = False
    this.ValidationErrors = ""
    this.IsValidated = False
    this.IsValid = False
End Sub

' Инициализация команды
Public Sub Initialize(sheetName As String, cellAddress As String, newValue As Variant)
    this.sheetName = sheetName
    this.cellAddress = cellAddress
    this.newValue = newValue
    
    ' Сбрасываем состояние валидации при изменении параметров
    this.IsValidated = False
    this.ValidationErrors = ""
End Sub

' Реализация ICommandWithValidation
Private Sub ICommandWithValidation_Execute()
    On Error GoTo errorHandler
    
    ' Проверяем валидность перед выполнением
    If Not ICommandWithValidation_IsValid() Then
        err.Raise vbObjectError + 1002, "SetCellValueCommand", _
                 "Cannot execute invalid command: " & this.ValidationErrors
        Exit Sub
    End If
    
    ' Записываем время выполнения
    this.ExecutionTimestamp = Now
    
    ' Логируем начало выполнения
    this.logger.LogInfo "Выполнение команды", this.commandName
    
    ' Сохраняем старое значение для возможности отмены
    SaveOldValue
    
    ' Устанавливаем новое значение
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(this.sheetName)
    ws.Range(this.cellAddress).value = this.newValue
    
    ' Отмечаем успешное выполнение
    this.ExecutedSuccessfully = True
    
    ' Логируем успешное выполнение
    this.logger.LogInfo "Установлено значение '" & CStr(this.newValue) & "' в ячейку " & _
                        this.sheetName & "!" & this.cellAddress, this.commandName
    
    Exit Sub
    
errorHandler:
    this.ExecutedSuccessfully = False
    
    ' Логируем ошибку
    this.logger.LogError "Ошибка выполнения: " & err.description, this.commandName
    
    ' Пробрасываем ошибку дальше
    err.Raise err.Number, "SetCellValueCommand.Execute", err.description
End Sub

Private Sub ICommandWithValidation_Undo()
    On Error GoTo errorHandler
    
    ' Записываем время отмены
    this.UndoTimestamp = Now
    
    ' Логируем начало отмены
    this.logger.LogInfo "Отмена команды", this.commandName
    
    ' Восстанавливаем старое значение
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(this.sheetName)
    ws.Range(this.cellAddress).value = this.OldValue
    
    ' Отмечаем успешную отмену
    this.UndoneSuccessfully = True
    
    ' Логируем успешную отмену
    this.logger.LogInfo "Восстановлено значение '" & CStr(this.OldValue) & "' в ячейке " & _
                        this.sheetName & "!" & this.cellAddress, this.commandName
    
    Exit Sub
    
errorHandler:
    this.UndoneSuccessfully = False
    
    ' Логируем ошибку
    this.logger.LogError "Ошибка отмены: " & err.description, this.commandName
    
    ' Пробрасываем ошибку дальше
    err.Raise err.Number, "SetCellValueCommand.Undo", err.description
End Sub

Private Function ICommandWithValidation_Validate() As Boolean
    ' Очищаем результаты предыдущей валидации
    this.ValidationErrors = ""
    this.IsValidated = True
    this.IsValid = True
    
    ' Проверяем имя листа
    If Not this.ValidationManager.ValidateWorksheetName(this.sheetName) Then
        this.ValidationErrors = this.ValidationManager.GetErrorsAsString()
        this.IsValid = False
        ICommandWithValidation_Validate = False
        Exit Function
    End If
    
    ' Проверяем существование листа
    If Not ValidateWorksheetExists() Then
        this.IsValid = False
        ICommandWithValidation_Validate = False
        Exit Function
    End If
    
    ' Проверяем адрес ячейки
    If Not this.ValidationManager.ValidateCellAddress(this.cellAddress) Then
        If Len(this.ValidationErrors) > 0 Then this.ValidationErrors = this.ValidationErrors & vbCrLf
        this.ValidationErrors = this.ValidationErrors & this.ValidationManager.GetErrorsAsString()
        this.IsValid = False
        ICommandWithValidation_Validate = False
        Exit Function
    End If
    
    ' Все проверки пройдены
    ICommandWithValidation_Validate = True
End Function

Private Function ICommandWithValidation_GetValidationErrors() As String
    ' Если валидация еще не выполнялась, выполняем ее сейчас
    If Not this.IsValidated Then
        ICommandWithValidation_Validate
    End If
    
    ICommandWithValidation_GetValidationErrors = this.ValidationErrors
End Function

Private Function ICommandWithValidation_IsValid() As Boolean
    ' Если валидация еще не выполнялась, выполняем ее сейчас
    If Not this.IsValidated Then
        ICommandWithValidation_Validate
    End If
    
    ICommandWithValidation_IsValid = this.IsValid
End Function

Private Function ICommandWithValidation_GetCommandName() As String
    ICommandWithValidation_GetCommandName = this.commandName
End Function

Private Function ICommandWithValidation_WasExecutedSuccessfully() As Boolean
    ICommandWithValidation_WasExecutedSuccessfully = this.ExecutedSuccessfully
End Function

Private Function ICommandWithValidation_WasUndoneSuccessfully() As Boolean
    ICommandWithValidation_WasUndoneSuccessfully = this.UndoneSuccessfully
End Function

Private Function ICommandWithValidation_GetExecutionTimestamp() As Date
    ICommandWithValidation_GetExecutionTimestamp = this.ExecutionTimestamp
End Function

Private Function ICommandWithValidation_GetUndoTimestamp() As Date
    ICommandWithValidation_GetUndoTimestamp = this.UndoTimestamp
End Function

' Вспомогательные методы
Private Sub SaveOldValue()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(this.sheetName)
    this.OldValue = ws.Range(this.cellAddress).value
    
    If err.Number <> 0 Then
        this.logger.LogWarning "Не удалось получить старое значение для отмены: " & err.description, this.commandName
    End If
    
    On Error GoTo 0
End Sub

Private Function ValidateWorksheetExists() As Boolean
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(this.sheetName)
    
    If ws Is Nothing Or err.Number <> 0 Then
        If Len(this.ValidationErrors) > 0 Then this.ValidationErrors = this.ValidationErrors & vbCrLf
        this.ValidationErrors = this.ValidationErrors & "Лист '" & this.sheetName & "' не существует"
        ValidateWorksheetExists = False
    Else
        ValidateWorksheetExists = True
    End If
    
    On Error GoTo 0
End Function

' Геттеры для доступа к внутренним данным
Public Property Get sheetName() As String
    sheetName = this.sheetName
End Property

Public Property Get cellAddress() As String
    cellAddress = this.cellAddress
End Property

Public Property Get newValue() As Variant
    newValue = this.newValue
End Property

Public Property Get OldValue() As Variant
    OldValue = this.OldValue
End Property

Public Property Get commandName() As String
    commandName = this.commandName
End Property

Public Property Let commandName(value As String)
    this.commandName = value
End Property


'-------------------------------------------
' Component: CreateWorksheetCommand
'-------------------------------------------
'Ў Класс CreateWorksheetCommand Ў
'Ў Класс CreateWorksheetCommand Ў
Private Type TCommandParams
    TargetWorkbook As Workbook
    sheetName As String
    SheetIndex As Long
    TabColor As Long
    Visible As Boolean
    Protect As Boolean
    Gridlines As Boolean
    Headers As Boolean
    GroupWith As Worksheet
    UseTemplate As Boolean
    IsValid As Boolean
    OperationResult As Boolean
End Type

Private this As TCommandParams
Private Const TEMPLATE_SHEET As String = "TemplateSheet"

' Добавлено: Публичное свойство для результата операции
Public Property Get OperationResult() As Boolean
    OperationResult = this.OperationResult
End Property

' ... остальной код класса без изменений ...
'^ Конец класса ^

' Инициализация команды
Public Sub Initialize( _
    wb As Workbook, _
    sheetName As String, _
    Optional UseTemplate As Boolean = False, _
    Optional IsValid As Boolean = True)
    
    Set this.TargetWorkbook = wb
    this.sheetName = sheetName
    this.UseTemplate = UseTemplate
    this.IsValid = IsValid
    SetDefaults
End Sub

' Установка параметров по умолчанию
Private Sub SetDefaults()
    this.SheetIndex = 0
    this.TabColor = RGB(255, 255, 255)
    this.Visible = True
    this.Protect = False
    this.Gridlines = True
    this.Headers = True
    this.IsValid = True
    this.OperationResult = False
End Sub

' Свойства для дополнительных параметров
Public Property Let SheetIndex(value As Long): this.SheetIndex = value: End Property
Public Property Let TabColor(value As Long): this.TabColor = value: End Property
Public Property Let Visible(value As Boolean): this.Visible = value: End Property
Public Property Let Protect(value As Boolean): this.Protect = value: End Property
Public Property Let Gridlines(value As Boolean): this.Gridlines = value: End Property
Public Property Let Headers(value As Boolean): this.Headers = value: End Property
Public Property Let GroupWith(value As Worksheet): Set this.GroupWith = value: End Property

' Главный метод выполнения команды
Public Sub Execute()
    On Error GoTo errorHandler
    Dim ws As Worksheet
    
    Log "Command execution started"
    
    If this.IsValid Then
        ValidateParameters
        If this.UseTemplate Then
            Set ws = CreateFromTemplate
        Else
            Set ws = CreateDefault
        End If
        
        ConfigureSheet ws
        this.OperationResult = True
        Log "Sheet '" & this.sheetName & "' created successfully"
    Else
        ValidateParameters
        this.OperationResult = True
        Log "Validation passed"
    End If
    
    Exit Sub
    
errorHandler:
    this.OperationResult = False
    Log "Error " & err.Number & ": " & err.description
End Sub

' Создание из шаблона
Private Function CreateFromTemplate() As Worksheet
    If Not SheetExists(TEMPLATE_SHEET, this.TargetWorkbook) Then
        err.Raise vbObjectError + 1000, , "Template sheet not found"
    End If
    
    this.TargetWorkbook.Worksheets(TEMPLATE_SHEET).copy _
        Before:=this.TargetWorkbook.Worksheets(this.SheetIndex)
    Set CreateFromTemplate = ActiveSheet
    CreateFromTemplate.name = this.sheetName
End Function

' Обычное создание листа
Private Function CreateDefault() As Worksheet
    Set CreateDefault = this.TargetWorkbook.Worksheets.Add( _
        Before:=this.TargetWorkbook.Worksheets(this.SheetIndex))
    CreateDefault.name = this.sheetName
End Function

' Настройка параметров листа
Private Sub ConfigureSheet(ws As Worksheet)
    With ws
        .Tab.Color = this.TabColor
        .Visible = this.Visible
        .Activate
        
        ' Настройка отображения
        .DisplayGridlines = this.Gridlines
        .DisplayHeadings = this.Headers
        
        ' Группировка
        If Not this.GroupWith Is Nothing Then
            .Group = this.GroupWith.Group
        End If
        
        ' Защита
        If this.Protect Then
            .Protect Password:="", DrawingObjects:=True, Contents:=True
        End If
    End With
End Sub

' Валидация параметров
Private Sub ValidateParameters()
    ValidateWorkbook
    ValidateSheetName
    ValidateTemplate
End Sub

Private Sub ValidateWorkbook()
    If this.TargetWorkbook Is Nothing Then
        err.Raise vbObjectError + 1001, , "Target workbook not specified"
    End If
End Sub

Private Sub ValidateSheetName()
    If Not ValidationManager.ValidateName(this.TargetWorkbook, this.sheetName) Then
        err.Raise vbObjectError + 1002, , "Invalid sheet name"
    End If
End Sub

Private Sub ValidateTemplate()
    If this.UseTemplate And Not SheetExists(TEMPLATE_SHEET, this.TargetWorkbook) Then
        err.Raise vbObjectError + 1003, , "Template sheet missing"
    End If
End Sub

' Логирование
Private Sub Log(message As String)
    Debug.Print Format(Now, "yyyy-mm-dd hh:mm:ss") & " | " & message
End Sub

' Вспомогательные функции
Private Function SheetExists(sName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    SheetExists = (Not wb.Worksheets(sName) Is Nothing)
End Function
'^ Конец класса ^


'-------------------------------------------
' Component: ICommand
'-------------------------------------------
' Интерфейс ICommand
' Базовый интерфейс для всех команд
Option Explicit

' Выполнить команду
Public Sub Execute()
End Sub

' Отменить команду
Public Sub Undo()
End Sub

' Получить имя команды
Public Function GetCommandName() As String
End Function

' Проверить, была ли команда успешно выполнена
Public Function WasExecutedSuccessfully() As Boolean
End Function

' Проверить, была ли команда успешно отменена
Public Function WasUndoneSuccessfully() As Boolean
End Function

' Получить времененную метку выполнения
Public Function GetExecutionTimestamp() As Date
End Function

' Получить временную метку отмены
Public Function GetUndoTimestamp() As Date
End Function

'-------------------------------------------
' Component: Receiver
'-------------------------------------------
Option Explicit

' Класс-получатель, который выполняет фактическую работу
Public Sub DoSomething()
    MsgBox "Выполняется действие в получателе", vbInformation, "Command Pattern"
End Sub

Public Sub DoSomethingElse()
    MsgBox "Выполняется другое действие в получателе", vbInformation, "Command Pattern"
End Sub

Public Sub UndoSomething()
    MsgBox "Отмена действия в получателе", vbInformation, "Command Pattern"
End Sub

Public Sub UndoSomethingElse()
    MsgBox "Отмена другого действия в получателе", vbInformation, "Command Pattern"
End Sub

'-------------------------------------------
' Component: TestModule
'-------------------------------------------
' Модуль TestModule
' Содержит тестовые процедуры для проверки архитектуры паттерна Command
Option Explicit

' Главная тестовая процедура
Public Sub TestCommandPattern()
    ' Инициализация логгера с настройками для тестирования
    InitializeLogger
    
    ' Получаем ссылку на логгер для использования в тестах
    Dim logger As logger
    Set logger = GetLogger ' Используем функцию из модуля LogLevel
    
    ' Логируем начало тестирования
    logger.LogInfo "Начало тестирования паттерна Command", "TestModule"
    
    ' Тестируем базовые компоненты
    TestLogger
    TestValidation
    TestCommandExecution
    TestCommandUndo
    
    ' Логируем успешное завершение тестирования
    logger.LogInfo "Тестирование паттерна Command успешно завершено", "TestModule"
    
    ' Выводим результаты в MsgBox
    MsgBox "Тестирование успешно завершено! Проверьте журнал для получения подробной информации.", _
           vbInformation, "Тестирование командного паттерна"
End Sub

' Инициализация логгера
Private Sub InitializeLogger()
    Dim logger As logger
    Set logger = GetLogger ' Используем функцию из модуля LogLevel
    
    ' Настройка файла лога в папке документов пользователя
    Dim logPath As String
    logPath = Environ("USERPROFILE") & "\AppData\Local\Excellent VBA\Debug\Logs"
    
    ' Создаем папку, если она не существует
    On Error Resume Next
    If Dir(logPath, vbDirectory) = "" Then
        MkDir logPath
    End If
    On Error GoTo 0
    
    ' Устанавливаем файл лога
    logger.LogFile = logPath & "CommandPattern_" & Format(Now, "yyyymmdd_hhmmss") & ".log"
    
    ' Включаем все уровни логирования для тестирования
    logger.MinimumLevel = LogLevel.LogDebug
    
    ' Очищаем историю логов
    logger.ClearHistory
    
    ' Логируем информацию о начале тестирования
    logger.LogInfo "Логгер инициализирован. Файл лога: " & logger.LogFile, "TestModule"
End Sub

' Тестирование системы логирования
Private Sub TestLogger()
    Dim logger As logger
    Set logger = GetLogger ' Используем функцию из модуля LogLevel
    
    logger.LogInfo "Начало тестирования системы логирования", "TestLogger"
    
    ' Тест разных уровней логирования
    logger.LogDebug "Это отладочное сообщение", "TestLogger"
    logger.LogInfo "Это информационное сообщение", "TestLogger"
    logger.LogWarning "Это предупреждение", "TestLogger"
    logger.LogError "Это сообщение об ошибке", "TestLogger"
    logger.LogCritical "Это критическая ошибка", "TestLogger"
    
    ' Проверка получения истории логов
    Dim LogHistory As Collection
    Set LogHistory = logger.GetHistory
    
    logger.LogInfo "Количество записей в истории логов: " & LogHistory.Count, "TestLogger"
    
    ' Проверка установки минимального уровня логирования
    logger.MinimumLevel = LogLevel.LogWarning
    logger.LogInfo "Это сообщение не должно отображаться из-за уровня логирования", "TestLogger"
    logger.LogError "Это сообщение об ошибке должно отображаться", "TestLogger"
    
    ' Восстановление уровня логирования
    logger.MinimumLevel = LogLevel.LogDebug
    
    logger.LogInfo "Тестирование системы логирования успешно завершено", "TestLogger"
End Sub

' Тестирование системы валидации
Private Sub TestValidation()
    Dim logger As logger
    Set logger = GetLogger
    logger.LogInfo "Начало тестирования системы валидации", "TestValidation"
    
    ' Получаем экземпляр менеджера валидации
    Dim ValidationManager As ValidationManager
    Set ValidationManager = GetValidationManager ' Используем функцию из модуля LogLevel
    
    ' Тестирование валидации имени листа
    Dim validSheetName As String
    Dim invalidSheetName As String
    
    validSheetName = "TestSheet"
    invalidSheetName = "Test*Sheet" ' Содержит недопустимый символ
    
    ' Проверяем валидное имя
    logger.LogInfo "Проверка валидного имени листа: " & validSheetName, "TestValidation"
    If ValidationManager.ValidateWorksheetName(validSheetName) Then
        logger.LogInfo "Валидация успешна", "TestValidation"
    Else
        logger.LogWarning "Валидация не прошла: " & ValidationManager.GetErrorsAsString(), "TestValidation"
    End If
    
    ' Проверяем невалидное имя
    logger.LogInfo "Проверка невалидного имени листа: " & invalidSheetName, "TestValidation"
    If Not ValidationManager.ValidateWorksheetName(invalidSheetName) Then
        logger.LogInfo "Валидация корректно не прошла: " & ValidationManager.GetErrorsAsString(), "TestValidation"
    Else
        logger.LogWarning "Валидация некорректно прошла для невалидного имени", "TestValidation"
    End If
    
    ' Тестирование валидации адреса ячейки
    Dim validCellAddress As String
    Dim invalidCellAddress As String
    
    validCellAddress = "A1"
    invalidCellAddress = "ZZ99999" ' Предположим, что это невалидный адрес
    
    ' Проверяем валидный адрес
    logger.LogInfo "Проверка валидного адреса ячейки: " & validCellAddress, "TestValidation"
    If ValidationManager.ValidateCellAddress(validCellAddress) Then
        logger.LogInfo "Валидация успешна", "TestValidation"
    Else
        logger.LogWarning "Валидация не прошла: " & ValidationManager.GetErrorsAsString(), "TestValidation"
    End If
    
    ' Проверяем невалидный адрес (для адреса ZZ99999 валидация может и пройти, зависит от версии Excel)
    logger.LogInfo "Проверка потенциально невалидного адреса ячейки: " & invalidCellAddress, "TestValidation"
    If ValidationManager.ValidateCellAddress(invalidCellAddress) Then
        logger.LogInfo "Валидация прошла (возможно, адрес валиден в текущей версии Excel)", "TestValidation"
    Else
        logger.LogInfo "Валидация корректно не прошла: " & ValidationManager.GetErrorsAsString(), "TestValidation"
    End If
    
    logger.LogInfo "Тестирование системы валидации успешно завершено", "TestValidation"
End Sub

' Тестирование выполнения команды
Private Sub TestCommandExecution()
    Dim logger As logger
    Set logger = GetLogger
    
    logger.LogInfo "Начало тестирования выполнения команды", "TestCommandExecution"
    
    ' Убедимся, что существует хотя бы один лист
    EnsureTestSheet
    
    ' Создаем команду установки значения ячейки
    Dim SetCellValueCommand As New SetCellValueCommand
    SetCellValueCommand.Initialize "Sheet1", "A1", "Тестовое значение"
    
    ' Получаем инвокер
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker ' Используем функцию из модуля LogLevel
    
    ' Устанавливаем команду в инвокер
    invoker.SetCommand SetCellValueCommand
    
    ' Выполняем команду
    logger.LogInfo "Выполнение команды установки значения в ячейку A1", "TestCommandExecution"
    If invoker.ExecuteCommand() Then
        logger.LogInfo "Команда успешно выполнена", "TestCommandExecution"
    Else
        logger.LogError "Ошибка выполнения команды: " & invoker.LastError, "TestCommandExecution"
    End If
    
    ' Проверяем, что значение действительно установлено
    Dim actualValue As String
    actualValue = ThisWorkbook.Worksheets("Sheet1").Range("A1").value
    
    logger.LogInfo "Значение в ячейке A1: " & actualValue, "TestCommandExecution"
    
    If actualValue = "Тестовое значение" Then
        logger.LogInfo "Значение установлено корректно", "TestCommandExecution"
    Else
        logger.LogError "Значение установлено некорректно", "TestCommandExecution"
    End If
    
    ' Проверяем количество команд в истории
    logger.LogInfo "Количество команд в истории: " & invoker.GetHistoryCount(), "TestCommandExecution"
    
    logger.LogInfo "Тестирование выполнения команды успешно завершено", "TestCommandExecution"
End Sub

' Тестирование отмены команды
Private Sub TestCommandUndo()
    Dim logger As logger
    Set logger = GetLogger
    
    logger.LogInfo "Начало тестирования отмены команды", "TestCommandUndo"
    
    ' Убедимся, что существует хотя бы один лист
    EnsureTestSheet
    
    ' Создаем команду установки значения ячейки
    Dim SetCellValueCommand As New SetCellValueCommand
    SetCellValueCommand.Initialize "Sheet1", "B1", "Значение для отмены"
    
    ' Получаем инвокер
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker ' Используем функцию из модуля LogLevel
    
    ' Устанавливаем команду в инвокер
    invoker.SetCommand SetCellValueCommand
    
    ' Запоминаем текущее значение
    Dim originalValue As Variant
    originalValue = ThisWorkbook.Worksheets("Sheet1").Range("B1").value
    
    logger.LogInfo "Исходное значение в ячейке B1: " & IIf(IsEmpty(originalValue), "[EMPTY]", originalValue), "TestCommandUndo"
    
    ' Выполняем команду
    logger.LogInfo "Выполнение команды установки значения в ячейку B1", "TestCommandUndo"
    If invoker.ExecuteCommand() Then
        logger.LogInfo "Команда успешно выполнена", "TestCommandUndo"
    Else
        logger.LogError "Ошибка выполнения команды: " & invoker.LastError, "TestCommandUndo"
        Exit Sub
    End If
    
    ' Проверяем, что значение действительно установлено
    Dim newValue As String
    newValue = ThisWorkbook.Worksheets("Sheet1").Range("B1").value
    
    logger.LogInfo "Новое значение в ячейке B1: " & newValue, "TestCommandUndo"
    
    ' Отменяем команду
    logger.LogInfo "Отмена последней команды", "TestCommandUndo"
    If invoker.UndoLastCommand() Then
        logger.LogInfo "Команда успешно отменена", "TestCommandUndo"
    Else
        logger.LogError "Ошибка отмены команды: " & invoker.LastError, "TestCommandUndo"
        Exit Sub
    End If
    
    ' Проверяем, что значение восстановлено
    Dim restoredValue As Variant
    restoredValue = ThisWorkbook.Worksheets("Sheet1").Range("B1").value
    
    logger.LogInfo "Восстановленное значение в ячейке B1: " & IIf(IsEmpty(restoredValue), "[EMPTY]", restoredValue), "TestCommandUndo"
    
    ' Сравниваем с исходным значением (учитывая возможность Empty)
    If IsEmpty(originalValue) And IsEmpty(restoredValue) Then
        logger.LogInfo "Значение успешно восстановлено (пустое)", "TestCommandUndo"
    ElseIf originalValue = restoredValue Then
        logger.LogInfo "Значение успешно восстановлено", "TestCommandUndo"
    Else
        logger.LogError "Значение восстановлено некорректно", "TestCommandUndo"
    End If
    
    ' Проверяем количество команд в истории (должно уменьшиться после отмены)
    logger.LogInfo "Количество команд в истории: " & invoker.GetHistoryCount(), "TestCommandUndo"
    
    logger.LogInfo "Тестирование отмены команды успешно завершено", "TestCommandUndo"
End Sub

' В модуле ErrorTestModule модифицировать метод TestUncaughtErrorHandling
Private Sub TestUncaughtErrorHandling()
    Dim logger As Object
    Set logger = GetLogger
    
    Dim ErrorManager As ErrorManager
    Set ErrorManager = GetErrorManager
    
    logger.LogInfo "Начало тестирования обработки необработанных ошибок", "TestUncaughtErrorHandling"
    
    ' Сохраняем текущие обработчики
    Dim originalHandlers As Collection
    Set originalHandlers = ErrorManager.GetHandlers
    
    ' Временно очищаем обработчики для теста
    ErrorManager.ClearHandlers
    
    ' Включаем выбрасывание необработанных ошибок
    Dim originalThrowSetting As Boolean
    originalThrowSetting = ErrorManager.ThrowUnhandledErrors
    ErrorManager.ThrowUnhandledErrors = True
    
    ' Создаем нестандартную ошибку
    Dim errorNumber As Long
    errorNumber = vbObjectError + 99999
    
    ' Пытаемся обработать ошибку без обработчиков
    On Error Resume Next
    ErrorManager.HandleCustomError errorNumber, "Нестандартная ошибка для проверки", _
                              "TestUncaughtErrorHandling", "Проверка необработанных ошибок", "", 4
    
    ' Проверяем, была ли выброшена ошибка
    Dim uncaughtError As Boolean
    uncaughtError = (err.Number <> 0)
    
    ' Восстанавливаем настройки
    ErrorManager.ThrowUnhandledErrors = originalThrowSetting
    
    ' Восстанавливаем обработчики
    ErrorManager.RestoreHandlers originalHandlers
    
    On Error GoTo 0
    
    ' Логируем результат
    If uncaughtError Then
        logger.LogInfo "Тест пройден: необработанная ошибка была выброшена", "TestUncaughtErrorHandling"
    Else
        logger.LogWarning "Тест не пройден: необработанная ошибка не была выброшена", "TestUncaughtErrorHandling"
    End If
    
    logger.LogInfo "Тестирование обработки необработанных ошибок завершено", "TestUncaughtErrorHandling"
End Sub


' Вспомогательная функция для обеспечения наличия тестового листа
Private Sub EnsureTestSheet()
    On Error Resume Next
    
    ' Проверяем, существует ли лист Sheet1
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    ' Если не существует, создаем его
    If ws Is Nothing Then
        ' Добавляем новый лист с именем Sheet1
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = "Sheet1"
        
    Dim logger As logger
    Set logger = GetLogger
        logger.LogInfo "Создан тестовый лист Sheet1", "EnsureTestSheet"
    End If
    
    On Error GoTo 0
End Sub


'-------------------------------------------
' Component: Invoker
'-------------------------------------------
Option Explicit

Private mCommand As ICommandWithValidation
Private mCommandHistory As Collection

Private Sub Class_Initialize()
    Set mCommandHistory = New Collection
End Sub

' Установка команды
Public Sub SetCommand(command As ICommandWithValidation)
    Set mCommand = command
End Sub

' Выполнение команды с валидацией
Public Function ExecuteCommand() As Boolean
    If mCommand Is Nothing Then
        ExecuteCommand = False
        Exit Function
    End If
    
    ' Валидация перед выполнением
    If mCommand.Validate() Then
        mCommand.Execute
        mCommandHistory.Add mCommand
        ExecuteCommand = True
    Else
        ' Если команда не проходит валидацию
        If TypeOf mCommand Is ValidatedSetValueCommand Then
            Dim cmd As ValidatedSetValueCommand
            Set cmd = mCommand
            MsgBox "Ошибка валидации: " & cmd.GetErrorMessage(), vbExclamation, "Ошибка"
        Else
            MsgBox "Команда не прошла валидацию", vbExclamation, "Ошибка"
        End If
        ExecuteCommand = False
    End If
End Function

' Отмена последней команды
Public Sub UndoLastCommand()
    If mCommandHistory.Count > 0 Then
        Dim lastCommand As ICommandWithValidation
        Set lastCommand = mCommandHistory(mCommandHistory.Count)
        lastCommand.Undo
        mCommandHistory.Remove mCommandHistory.Count
    End If
End Sub

' Получение количества команд в истории
Public Function GetHistoryCount() As Integer
    GetHistoryCount = mCommandHistory.Count
End Function

'-------------------------------------------
' Component: ExcelReceiver
'-------------------------------------------
Option Explicit

' Данные для отмены операций
Private Type CellData
    value As Variant
    address As String
    Sheet As String
End Type

Private mLastCell As CellData

' Класс для работы с Excel
Public Sub SetCellValue(sheetName As String, cellAddress As String, value As Variant)
    ' Сохраняем текущее значение для отмены
    SaveCellState sheetName, cellAddress
    
    ' Устанавливаем новое значение
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Range(cellAddress).value = value
End Sub

Public Sub SetCellColor(sheetName As String, cellAddress As String, colorIndex As Integer)
    ' Сохраняем текущее значение для отмены (в данном случае только адрес)
    mLastCell.Sheet = sheetName
    mLastCell.address = cellAddress
    
    ' Устанавливаем новый цвет
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Range(cellAddress).Interior.colorIndex = colorIndex
End Sub

Public Sub UndoCellValue()
    If mLastCell.address <> "" Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(mLastCell.Sheet)
        ws.Range(mLastCell.address).value = mLastCell.value
    End If
End Sub

Public Sub UndoCellColor()
    If mLastCell.address <> "" Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets(mLastCell.Sheet)
        ws.Range(mLastCell.address).Interior.colorIndex = xlNone
    End If
End Sub

Private Sub SaveCellState(sheetName As String, cellAddress As String)
    mLastCell.Sheet = sheetName
    mLastCell.address = cellAddress
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(sheetName)
    mLastCell.value = ws.Range(cellAddress).value
End Sub

'-------------------------------------------
' Component: SetColorCommand
'-------------------------------------------
Option Explicit

Implements ICommand

Private mReceiver As ExcelReceiver
Private mSheetName As String
Private mCellAddress As String
Private mColorIndex As Integer

' Конструктор - также исправлено на ByVal
Public Sub Initialize(receiver As ExcelReceiver, sheetName As String, cellAddress As String, ByVal colorIndex As Integer)
    Set mReceiver = receiver
    mSheetName = sheetName
    mCellAddress = cellAddress
    mColorIndex = colorIndex
End Sub

' Реализация интерфейса ICommand
Private Sub ICommand_Execute()
    mReceiver.SetCellColor mSheetName, mCellAddress, mColorIndex
End Sub

Private Sub ICommand_Undo()
    mReceiver.UndoCellColor
End Sub

' Геттеры для сериализации
Public Function GetSheetName() As String
    GetSheetName = mSheetName
End Function

Public Function GetCellAddress() As String
    GetCellAddress = mCellAddress
End Function

Public Function GetColorIndex() As Integer
    GetColorIndex = mColorIndex
End Function

'-------------------------------------------
' Component: MacroCommand
'-------------------------------------------
Option Explicit

Implements ICommand

Private mCommands As Collection

Private Sub Class_Initialize()
    Set mCommands = New Collection
End Sub

' Добавление команды в макрокоманду
Public Sub AddCommand(cmd As ICommand)
    mCommands.Add cmd
End Sub

' Реализация интерфейса ICommand
Private Sub ICommand_Execute()
    Dim cmd As ICommand
    For Each cmd In mCommands
        cmd.Execute
    Next cmd
End Sub

Private Sub ICommand_Undo()
    Dim i As Integer
    ' Отмена команд в обратном порядке
    For i = mCommands.Count To 1 Step -1
        Dim cmd As ICommand
        Set cmd = mCommands(i)
        cmd.Undo
    Next i
End Sub

'-------------------------------------------
' Component: ICommandWithValidation
'-------------------------------------------
' Интерфейс ICommandWithValidation
' Расширяет ICommand добавляя валидацию
Option Explicit

' Наследуемые методы ICommand
Public Sub Execute()
End Sub

Public Sub Undo()
End Sub

Public Function GetCommandName() As String
End Function

Public Function WasExecutedSuccessfully() As Boolean
End Function

Public Function WasUndoneSuccessfully() As Boolean
End Function

Public Function GetExecutionTimestamp() As Date
End Function

Public Function GetUndoTimestamp() As Date
End Function

' Дополнительные методы для валидации
Public Function Validate() As Boolean
End Function

Public Function GetValidationErrors() As String
End Function

Public Function IsValid() As Boolean
End Function

'-------------------------------------------
' Component: ValidatedSetValueCommand
'-------------------------------------------
Option Explicit

Implements ICommandWithValidation

Private mReceiver As ExcelReceiver
Private mSheetName As String
Private mCellAddress As String
Private mValue As Variant
Private mErrorMessage As String

' Конструктор
Public Sub Initialize(receiver As ExcelReceiver, sheetName As String, cellAddress As String, value As Variant)
    Set mReceiver = receiver
    mSheetName = sheetName
    mCellAddress = cellAddress
    mValue = value
End Sub

' Валидация команды
Private Function ICommandWithValidation_Validate() As Boolean
    mErrorMessage = ""
    
    ' Проверяем существование листа
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(mSheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        mErrorMessage = "Лист '" & mSheetName & "' не найден"
        ICommandWithValidation_Validate = False
        Exit Function
    End If
    
    ' Проверяем корректность адреса ячейки
    On Error Resume Next
    Dim rng As Range
    Set rng = ws.Range(mCellAddress)
    On Error GoTo 0
    
    If rng Is Nothing Then
        mErrorMessage = "Некорректный адрес ячейки: " & mCellAddress
        ICommandWithValidation_Validate = False
        Exit Function
    End If
    
    ICommandWithValidation_Validate = True
End Function

' Получение сообщения об ошибке валидации
Public Function GetErrorMessage() As String
    GetErrorMessage = mErrorMessage
End Function

' Реализация интерфейса ICommandWithValidation
Private Sub ICommandWithValidation_Execute()
    mReceiver.SetCellValue mSheetName, mCellAddress, mValue
End Sub

Private Sub ICommandWithValidation_Undo()
    mReceiver.UndoCellValue
End Sub

'-------------------------------------------
' Component: ValidatingInvoker
'-------------------------------------------
Option Explicit

Private mCommand As ICommandWithValidation
Private mCommandHistory As Collection

Private Sub Class_Initialize()
    Set mCommandHistory = New Collection
End Sub

' Установка команды
Public Sub SetCommand(command As ICommandWithValidation)
    Set mCommand = command
End Sub

' Выполнение команды с валидацией
Public Function ExecuteCommand() As Boolean
    If mCommand Is Nothing Then
        ExecuteCommand = False
        Exit Function
    End If
    
    ' Валидация перед выполнением
    If mCommand.Validate() Then
        mCommand.Execute
        mCommandHistory.Add mCommand
        ExecuteCommand = True
    Else
        ' Если команда не проходит валидацию
        If TypeOf mCommand Is ValidatedSetValueCommand Then
            Dim cmd As ValidatedSetValueCommand
            Set cmd = mCommand
            MsgBox "Ошибка валидации: " & cmd.GetErrorMessage(), vbExclamation, "Ошибка"
        Else
            MsgBox "Команда не прошла валидацию", vbExclamation, "Ошибка"
        End If
        ExecuteCommand = False
    End If
End Function

' Отмена последней команды
Public Sub UndoLastCommand()
    If mCommandHistory.Count > 0 Then
        Dim lastCommand As ICommandWithValidation
        Set lastCommand = mCommandHistory(mCommandHistory.Count)
        lastCommand.Undo
        mCommandHistory.Remove mCommandHistory.Count
    End If
End Sub

' Получение количества команд в истории
Public Function GetHistoryCount() As Integer
    GetHistoryCount = mCommandHistory.Count
End Function

'-------------------------------------------
' Component: StandartInvoker
'-------------------------------------------
Option Explicit

Private mCommand As ICommand
Private mCommandHistory As Collection
Private mHistoryManager As CommandHistoryManager
Private mLogger As CommandLogger

Private Sub Class_Initialize()
    Set mCommandHistory = New Collection
    Set mHistoryManager = New CommandHistoryManager
    Set mLogger = New CommandLogger
End Sub

' Установка команды
Public Sub SetCommand(command As ICommand)
    Set mCommand = command
End Sub

' Выполнение команды с сохранением в истории
Public Sub ExecuteCommand()
    If Not mCommand Is Nothing Then
        ' Логируем выполнение
        mLogger.LogCommandExecution mCommand
        
        ' Выполняем команду
        mCommand.Execute
        
        ' Добавляем в историю
        mCommandHistory.Add mCommand
        mHistoryManager.AddToHistory mCommand
    End If
End Sub

' Отмена последней команды
Public Sub UndoLastCommand()
    If mCommandHistory.Count > 0 Then
        Dim lastCommand As ICommand
        Set lastCommand = mCommandHistory(mCommandHistory.Count)
        
        ' Логируем отмену
        mLogger.LogCommandUndo lastCommand
        
        ' Выполняем отмену
        lastCommand.Undo
        
        mCommandHistory.Remove mCommandHistory.Count
    End If
End Sub

' Получение количества команд в истории
Public Function GetHistoryCount() As Integer
    GetHistoryCount = mCommandHistory.Count
End Function

' Сохранить историю команд в файл
Public Sub SaveHistory()
    mHistoryManager.SaveHistoryToFile
End Sub

' Загрузить и выполнить историю команд из файла
Public Sub LoadAndReplayHistory(receiver As ExcelReceiver)
    mHistoryManager.LoadHistoryFromFile receiver
    mHistoryManager.ReplayHistory
End Sub

' Очистить историю команд
Public Sub ClearHistory()
    Set mCommandHistory = New Collection
    mHistoryManager.ClearHistory
End Sub

' Получить ссылку на логгер для настройки
Public Property Get logger() As CommandLogger
    Set logger = mLogger
End Property

'-------------------------------------------
' Component: CommandSerializer
'-------------------------------------------
Option Explicit

' Класс для сериализации и десериализации команд
Private Const DELIMITER As String = "|"

' Сериализация команды в строку
Public Function SerializeCommand(cmd As Object) As String
    Dim result As String
    
    If TypeOf cmd Is SetValueCommand Then
        Dim valueCmd As SetValueCommand
        Set valueCmd = cmd
        result = "SetValueCommand" & DELIMITER & _
                 valueCmd.GetSheetName() & DELIMITER & _
                 valueCmd.GetCellAddress() & DELIMITER & _
                 CStr(valueCmd.GetValue())
                 
    ElseIf TypeOf cmd Is SetColorCommand Then
        Dim colorCmd As SetColorCommand
        Set colorCmd = cmd
        result = "SetColorCommand" & DELIMITER & _
                 colorCmd.GetSheetName() & DELIMITER & _
                 colorCmd.GetCellAddress() & DELIMITER & _
                 CStr(colorCmd.GetColorIndex())
    Else
        result = "UnknownCommand"
    End If
    
    SerializeCommand = result
End Function

' Десериализация команды из строки
Public Function DeserializeCommand(serialized As String, receiver As ExcelReceiver) As ICommand
    Dim parts() As String
    parts = Split(serialized, DELIMITER)
    
    If UBound(parts) < 3 Then
        Set DeserializeCommand = Nothing
        Exit Function
    End If
    
    Dim cmdType As String
    cmdType = parts(0)
    
    Dim sheetName As String
    sheetName = parts(1)
    
    Dim cellAddress As String
    cellAddress = parts(2)
    
    ' Исправление ошибки ByRef Mismatch
    If cmdType = "SetValueCommand" Then
        Dim valueCmd As New SetValueCommand
        ' Создаем временную переменную для хранения значения из строки
        Dim valueToSet As Variant
        valueToSet = parts(3)
        valueCmd.Initialize receiver, sheetName, cellAddress, valueToSet
        Set DeserializeCommand = valueCmd
        
    ElseIf cmdType = "SetColorCommand" Then
        Dim colorCmd As New SetColorCommand
        ' Создаем временную переменную для хранения числа из строки
        Dim colorIndex As Integer
        colorIndex = CInt(parts(3))
        colorCmd.Initialize receiver, sheetName, cellAddress, colorIndex
        Set DeserializeCommand = colorCmd
        
    Else
        Set DeserializeCommand = Nothing
    End If
End Function


'-------------------------------------------
' Component: CommandHistoryManager
'-------------------------------------------
Option Explicit

Private mSerializer As CommandSerializer
Private mHistory As Collection
Private mHistoryFile As String
Private Const HistoryFilePath = "C:\Users\dalis\AppData\Local\Excellent VBA\Debug\Logs"


Private Sub Class_Initialize()
    Set mSerializer = New CommandSerializer
    Set mHistory = New Collection
    mHistoryFile = HistoryFilePath & "\command_history.txt"
End Sub

' Добавить команду в историю
Public Sub AddToHistory(cmd As ICommand)
    mHistory.Add cmd
End Sub

' Очистить историю
Public Sub ClearHistory()
    Set mHistory = New Collection
End Sub

' Получить количество команд в истории
Public Function GetHistoryCount() As Integer
    GetHistoryCount = mHistory.Count
End Function

' Сохранить историю в файл
Public Sub SaveHistoryToFile()
    Dim fileNum As Integer
    fileNum = FreeFile
    
    On Error Resume Next
    Open mHistoryFile For Output As #fileNum
    
    If err.Number <> 0 Then
        MsgBox "Не удалось открыть файл для сохранения истории: " & err.description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim i As Integer
    Dim cmd As Object
    
    For i = 1 To mHistory.Count
        Set cmd = mHistory(i)
        Print #fileNum, mSerializer.SerializeCommand(cmd)
    Next i
    
    Close #fileNum
    
    MsgBox "История команд сохранена в файл: " & mHistoryFile, vbInformation
End Sub

' Загрузить историю из файла
Public Function LoadHistoryFromFile(receiver As ExcelReceiver) As Collection
    Dim result As New Collection
    
    If Dir(mHistoryFile) = "" Then
        MsgBox "Файл истории команд не найден: " & mHistoryFile, vbExclamation
        Set LoadHistoryFromFile = result
        Exit Function
    End If
    
    Dim fileNum As Integer
    Dim line As String
    Dim cmd As ICommand
    
    fileNum = FreeFile
    
    On Error Resume Next
    Open mHistoryFile For Input As #fileNum
    
    If err.Number <> 0 Then
        MsgBox "Не удалось открыть файл истории: " & err.description, vbExclamation
        Set LoadHistoryFromFile = result
        Exit Function
    End If
    On Error GoTo 0
    
    While Not EOF(fileNum)
        Line Input #fileNum, line
        If line <> "" Then
            Set cmd = mSerializer.DeserializeCommand(line, receiver)
            If Not cmd Is Nothing Then
                result.Add cmd
            End If
        End If
    Wend
    
    Close #fileNum
    
    Set mHistory = result
    Set LoadHistoryFromFile = result
End Function

' Воспроизвести всю историю
Public Sub ReplayHistory()
    Dim i As Integer
    Dim cmd As ICommand
    
    For i = 1 To mHistory.Count
        Set cmd = mHistory(i)
        cmd.Execute
        ' Можно добавить задержку для визуализации
        ' Application.Wait Now + TimeValue("00:00:01")
    Next i
End Sub

'-------------------------------------------
' Component: CommandLogger
'-------------------------------------------
Option Explicit

Private mLogFile As String
Private mEnableLogging As Boolean

Private Const LogFilePath = "C:\Users\dalis\AppData\Local\Excellent VBA\Debug\Logs"



Private Sub Class_Initialize()
    mLogFile = LogFilePath & "\command_log.txt"
    mEnableLogging = True
End Sub

' Включить/выключить логирование
Public Property Let EnableLogging(value As Boolean)
    mEnableLogging = value
End Property

Public Property Get EnableLogging() As Boolean
    EnableLogging = mEnableLogging
End Property

' Установить путь к файлу лога
Public Property Let LogFile(path As String)
    mLogFile = path
End Property

' Очистить файл лога
Public Sub ClearLog()
    If mEnableLogging Then
        Dim fileNum As Integer
        fileNum = FreeFile
        
        On Error Resume Next
        Open mLogFile For Output As #fileNum
        Close #fileNum
        On Error GoTo 0
    End If
End Sub

' Записать сообщение в лог
Public Sub LogMessage(message As String)
    If Not mEnableLogging Then Exit Sub
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    On Error Resume Next
    Open mLogFile For Append As #fileNum
    
    If err.Number <> 0 Then
        MsgBox "Не удалось открыть файл лога: " & err.description, vbExclamation
        Debug.Print "Не удалось открыть файл лога: " & err.description, vbExclamation
        
        Exit Sub
    End If
    On Error GoTo 0
    Debug.Print Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & message
    Print #fileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & message
    Close #fileNum
End Sub

' Логирование выполнения команды
Public Sub LogCommandExecution(cmd As Object)
    Dim message As String
    
    If TypeOf cmd Is SetValueCommand Then
        message = "Выполнение команды: Установка значения в ячейке"
    ElseIf TypeOf cmd Is SetColorCommand Then
        message = "Выполнение команды: Установка цвета ячейки"
    ElseIf TypeOf cmd Is MacroCommand Then
        message = "Выполнение макрокоманды"
    Else
        message = "Выполнение неизвестной команды"
    End If
    
    LogMessage message
End Sub

' Логирование отмены команды
Public Sub LogCommandUndo(cmd As Object)
    Dim message As String
    
    If TypeOf cmd Is SetValueCommand Then
        message = "Отмена команды: Установка значения в ячейке"
    ElseIf TypeOf cmd Is SetColorCommand Then
        message = "Отмена команды: Установка цвета ячейки"
    ElseIf TypeOf cmd Is MacroCommand Then
        message = "Отмена макрокоманды"
    Else
        message = "Отмена неизвестной команды"
    End If
    
    LogMessage message
End Sub

'-------------------------------------------
' Component: DelayedCommand
'-------------------------------------------
Option Explicit

Implements ICommand

Private mInnerCommand As ICommand
Private mDelaySeconds As Integer

' Конструктор
Public Sub Initialize(command As ICommand, delaySeconds As Integer)
    Set mInnerCommand = command
    mDelaySeconds = delaySeconds
End Sub

' Реализация интерфейса ICommand
Private Sub ICommand_Execute()
    ' Показываем сообщение о задержке
    MsgBox "Команда будет выполнена через " & mDelaySeconds & " секунд", vbInformation
    
    ' Ждем указанное количество секунд
    Application.Wait Now + TimeValue("00:00:" & mDelaySeconds)
    
    ' Выполняем команду
    mInnerCommand.Execute
End Sub

Private Sub ICommand_Undo()
    ' Отмена не имеет задержки
    mInnerCommand.Undo
End Sub

'-------------------------------------------
' Component: Guard
'-------------------------------------------
' Класс Guard
' Предоставляет набор методов для проверки условий и генерации исключений при их нарушении
Option Explicit

' Примечание: Функция GetInstance заменена на GetGuard в модуле LogLevel

' Проверяет, что объект не равен Nothing
Public Sub NotNull(obj As Variant, Optional message As String = "Объект не может быть пустым (Nothing)")
    If obj Is Nothing Then
        err.Raise vbObjectError + 10000, "Guard.NotNull", message
    End If
End Sub

' Проверяет, что строка не пуста
Public Sub NotNullOrEmpty(str As String, Optional message As String = "Строка не может быть пустой")
    If Len(Trim(str)) = 0 Then
        err.Raise vbObjectError + 10001, "Guard.NotNullOrEmpty", message
    End If
End Sub

' Проверяет, что число находится в допустимом диапазоне
Public Sub InRange(value As Variant, minValue As Variant, maxValue As Variant, _
                   Optional message As String = "Значение должно быть в диапазоне от {min} до {max}")
    If value < minValue Or value > maxValue Then
        ' Заменяем плейсхолдеры в сообщении
        message = Replace(message, "{min}", CStr(minValue))
        message = Replace(message, "{max}", CStr(maxValue))
        message = Replace(message, "{value}", CStr(value))
        
        err.Raise vbObjectError + 10002, "Guard.InRange", message
    End If
End Sub

' Проверяет, что условие истинно
Public Sub IsTrue(condition As Boolean, Optional message As String = "Условие должно быть истинным")
    If Not condition Then
        err.Raise vbObjectError + 10003, "Guard.IsTrue", message
    End If
End Sub

' Проверяет, что условие ложно
Public Sub IsFalse(condition As Boolean, Optional message As String = "Условие должно быть ложным")
    If condition Then
        err.Raise vbObjectError + 10004, "Guard.IsFalse", message
    End If
End Sub

' Проверяет, что строка имеет допустимую длину
Public Sub StringLength(str As String, minLength As Long, maxLength As Long, _
                        Optional message As String = "Длина строки должна быть от {min} до {max} символов")
    If Len(str) < minLength Or Len(str) > maxLength Then
        ' Заменяем плейсхолдеры в сообщении
        message = Replace(message, "{min}", CStr(minLength))
        message = Replace(message, "{max}", CStr(maxLength))
        message = Replace(message, "{actual}", CStr(Len(str)))
        
        err.Raise vbObjectError + 10005, "Guard.StringLength", message
    End If
End Sub

' Проверяет, что строка соответствует формату (используя Like оператор)
Public Sub StringFormat(str As String, pattern As String, _
                        Optional message As String = "Строка не соответствует формату")
    If Not str Like pattern Then
        err.Raise vbObjectError + 10006, "Guard.StringFormat", message
    End If
End Sub

' Проверяет, что файл существует
Public Sub FileExists(filePath As String, _
                     Optional message As String = "Файл не существует: {path}")
    If Dir(filePath) = "" Then
        message = Replace(message, "{path}", filePath)
        err.Raise vbObjectError + 10007, "Guard.FileExists", message
    End If
End Sub

' Проверяет, что лист существует в книге
Public Sub WorksheetExists(wb As Workbook, sheetName As String, _
                          Optional message As String = "Лист '{sheet}' не существует в книге")
    Dim ws As Worksheet
    Dim exists As Boolean
    
    exists = False
    
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    If Not ws Is Nothing Then exists = True
    On Error GoTo 0
    
    If Not exists Then
        message = Replace(message, "{sheet}", sheetName)
        err.Raise vbObjectError + 10008, "Guard.WorksheetExists", message
    End If
End Sub

' Проверяет, что ячейка существует на листе
Public Sub CellExists(ws As Worksheet, cellAddress As String, _
                     Optional message As String = "Ячейка '{cell}' не существует на листе '{sheet}'")
    Dim cell As Range
    Dim exists As Boolean
    
    exists = False
    
    On Error Resume Next
    Set cell = ws.Range(cellAddress)
    If Not cell Is Nothing Then exists = True
    On Error GoTo 0
    
    If Not exists Then
        message = Replace(message, "{cell}", cellAddress)
        message = Replace(message, "{sheet}", ws.name)
        err.Raise vbObjectError + 10009, "Guard.CellExists", message
    End If
End Sub


'-------------------------------------------
' Component: ValidationManager
'-------------------------------------------
' Класс ValidationManager
' Централизованный управляющий класс для выполнения валидации
Option Explicit

' Внутренняя структура для хранения данных
Private Type TValidationManager
    Guard As Guard
    ValidationErrors As Collection  ' Коллекция ошибок валидации
End Type

Private this As TValidationManager

' Инициализация
Private Sub Class_Initialize()
    Set this.Guard = GetGuard ' Используем функцию из модуля LogLevel
    Set this.ValidationErrors = New Collection
End Sub

' Примечание: метод GetInstance заменен на GetValidationManager в модуле LogLevel

' Доступ к экземпляру Guard
Public Property Get Guard() As Guard
    Set Guard = this.Guard
End Property

' Очистка ошибок валидации
Public Sub ClearErrors()
    Set this.ValidationErrors = New Collection
End Sub

' Добавление ошибки валидации
Public Sub AddError(errorMessage As String)
    this.ValidationErrors.Add errorMessage
End Sub

' Проверка наличия ошибок
Public Function HasErrors() As Boolean
    HasErrors = (this.ValidationErrors.Count > 0)
End Function

' Получение всех ошибок валидации в виде строки
Public Function GetErrorsAsString() As String
    Dim result As String
    Dim i As Long
    
    result = ""
    
    For i = 1 To this.ValidationErrors.Count
        If Len(result) > 0 Then
            result = result & vbCrLf
        End If
        result = result & this.ValidationErrors(i)
    Next i
    
    GetErrorsAsString = result
End Function

' Получение всех ошибок валидации в виде коллекции
Public Function GetErrors() As Collection
    Set GetErrors = this.ValidationErrors
End Function

' Количество ошибок
Public Function ErrorCount() As Long
    ErrorCount = this.ValidationErrors.Count
End Function

' Проверка имени листа
Public Function ValidateWorksheetName(name As String) As Boolean
    Dim result As Boolean
    result = True
    
    ' Очищаем предыдущие ошибки
    ClearErrors
    
    ' Проверка на пустое имя
    If Len(Trim(name)) = 0 Then
        AddError "Имя листа не может быть пустым"
        result = False
    End If
    
    ' Проверка длины имени (Excel ограничивает 31 символом)
    If Len(name) > 31 Then
        AddError "Имя листа не может быть длиннее 31 символа"
        result = False
    End If
    
    ' Проверка на недопустимые символы
    Dim invalidChars As String
    invalidChars = "\/[]?*:"
    
    Dim i As Integer
    For i = 1 To Len(invalidChars)
        If InStr(name, Mid(invalidChars, i, 1)) > 0 Then
            AddError "Имя листа содержит недопустимый символ: " & Mid(invalidChars, i, 1)
            result = False
        End If
    Next i
    
    ValidateWorksheetName = result
End Function

' Проверка существования листа в книге
Public Function ValidateWorksheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim result As Boolean
    result = True
    
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    
    If ws Is Nothing Then
        AddError "Лист с именем '" & sheetName & "' не существует в книге"
        result = False
    End If
    On Error GoTo 0
    
    ValidateWorksheetExists = result
End Function

' Проверка адреса ячейки
Public Function ValidateCellAddress(address As String) As Boolean
    Dim result As Boolean
    result = True
    
    ' Очищаем предыдущие ошибки
    ClearErrors
    
    ' Проверка на пустой адрес
    If Len(Trim(address)) = 0 Then
        AddError "Адрес ячейки не может быть пустым"
        result = False
    End If
    
    ' Попытка создать объект Range с этим адресом
    On Error Resume Next
    Dim tempSheet As Worksheet
    Set tempSheet = ThisWorkbook.Worksheets(1) ' Используем любой существующий лист для проверки
    
    Dim tempRange As Range
    Set tempRange = tempSheet.Range(address)
    
    If err.Number <> 0 Then
        AddError "Неверный формат адреса ячейки: " & address
        result = False
    End If
    On Error GoTo 0
    
    ValidateCellAddress = result
End Function

' Проверка значения на допустимый тип данных
Public Function ValidateValueType(value As Variant, expectedType As String) As Boolean
    Dim result As Boolean
    result = True
    
    ' Очищаем предыдущие ошибки
    ClearErrors
    
    Select Case LCase(expectedType)
        Case "string"
            If Not IsString(value) Then
                AddError "Значение должно быть строкой"
                result = False
            End If
        
        Case "number", "numeric"
            If Not IsNumeric(value) Then
                AddError "Значение должно быть числом"
                result = False
            End If
        
        Case "date"
            If Not IsDate(value) Then
                AddError "Значение должно быть датой"
                result = False
            End If
        
        Case "boolean"
            If Not TypeName(value) = "Boolean" Then
                AddError "Значение должно быть логическим типом (True/False)"
                result = False
            End If
        
        Case Else
            AddError "Неизвестный тип данных для валидации: " & expectedType
            result = False
    End Select
    
    ValidateValueType = result
End Function

' Вспомогательные функции для определения типа данных
Private Function IsString(value As Variant) As Boolean
    IsString = (TypeName(value) = "String")
End Function

'-------------------------------------------
' Component: ErrorInfo
'-------------------------------------------
' Класс ErrorInfo
' Представляет информацию об ошибке, возникшей в системе
Option Explicit

' Внутренняя структура для хранения данных
Private Type TErrorInfo
    Number As Long         ' Номер ошибки
    description As String  ' Описание ошибки
    source As String       ' Источник ошибки (модуль, класс, процедура)
    LineNumber As Long     ' Номер строки, где произошла ошибка (Erl)
    context As String      ' Контекст, в котором произошла ошибка
    Timestamp As Date      ' Время возникновения ошибки
    commandName As String  ' Имя команды, при выполнении которой произошла ошибка
    severity As Long       ' Серьезность ошибки (1-5, где 5 - критическая)
    IsHandled As Boolean   ' Флаг, указывающий, была ли ошибка обработана
End Type

Private this As TErrorInfo

' Инициализация
Private Sub Class_Initialize()
    this.Timestamp = Now
    this.severity = 3 ' По умолчанию средняя серьезность
    this.IsHandled = False
End Sub

' Свойства для доступа к данным
Public Property Let Number(value As Long)
    this.Number = value
End Property

Public Property Get Number() As Long
    Number = this.Number
End Property

Public Property Let description(value As String)
    this.description = value
End Property

Public Property Get description() As String
    description = this.description
End Property

Public Property Let source(value As String)
    this.source = value
End Property

Public Property Get source() As String
    source = this.source
End Property

Public Property Let LineNumber(value As Long)
    this.LineNumber = value
End Property

Public Property Get LineNumber() As Long
    LineNumber = this.LineNumber
End Property

Public Property Let context(value As String)
    this.context = value
End Property

Public Property Get context() As String
    context = this.context
End Property

Public Property Let Timestamp(value As Date)
    this.Timestamp = value
End Property

Public Property Get Timestamp() As Date
    Timestamp = this.Timestamp
End Property

Public Property Let commandName(value As String)
    this.commandName = value
End Property

Public Property Get commandName() As String
    commandName = this.commandName
End Property

Public Property Let severity(value As Long)
    ' Убедимся, что серьезность находится в допустимом диапазоне
    If value < 1 Then
        this.severity = 1
    ElseIf value > 5 Then
        this.severity = 5
    Else
        this.severity = value
    End If
End Property

Public Property Get severity() As Long
    severity = this.severity
End Property

Public Property Let IsHandled(value As Boolean)
    this.IsHandled = value
End Property

Public Property Get IsHandled() As Boolean
    IsHandled = this.IsHandled
End Property

' Методы для работы с объектом
Public Sub InitFromErr(err As ErrObject, Optional source As String = "", Optional context As String = "", Optional commandName As String = "")
    this.Number = err.Number
    this.description = err.description
    
    ' Если источник не задан, используем источник из объекта ошибки
    If Len(source) > 0 Then
        this.source = source
    Else
        this.source = err.source
    End If
    
    ' Получаем номер строки, если доступно
    this.LineNumber = Erl
    
    this.context = context
    this.commandName = commandName
    this.Timestamp = Now
    
    ' Определяем серьезность ошибки на основе ее номера
    ' VBA системные ошибки имеют отрицательные номера или меньше 1000, считаем их серьезными
    If err.Number < 0 Or (err.Number > 0 And err.Number < 1000) Then
        this.severity = 4 ' Серьезная ошибка
    ' Пользовательские ошибки обычно имеют номера, начиная с vbObjectError
    ElseIf err.Number >= vbObjectError Then
        this.severity = 3 ' Средняя серьезность
    Else
        this.severity = 2 ' Несерьезная ошибка
    End If
    
    this.IsHandled = False
End Sub

' Преобразование в строку для вывода
Public Function ToString() As String
    Dim result As String
    
    ' Формат: [Timestamp] [Severity] [Number] [Source:LineNumber] [CommandName]: Description (Context)
    result = Format(this.Timestamp, "yyyy-mm-dd hh:mm:ss")
    result = result & " [" & SeverityToString() & "]"
    result = result & " [#" & this.Number & "]"
    
    If Len(this.source) > 0 Then
        result = result & " [" & this.source
        If this.LineNumber > 0 Then
            result = result & ":" & this.LineNumber
        End If
        result = result & "]"
    End If
    
    If Len(this.commandName) > 0 Then
        result = result & " [Command: " & this.commandName & "]"
    End If
    
    result = result & ": " & this.description
    
    If Len(this.context) > 0 Then
        result = result & " (" & this.context & ")"
    End If
    
    If this.IsHandled Then
        result = result & " [Обработано]"
    End If
    
    ToString = result
End Function

' Создание копии
Public Function Clone() As errorInfo
    Dim copy As New errorInfo
    
    copy.Number = this.Number
    copy.description = this.description
    copy.source = this.source
    copy.LineNumber = this.LineNumber
    copy.context = this.context
    copy.Timestamp = this.Timestamp
    copy.commandName = this.commandName
    copy.severity = this.severity
    copy.IsHandled = this.IsHandled
    
    Set Clone = copy
End Function

' Преобразование серьезности ошибки в текст
Private Function SeverityToString() As String
    Select Case this.severity
        Case 1: SeverityToString = "MINOR"
        Case 2: SeverityToString = "LOW"
        Case 3: SeverityToString = "MEDIUM"
        Case 4: SeverityToString = "HIGH"
        Case 5: SeverityToString = "CRITICAL"
        Case Else: SeverityToString = "UNKNOWN"
    End Select
End Function

'-------------------------------------------
' Component: IErrorObserver
'-------------------------------------------
' Интерфейс IErrorObserver
' Наблюдатель за событиями ошибок в системе (паттерн Observer)
Option Explicit

' Метод вызывается при возникновении ошибки
Public Sub ErrorOccurred(errorInfo As errorInfo)
End Sub

' Метод вызывается после обработки ошибки
Public Sub ErrorHandled(errorInfo As errorInfo)
End Sub


'-------------------------------------------
' Component: ErrorHandler
'-------------------------------------------
' ErrorHandler.cls
Option Explicit

Private Type TErrorHandler
    Observers As Collection
    logger As logger
    LastError As errorInfo
End Type

Private this As TErrorHandler
Private mInstance As errorHandler

Private Sub Class_Initialize()
    Set this.Observers = New Collection
    Set this.logger = logger.GetInstance
    Set this.LastError = New errorInfo
End Sub

Public Function GetInstance() As errorHandler
    If mInstance Is Nothing Then
        Set mInstance = New errorHandler
    End If
    Set GetInstance = mInstance
End Function

Public Sub AddObserver(observer As IErrorObserver)
    this.Observers.Add observer
End Sub

Public Sub RemoveObserver(observer As IErrorObserver)
    Dim i As Integer
    For i = 1 To this.Observers.Count
        If this.Observers(i) Is observer Then
            this.Observers.Remove i
            Exit Sub
        End If
    Next i
End Sub

Public Sub HandleError(err As ErrObject, Optional context As String = "")
    ' Создаем объект ошибки
    Set this.LastError = New errorInfo
    this.LastError.Number = err.Number
    this.LastError.description = err.description
    this.LastError.source = err.source
    this.LastError.LineNumber = Erl
    this.LastError.context = context
    this.LastError.Timestamp = Now
    
    ' Логируем ошибку
    this.logger.LogError this.LastError.ToString()
    
    ' Уведомляем наблюдателей
    Dim observer As IErrorObserver
    For Each observer In this.Observers
        observer.ErrorOccurred this.LastError
    Next observer
End Sub

Public Property Get LastError() As errorInfo
    Set LastError = this.LastError
End Property

'-------------------------------------------
' Component: LogLevels
'-------------------------------------------
' Модуль LogLevel
' Определяет уровни логирования для системы и предоставляет доступ к экземпляру Logger
Option Explicit

' Переменная для хранения единственного экземпляра Logger
Private mLoggerInstance As logger

' Перечисление уровней логирования в порядке увеличения важности
Public Enum LogLevel
    LogDebug = 0     ' Подробная информация для отладки
    LogInfo = 1      ' Общая информация о работе системы
    LogWarning = 2   ' Предупреждения, не критичные для работы
    LogError = 3     ' Ошибки, которые позволяют продолжить работу
    LogCritical = 4  ' Критические ошибки, требующие остановки
End Enum
' Переменная для хранения единственного экземпляра ValidationManager
Private mValidationManagerInstance As ValidationManager
' Переменная для хранения единственного экземпляра CommandInvoker
Private mCommandInvokerInstance As CommandInvoker
' Переменная для хранения единственного экземпляра ErrorManager
Private mErrorManagerInstance As ErrorManager
' Переменная для хранения единственного экземпляра CommandMediator
Private mCommandMediatorInstance As CommandMediator
' Переменная для хранения единственного экземпляра Guard
Private mGuardInstance As Guard

' Функция для получения названия уровня логирования по его значению
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
        Case LogLevel.LogCritical
            LogLevelToString = "CRITICAL"
        Case Else
            LogLevelToString = "UNKNOWN"
    End Select
End Function

' Функция для получения экземпляра Logger (реализация паттерна Singleton)
Public Function GetLogger() As logger
    If mLoggerInstance Is Nothing Then
        Set mLoggerInstance = New logger
    End If
    Set GetLogger = mLoggerInstance
End Function



' Функция для получения экземпляра Guard (реализация паттерна Singleton)
Public Function GetGuard() As Guard
    If mGuardInstance Is Nothing Then
        Set mGuardInstance = New Guard
    End If
    Set GetGuard = mGuardInstance
End Function


' Функция для получения экземпляра ValidationManager (реализация паттерна Singleton)
Public Function GetValidationManager() As ValidationManager
    If mValidationManagerInstance Is Nothing Then
        Set mValidationManagerInstance = New ValidationManager
    End If
    Set GetValidationManager = mValidationManagerInstance
End Function


' Функция для получения экземпляра CommandInvoker (реализация паттерна Singleton)
Public Function GetCommandInvoker() As CommandInvoker
    If mCommandInvokerInstance Is Nothing Then
        Set mCommandInvokerInstance = New CommandInvoker
    End If
    Set GetCommandInvoker = mCommandInvokerInstance
End Function



' Функция для получения экземпляра ErrorManager (реализация паттерна Singleton)
Public Function GetErrorManager() As ErrorManager
    If mErrorManagerInstance Is Nothing Then
        Set mErrorManagerInstance = New ErrorManager
    End If
    Set GetErrorManager = mErrorManagerInstance
End Function


' Функция для получения экземпляра CommandMediator (реализация паттерна Singleton)
Public Function GetCommandMediator() As CommandMediator
    If mCommandMediatorInstance Is Nothing Then
        Set mCommandMediatorInstance = New CommandMediator
    End If
    Set GetCommandMediator = mCommandMediatorInstance
End Function

' Функция для получения уровня логирования по его названию
Public Function StringToLogLevel(levelName As String) As LogLevel
    Select Case UCase(levelName)
        Case "DEBUG"
            StringToLogLevel = LogLevel.LogDebug
        Case "INFO"
            StringToLogLevel = LogLevel.LogInfo
        Case "WARNING"
            StringToLogLevel = LogLevel.LogWarning
        Case "ERROR"
            StringToLogLevel = LogLevel.LogError
        Case "CRITICAL"
            StringToLogLevel = LogLevel.LogCritical
        Case Else
            ' По умолчанию используем Info
            StringToLogLevel = LogLevel.LogInfo
    End Select
End Function

'-------------------------------------------
' Component: LogEntry
'-------------------------------------------
' Класс LogEntry
' Представляет одну запись в логе системы
Option Explicit

' Внутренняя структура для хранения данных
Private Type TLogEntry
    level As LogLevel     ' Уровень логирования
    message As String     ' Сообщение
    Timestamp As Date     ' Время создания записи
    source As String      ' Источник (модуль, класс, процедура)
    commandName As String ' Связанная команда (если есть)
    UserName As String    ' Имя пользователя (если доступно)
End Type

Private this As TLogEntry

' Инициализация
Private Sub Class_Initialize()
    this.Timestamp = Now
    this.level = LogLevel.LogInfo
    this.source = ""
    this.commandName = ""
    this.UserName = ""
End Sub

' Свойства для доступа к данным
Public Property Let level(value As LogLevel)
    this.level = value
End Property

Public Property Get level() As LogLevel
    level = this.level
End Property

Public Property Let message(value As String)
    this.message = value
End Property

Public Property Get message() As String
    message = this.message
End Property

Public Property Let Timestamp(value As Date)
    this.Timestamp = value
End Property

Public Property Get Timestamp() As Date
    Timestamp = this.Timestamp
End Property

Public Property Let source(value As String)
    this.source = value
End Property

Public Property Get source() As String
    source = this.source
End Property

Public Property Let commandName(value As String)
    this.commandName = value
End Property

Public Property Get commandName() As String
    commandName = this.commandName
End Property

Public Property Let UserName(value As String)
    this.UserName = value
End Property

Public Property Get UserName() As String
    UserName = this.UserName
End Property

' Преобразование записи лога в строку для вывода
Public Function ToString() As String
    Dim result As String
    
    ' Формат: [Timestamp] [Level] [Source] [CommandName] [UserName]: Message
    result = Format(this.Timestamp, "yyyy-mm-dd hh:mm:ss")
    result = result & " [" & LogLevelToString(this.level) & "]"
    
    If Len(this.source) > 0 Then
        result = result & " [" & this.source & "]"
    End If
    
    If Len(this.commandName) > 0 Then
        result = result & " [Command: " & this.commandName & "]"
    End If
    
    If Len(this.UserName) > 0 Then
        result = result & " [User: " & this.UserName & "]"
    End If
    
    result = result & ": " & this.message
    
    ToString = result
End Function

' Создание копии записи лога
Public Function Clone() As logEntry
    Dim copy As New logEntry
    
    copy.level = this.level
    copy.message = this.message
    copy.Timestamp = this.Timestamp
    copy.source = this.source
    copy.commandName = this.commandName
    copy.UserName = this.UserName
    
    Set Clone = copy
End Function

'-------------------------------------------
' Component: ILogObserver
'-------------------------------------------
' Интерфейс ILogObserver
' Наблюдатель за событиями логирования (паттерн Observer)
Option Explicit

' Метод вызывается при добавлении новой записи в лог
Public Sub LogEntryAdded(entry As logEntry)
End Sub

'-------------------------------------------
' Component: Logger
'-------------------------------------------
' Класс Logger
' Централизованная система логирования с поддержкой паттерна Observer
Option Explicit

' Внутренняя структура для хранения данных
Private Type TLogger
    Observers As Collection   ' Коллекция наблюдателей
    LogFile As String         ' Путь к файлу лога
    EnableFileLogging As Boolean ' Включено ли логирование в файл
    EnableConsoleLogging As Boolean ' Включено ли логирование в консоль (Immediate Window)
    MinimumLevel As LogLevel  ' Минимальный уровень сообщений для логирования
    CurrentUser As String     ' Текущий пользователь
    LogHistory As Collection  ' История сообщений (для хранения в памяти)
    MaxHistorySize As Long    ' Максимальный размер истории
End Type

Private this As TLogger

' Инициализация
Private Sub Class_Initialize()
    ' Создаем коллекции
    Set this.Observers = New Collection
    Set this.LogHistory = New Collection
    
    ' Настройки по умолчанию
    this.LogFile = "C:\Logs\CommandPattern.log" ' Путь по умолчанию
    this.EnableFileLogging = True
    this.EnableConsoleLogging = True
    this.MinimumLevel = LogLevel.LogInfo
    this.CurrentUser = Application.UserName ' Имя пользователя Excel
    this.MaxHistorySize = 100 ' Хранить последние 100 сообщений
    
    ' Создаем папку для логов, если она не существует
    EnsureLogDirectoryExists
End Sub

' Примечание: метод GetInstance заменен на GetLogger в модуле LogLevel

' Создание папки для логов
Private Sub EnsureLogDirectoryExists()
    Dim folderPath As String
    folderPath = Left(this.LogFile, InStrRev(this.LogFile, "\") - 1)
    
    On Error Resume Next
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    On Error GoTo 0
End Sub

' Настройка логгера
Public Property Let LogFile(value As String)
    this.LogFile = value
    EnsureLogDirectoryExists
End Property

Public Property Get LogFile() As String
    LogFile = this.LogFile
End Property

Public Property Let EnableFileLogging(value As Boolean)
    this.EnableFileLogging = value
End Property

Public Property Get EnableFileLogging() As Boolean
    EnableFileLogging = this.EnableFileLogging
End Property

Public Property Let EnableConsoleLogging(value As Boolean)
    this.EnableConsoleLogging = value
End Property

Public Property Get EnableConsoleLogging() As Boolean
    EnableConsoleLogging = this.EnableConsoleLogging
End Property

Public Property Let MinimumLevel(value As LogLevel)
    this.MinimumLevel = value
End Property

Public Property Get MinimumLevel() As LogLevel
    MinimumLevel = this.MinimumLevel
End Property

Public Property Let CurrentUser(value As String)
    this.CurrentUser = value
End Property

Public Property Get CurrentUser() As String
    CurrentUser = this.CurrentUser
End Property

Public Property Let MaxHistorySize(value As Long)
    this.MaxHistorySize = value
    TrimHistory ' Применяем ограничение сразу же
End Property

Public Property Get MaxHistorySize() As Long
    MaxHistorySize = this.MaxHistorySize
End Property

' Обрезать историю до максимального размера
Private Sub TrimHistory()
    ' Удаляем старые записи, если их больше максимального размера
    While this.LogHistory.Count > this.MaxHistorySize
        this.LogHistory.Remove 1
    Wend
End Sub

' Методы для паттерна Observer
Public Sub AddObserver(observer As ILogObserver)
    this.Observers.Add observer
End Sub

Public Sub RemoveObserver(observer As ILogObserver)
    Dim i As Long
    For i = 1 To this.Observers.Count
        If this.Observers(i) Is observer Then
            this.Observers.Remove i
            Exit Sub
        End If
    Next i
End Sub

' Основной метод логирования
Public Sub Log(level As LogLevel, message As String, Optional source As String = "", Optional commandName As String = "")
    ' Проверяем, нужно ли логировать сообщение с этим уровнем
    If level < this.MinimumLevel Then
        Exit Sub
    End If
    
    ' Создаем запись лога
    Dim entry As New logEntry
    entry.level = level
    entry.message = message
    entry.Timestamp = Now
    entry.source = source
    entry.commandName = commandName
    entry.UserName = this.CurrentUser
    
    ' Сохраняем в историю
    this.LogHistory.Add entry
    TrimHistory
    
    ' Логируем в файл
    If this.EnableFileLogging Then
        WriteToFile entry
    End If
    
    ' Логируем в консоль
    If this.EnableConsoleLogging Then
        Debug.Print entry.ToString()
    End If
    
    ' Уведомляем наблюдателей
    NotifyObservers entry
End Sub

' Уведомить всех наблюдателей о новой записи
Private Sub NotifyObservers(entry As logEntry)
    Dim observer As ILogObserver
    Dim i As Long
    
    ' Для каждого наблюдателя вызываем метод LogEntryAdded
    For i = 1 To this.Observers.Count
        Set observer = this.Observers(i)
        observer.LogEntryAdded entry.Clone() ' Передаем копию записи
    Next i
End Sub

' Запись в файл
Private Sub WriteToFile(entry As logEntry)
    On Error Resume Next
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open this.LogFile For Append As #fileNum
    
    If err.Number <> 0 Then
        ' Если не удалось открыть файл, выводим ошибку только в консоль
        Debug.Print "Не удалось открыть файл лога: " & err.description
        Exit Sub
    End If
    
    Print #fileNum, entry.ToString()
    Close #fileNum
    
    On Error GoTo 0
End Sub

' Удобные методы для разных уровней логирования
Public Sub LogDebug(message As String, Optional source As String = "", Optional commandName As String = "")
    Log LogLevel.LogDebug, message, source, commandName
End Sub

Public Sub LogInfo(message As String, Optional source As String = "", Optional commandName As String = "")
    Log LogLevel.LogInfo, message, source, commandName
End Sub

Public Sub LogWarning(message As String, Optional source As String = "", Optional commandName As String = "")
    Log LogLevel.LogWarning, message, source, commandName
End Sub

Public Sub LogError(message As String, Optional source As String = "", Optional commandName As String = "")
    Log LogLevel.LogError, message, source, commandName
End Sub

Public Sub LogCritical(message As String, Optional source As String = "", Optional commandName As String = "")
    Log LogLevel.LogCritical, message, source, commandName
End Sub

' Получение истории логов
Public Function GetHistory() As Collection
    Set GetHistory = this.LogHistory
End Function

' Очистка истории логов
Public Sub ClearHistory()
    Set this.LogHistory = New Collection
End Sub

' Очистка файла лога
Public Sub ClearLogFile()
    On Error Resume Next
    
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open this.LogFile For Output As #fileNum
    Close #fileNum
    
    If err.Number = 0 Then
        LogInfo "Файл лога очищен", "Logger"
    Else
        LogError "Не удалось очистить файл лога: " & err.description, "Logger"
    End If
    
    On Error GoTo 0
End Sub

'-------------------------------------------
' Component: CommandMediator
'-------------------------------------------
' Класс CommandMediator
' Центральный класс для координации взаимодействия между компонентами системы
Option Explicit

' Внутренняя структура для хранения данных
Private Type TCommandMediator
    Components As Object         ' Коллекция компонентов (Dictionary)
    logger As Object             ' Объект логгера
    ErrorManager As Object       ' Менеджер обработки ошибок
    CommandInvoker As Object     ' Инвокер команд
    ValidationManager As Object  ' Менеджер валидации
    IsInitialized As Boolean     ' Флаг инициализации
End Type

Private this As TCommandMediator

' Инициализация
Private Sub Class_Initialize()
    ' Создание словаря для хранения компонентов
    Set this.Components = CreateObject("Scripting.Dictionary")
    this.IsInitialized = False
End Sub

' Инициализация медиатора и получение ссылок на основные компоненты системы
Public Sub Initialize()
    If this.IsInitialized Then Exit Sub
    
    ' Получение экземпляров основных компонентов
    Set this.logger = GetLogger()
    Set this.ErrorManager = GetErrorManager()
    Set this.CommandInvoker = GetCommandInvoker()
    Set this.ValidationManager = GetValidationManager()
    
    this.logger.LogInfo "CommandMediator инициализирован", "CommandMediator"
    this.IsInitialized = True
End Sub

' Регистрация компонента в медиаторе
Public Sub RegisterComponent(component As IComponent)
    Dim componentID As String
    componentID = component.GetComponentID()
    
    ' Проверка на уникальность идентификатора компонента
    If this.Components.exists(componentID) Then
        this.logger.LogWarning "Компонент с идентификатором '" & componentID & "' уже зарегистрирован", "CommandMediator"
        Exit Sub
    End If
    
    ' Добавление компонента в коллекцию
    this.Components.Add componentID, component
    
    ' Установка ссылки на медиатор в компоненте
    component.SetMediator Me
    
    this.logger.LogInfo "Компонент '" & componentID & "' зарегистрирован в медиаторе", "CommandMediator"
End Sub

' Удаление компонента из медиатора
Public Sub UnregisterComponent(componentID As String)
    If Not this.Components.exists(componentID) Then
        this.logger.LogWarning "Компонент с идентификатором '" & componentID & "' не найден", "CommandMediator"
        Exit Sub
    End If
    
    this.Components.Remove componentID
    this.logger.LogInfo "Компонент '" & componentID & "' удален из медиатора", "CommandMediator"
End Sub

' Отправка сообщения всем компонентам
Public Function BroadcastMessage(messageType As String, data As Variant) As Long
    Dim component As IComponent
    Dim componentsProcessed As Long
    Dim key As Variant
    
    componentsProcessed = 0
    
    ' Логирование отправки сообщения
    this.logger.LogDebug "Отправка широковещательного сообщения типа '" & messageType & "'", "CommandMediator"
    
    ' Проверка инициализации словаря компонентов
    If this.Components Is Nothing Then
        this.logger.LogError "Словарь компонентов не инициализирован", "CommandMediator"
        BroadcastMessage = 0
        Exit Function
    End If
    
    ' Отправка сообщения всем компонентам
    On Error Resume Next
    For Each key In this.Components.Keys
        Set component = this.Components(key)
        
        If Not component Is Nothing Then
            If component.ProcessMessage(messageType, data) Then
                componentsProcessed = componentsProcessed + 1
            End If
        End If
    Next key
    On Error GoTo 0
    
    this.logger.LogDebug "Сообщение обработано " & componentsProcessed & " компонентами", "CommandMediator"
    BroadcastMessage = componentsProcessed
End Function

' Отправка сообщения конкретному компоненту
Public Function SendMessage(targetComponentID As String, messageType As String, data As Variant) As Boolean
    ' Проверка инициализации словаря компонентов
    If this.Components Is Nothing Then
        this.logger.LogError "Словарь компонентов не инициализирован", "CommandMediator"
        SendMessage = False
        Exit Function
    End If

    On Error Resume Next
    ' Проверка существования компонента
    If Not this.Components.exists(targetComponentID) Then
        this.logger.LogWarning "Компонент с идентификатором '" & targetComponentID & "' не найден", "CommandMediator"
        SendMessage = False
        Exit Function
    End If
    
    Dim component As IComponent
    Set component = this.Components(targetComponentID)
    
    ' Логирование отправки сообщения
    this.logger.LogDebug "Отправка сообщения типа '" & messageType & "' компоненту '" & targetComponentID & "'", "CommandMediator"
    
    ' Отправка сообщения компоненту
    SendMessage = component.ProcessMessage(messageType, data)
    
    If err.Number <> 0 Then
        this.logger.LogError "Ошибка при отправке сообщения: " & err.description, "CommandMediator"
        SendMessage = False
    End If
    On Error GoTo 0
    
    If SendMessage Then
        this.logger.LogDebug "Сообщение успешно обработано компонентом '" & targetComponentID & "'", "CommandMediator"
    Else
        this.logger.LogWarning "Сообщение не обработано компонентом '" & targetComponentID & "'", "CommandMediator"
    End If
End Function

' Уведомление о выполнении команды
Public Sub NotifyCommandExecuted(command As Object, success As Boolean)
    ' Преобразование результата в строку для передачи в сообщении
    Dim resultStr As String
    resultStr = IIf(success, "успешно", "неуспешно")
    
    ' Логирование события
    this.logger.LogInfo "Команда " & TypeName(command) & " выполнена " & resultStr, "CommandMediator"
    
    ' Создание словаря с информацией для передачи компонентам
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    data.Add "Command", command
    data.Add "Success", success
    
    ' Отправка широковещательного сообщения о выполнении команды
    BroadcastMessage "CommandExecuted", data
End Sub

' Метод для выполнения команды через медиатор
Public Function ExecuteCommand(command As Object) As Boolean
    ' Проверка инициализации
    If Not this.IsInitialized Then
        this.logger.LogError "Медиатор не инициализирован", "CommandMediator"
        ExecuteCommand = False
        Exit Function
    End If
    
    ' Проверка наличия компонента CommandManager
    If Not this.Components.exists("CommandManager") Then
        this.logger.LogError "Компонент CommandManager не зарегистрирован", "CommandMediator"
        ExecuteCommand = False
        Exit Function
    End If
    
    ' Отправка сообщения компоненту CommandManager для выполнения команды
    ' Создаем объект Dictionary для передачи параметров
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    data.Add "Command", command
    
    ExecuteCommand = SendMessage("CommandManager", "ExecuteCommand", data)
End Function

' Метод для отмены последней команды через медиатор
Public Function UndoLastCommand() As Boolean
    ' Проверка инициализации
    If Not this.IsInitialized Then
        this.logger.LogError "Медиатор не инициализирован", "CommandMediator"
        UndoLastCommand = False
        Exit Function
    End If
    
    ' Проверка наличия компонента CommandManager
    If Not this.Components.exists("CommandManager") Then
        this.logger.LogError "Компонент CommandManager не зарегистрирован", "CommandMediator"
        UndoLastCommand = False
        Exit Function
    End If
    
    ' Отправка сообщения компоненту CommandManager для отмены команды
    UndoLastCommand = SendMessage("CommandManager", "UndoCommand", Nothing)
End Function


' Уведомление об отмене команды
Public Sub NotifyCommandUndone(command As Object, success As Boolean)
    ' Преобразование результата в строку для передачи в сообщении
    Dim resultStr As String
    resultStr = IIf(success, "успешно", "неуспешно")
    
    ' Логирование события
    this.logger.LogInfo "Команда " & TypeName(command) & " отменена " & resultStr, "CommandMediator"
    
    ' Создание словаря с информацией для передачи компонентам
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    data.Add "Command", command
    data.Add "Success", success
    
    ' Отправка широковещательного сообщения об отмене команды
    BroadcastMessage "CommandUndone", data
End Sub

' Уведомление о возникновении ошибки
Public Sub NotifyErrorOccurred(errorInfo As errorInfo)
    ' Логирование события
    this.logger.LogError "Произошла ошибка: " & errorInfo.description, "CommandMediator"
    
    ' Отправка широковещательного сообщения о возникновении ошибки
    BroadcastMessage "ErrorOccurred", errorInfo
End Sub

' Уведомление об обработке ошибки
Public Sub NotifyErrorHandled(errorInfo As errorInfo)
    ' Логирование события
    this.logger.LogInfo "Ошибка обработана: " & errorInfo.description, "CommandMediator"
    
    ' Отправка широковещательного сообщения об обработке ошибки
    BroadcastMessage "ErrorHandled", errorInfo
End Sub

' Уведомление о создании записи в логе
Public Sub NotifyLogEntryAdded(entry As logEntry)
    ' Отправка широковещательного сообщения о создании записи в логе
    BroadcastMessage "LogEntryAdded", entry
End Sub

' Доступ к основным компонентам системы
Public Property Get logger() As Object
    Set logger = this.logger
End Property

Public Property Get ErrorManager() As Object
    Set ErrorManager = this.ErrorManager
End Property

Public Property Get CommandInvoker() As Object
    Set CommandInvoker = this.CommandInvoker
End Property

Public Property Get ValidationManager() As Object
    Set ValidationManager = this.ValidationManager
End Property

' Проверка инициализации медиатора
Public Property Get IsInitialized() As Boolean
    IsInitialized = this.IsInitialized
End Property

' Получение коллекции компонентов (для тестирования)
Public Property Get Components() As Object
    Set Components = this.Components
End Property

'-------------------------------------------
' Component: Security
'-------------------------------------------
' Security.cls
Option Explicit

Private Type TSecurity
    logger As logger
    UserRoles As Collection ' Коллекция, где ключ - имя пользователя, значение - роль
    OperationRoles As Collection ' Коллекция, где ключ - операция, значение - минимальная роль
    EncryptionKey As String
End Type

Private this As TSecurity
Private mInstance As security

Private Sub Class_Initialize()
    Set this.logger = logger.GetInstance
    Set this.UserRoles = New Collection
    Set this.OperationRoles = New Collection
    
    ' Инициализация по умолчанию
    InitializeDefaultSecurity
End Sub

Public Function GetInstance() As security
    If mInstance Is Nothing Then
        Set mInstance = New security
    End If
    Set GetInstance = mInstance
End Function

Private Sub InitializeDefaultSecurity()
    ' Добавляем несколько стандартных ролей пользователей
    SetUserRole "Admin", "Administrator"
    SetUserRole "User", "RegularUser"
    
    ' Настраиваем доступ к операциям
    SetOperationRole "CreateWorksheet", "RegularUser"
    SetOperationRole "DeleteWorksheet", "Administrator"
    SetOperationRole "SetCellValue", "RegularUser"
    SetOperationRole "SetCellColor", "RegularUser"
    
    ' Генерируем ключ шифрования
    this.EncryptionKey = GenerateEncryptionKey()
End Sub

Public Sub SetUserRole(UserName As String, roleName As String)
    On Error Resume Next
    this.UserRoles.Add roleName, UserName
    If err.Number = 457 Then ' Уже существует
        this.UserRoles.Remove UserName
        this.UserRoles.Add roleName, UserName
    End If
    On Error GoTo 0
    
    this.logger.LogInfo "User '" & UserName & "' assigned role '" & roleName & "'", "Security"
End Sub

Public Sub SetOperationRole(operationName As String, minRoleName As String)
    On Error Resume Next
    this.OperationRoles.Add minRoleName, operationName
    If err.Number = 457 Then ' Уже существует
        this.OperationRoles.Remove operationName
        this.OperationRoles.Add minRoleName, operationName
    End If
    On Error GoTo 0
    
    this.logger.LogInfo "Operation '" & operationName & "' requires role '" & minRoleName & "'", "Security"
End Sub

Public Function CheckAccess(UserName As String, operationName As String) As Boolean
    On Error GoTo errorHandler
    
    ' Получаем роль пользователя
    Dim userRole As String
    userRole = this.UserRoles(UserName)
    
    ' Получаем минимальную роль для операции
    Dim operationMinRole As String
    operationMinRole = this.OperationRoles(operationName)
    
    ' Проверяем доступ (простая реализация - только точное соответствие или Admin)
    If userRole = "Administrator" Or userRole = operationMinRole Then
        CheckAccess = True
        this.logger.LogInfo "Access granted for user '" & UserName & "' to operation '" & operationName & "'", "Security"
    Else
        CheckAccess = False
        this.logger.LogWarning "Access denied for user '" & UserName & "' to operation '" & operationName & "'", "Security"
    End If
    
    Exit Function
    
errorHandler:
    CheckAccess = False
    this.logger.LogError "Error checking access: " & err.description, "Security"
End Function

Private Function GenerateEncryptionKey() As String
    ' Простая реализация для демонстрации
    GenerateEncryptionKey = "SecretKey" & Format(Now, "yyyymmddhhmmss")
End Function

Public Function Encrypt(data As String) As String
    ' Простая реализация шифрования для демонстрации
    Dim result As String
    Dim i As Integer, j As Integer
    Dim key As String
    
    key = this.EncryptionKey
    result = ""
    
    For i = 1 To Len(data)
        j = ((i - 1) Mod Len(key)) + 1
        result = result & Chr(Asc(Mid(data, i, 1)) Xor Asc(Mid(key, j, 1)))
    Next i
    
    Encrypt = result
End Function

Public Function Decrypt(data As String) As String
    ' Для XOR-шифрования операция дешифрования идентична шифрованию
    Decrypt = Encrypt(data)
End Function

'-------------------------------------------
' Component: AbstractCommand
'-------------------------------------------
' Абстрактный класс AbstractCommand
' Базовая реализация ICommand
Option Explicit

Implements ICommand

' Внутренняя структура для хранения данных
Private Type TAbstractCommand
    commandName As String
    ExecutionTimestamp As Date
    UndoTimestamp As Date
    ExecutedSuccessfully As Boolean
    UndoneSuccessfully As Boolean
End Type

Private this As TAbstractCommand

' Инициализация
Private Sub Class_Initialize()
    this.commandName = TypeName(Me)
    this.ExecutedSuccessfully = False
    this.UndoneSuccessfully = False
End Sub

' Реализация ICommand
Private Sub ICommand_Execute()
    On Error GoTo errorHandler
    
    ' Записываем время выполнения
    this.ExecutionTimestamp = Now
    
    ' Вызываем метод, который должен быть переопределен в подклассах
    ExecuteCore
    
    ' Отмечаем успешное выполнение
    this.ExecutedSuccessfully = True
    Exit Sub
    
errorHandler:
    this.ExecutedSuccessfully = False
    ' Пока просто пробрасываем ошибку дальше
    err.Raise err.Number, "AbstractCommand.Execute", err.description
End Sub

Private Sub ICommand_Undo()
    On Error GoTo errorHandler
    
    ' Записываем время отмены
    this.UndoTimestamp = Now
    
    ' Вызываем метод, который должен быть переопределен в подклассах
    UndoCore
    
    ' Отмечаем успешную отмену
    this.UndoneSuccessfully = True
    Exit Sub
    
errorHandler:
    this.UndoneSuccessfully = False
    ' Пока просто пробрасываем ошибку дальше
    err.Raise err.Number, "AbstractCommand.Undo", err.description
End Sub

Private Function ICommand_GetCommandName() As String
    ICommand_GetCommandName = this.commandName
End Function

Private Function ICommand_WasExecutedSuccessfully() As Boolean
    ICommand_WasExecutedSuccessfully = this.ExecutedSuccessfully
End Function

Private Function ICommand_WasUndoneSuccessfully() As Boolean
    ICommand_WasUndoneSuccessfully = this.UndoneSuccessfully
End Function

Private Function ICommand_GetExecutionTimestamp() As Date
    ICommand_GetExecutionTimestamp = this.ExecutionTimestamp
End Function

Private Function ICommand_GetUndoTimestamp() As Date
    ICommand_GetUndoTimestamp = this.UndoTimestamp
End Function

' Методы для переопределения в наследниках
Public Sub ExecuteCore()
    ' Подклассы должны переопределить этот метод
    err.Raise vbObjectError + 1000, "AbstractCommand", "ExecuteCore must be overridden by subclass"
End Sub

Public Sub UndoCore()
    ' Подклассы должны переопределить этот метод
    err.Raise vbObjectError + 1001, "AbstractCommand", "UndoCore must be overridden by subclass"
End Sub

' Свойства для наследников
Public Property Let commandName(value As String)
    this.commandName = value
End Property

Public Property Get commandName() As String
    commandName = this.commandName
End Property

Public Property Get ExecutedSuccessfully() As Boolean
    ExecutedSuccessfully = this.ExecutedSuccessfully
End Property

Public Property Get UndoneSuccessfully() As Boolean
    UndoneSuccessfully = this.UndoneSuccessfully
End Property

'-------------------------------------------
' Component: ApplicationInitializer
'-------------------------------------------
' ApplicationInitializer.cls
Option Explicit

Public Sub InitializeApplication()
    ' Инициализация медиатора и всех компонентов
    Dim mediator As CommandMediator
    Set mediator = CommandMediator.GetInstance
    mediator.InitializeComponents
    
    ' Настройка логгера
    Dim logger As logger
    Set logger = logger.GetInstance
    logger.LogFile = "C:\Users\dalis\AppData\Local\Excellent VBA\Debug\Logs\CommandApp.log"
    logger.MinimumLevel = LogLevel.LogDebug
    
    ' Инициализация системы безопасности
    Dim security As security
    Set security = security.GetInstance
    
    ' Дополнительные настройки...
    
    logger.LogInfo "Application initialized successfully", "ApplicationInitializer"
End Sub

'-------------------------------------------
' Component: AbstractCommandWithValidation
'-------------------------------------------
' Абстрактный класс AbstractCommandWithValidation
' Базовая реализация ICommandWithValidation
Option Explicit

Implements ICommandWithValidation

' Внутренняя структура для хранения данных
Private Type TAbstractCommandWithValidation
    commandName As String
    ExecutionTimestamp As Date
    UndoTimestamp As Date
    ExecutedSuccessfully As Boolean
    UndoneSuccessfully As Boolean
    ValidationErrors As String
    IsValidated As Boolean
    IsValid As Boolean
End Type

Private this As TAbstractCommandWithValidation

' Инициализация
Private Sub Class_Initialize()
    this.commandName = TypeName(Me)
    this.ExecutedSuccessfully = False
    this.UndoneSuccessfully = False
    this.ValidationErrors = ""
    this.IsValidated = False
    this.IsValid = False
End Sub

' Реализация ICommandWithValidation
Private Sub ICommandWithValidation_Execute()
    On Error GoTo errorHandler
    
    ' Проверяем валидность перед выполнением
    If Not ICommandWithValidation_IsValid() Then
        err.Raise vbObjectError + 1002, "AbstractCommandWithValidation", _
                 "Cannot execute invalid command: " & this.ValidationErrors
        Exit Sub
    End If
    
    ' Записываем время выполнения
    this.ExecutionTimestamp = Now
    
    ' Вызываем метод, который должен быть переопределен в подклассах
    ExecuteCore
    
    ' Отмечаем успешное выполнение
    this.ExecutedSuccessfully = True
    Exit Sub
    
errorHandler:
    this.ExecutedSuccessfully = False
    ' Пока просто пробрасываем ошибку дальше
    err.Raise err.Number, "AbstractCommandWithValidation.Execute", err.description
End Sub

Private Sub ICommandWithValidation_Undo()
    On Error GoTo errorHandler
    
    ' Записываем время отмены
    this.UndoTimestamp = Now
    
    ' Вызываем метод, который должен быть переопределен в подклассах
    UndoCore
    
    ' Отмечаем успешную отмену
    this.UndoneSuccessfully = True
    Exit Sub
    
errorHandler:
    this.UndoneSuccessfully = False
    ' Пока просто пробрасываем ошибку дальше
    err.Raise err.Number, "AbstractCommandWithValidation.Undo", err.description
End Sub

Private Function ICommandWithValidation_GetCommandName() As String
    ICommandWithValidation_GetCommandName = this.commandName
End Function

Private Function ICommandWithValidation_WasExecutedSuccessfully() As Boolean
    ICommandWithValidation_WasExecutedSuccessfully = this.ExecutedSuccessfully
End Function

Private Function ICommandWithValidation_WasUndoneSuccessfully() As Boolean
    ICommandWithValidation_WasUndoneSuccessfully = this.UndoneSuccessfully
End Function

Private Function ICommandWithValidation_GetExecutionTimestamp() As Date
    ICommandWithValidation_GetExecutionTimestamp = this.ExecutionTimestamp
End Function

Private Function ICommandWithValidation_GetUndoTimestamp() As Date
    ICommandWithValidation_GetUndoTimestamp = this.UndoTimestamp
End Function

Private Function ICommandWithValidation_Validate() As Boolean
    ' Очищаем результаты предыдущей валидации
    this.ValidationErrors = ""
    this.IsValid = False
    this.IsValidated = False
    
    ' Вызываем метод, который должен быть переопределен в подклассах
    Dim result As Boolean
    result = ValidateCore()
    
    ' Сохраняем результат валидации
    this.IsValid = result
    this.IsValidated = True
    
    ICommandWithValidation_Validate = result
End Function

Private Function ICommandWithValidation_GetValidationErrors() As String
    ' Если валидация еще не выполнялась, выполняем ее сейчас
    If Not this.IsValidated Then
        ICommandWithValidation_Validate
    End If
    
    ICommandWithValidation_GetValidationErrors = this.ValidationErrors
End Function

Private Function ICommandWithValidation_IsValid() As Boolean
    ' Если валидация еще не выполнялась, выполняем ее сейчас
    If Not this.IsValidated Then
        ICommandWithValidation_Validate
    End If
    
    ICommandWithValidation_IsValid = this.IsValid
End Function

' Методы для переопределения в наследниках
Public Sub ExecuteCore()
    ' Подклассы должны переопределить этот метод
    err.Raise vbObjectError + 1000, "AbstractCommandWithValidation", "ExecuteCore must be overridden by subclass"
End Sub

Public Sub UndoCore()
    ' Подклассы должны переопределить этот метод
    err.Raise vbObjectError + 1001, "AbstractCommandWithValidation", "UndoCore must be overridden by subclass"
End Sub

Public Function ValidateCore() As Boolean
    ' Подклассы должны переопределить этот метод
    ' По умолчанию считаем команду валидной
    ValidateCore = True
End Function

' Свойства для наследников
Public Property Let commandName(value As String)
    this.commandName = value
End Property

Public Property Get commandName() As String
    commandName = this.commandName
End Property

Public Property Get ExecutedSuccessfully() As Boolean
    ExecutedSuccessfully = this.ExecutedSuccessfully
End Property

Public Property Get UndoneSuccessfully() As Boolean
    UndoneSuccessfully = this.UndoneSuccessfully
End Property

' Методы для работы с ошибками валидации
Protected Sub AddValidationError(errorMessage As String)
    If Len(this.ValidationErrors) > 0 Then
        this.ValidationErrors = this.ValidationErrors & vbCrLf
    End If
    this.ValidationErrors = this.ValidationErrors & errorMessage
    this.IsValid = False
End Sub

Public Property Get ValidationErrors() As String
    ValidationErrors = this.ValidationErrors
End Property


'-------------------------------------------
' Component: CommandInvoker
'-------------------------------------------
' Класс CommandInvoker
' Управляет выполнением команд и историей отмены
Option Explicit

' Внутренняя структура для хранения данных
Private Type TCommandInvoker
    CommandHistory As Collection ' История выполненных команд
    logger As logger            ' Логгер
    MaxHistorySize As Long      ' Максимальный размер истории
    CurrentCommand As Object    ' Текущая команда (ICommand или ICommandWithValidation)
    LastError As String         ' Последняя ошибка
End Type

Private this As TCommandInvoker

' Инициализация
Private Sub Class_Initialize()
    Set this.CommandHistory = New Collection
    Set this.logger = GetLogger ' Используем функцию из модуля LogLevel
    this.MaxHistorySize = 50 ' По умолчанию храним 50 последних команд
    this.LastError = ""
End Sub

' Примечание: метод GetInstance заменен на GetCommandInvoker в модуле LogLevel

' Установка текущей команды
Public Sub SetCommand(command As Object)
    Set this.CurrentCommand = command
    this.LastError = ""
End Sub

' Получение текущей команды
Public Function GetCommand() As Object
    Set GetCommand = this.CurrentCommand
End Function

' Выполнение текущей команды
Public Function ExecuteCommand() As Boolean
    ' Проверяем, что команда установлена
    If this.CurrentCommand Is Nothing Then
        this.LastError = "Команда не установлена"
        this.logger.LogError this.LastError, "CommandInvoker"
        ExecuteCommand = False
        Exit Function
    End If
    
    On Error GoTo errorHandler
    
    ' Получаем имя команды для логирования
    Dim commandName As String
    If TypeOf this.CurrentCommand Is ICommand Then
        ' Команда реализует интерфейс ICommand
        Dim cmd As ICommand
        Set cmd = this.CurrentCommand
        commandName = cmd.GetCommandName()
        
        ' Логируем начало выполнения
        this.logger.LogInfo "Выполнение команды", commandName
        
        ' Выполняем команду
        cmd.Execute
        
        ' Проверяем успешность выполнения
        If cmd.WasExecutedSuccessfully() Then
            ' Добавляем в историю
            this.CommandHistory.Add this.CurrentCommand
            TrimHistory
            
            ' Логируем успешное выполнение
            this.logger.LogInfo "Команда успешно выполнена", commandName
            ExecuteCommand = True
        Else
            ' Логируем неудачное выполнение
            this.LastError = "Команда не была выполнена успешно"
            this.logger.LogWarning this.LastError, commandName
            ExecuteCommand = False
        End If
        
    ElseIf TypeOf this.CurrentCommand Is ICommandWithValidation Then
        ' Команда реализует интерфейс ICommandWithValidation
        Dim cmdWithValidation As ICommandWithValidation
        Set cmdWithValidation = this.CurrentCommand
        commandName = cmdWithValidation.GetCommandName()
        
        ' Проверяем валидность
        If cmdWithValidation.IsValid() Then
            ' Логируем начало выполнения
            this.logger.LogInfo "Выполнение команды", commandName
            
            ' Выполняем команду
            cmdWithValidation.Execute
            
            ' Проверяем успешность выполнения
            If cmdWithValidation.WasExecutedSuccessfully() Then
                ' Добавляем в историю
                this.CommandHistory.Add this.CurrentCommand
                TrimHistory
                
                ' Логируем успешное выполнение
                this.logger.LogInfo "Команда успешно выполнена", commandName
                ExecuteCommand = True
            Else
                ' Логируем неудачное выполнение
                this.LastError = "Команда не была выполнена успешно"
                this.logger.LogWarning this.LastError, commandName
                ExecuteCommand = False
            End If
        Else
            ' Логируем ошибку валидации
            this.LastError = "Команда не прошла валидацию: " & cmdWithValidation.GetValidationErrors()
            this.logger.LogWarning this.LastError, commandName
            ExecuteCommand = False
        End If
    Else
        ' Неизвестный тип команды
        this.LastError = "Неизвестный тип команды"
        this.logger.LogError this.LastError, "CommandInvoker"
        ExecuteCommand = False
    End If
    
    Exit Function
    
errorHandler:
    this.LastError = "Ошибка при выполнении команды: " & err.description
    this.logger.LogError this.LastError, IIf(commandName <> "", commandName, "CommandInvoker")
    ExecuteCommand = False
End Function

' Отмена последней команды
Public Function UndoLastCommand() As Boolean
    ' Проверяем, есть ли команды в истории
    If this.CommandHistory.Count = 0 Then
        this.LastError = "История команд пуста"
        this.logger.LogWarning this.LastError, "CommandInvoker"
        UndoLastCommand = False
        Exit Function
    End If
    
    On Error GoTo errorHandler
    
    ' Получаем последнюю команду из истории
    Dim lastCommand As Object
    Set lastCommand = this.CommandHistory(this.CommandHistory.Count)
    
    ' Получаем имя команды для логирования
    Dim commandName As String
    If TypeOf lastCommand Is ICommand Then
        Dim cmd As ICommand
        Set cmd = lastCommand
        commandName = cmd.GetCommandName()
        
        ' Логируем начало отмены
        this.logger.LogInfo "Отмена команды", commandName
        
        ' Выполняем отмену
        cmd.Undo
        
        ' Проверяем успешность отмены
        If cmd.WasUndoneSuccessfully() Then
            ' Удаляем из истории
            this.CommandHistory.Remove this.CommandHistory.Count
            
            ' Логируем успешную отмену
            this.logger.LogInfo "Команда успешно отменена", commandName
            UndoLastCommand = True
        Else
            ' Логируем неудачную отмену
            this.LastError = "Команда не была отменена успешно"
            this.logger.LogWarning this.LastError, commandName
            UndoLastCommand = False
        End If
        
    ElseIf TypeOf lastCommand Is ICommandWithValidation Then
        Dim cmdWithValidation As ICommandWithValidation
        Set cmdWithValidation = lastCommand
        commandName = cmdWithValidation.GetCommandName()
        
        ' Логируем начало отмены
        this.logger.LogInfo "Отмена команды", commandName
        
        ' Выполняем отмену
        cmdWithValidation.Undo
        
        ' Проверяем успешность отмены
        If cmdWithValidation.WasUndoneSuccessfully() Then
            ' Удаляем из истории
            this.CommandHistory.Remove this.CommandHistory.Count
            
            ' Логируем успешную отмену
            this.logger.LogInfo "Команда успешно отменена", commandName
            UndoLastCommand = True
        Else
            ' Логируем неудачную отмену
            this.LastError = "Команда не была отменена успешно"
            this.logger.LogWarning this.LastError, commandName
            UndoLastCommand = False
        End If
    Else
        ' Неизвестный тип команды
        this.LastError = "Неизвестный тип команды в истории"
        this.logger.LogError this.LastError, "CommandInvoker"
        UndoLastCommand = False
    End If
    
    Exit Function
    
errorHandler:
    this.LastError = "Ошибка при отмене команды: " & err.description
    this.logger.LogError this.LastError, IIf(commandName <> "", commandName, "CommandInvoker")
    UndoLastCommand = False
End Function

' Обрезать историю до максимального размера
Private Sub TrimHistory()
    ' Удаляем старые команды, если их больше максимального размера
    While this.CommandHistory.Count > this.MaxHistorySize
        this.CommandHistory.Remove 1
    Wend
End Sub

' Очистка истории команд
Public Sub ClearHistory()
    Set this.CommandHistory = New Collection
    this.logger.LogInfo "История команд очищена", "CommandInvoker"
End Sub

' Получение количества команд в истории
Public Function GetHistoryCount() As Long
    GetHistoryCount = this.CommandHistory.Count
End Function

' Получение последней ошибки
Public Property Get LastError() As String
    LastError = this.LastError
End Property

' Настройка максимального размера истории
Public Property Let MaxHistorySize(value As Long)
    this.MaxHistorySize = value
    TrimHistory
End Property

Public Property Get MaxHistorySize() As Long
    MaxHistorySize = this.MaxHistorySize
End Property

' Получение команды из истории по индексу (1-based)
Public Function GetHistoryCommand(index As Long) As Object
    If index < 1 Or index > this.CommandHistory.Count Then
        Set GetHistoryCommand = Nothing
    Else
        Set GetHistoryCommand = this.CommandHistory(index)
    End If
End Function

' Получение всей истории команд
Public Function GetHistory() As Collection
    Set GetHistory = this.CommandHistory
End Function

'-------------------------------------------
' Component: SetCellValueCommand
'-------------------------------------------
' Класс SetCellValueCommand
' Команда для установки значения ячейки с использованием новой архитектуры
Option Explicit

Implements ICommandWithValidation

' Внутренняя структура для хранения данных команды
Private Type TSetCellValueCommand
    sheetName As String          ' Имя листа
    cellAddress As String        ' Адрес ячейки
    newValue As Variant          ' Новое значение
    OldValue As Variant          ' Старое значение для отмены
    ValidationManager As ValidationManager ' Менеджер валидации
    logger As logger             ' Логгер
    commandName As String        ' Имя команды
    ExecutionTimestamp As Date   ' Время выполнения
    UndoTimestamp As Date        ' Время отмены
    ExecutedSuccessfully As Boolean ' Успешно ли выполнена
    UndoneSuccessfully As Boolean ' Успешно ли отменена
    ValidationErrors As String   ' Ошибки валидации
    IsValidated As Boolean       ' Была ли выполнена валидация
    IsValid As Boolean           ' Валидна ли команда
End Type

Private this As TSetCellValueCommand

' Инициализация
Private Sub Class_Initialize()
    Set this.ValidationManager = GetValidationManager ' Используем функцию из модуля LogLevel
    Set this.logger = GetLogger ' Используем функцию из модуля LogLevel
    this.commandName = "SetCellValueCommand"
    this.ExecutedSuccessfully = False
    this.UndoneSuccessfully = False
    this.ValidationErrors = ""
    this.IsValidated = False
    this.IsValid = False
End Sub

' Инициализация команды
Public Sub Initialize(sheetName As String, cellAddress As String, newValue As Variant)
    this.sheetName = sheetName
    this.cellAddress = cellAddress
    this.newValue = newValue
    
    ' Сбрасываем состояние валидации при изменении параметров
    this.IsValidated = False
    this.ValidationErrors = ""
End Sub

' Реализация ICommandWithValidation
Private Sub ICommandWithValidation_Execute()
    On Error GoTo errorHandler
    
    ' Проверяем валидность перед выполнением
    If Not ICommandWithValidation_IsValid() Then
        err.Raise vbObjectError + 1002, "SetCellValueCommand", _
                 "Cannot execute invalid command: " & this.ValidationErrors
        Exit Sub
    End If
    
    ' Записываем время выполнения
    this.ExecutionTimestamp = Now
    
    ' Логируем начало выполнения
    this.logger.LogInfo "Выполнение команды", this.commandName
    
    ' Сохраняем старое значение для возможности отмены
    SaveOldValue
    
    ' Устанавливаем новое значение
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(this.sheetName)
    ws.Range(this.cellAddress).value = this.newValue
    
    ' Отмечаем успешное выполнение
    this.ExecutedSuccessfully = True
    
    ' Логируем успешное выполнение
    this.logger.LogInfo "Установлено значение '" & CStr(this.newValue) & "' в ячейку " & _
                        this.sheetName & "!" & this.cellAddress, this.commandName
    
    Exit Sub
    
errorHandler:
    this.ExecutedSuccessfully = False
    
    ' Логируем ошибку
    this.logger.LogError "Ошибка выполнения: " & err.description, this.commandName
    
    ' Пробрасываем ошибку дальше
    err.Raise err.Number, "SetCellValueCommand.Execute", err.description
End Sub

Private Sub ICommandWithValidation_Undo()
    On Error GoTo errorHandler
    
    ' Записываем время отмены
    this.UndoTimestamp = Now
    
    ' Логируем начало отмены
    this.logger.LogInfo "Отмена команды", this.commandName
    
    ' Восстанавливаем старое значение
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(this.sheetName)
    ws.Range(this.cellAddress).value = this.OldValue
    
    ' Отмечаем успешную отмену
    this.UndoneSuccessfully = True
    
    ' Логируем успешную отмену
    this.logger.LogInfo "Восстановлено значение '" & CStr(this.OldValue) & "' в ячейке " & _
                        this.sheetName & "!" & this.cellAddress, this.commandName
    
    Exit Sub
    
errorHandler:
    this.UndoneSuccessfully = False
    
    ' Логируем ошибку
    this.logger.LogError "Ошибка отмены: " & err.description, this.commandName
    
    ' Пробрасываем ошибку дальше
    err.Raise err.Number, "SetCellValueCommand.Undo", err.description
End Sub

Private Function ICommandWithValidation_Validate() As Boolean
    ' Очищаем результаты предыдущей валидации
    this.ValidationErrors = ""
    this.IsValidated = True
    this.IsValid = True
    
    ' Проверяем имя листа
    If Not this.ValidationManager.ValidateWorksheetName(this.sheetName) Then
        this.ValidationErrors = this.ValidationManager.GetErrorsAsString()
        this.IsValid = False
        ICommandWithValidation_Validate = False
        Exit Function
    End If
    
    ' Проверяем существование листа
    If Not ValidateWorksheetExists() Then
        this.IsValid = False
        ICommandWithValidation_Validate = False
        Exit Function
    End If
    
    ' Проверяем адрес ячейки
    If Not this.ValidationManager.ValidateCellAddress(this.cellAddress) Then
        If Len(this.ValidationErrors) > 0 Then this.ValidationErrors = this.ValidationErrors & vbCrLf
        this.ValidationErrors = this.ValidationErrors & this.ValidationManager.GetErrorsAsString()
        this.IsValid = False
        ICommandWithValidation_Validate = False
        Exit Function
    End If
    
    ' Все проверки пройдены
    ICommandWithValidation_Validate = True
End Function

Private Function ICommandWithValidation_GetValidationErrors() As String
    ' Если валидация еще не выполнялась, выполняем ее сейчас
    If Not this.IsValidated Then
        ICommandWithValidation_Validate
    End If
    
    ICommandWithValidation_GetValidationErrors = this.ValidationErrors
End Function

Private Function ICommandWithValidation_IsValid() As Boolean
    ' Если валидация еще не выполнялась, выполняем ее сейчас
    If Not this.IsValidated Then
        ICommandWithValidation_Validate
    End If
    
    ICommandWithValidation_IsValid = this.IsValid
End Function

Private Function ICommandWithValidation_GetCommandName() As String
    ICommandWithValidation_GetCommandName = this.commandName
End Function

Private Function ICommandWithValidation_WasExecutedSuccessfully() As Boolean
    ICommandWithValidation_WasExecutedSuccessfully = this.ExecutedSuccessfully
End Function

Private Function ICommandWithValidation_WasUndoneSuccessfully() As Boolean
    ICommandWithValidation_WasUndoneSuccessfully = this.UndoneSuccessfully
End Function

Private Function ICommandWithValidation_GetExecutionTimestamp() As Date
    ICommandWithValidation_GetExecutionTimestamp = this.ExecutionTimestamp
End Function

Private Function ICommandWithValidation_GetUndoTimestamp() As Date
    ICommandWithValidation_GetUndoTimestamp = this.UndoTimestamp
End Function

' В классе SetCellValueCommand модифицировать метод ExecuteCore
Public Sub ExecuteCore()
    ' Проверяем существование листа
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(this.sheetName)
    Dim errNumber As Long
    errNumber = err.Number
    On Error GoTo 0
    
    ' Если лист не существует, генерируем ошибку
    If errNumber <> 0 Or ws Is Nothing Then
        err.Raise vbObjectError + 5000, "SetCellValueCommand.ExecuteCore", _
                "Лист '" & this.sheetName & "' не существует в книге"
        Exit Sub
    End If
    
    ' Сохраняем старое значение для возможности отмены
    SaveOldValue
    
    ' Устанавливаем новое значение
    ws.Range(this.cellAddress).value = this.newValue
End Sub


' Вспомогательные методы
Private Sub SaveOldValue()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(this.sheetName)
    this.OldValue = ws.Range(this.cellAddress).value
    
    If err.Number <> 0 Then
        this.logger.LogWarning "Не удалось получить старое значение для отмены: " & err.description, this.commandName
    End If
    
    On Error GoTo 0
End Sub

' В методе ValidateWorksheetExists класса SetCellValueCommand
Private Function ValidateWorksheetExists() As Boolean
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(this.sheetName)
    
    ' Строгая проверка существования листа
    If ws Is Nothing Or err.Number <> 0 Then
        If Len(this.ValidationErrors) > 0 Then this.ValidationErrors = this.ValidationErrors & vbCrLf
        this.ValidationErrors = this.ValidationErrors & "Лист '" & this.sheetName & "' не существует"
        ValidateWorksheetExists = False
    Else
        ValidateWorksheetExists = True
    End If
    
    On Error GoTo 0
End Function

' Геттеры для доступа к внутренним данным
Public Property Get sheetName() As String
    sheetName = this.sheetName
End Property

Public Property Get cellAddress() As String
    cellAddress = this.cellAddress
End Property

Public Property Get newValue() As Variant
    newValue = this.newValue
End Property

Public Property Get OldValue() As Variant
    OldValue = this.OldValue
End Property

Public Property Get commandName() As String
    commandName = this.commandName
End Property

Public Property Let commandName(value As String)
    this.commandName = value
End Property


'-------------------------------------------
' Component: AbstractErrorHandler
'-------------------------------------------
' Класс AbstractErrorHandler
' Базовая реализация обработчика ошибок
Option Explicit

Implements IErrorHandler

' Внутренняя структура для хранения данных
Private Type TAbstractErrorHandler
    NextHandler As IErrorHandler
    logger As Object ' Будет Logger
End Type

Private this As TAbstractErrorHandler

' Инициализация
Private Sub Class_Initialize()
    Set this.logger = GetLogger
End Sub

' Реализация IErrorHandler
Private Sub IErrorHandler_SetNext(handler As IErrorHandler)
    Set this.NextHandler = handler
End Sub

Private Function IErrorHandler_GetNext() As IErrorHandler
    Set IErrorHandler_GetNext = this.NextHandler
End Function

Private Function IErrorHandler_HandleError(errorInfo As errorInfo) As Boolean
    ' Проверяем, можем ли мы обработать ошибку
    If IErrorHandler_CanHandle(errorInfo) Then
        ' Выполняем специфичную для этого обработчика логику
        Dim handled As Boolean
        handled = ProcessError(errorInfo)
        
        ' Если обработка была успешной, помечаем ошибку как обработанную
        If handled Then
            errorInfo.IsHandled = True
            IErrorHandler_HandleError = True
            Exit Function
        End If
    End If
    
    ' Если мы не смогли обработать ошибку, передаем ее следующему обработчику
    If Not this.NextHandler Is Nothing Then
        IErrorHandler_HandleError = this.NextHandler.HandleError(errorInfo)
    Else
        ' Если следующего обработчика нет, ошибка остается необработанной
        IErrorHandler_HandleError = False
    End If
End Function

Private Function IErrorHandler_CanHandle(errorInfo As errorInfo) As Boolean
    ' По умолчанию определяем возможность обработки на основе номера ошибки
    ' Но подклассы должны переопределить этот метод
    IErrorHandler_CanHandle = False
End Function

' Методы для переопределения в подклассах
Public Function ProcessError(errorInfo As errorInfo) As Boolean
    ' Должен быть переопределен в подклассах
    ' Возвращает True, если ошибка обработана успешно
    ProcessError = False
End Function

Public Function CanHandle(errorInfo As errorInfo) As Boolean
    ' Должен быть переопределен в подклассах
    CanHandle = False
End Function

' Свойства и методы для подклассов
Private Property Get logger() As Object
    Set logger = this.logger
End Property

' Установка следующего обработчика в цепочке
Public Sub SetNext(handler As IErrorHandler)
    Set this.NextHandler = handler
End Sub

' Получение следующего обработчика
Public Function GetNext() As IErrorHandler
    Set GetNext = this.NextHandler
End Function

'-------------------------------------------
' Component: IErrorHandler
'-------------------------------------------
' Интерфейс IErrorHandler
' Обработчик ошибок (используется в паттерне Chain of Responsibility)
Option Explicit

' Установка следующего обработчика в цепочке
Public Sub SetNext(handler As IErrorHandler)
End Sub

' Получение следующего обработчика в цепочке
Public Function GetNext() As IErrorHandler
End Function

' Обработка ошибки
' Возвращает True, если ошибка обработана, иначе False
Public Function HandleError(errorInfo As errorInfo) As Boolean
End Function

' Проверка, может ли обработчик обработать данную ошибку
Public Function CanHandle(errorInfo As errorInfo) As Boolean
End Function

'-------------------------------------------
' Component: ErrorManager
'-------------------------------------------
' Класс ErrorManager
' Централизованное управление обработкой ошибок
Option Explicit

' Внутренняя структура для хранения данных
Private Type TErrorManager
    Observers As Collection   ' Коллекция наблюдателей
    handlers As Collection    ' Коллекция обработчиков ошибок
    RootHandler As IErrorHandler ' Корневой обработчик в цепочке
    logger As Object ' Будет Logger
    LastError As errorInfo    ' Последняя обработанная ошибка
    errorHistory As Collection ' История ошибок
    MaxHistorySize As Long    ' Максимальный размер истории
    EnableAutomaticLogging As Boolean ' Автоматически логировать ошибки
    ThrowUnhandledErrors As Boolean ' Вызывать исключение для необработанных ошибок
End Type

Private this As TErrorManager

' Инициализация
Private Sub Class_Initialize()
    Set this.Observers = New Collection
    Set this.handlers = New Collection
    Set this.errorHistory = New Collection
    Set this.logger = GetLogger
    
    this.MaxHistorySize = 50 ' По умолчанию храним 50 последних ошибок
    this.EnableAutomaticLogging = True
    this.ThrowUnhandledErrors = False
End Sub

' Добавление наблюдателя
' В методе AddObserver класса ErrorManager добавить проверку дублирования
Public Sub AddObserver(observer As IErrorObserver)
    Dim alreadyExists As Boolean
    alreadyExists = False
    
    ' Проверяем, не добавлен ли уже такой наблюдатель
    Dim i As Long
    For i = 1 To this.Observers.Count
        If this.Observers(i) Is observer Then
            alreadyExists = True
            Exit For
        End If
    Next i
    
    ' Добавляем наблюдателя, если его нет в коллекции
    If Not alreadyExists Then
        this.Observers.Add observer
    End If
End Sub

' Удаление наблюдателя
Public Sub RemoveObserver(observer As IErrorObserver)
    Dim i As Long
    For i = 1 To this.Observers.Count
        If this.Observers(i) Is observer Then
            this.Observers.Remove i
            Exit Sub
        End If
    Next i
End Sub

' Добавление обработчика ошибок
Public Sub AddHandler(handler As IErrorHandler)
    ' Проверяем, не добавлен ли уже такой обработчик
    Dim alreadyExists As Boolean
    alreadyExists = False
    
    Dim i As Long
    For i = 1 To this.handlers.Count
        If this.handlers(i) Is handler Then
            alreadyExists = True
            Exit For
        End If
    Next i
    
    ' Добавляем обработчик, если его нет в коллекции
    If Not alreadyExists Then
        this.handlers.Add handler
        
        ' Если это первый обработчик, устанавливаем его как корневой
        If this.handlers.Count = 1 Then
            Set this.RootHandler = handler
        Else
            ' Иначе добавляем его в конец цепочки
            Dim lastHandler As IErrorHandler
            Set lastHandler = this.handlers(this.handlers.Count - 1)
            lastHandler.SetNext handler
        End If
    End If
End Sub

' Очистка цепочки обработчиков
Public Sub ClearHandlers()
    Set this.handlers = New Collection
    Set this.RootHandler = Nothing
End Sub

' Настройки
Public Property Let MaxHistorySize(value As Long)
    this.MaxHistorySize = value
    TrimHistory
End Property

Public Property Get MaxHistorySize() As Long
    MaxHistorySize = this.MaxHistorySize
End Property

Public Property Let EnableAutomaticLogging(value As Boolean)
    this.EnableAutomaticLogging = value
End Property

Public Property Get EnableAutomaticLogging() As Boolean
    EnableAutomaticLogging = this.EnableAutomaticLogging
End Property

Public Property Let ThrowUnhandledErrors(value As Boolean)
    this.ThrowUnhandledErrors = value
End Property

Public Property Get ThrowUnhandledErrors() As Boolean
    ThrowUnhandledErrors = this.ThrowUnhandledErrors
End Property

' Основной метод обработки ошибки
Public Function HandleError(errorInfo As errorInfo) As Boolean
    ' Сохраняем ошибку в истории
    this.errorHistory.Add errorInfo.Clone()
    TrimHistory
    
    ' Запоминаем последнюю ошибку
    Set this.LastError = errorInfo
    
    ' Логируем ошибку, если включено автоматическое логирование
    If this.EnableAutomaticLogging Then
        this.logger.LogError errorInfo.ToString()
    End If
    
    ' Уведомляем наблюдателей о возникновении ошибки
    NotifyObserversErrorOccurred errorInfo
    
    ' Если есть корневой обработчик, передаем ему ошибку
    Dim handled As Boolean
    handled = False
    
    If Not this.RootHandler Is Nothing Then
        handled = this.RootHandler.HandleError(errorInfo)
    End If
    
    ' Уведомляем наблюдателей о результате обработки
    If handled Then
        NotifyObserversErrorHandled errorInfo
    End If
    
    ' Если ошибка не обработана и включен флаг выбрасывания исключений
    If Not handled And this.ThrowUnhandledErrors Then
        err.Raise errorInfo.Number, errorInfo.source, errorInfo.description
    End If
    
    HandleError = handled
End Function

' Обработка стандартного объекта ошибки VBA
Public Function HandleVBAError(err As ErrObject, Optional source As String = "", Optional context As String = "", Optional commandName As String = "") As Boolean
    Dim errorInfo As New errorInfo
    errorInfo.InitFromErr err, source, context, commandName
    
    HandleVBAError = HandleError(errorInfo)
End Function

' Обработка ошибки с заданными параметрами
Public Function HandleCustomError(errorNumber As Long, description As String, Optional source As String = "", Optional context As String = "", Optional commandName As String = "", Optional severity As Long = 3) As Boolean
    Dim errorInfo As New errorInfo
    
    errorInfo.Number = errorNumber
    errorInfo.description = description
    errorInfo.source = source
    errorInfo.context = context
    errorInfo.commandName = commandName
    errorInfo.severity = severity
    errorInfo.Timestamp = Now
    
    HandleCustomError = HandleError(errorInfo)
End Function

' Получение последней ошибки
Public Property Get LastError() As errorInfo
    Set LastError = this.LastError
End Property

' Получение истории ошибок
Public Function GetErrorHistory() As Collection
    Set GetErrorHistory = this.errorHistory
End Function

' Очистка истории ошибок
Public Sub ClearErrorHistory()
    Set this.errorHistory = New Collection
End Sub

' Проверка наличия корневого обработчика
Public Function HasHandlers() As Boolean
    HasHandlers = Not (this.RootHandler Is Nothing)
End Function

' Вспомогательные методы
Private Sub TrimHistory()
    ' Удаляем старые записи, если их больше максимального размера
    While this.errorHistory.Count > this.MaxHistorySize
        this.errorHistory.Remove 1
    Wend
End Sub

' Уведомление наблюдателей о возникновении ошибки
Private Sub NotifyObserversErrorOccurred(errorInfo As errorInfo)
    Dim i As Long
    For i = 1 To this.Observers.Count
        Dim observer As IErrorObserver
        Set observer = this.Observers(i)
        observer.ErrorOccurred errorInfo.Clone()
    Next i
End Sub

' Уведомление наблюдателей об обработке ошибки
Private Sub NotifyObserversErrorHandled(errorInfo As errorInfo)
    Dim i As Long
    For i = 1 To this.Observers.Count
        Dim observer As IErrorObserver
        Set observer = this.Observers(i)
        observer.ErrorHandled errorInfo.Clone()
    Next i
End Sub
' Добавить в класс ErrorManager
Public Function GetHandlersCount() As Long
    GetHandlersCount = this.handlers.Count
End Function

Public Function GetHandlerAt(index As Long) As IErrorHandler
    If index < 1 Or index > this.handlers.Count Then
        Set GetHandlerAt = Nothing
    Else
        Set GetHandlerAt = this.handlers(index)
    End If
End Function

' В класс ErrorManager добавить методы
' Получение списка зарегистрированных обработчиков
Public Function GetHandlers() As Collection
    Dim result As New Collection
    Dim i As Long
    
    For i = 1 To this.handlers.Count
        result.Add this.handlers(i)
    Next i
    
    Set GetHandlers = result
End Function

' Восстановление набора обработчиков
Public Sub RestoreHandlers(handlers As Collection)
    ClearHandlers
    
    Dim i As Long
    For i = 1 To handlers.Count
        AddHandler handlers(i)
    Next i
End Sub

'-------------------------------------------
' Component: GeneralErrorHandler
'-------------------------------------------
' Класс GeneralErrorHandler
' Обработчик общих ошибок, которые не обработаны другими обработчиками
Option Explicit

Implements IErrorHandler

' Внутренняя структура для хранения данных
Private Type TGeneralErrorHandler
    NextHandler As IErrorHandler
    logger As Object ' Будет Logger
    ShowMessageBox As Boolean ' Показывать ли сообщение пользователю
    LogToFile As Boolean      ' Логировать ли ошибку в файл
End Type

Private this As TGeneralErrorHandler

' Инициализация
Private Sub Class_Initialize()
    Set this.logger = GetLogger
    this.ShowMessageBox = True
    this.LogToFile = True
End Sub

' Реализация IErrorHandler
Private Sub IErrorHandler_SetNext(handler As IErrorHandler)
    Set this.NextHandler = handler
End Sub

Private Function IErrorHandler_GetNext() As IErrorHandler
    Set IErrorHandler_GetNext = this.NextHandler
End Function

Private Function IErrorHandler_HandleError(errorInfo As errorInfo) As Boolean
    ' Если ошибка уже обработана, просто возвращаем True
    If errorInfo.IsHandled Then
        IErrorHandler_HandleError = True
        Exit Function
    End If
    
    ' Проверяем, можем ли мы обработать эту ошибку
    If IErrorHandler_CanHandle(errorInfo) Then
        ' Логируем ошибку
        If this.LogToFile Then
            this.logger.LogError "Обрабатывается общая ошибка: " & errorInfo.ToString(), "GeneralErrorHandler"
        End If
        
        ' Показываем сообщение пользователю
        If this.ShowMessageBox Then
            MsgBox "Произошла ошибка: " & errorInfo.description, _
                   vbExclamation + vbOKOnly, "Ошибка " & errorInfo.Number
        End If
        
        ' Помечаем ошибку как обработанную
        errorInfo.IsHandled = True
        IErrorHandler_HandleError = True
    Else
        ' Если мы не можем обработать ошибку, передаем ее следующему обработчику
        If Not this.NextHandler Is Nothing Then
            IErrorHandler_HandleError = this.NextHandler.HandleError(errorInfo)
        Else
            IErrorHandler_HandleError = False
        End If
    End If
End Function

Private Function IErrorHandler_CanHandle(errorInfo As errorInfo) As Boolean
    ' Общий обработчик может обработать любую ошибку
    ' Обычно он должен быть последним в цепочке
    IErrorHandler_CanHandle = True
End Function

' Публичные методы и свойства
Public Property Let ShowMessageBox(value As Boolean)
    this.ShowMessageBox = value
End Property

Public Property Get ShowMessageBox() As Boolean
    ShowMessageBox = this.ShowMessageBox
End Property

Public Property Let LogToFile(value As Boolean)
    this.LogToFile = value
End Property

Public Property Get LogToFile() As Boolean
    LogToFile = this.LogToFile
End Property

' Установка следующего обработчика в цепочке
Public Sub SetNext(handler As IErrorHandler)
    Set this.NextHandler = handler
End Sub

' Получение следующего обработчика
Public Function GetNext() As IErrorHandler
    Set GetNext = this.NextHandler
End Function


'-------------------------------------------
' Component: ValidationErrorHandler
'-------------------------------------------
' Класс ValidationErrorHandler
' Обработчик ошибок валидации
Option Explicit

Implements IErrorHandler

' Внутренняя структура для хранения данных
Private Type TValidationErrorHandler
    NextHandler As IErrorHandler
    logger As Object ' Будет Logger
    ShowMessageBox As Boolean ' Показывать ли сообщение пользователю
    ErrorPrefix As Long ' Префикс кодов ошибок валидации
End Type

Private this As TValidationErrorHandler

' Инициализация
Private Sub Class_Initialize()
    Set this.logger = GetLogger
    this.ShowMessageBox = True
    this.ErrorPrefix = 10000 ' Пользовательские ошибки валидации начинаются с 10000 + vbObjectError
End Sub

' Реализация IErrorHandler
Private Sub IErrorHandler_SetNext(handler As IErrorHandler)
    Set this.NextHandler = handler
End Sub

Private Function IErrorHandler_GetNext() As IErrorHandler
    Set IErrorHandler_GetNext = this.NextHandler
End Function

Private Function IErrorHandler_HandleError(errorInfo As errorInfo) As Boolean
    ' Если ошибка уже обработана, просто возвращаем True
    If errorInfo.IsHandled Then
        IErrorHandler_HandleError = True
        Exit Function
    End If
    
    ' Проверяем, можем ли мы обработать эту ошибку
    If IErrorHandler_CanHandle(errorInfo) Then
        ' Логируем ошибку
        this.logger.LogWarning "Ошибка валидации: " & errorInfo.description, "ValidationErrorHandler"
        
        ' Показываем сообщение пользователю
        If this.ShowMessageBox Then
            MsgBox "Ошибка валидации: " & errorInfo.description, _
                   vbExclamation + vbOKOnly, "Ошибка валидации"
        End If
        
        ' Помечаем ошибку как обработанную
        errorInfo.IsHandled = True
        IErrorHandler_HandleError = True
    Else
        ' Если мы не можем обработать ошибку, передаем ее следующему обработчику
        If Not this.NextHandler Is Nothing Then
            IErrorHandler_HandleError = this.NextHandler.HandleError(errorInfo)
        Else
            IErrorHandler_HandleError = False
        End If
    End If
End Function

Private Function IErrorHandler_CanHandle(errorInfo As errorInfo) As Boolean
    ' Определяем, является ли ошибка ошибкой валидации по ее номеру
    ' Ошибки валидации имеют номера в диапазоне vbObjectError + 10000 до vbObjectError + 10999
    Dim errorNumber As Long
    errorNumber = errorInfo.Number
    
    If errorNumber >= vbObjectError + this.ErrorPrefix And _
       errorNumber < vbObjectError + this.ErrorPrefix + 1000 Then
        IErrorHandler_CanHandle = True
    Else
        ' Проверяем по источнику и контексту
        Dim isValidationError As Boolean
        isValidationError = False
        
        ' Если источник содержит "Validation" или контекст содержит "Validation"
        If InStr(1, errorInfo.source, "Validation", vbTextCompare) > 0 Or _
           InStr(1, errorInfo.context, "Validation", vbTextCompare) > 0 Then
            isValidationError = True
        End If
        
        ' Если описание ошибки указывает на валидацию
        If InStr(1, errorInfo.description, "валидаци", vbTextCompare) > 0 Or _
           InStr(1, errorInfo.description, "invalid", vbTextCompare) > 0 Or _
           InStr(1, errorInfo.description, "validation", vbTextCompare) > 0 Then
            isValidationError = True
        End If
        
        IErrorHandler_CanHandle = isValidationError
    End If
End Function

' Публичные методы и свойства
Public Property Let ShowMessageBox(value As Boolean)
    this.ShowMessageBox = value
End Property

Public Property Get ShowMessageBox() As Boolean
    ShowMessageBox = this.ShowMessageBox
End Property

Public Property Let ErrorPrefix(value As Long)
    this.ErrorPrefix = value
End Property

Public Property Get ErrorPrefix() As Long
    ErrorPrefix = this.ErrorPrefix
End Property

' Установка следующего обработчика в цепочке
Public Sub SetNext(handler As IErrorHandler)
    Set this.NextHandler = handler
End Sub

' Получение следующего обработчика
Public Function GetNext() As IErrorHandler
    Set GetNext = this.NextHandler
End Function


'-------------------------------------------
' Component: FileErrorHandler
'-------------------------------------------
' Класс FileErrorHandler
' Обработчик ошибок файловой системы
Option Explicit

Implements IErrorHandler

' Внутренняя структура для хранения данных
Private Type TFileErrorHandler
    NextHandler As IErrorHandler
    logger As Object ' Будет Logger
    ShowMessageBox As Boolean ' Показывать ли сообщение пользователю
    RetryCount As Long ' Количество попыток повторения операции
    AutoCreateDirectories As Boolean ' Автоматически создавать директории при ошибках доступа
End Type

Private this As TFileErrorHandler

' Инициализация
Private Sub Class_Initialize()
    Set this.logger = GetLogger
    this.ShowMessageBox = True
    this.RetryCount = 3 ' По умолчанию пытаемся повторить операцию 3 раза
    this.AutoCreateDirectories = True ' По умолчанию автоматически создаем директории
End Sub

' Реализация IErrorHandler
Private Sub IErrorHandler_SetNext(handler As IErrorHandler)
    Set this.NextHandler = handler
End Sub

Private Function IErrorHandler_GetNext() As IErrorHandler
    Set IErrorHandler_GetNext = this.NextHandler
End Function

Private Function IErrorHandler_HandleError(errorInfo As errorInfo) As Boolean
    ' Если ошибка уже обработана, просто возвращаем True
    If errorInfo.IsHandled Then
        IErrorHandler_HandleError = True
        Exit Function
    End If
    
    ' Проверяем, можем ли мы обработать эту ошибку
    If IErrorHandler_CanHandle(errorInfo) Then
        ' Логируем ошибку
        this.logger.LogWarning "Ошибка файловой системы: " & errorInfo.description, "FileErrorHandler"
        
        ' Определяем тип ошибки и пытаемся ее обработать
        Dim handled As Boolean
        handled = False
        
        Select Case errorInfo.Number
            ' Файл не найден
            Case 53, 75, 76
                handled = HandleFileNotFoundError(errorInfo)
                
            ' Путь не найден
            Case 76
                handled = HandlePathNotFoundError(errorInfo)
                
            ' Нет доступа к файлу
            Case 70, 75
                handled = HandleFileAccessError(errorInfo)
                
            ' Диск не готов
            Case 71
                handled = HandleDriveNotReadyError(errorInfo)
                
            ' Другие ошибки ввода-вывода
            Case Else
                handled = HandleGenericFileError(errorInfo)
        End Select
        
        ' Если мы смогли обработать ошибку, помечаем ее как обработанную
        If handled Then
            errorInfo.IsHandled = True
            IErrorHandler_HandleError = True
        Else
            ' Если не смогли обработать, передаем следующему обработчику
            If Not this.NextHandler Is Nothing Then
                IErrorHandler_HandleError = this.NextHandler.HandleError(errorInfo)
            Else
                IErrorHandler_HandleError = False
            End If
        End If
    Else
        ' Если это не ошибка файловой системы, передаем ее следующему обработчику
        If Not this.NextHandler Is Nothing Then
            IErrorHandler_HandleError = this.NextHandler.HandleError(errorInfo)
        Else
            IErrorHandler_HandleError = False
        End If
    End If
End Function

Private Function IErrorHandler_CanHandle(errorInfo As errorInfo) As Boolean
    ' Проверяем, является ли ошибка файловой ошибкой по ее номеру
    Dim errorNumber As Long
    errorNumber = errorInfo.Number
    
    ' Типичные ошибки файловой системы
    Select Case errorNumber
        ' Файл не найден
        Case 53, 75, 76
            IErrorHandler_CanHandle = True
            
        ' Путь не найден
        Case 76
            IErrorHandler_CanHandle = True
            
        ' Нет доступа к файлу
        Case 70, 75
            IErrorHandler_CanHandle = True
            
        ' Диск не готов
        Case 71
            IErrorHandler_CanHandle = True
            
        ' Другие ошибки ввода-вывода
        Case 55, 57, 58, 59, 61, 62, 63, 67, 68, 74
            IErrorHandler_CanHandle = True
            
        Case Else
            ' Если номер ошибки не соответствует известным ошибкам файловой системы,
            ' проверяем по описанию и контексту
            Dim isFileError As Boolean
            isFileError = False
            
            ' Проверяем по ключевым словам в описании ошибки
            If InStr(1, errorInfo.description, "файл", vbTextCompare) > 0 Or _
               InStr(1, errorInfo.description, "директор", vbTextCompare) > 0 Or _
               InStr(1, errorInfo.description, "каталог", vbTextCompare) > 0 Or _
               InStr(1, errorInfo.description, "папк", vbTextCompare) > 0 Or _
               InStr(1, errorInfo.description, "диск", vbTextCompare) > 0 Or _
               InStr(1, errorInfo.description, "file", vbTextCompare) > 0 Or _
               InStr(1, errorInfo.description, "directory", vbTextCompare) > 0 Or _
               InStr(1, errorInfo.description, "folder", vbTextCompare) > 0 Or _
               InStr(1, errorInfo.description, "drive", vbTextCompare) > 0 Or _
               InStr(1, errorInfo.description, "path", vbTextCompare) > 0 Then
                isFileError = True
            End If
            
            IErrorHandler_CanHandle = isFileError
    End Select
End Function

' Обработчики различных типов ошибок
Private Function HandleFileNotFoundError(errorInfo As errorInfo) As Boolean
    ' Показываем пользователю сообщение о том, что файл не найден
    If this.ShowMessageBox Then
        MsgBox "Файл не найден: " & ExtractFilePathFromDescription(errorInfo.description), _
               vbExclamation + vbOKOnly, "Ошибка файловой системы"
    End If
    
    ' Для этого типа ошибок не предпринимаем автоматических действий
    HandleFileNotFoundError = True
End Function

Private Function HandlePathNotFoundError(errorInfo As errorInfo) As Boolean
    ' Извлекаем путь из описания ошибки
    Dim path As String
    path = ExtractFilePathFromDescription(errorInfo.description)
    
    ' Если включено автоматическое создание директорий, пытаемся создать путь
    If this.AutoCreateDirectories And Len(path) > 0 Then
        ' Проверяем, это путь к файлу или к директории
        If Right(path, 1) <> "\" Then
            ' Это путь к файлу, получаем путь к директории
            path = Left(path, InStrRev(path, "\") - 1)
        End If
        
        ' Пытаемся создать директорию
        On Error Resume Next
        MkDir path
        Dim createError As Long
        createError = err.Number
        On Error GoTo 0
        
        ' Если удалось создать директорию
        If createError = 0 Then
            this.logger.LogInfo "Автоматически создана директория: " & path, "FileErrorHandler"
            HandlePathNotFoundError = True
            Exit Function
        End If
    End If
    
    ' Если не удалось автоматически создать директорию или это отключено
    If this.ShowMessageBox Then
        MsgBox "Путь не найден: " & path, vbExclamation + vbOKOnly, "Ошибка файловой системы"
    End If
    
    HandlePathNotFoundError = (Not this.ShowMessageBox) ' Считаем ошибку обработанной, если не показываем сообщение
End Function

Private Function HandleFileAccessError(errorInfo As errorInfo) As Boolean
    ' Показываем пользователю сообщение о проблеме с доступом к файлу
    If this.ShowMessageBox Then
        MsgBox "Нет доступа к файлу: " & ExtractFilePathFromDescription(errorInfo.description) & vbCrLf & _
               "Возможно, файл уже открыт в другой программе или у вас нет прав доступа.", _
               vbExclamation + vbOKOnly, "Ошибка доступа к файлу"
    End If
    
    HandleFileAccessError = True
End Function

Private Function HandleDriveNotReadyError(errorInfo As errorInfo) As Boolean
    ' Показываем пользователю сообщение о том, что диск не готов
    If this.ShowMessageBox Then
        MsgBox "Диск не готов. Убедитесь, что диск вставлен и работает корректно.", _
               vbExclamation + vbOKOnly, "Ошибка диска"
    End If
    
    HandleDriveNotReadyError = True
End Function

Private Function HandleGenericFileError(errorInfo As errorInfo) As Boolean
    ' Обработка других ошибок файловой системы
    If this.ShowMessageBox Then
        MsgBox "Произошла ошибка при работе с файлом: " & errorInfo.description, _
               vbExclamation + vbOKOnly, "Ошибка файловой системы"
    End If
    
    HandleGenericFileError = True
End Function

' Вспомогательные методы
Private Function ExtractFilePathFromDescription(description As String) As String
    ' Пытаемся извлечь путь к файлу из описания ошибки
    ' Обычно путь заключен в кавычки или идет после определенных фраз
    Dim path As String
    path = ""
    
    ' Проверяем наличие пути в кавычках
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(1, description, """")
    If startPos > 0 Then
        endPos = InStr(startPos + 1, description, """")
        If endPos > startPos Then
            path = Mid(description, startPos + 1, endPos - startPos - 1)
            ExtractFilePathFromDescription = path
            Exit Function
        End If
    End If
    
    ' Проверяем наличие пути после двоеточия
    startPos = InStr(1, description, ":")
    If startPos > 0 Then
        ' Предполагаем, что после двоеточия идет путь
        path = Trim(Mid(description, startPos + 1))
        ExtractFilePathFromDescription = path
        Exit Function
    End If
    
    ' Если не удалось извлечь путь, возвращаем пустую строку
    ExtractFilePathFromDescription = ""
End Function

' Публичные методы и свойства
Public Property Let ShowMessageBox(value As Boolean)
    this.ShowMessageBox = value
End Property

Public Property Get ShowMessageBox() As Boolean
    ShowMessageBox = this.ShowMessageBox
End Property

Public Property Let RetryCount(value As Long)
    this.RetryCount = value
End Property

Public Property Get RetryCount() As Long
    RetryCount = this.RetryCount
End Property

Public Property Let AutoCreateDirectories(value As Boolean)
    this.AutoCreateDirectories = value
End Property

Public Property Get AutoCreateDirectories() As Boolean
    AutoCreateDirectories = this.AutoCreateDirectories
End Property

' Установка следующего обработчика в цепочке
Public Sub SetNext(handler As IErrorHandler)
    Set this.NextHandler = handler
End Sub

' Получение следующего обработчика
Public Function GetNext() As IErrorHandler
    Set GetNext = this.NextHandler
End Function


'-------------------------------------------
' Component: LoggingErrorObserver
'-------------------------------------------
' Класс LoggingErrorObserver
' Наблюдатель за ошибками, который логирует их в систему логирования
Option Explicit

Implements IErrorObserver

' Внутренняя структура для хранения данных
Private Type TLoggingErrorObserver
    logger As Object ' Будет Logger
    LogErrorOccurrence As Boolean ' Логировать возникновение ошибки
    LogErrorHandling As Boolean ' Логировать обработку ошибки
End Type

Private this As TLoggingErrorObserver

' Инициализация
Private Sub Class_Initialize()
    Set this.logger = GetLogger
    this.LogErrorOccurrence = True
    this.LogErrorHandling = True
End Sub

' Реализация IErrorObserver
Private Sub IErrorObserver_ErrorOccurred(errorInfo As errorInfo)
    If this.LogErrorOccurrence Then
        ' Определяем уровень логирования на основе серьезности ошибки
        Select Case errorInfo.severity
            Case 1
                this.logger.LogDebug "Произошла незначительная ошибка: " & errorInfo.ToString(), "ErrorObserver"
            Case 2
                this.logger.LogInfo "Произошла ошибка низкого уровня: " & errorInfo.ToString(), "ErrorObserver"
            Case 3
                this.logger.LogWarning "Произошла ошибка среднего уровня: " & errorInfo.ToString(), "ErrorObserver"
            Case 4
                this.logger.LogError "Произошла серьезная ошибка: " & errorInfo.ToString(), "ErrorObserver"
            Case 5
                this.logger.LogCritical "Произошла критическая ошибка: " & errorInfo.ToString(), "ErrorObserver"
            Case Else
                this.logger.LogWarning "Произошла ошибка: " & errorInfo.ToString(), "ErrorObserver"
        End Select
    End If
End Sub

Private Sub IErrorObserver_ErrorHandled(errorInfo As errorInfo)
    If this.LogErrorHandling Then
        this.logger.LogInfo "Ошибка обработана: " & errorInfo.ToString(), "ErrorObserver"
    End If
End Sub

' Публичные методы и свойства
Public Property Let LogErrorOccurrence(value As Boolean)
    this.LogErrorOccurrence = value
End Property

Public Property Get LogErrorOccurrence() As Boolean
    LogErrorOccurrence = this.LogErrorOccurrence
End Property

Public Property Let LogErrorHandling(value As Boolean)
    this.LogErrorHandling = value
End Property

Public Property Get LogErrorHandling() As Boolean
    LogErrorHandling = this.LogErrorHandling
End Property

'-------------------------------------------
' Component: Error
'-------------------------------------------
' Класс LoggingErrorObserver
' Наблюдатель за ошибками, который логирует их в систему логирования
Option Explicit

Implements IErrorObserver

' Внутренняя структура для хранения данных
Private Type TLoggingErrorObserver
    logger As Object ' Будет Logger
    LogErrorOccurrence As Boolean ' Логировать возникновение ошибки
    LogErrorHandling As Boolean ' Логировать обработку ошибки
End Type

Private this As TLoggingErrorObserver

' Инициализация
Private Sub Class_Initialize()
    Set this.logger = GetLogger
    this.LogErrorOccurrence = True
    this.LogErrorHandling = True
End Sub

' Реализация IErrorObserver
Private Sub IErrorObserver_ErrorOccurred(errorInfo As errorInfo)
    If this.LogErrorOccurrence Then
        ' Определяем уровень логирования на основе серьезности ошибки
        Select Case errorInfo.severity
            Case 1
                this.logger.LogDebug "Произошла незначительная ошибка: " & errorInfo.ToString(), "ErrorObserver"
            Case 2
                this.logger.LogInfo "Произошла ошибка низкого уровня: " & errorInfo.ToString(), "ErrorObserver"
            Case 3
                this.logger.LogWarning "Произошла ошибка среднего уровня: " & errorInfo.ToString(), "ErrorObserver"
            Case 4
                this.logger.LogError "Произошла серьезная ошибка: " & errorInfo.ToString(), "ErrorObserver"
            Case 5
                this.logger.LogCritical "Произошла критическая ошибка: " & errorInfo.ToString(), "ErrorObserver"
            Case Else
                this.logger.LogWarning "Произошла ошибка: " & errorInfo.ToString(), "ErrorObserver"
        End Select
    End If
End Sub

Private Sub IErrorObserver_ErrorHandled(errorInfo As errorInfo)
    If this.LogErrorHandling Then
        this.logger.LogInfo "Ошибка обработана: " & errorInfo.ToString(), "ErrorObserver"
    End If
End Sub

' Публичные методы и свойства
Public Property Let LogErrorOccurrence(value As Boolean)
    this.LogErrorOccurrence = value
End Property

Public Property Get LogErrorOccurrence() As Boolean
    LogErrorOccurrence = this.LogErrorOccurrence
End Property

Public Property Let LogErrorHandling(value As Boolean)
    this.LogErrorHandling = value
End Property

Public Property Get LogErrorHandling() As Boolean
    LogErrorHandling = this.LogErrorHandling
End Property

'-------------------------------------------
' Component: ErrorSystemInitializer
'-------------------------------------------
' Модуль ErrorSystemInitializer
' Инициализация системы обработки ошибок
Option Explicit

' Инициализация системы обработки ошибок
Public Sub InitializeErrorSystem()
    ' Получаем экземпляр менеджера ошибок
    Dim ErrorManager As ErrorManager
    Set ErrorManager = GetErrorManager
    
    ' Создаем наблюдателя для логирования ошибок
    Dim loggingObserver As New LoggingErrorObserver
    
    ' Добавляем наблюдателя в менеджер ошибок
    ErrorManager.AddObserver loggingObserver
    
    ' Создаем обработчики ошибок
    
    ' 1. Обработчик ошибок валидации (первый в цепочке)
    Dim validationHandler As New ValidationErrorHandler
    validationHandler.ShowMessageBox = True
    
    ' 2. Обработчик ошибок файловой системы
    Dim fileHandler As New FileErrorHandler
    fileHandler.ShowMessageBox = True
    fileHandler.AutoCreateDirectories = True
    
    ' 3. Общий обработчик ошибок (последний в цепочке)
    Dim generalHandler As New GeneralErrorHandler
    generalHandler.ShowMessageBox = True
    generalHandler.LogToFile = True
    
    ' Строим цепочку обработчиков
    ErrorManager.ClearHandlers
    ErrorManager.AddHandler validationHandler
    ErrorManager.AddHandler fileHandler
    ErrorManager.AddHandler generalHandler
    
    ' Настраиваем менеджер ошибок
    ErrorManager.EnableAutomaticLogging = True
    ErrorManager.ThrowUnhandledErrors = False
    ErrorManager.MaxHistorySize = 50
    
    ' Логируем успешную инициализацию
    Dim logger As Object
    Set logger = GetLogger
    logger.LogInfo "Система обработки ошибок успешно инициализирована", "ErrorSystemInitializer"
End Sub

' Тестирование системы обработки ошибок
Public Sub TestErrorSystem()
    ' Получаем экземпляр менеджера ошибок
    Dim ErrorManager As ErrorManager
    Set ErrorManager = GetErrorManager
    
    ' Получаем логгер
    Dim logger As Object
    Set logger = GetLogger
    
    logger.LogInfo "Начало тестирования системы обработки ошибок", "TestErrorSystem"
    
    ' Тестируем обработку ошибки валидации
    logger.LogInfo "Тест обработки ошибки валидации:", "TestErrorSystem"
    ErrorManager.HandleCustomError vbObjectError + 10001, "Недопустимое имя листа: содержит запрещенные символы", _
                                  "ValidationManager", "Валидация имени листа", "", 2
    
    ' Тестируем обработку ошибки файловой системы
    logger.LogInfo "Тест обработки ошибки файловой системы:", "TestErrorSystem"
    ErrorManager.HandleCustomError 53, "Файл не найден: ""C:\Test\missing_file.txt""", _
                                 "FileManager", "Попытка открыть файл", "", 3
    
    ' Тестируем обработку общей ошибки
    logger.LogInfo "Тест обработки общей ошибки:", "TestErrorSystem"
    ErrorManager.HandleCustomError 5, "Неправильный вызов процедуры", _
                                 "CommandManager", "Выполнение команды", "SetValueCommand", 4
    
    ' Выводим сводку
    logger.LogInfo "Тестирование системы обработки ошибок завершено", "TestErrorSystem"
    logger.LogInfo "Количество ошибок в истории: " & ErrorManager.GetErrorHistory().Count, "TestErrorSystem"
End Sub


'-------------------------------------------
' Component: ErrorTestModule
'-------------------------------------------
' Модуль ErrorTestModule
' Тестирование системы обработки ошибок в различных сценариях
Option Explicit

' Главная тестовая процедура
Public Sub TestErrorHandlingSystem()
    ' Инициализация логгера
    InitializeLogger
    
    ' Получаем ссылку на логгер для использования в тестах
    Dim logger As Object
    Set logger = GetLogger
    
    ' Инициализация системы обработки ошибок
    InitializeErrorSystem
    
    ' Логируем начало тестирования
    logger.LogInfo "Начало комплексного тестирования системы обработки ошибок", "ErrorTestModule"
    
    ' Выполняем различные тесты
    TestBasicErrorHandling
    TestValidationErrorHandling
    TestFileErrorHandling
    TestCommandErrorHandling
    TestUncaughtErrorHandling
    
    ' Логируем завершение тестирования
    logger.LogInfo "Комплексное тестирование системы обработки ошибок успешно завершено", "ErrorTestModule"
    
    ' Выводим результаты в MsgBox
    MsgBox "Тестирование системы обработки ошибок успешно завершено! " & _
           "Проверьте журнал для получения подробной информации.", _
           vbInformation, "Тестирование системы обработки ошибок"
End Sub

' Инициализация логгера
Private Sub InitializeLogger()
    Dim logger As Object
    Set logger = GetLogger
    
    ' Настройка файла лога в папке документов пользователя
    Dim logPath As String
    logPath = Environ("USERPROFILE") & "\Documents\CommandPatternLogs\"
    
    ' Создаем папку, если она не существует
    On Error Resume Next
    If Dir(logPath, vbDirectory) = "" Then
        MkDir logPath
    End If
    On Error GoTo 0
    
    ' Устанавливаем файл лога
    logger.LogFile = logPath & "ErrorHandling_" & Format(Now, "yyyymmdd_hhmmss") & ".log"
    
    ' Включаем все уровни логирования для тестирования
    logger.MinimumLevel = LogLevel.LogDebug
    
    ' Очищаем историю логов
    logger.ClearHistory
    
    ' Логируем информацию о начале тестирования
    logger.LogInfo "Логгер инициализирован. Файл лога: " & logger.LogFile, "ErrorTestModule"
End Sub

' Тестирование базовой обработки ошибок
Private Sub TestBasicErrorHandling()
    Dim logger As Object
    Set logger = GetLogger
    
    Dim ErrorManager As ErrorManager
    Set ErrorManager = GetErrorManager
    
    logger.LogInfo "Начало тестирования базовой обработки ошибок", "TestBasicErrorHandling"
    
    ' Создаем объект ошибки
    Dim errorInfo As New errorInfo
    errorInfo.Number = 1001
    errorInfo.description = "Тестовая ошибка для проверки обработки"
    errorInfo.source = "TestBasicErrorHandling"
    errorInfo.context = "Тестирование базовой функциональности"
    errorInfo.severity = 2 ' Низкая серьезность
    
    ' Обрабатываем ошибку
    Dim handled As Boolean
    handled = ErrorManager.HandleError(errorInfo)
    
    ' Проверяем результат
    If handled Then
        logger.LogInfo "Ошибка успешно обработана", "TestBasicErrorHandling"
    Else
        logger.LogWarning "Ошибка не была обработана", "TestBasicErrorHandling"
    End If
    
    ' Проверяем историю ошибок
    Dim errorHistory As Collection
    Set errorHistory = ErrorManager.GetErrorHistory()
    
    logger.LogInfo "Количество ошибок в истории: " & errorHistory.Count, "TestBasicErrorHandling"
    logger.LogInfo "Последняя ошибка: " & ErrorManager.LastError.ToString(), "TestBasicErrorHandling"
    
    logger.LogInfo "Тестирование базовой обработки ошибок завершено", "TestBasicErrorHandling"
End Sub

' Тестирование обработки ошибок валидации
Private Sub TestValidationErrorHandling()
    Dim logger As Object
    Set logger = GetLogger
    
    Dim ErrorManager As ErrorManager
    Set ErrorManager = GetErrorManager
    
    logger.LogInfo "Начало тестирования обработки ошибок валидации", "TestValidationErrorHandling"
    
    ' Создаем объект ошибки валидации
    Dim errorNumber As Long
    errorNumber = vbObjectError + 10003 ' Ошибка валидации
    
    ' Обрабатываем ошибку валидации через HandleCustomError
    Dim handled As Boolean
    handled = ErrorManager.HandleCustomError(errorNumber, "Недопустимое имя листа: содержит запрещенные символы", _
                                          "ValidationManager", "Валидация имени листа", "", 2)
    
    ' Проверяем результат
    If handled Then
        logger.LogInfo "Ошибка валидации успешно обработана", "TestValidationErrorHandling"
    Else
        logger.LogWarning "Ошибка валидации не была обработана", "TestValidationErrorHandling"
    End If
    
    logger.LogInfo "Тестирование обработки ошибок валидации завершено", "TestValidationErrorHandling"
End Sub

' Тестирование обработки ошибок файловой системы
Private Sub TestFileErrorHandling()
    Dim logger As Object
    Set logger = GetLogger
    
    Dim ErrorManager As ErrorManager
    Set ErrorManager = GetErrorManager
    
    logger.LogInfo "Начало тестирования обработки ошибок файловой системы", "TestFileErrorHandling"
    
    ' Тест обработки ошибки "Файл не найден"
    On Error Resume Next
    Open "C:\несуществующий_файл.txt" For Input As #1
    Dim fileError As Long
    fileError = err.Number
    Dim fileErrorDesc As String
    fileErrorDesc = err.description
    On Error GoTo 0
    
    ' Обрабатываем ошибку файловой системы
    Dim handled As Boolean
    handled = ErrorManager.HandleCustomError(fileError, fileErrorDesc, _
                                          "TestFileErrorHandling", "Попытка открыть файл", "", 3)
    
    ' Проверяем результат
    If handled Then
        logger.LogInfo "Ошибка файловой системы успешно обработана", "TestFileErrorHandling"
    Else
        logger.LogWarning "Ошибка файловой системы не была обработана", "TestFileErrorHandling"
    End If
    
    logger.LogInfo "Тестирование обработки ошибок файловой системы завершено", "TestFileErrorHandling"
End Sub

' Тестирование обработки ошибок команд
Private Sub TestCommandErrorHandling()
    Dim logger As Object
    Set logger = GetLogger
    
    Dim ErrorManager As ErrorManager
    Set ErrorManager = GetErrorManager
    
    logger.LogInfo "Начало тестирования обработки ошибок команд", "TestCommandErrorHandling"
    
    ' Создаем команду, которая вызовет ошибку
    Dim cmd As New SetCellValueCommand
    cmd.Initialize "НесуществующийЛист", "A1", "Тестовое значение"
    
    ' Получаем инвокер
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker
    
    ' Устанавливаем команду
    invoker.SetCommand cmd
    
    ' Пытаемся выполнить команду, которая должна вызвать ошибку
    On Error Resume Next
    invoker.ExecuteCommand
    
    ' Получаем информацию об ошибке
    Dim errNumber As Long
    errNumber = err.Number
    Dim errDescription As String
    errDescription = err.description
    On Error GoTo 0
    
    ' Если произошла ошибка, обрабатываем ее
    If errNumber <> 0 Then
        ' Обрабатываем ошибку
        Dim handled As Boolean
        handled = ErrorManager.HandleCustomError(errNumber, errDescription, _
                                              "TestCommandErrorHandling", "Выполнение команды", "SetCellValueCommand", 3)
        
        ' Проверяем результат
        If handled Then
            logger.LogInfo "Ошибка команды успешно обработана", "TestCommandErrorHandling"
        Else
            logger.LogWarning "Ошибка команды не была обработана", "TestCommandErrorHandling"
        End If
    Else
        logger.LogWarning "Ошибка не произошла, хотя ожидалась", "TestCommandErrorHandling"
    End If
    
    logger.LogInfo "Тестирование обработки ошибок команд завершено", "TestCommandErrorHandling"
End Sub

' Тестирование обработки необработанных ошибок
Private Sub TestUncaughtErrorHandling()
    Dim logger As Object
    Set logger = GetLogger
    
    Dim ErrorManager As ErrorManager
    Set ErrorManager = GetErrorManager
    
    logger.LogInfo "Начало тестирования обработки необработанных ошибок", "TestUncaughtErrorHandling"
    
    ' Временно включаем выбрасывание необработанных ошибок
    Dim prevThrowUnhandledErrors As Boolean
    prevThrowUnhandledErrors = ErrorManager.ThrowUnhandledErrors
    ErrorManager.ThrowUnhandledErrors = True
    
    ' Создаем нестандартную ошибку, которая не будет обработана ни одним обработчиком
    ' (например, с очень большим номером)
    Dim errorNumber As Long
    errorNumber = vbObjectError + 99999
    
    ' Пытаемся обработать ошибку, которая не должна быть обработана
    On Error Resume Next
    ErrorManager.HandleCustomError errorNumber, "Нестандартная ошибка для проверки", _
                                "TestUncaughtErrorHandling", "Проверка необработанных ошибок", "", 4
    
    ' Проверяем, была ли выброшена ошибка
    Dim uncaughtError As Boolean
    uncaughtError = (err.Number <> 0)
    
    ' Восстанавливаем предыдущее значение
    ErrorManager.ThrowUnhandledErrors = prevThrowUnhandledErrors
    On Error GoTo 0
    
    ' Логируем результат
    If uncaughtError Then
        logger.LogInfo "Тест пройден: необработанная ошибка была выброшена", "TestUncaughtErrorHandling"
    Else
        logger.LogWarning "Тест не пройден: необработанная ошибка не была выброшена", "TestUncaughtErrorHandling"
    End If
    
    logger.LogInfo "Тестирование обработки необработанных ошибок завершено", "TestUncaughtErrorHandling"
End Sub


'-------------------------------------------
' Component: IComponent
'-------------------------------------------
' Интерфейс IComponent
' Базовый интерфейс для компонентов, взаимодействующих через Mediator
Option Explicit

' Установка ссылки на медиатор
Public Sub SetMediator(mediator As CommandMediator)
End Sub

' Получение идентификатора компонента
Public Function GetComponentID() As String
End Function

' Обработка сообщения от медиатора
Public Function ProcessMessage(messageType As String, data As Variant) As Boolean
End Function

'-------------------------------------------
' Component: ErrorHandlerComponent
'-------------------------------------------
' Класс ErrorHandlerComponent
' Компонент для обработки ошибок через медиатор
Option Explicit

Implements IComponent

' Внутренняя структура для хранения данных
Private Type TErrorHandlerComponent
    mediator As CommandMediator
    ErrorManager As Object
    logger As Object
End Type

Private this As TErrorHandlerComponent

' Инициализация
Private Sub Class_Initialize()
    Set this.ErrorManager = GetErrorManager
    Set this.logger = GetLogger
End Sub

' Реализация интерфейса IComponent
Private Sub IComponent_SetMediator(mediator As CommandMediator)
    Set this.mediator = mediator
End Sub

Private Function IComponent_GetComponentID() As String
    IComponent_GetComponentID = "ErrorHandler"
End Function

Private Function IComponent_ProcessMessage(messageType As String, data As Variant) As Boolean
    ' Проверка наличия медиатора
    If this.mediator Is Nothing Then
        this.logger.LogError "Медиатор не установлен", "ErrorHandlerComponent"
        IComponent_ProcessMessage = False
        Exit Function
    End If
    
    ' Обработка различных типов сообщений
    Select Case messageType
        Case "CommandExecuted"
            ' Обработка события выполнения команды
            IComponent_ProcessMessage = HandleCommandExecuted(data)
            
        Case "CommandUndone"
            ' Обработка события отмены команды
            IComponent_ProcessMessage = HandleCommandUndone(data)
            
        Case "ErrorOccurred"
            ' Обработка события возникновения ошибки
            IComponent_ProcessMessage = HandleErrorOccurred(data)
            
        Case "ErrorHandled"
            ' Обработка события обработки ошибки
            IComponent_ProcessMessage = HandleErrorHandled(data)
            
        Case Else
            ' Неизвестный тип сообщения
            this.logger.LogDebug "Получено неизвестное сообщение: " & messageType, "ErrorHandlerComponent"
            IComponent_ProcessMessage = False
    End Select
End Function

' Методы обработки событий
Private Function HandleCommandExecuted(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is Object  Then
        this.logger.LogWarning "Неверный формат данных для события CommandExecuted", "ErrorHandlerComponent"
        HandleCommandExecuted = False
        Exit Function
    End If
    
    ' Проверка наличия необходимых ключей
    If Not data.exists("Command") Or Not data.exists("Success") Then
        this.logger.LogWarning "Отсутствуют необходимые данные для события CommandExecuted", "ErrorHandlerComponent"
        HandleCommandExecuted = False
        Exit Function
    End If
    
    ' Получение данных
    Dim command As Object
    Dim success As Boolean
    
    Set command = data("Command")
    success = data("Success")
    
    ' Логирование события
    If success Then
        this.logger.LogInfo "Команда " & TypeName(command) & " выполнена успешно", "ErrorHandlerComponent"
    Else
        this.logger.LogWarning "Команда " & TypeName(command) & " не была выполнена успешно", "ErrorHandlerComponent"
    End If
    
    HandleCommandExecuted = True
End Function

Private Function HandleCommandUndone(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is Object  Then
        this.logger.LogWarning "Неверный формат данных для события CommandUndone", "ErrorHandlerComponent"
        HandleCommandUndone = False
        Exit Function
    End If
    
    ' Проверка наличия необходимых ключей
    If Not data.exists("Command") Or Not data.exists("Success") Then
        this.logger.LogWarning "Отсутствуют необходимые данные для события CommandUndone", "ErrorHandlerComponent"
        HandleCommandUndone = False
        Exit Function
    End If
    
    ' Получение данных
    Dim command As Object
    Dim success As Boolean
    
    Set command = data("Command")
    success = data("Success")
    
    ' Логирование события
    If success Then
        this.logger.LogInfo "Команда " & TypeName(command) & " отменена успешно", "ErrorHandlerComponent"
    Else
        this.logger.LogWarning "Команда " & TypeName(command) & " не была отменена успешно", "ErrorHandlerComponent"
    End If
    
    HandleCommandUndone = True
End Function

Private Function HandleErrorOccurred(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is errorInfo Then
        this.logger.LogWarning "Неверный формат данных для события ErrorOccurred", "ErrorHandlerComponent"
        HandleErrorOccurred = False
        Exit Function
    End If
    
    ' Получение данных об ошибке
    Dim errorInfo As errorInfo
    Set errorInfo = data
    
    ' Обработка ошибки через ErrorManager
    HandleErrorOccurred = this.ErrorManager.HandleError(errorInfo)
End Function

Private Function HandleErrorHandled(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is errorInfo Then
        this.logger.LogWarning "Неверный формат данных для события ErrorHandled", "ErrorHandlerComponent"
        HandleErrorHandled = False
        Exit Function
    End If
    
    ' Получение данных об ошибке
    Dim errorInfo As errorInfo
    Set errorInfo = data
    
    ' Логирование события
    this.logger.LogInfo "Ошибка обработана: " & errorInfo.description, "ErrorHandlerComponent"
    
    HandleErrorHandled = True
End Function

' Публичные методы
Public Sub HandleError(errorInfo As errorInfo)
    ' Отправка сообщения о возникновении ошибки через медиатор
    If Not this.mediator Is Nothing Then
        this.mediator.NotifyErrorOccurred errorInfo
    Else
        ' Если медиатор не установлен, обрабатываем ошибку напрямую
        this.ErrorManager.HandleError errorInfo
    End If
End Sub

Public Sub NotifyErrorHandled(errorInfo As errorInfo)
    ' Отправка сообщения об обработке ошибки через медиатор
    If Not this.mediator Is Nothing Then
        this.mediator.NotifyErrorHandled errorInfo
    End If
End Sub

'-------------------------------------------
' Component: LoggerComponent
'-------------------------------------------
' Класс LoggerComponent
' Компонент для логирования через медиатор
Option Explicit

Implements IComponent

' Внутренняя структура для хранения данных
Private Type TLoggerComponent
    mediator As CommandMediator
    logger As Object
    LogEntryHistory As Collection
    MaxHistorySize As Long
End Type

Private this As TLoggerComponent

' Инициализация
Private Sub Class_Initialize()
    Set this.logger = GetLogger
    Set this.LogEntryHistory = New Collection
    this.MaxHistorySize = 100 ' По умолчанию хранить последние 100 записей
End Sub

' Реализация интерфейса IComponent
Private Sub IComponent_SetMediator(mediator As CommandMediator)
    Set this.mediator = mediator
End Sub

Private Function IComponent_GetComponentID() As String
    IComponent_GetComponentID = "Logger"
End Function

Private Function IComponent_ProcessMessage(messageType As String, data As Variant) As Boolean
    ' Проверка наличия медиатора
    If this.mediator Is Nothing Then
        Debug.Print "LoggerComponent: Медиатор не установлен"
        IComponent_ProcessMessage = False
        Exit Function
    End If
    
    ' Обработка различных типов сообщений
    Select Case messageType
        Case "LogEntryAdded"
            ' Обработка события добавления записи в лог
            IComponent_ProcessMessage = HandleLogEntryAdded(data)
            
        Case "CommandExecuted"
            ' Обработка события выполнения команды
            IComponent_ProcessMessage = HandleCommandExecuted(data)
            
        Case "CommandUndone"
            ' Обработка события отмены команды
            IComponent_ProcessMessage = HandleCommandUndone(data)
            
        Case "ErrorOccurred"
            ' Обработка события возникновения ошибки
            IComponent_ProcessMessage = HandleErrorOccurred(data)
            
        Case "ErrorHandled"
            ' Обработка события обработки ошибки
            IComponent_ProcessMessage = HandleErrorHandled(data)
            
        Case Else
            ' Неизвестный тип сообщения
            Debug.Print "LoggerComponent: Получено неизвестное сообщение: " & messageType
            IComponent_ProcessMessage = False
    End Select
End Function

' Методы обработки событий
Private Function HandleLogEntryAdded(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is logEntry Then
        Debug.Print "LoggerComponent: Неверный формат данных для события LogEntryAdded"
        HandleLogEntryAdded = False
        Exit Function
    End If
    
    ' Получение данных
    Dim logEntry As logEntry
    Set logEntry = data
    
    ' Сохранение записи в истории
    this.LogEntryHistory.Add logEntry.Clone()
    TrimHistory
    
    ' Логирование события (в данном случае просто отладочная информация)
    Debug.Print "LoggerComponent: Добавлена запись в лог: " & logEntry.ToString()
    
    HandleLogEntryAdded = True
End Function

Private Function HandleCommandExecuted(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is Object  Then
        Debug.Print "LoggerComponent: Неверный формат данных для события CommandExecuted"
        HandleCommandExecuted = False
        Exit Function
    End If
    
    ' Проверка наличия необходимых ключей
    If Not data.exists("Command") Or Not data.exists("Success") Then
        Debug.Print "LoggerComponent: Отсутствуют необходимые данные для события CommandExecuted"
        HandleCommandExecuted = False
        Exit Function
    End If
    
    ' Получение данных
    Dim command As Object
    Dim success As Boolean
    
    Set command = data("Command")
    success = data("Success")
    
    ' Логирование события
    If success Then
        this.logger.LogInfo "Команда " & TypeName(command) & " выполнена успешно", "LoggerComponent"
    Else
        this.logger.LogWarning "Команда " & TypeName(command) & " не была выполнена успешно", "LoggerComponent"
    End If
    
    HandleCommandExecuted = True
End Function

Private Function HandleCommandUndone(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is Object  Then
        Debug.Print "LoggerComponent: Неверный формат данных для события CommandUndone"
        HandleCommandUndone = False
        Exit Function
    End If
    
    ' Проверка наличия необходимых ключей
    If Not data.exists("Command") Or Not data.exists("Success") Then
        Debug.Print "LoggerComponent: Отсутствуют необходимые данные для события CommandUndone"
        HandleCommandUndone = False
        Exit Function
    End If
    
    ' Получение данных
    Dim command As Object
    Dim success As Boolean
    
    Set command = data("Command")
    success = data("Success")
    
    ' Логирование события
    If success Then
        this.logger.LogInfo "Команда " & TypeName(command) & " отменена успешно", "LoggerComponent"
    Else
        this.logger.LogWarning "Команда " & TypeName(command) & " не была отменена успешно", "LoggerComponent"
    End If
    
    HandleCommandUndone = True
End Function

Private Function HandleErrorOccurred(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is errorInfo Then
        Debug.Print "LoggerComponent: Неверный формат данных для события ErrorOccurred"
        HandleErrorOccurred = False
        Exit Function
    End If
    
    ' Получение данных об ошибке
    Dim errorInfo As errorInfo
    Set errorInfo = data
    
    ' Логирование события
    this.logger.LogError "Произошла ошибка: " & errorInfo.ToString(), "LoggerComponent"
    
    HandleErrorOccurred = True
End Function

Private Function HandleErrorHandled(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is errorInfo Then
        Debug.Print "LoggerComponent: Неверный формат данных для события ErrorHandled"
        HandleErrorHandled = False
        Exit Function
    End If
    
    ' Получение данных об ошибке
    Dim errorInfo As errorInfo
    Set errorInfo = data
    
    ' Логирование события
    this.logger.LogInfo "Ошибка обработана: " & errorInfo.ToString(), "LoggerComponent"
    
    HandleErrorHandled = True
End Function

' Вспомогательные методы
Private Sub TrimHistory()
    ' Удаление старых записей, если их больше максимального размера
    While this.LogEntryHistory.Count > this.MaxHistorySize
        this.LogEntryHistory.Remove 1
    Wend
End Sub

' Публичные методы
Public Sub Log(level As LogLevel, message As String, Optional source As String = "")
    ' Создание записи лога
    Dim entry As New logEntry
    entry.level = level
    entry.message = message
    entry.source = source
    entry.Timestamp = Now
    
    ' Добавление записи в историю
    this.LogEntryHistory.Add entry.Clone()
    TrimHistory
    
    ' Запись через логгер
    this.logger.Log level, message, source
    
    ' Уведомление через медиатор
    If Not this.mediator Is Nothing Then
        this.mediator.NotifyLogEntryAdded entry
    End If
End Sub

Public Sub LogDebug(message As String, Optional source As String = "")
    Log LogLevel.LogDebug, message, source
End Sub

Public Sub LogInfo(message As String, Optional source As String = "")
    Log LogLevel.LogInfo, message, source
End Sub

Public Sub LogWarning(message As String, Optional source As String = "")
    Log LogLevel.LogWarning, message, source
End Sub

Public Sub LogError(message As String, Optional source As String = "")
    Log LogLevel.LogError, message, source
End Sub

Public Sub LogCritical(message As String, Optional source As String = "")
    Log LogLevel.LogCritical, message, source
End Sub

Public Property Let MaxHistorySize(value As Long)
    this.MaxHistorySize = value
    TrimHistory
End Property

Public Property Get MaxHistorySize() As Long
    MaxHistorySize = this.MaxHistorySize
End Property

Public Function GetHistory() As Collection
    Set GetHistory = this.LogEntryHistory
End Function

Public Sub ClearHistory()
    Set this.LogEntryHistory = New Collection
End Sub

'-------------------------------------------
' Component: CommandManagerComponent
'-------------------------------------------
' Класс CommandManagerComponent
' Компонент для управления командами через медиатор
Option Explicit

Implements IComponent

' Внутренняя структура для хранения данных
Private Type TCommandManagerComponent
    mediator As CommandMediator
    CommandInvoker As Object
    logger As Object
    LastErrorMessage As String
End Type

Private this As TCommandManagerComponent

' Инициализация
Private Sub Class_Initialize()
    Set this.CommandInvoker = GetCommandInvoker
    Set this.logger = GetLogger
    this.LastErrorMessage = ""
End Sub

' Реализация интерфейса IComponent
Private Sub IComponent_SetMediator(mediator As CommandMediator)
    Set this.mediator = mediator
End Sub

Private Function IComponent_GetComponentID() As String
    IComponent_GetComponentID = "CommandManager"
End Function

Private Function IComponent_ProcessMessage(messageType As String, data As Variant) As Boolean
    ' Проверка наличия медиатора
    If this.mediator Is Nothing Then
        this.logger.LogError "Медиатор не установлен", "CommandManagerComponent"
        IComponent_ProcessMessage = False
        Exit Function
    End If
    
    ' Обработка различных типов сообщений
    Select Case messageType
        Case "ExecuteCommand"
            ' Обработка запроса на выполнение команды
            IComponent_ProcessMessage = HandleExecuteCommand(data)
            
        Case "UndoCommand"
            ' Обработка запроса на отмену команды
            IComponent_ProcessMessage = HandleUndoCommand(data)
            
        Case "ErrorOccurred"
            ' Обработка события возникновения ошибки
            IComponent_ProcessMessage = HandleErrorOccurred(data)
            
        Case Else
            ' Неизвестный тип сообщения
            this.logger.LogDebug "Получено неизвестное сообщение: " & messageType, "CommandManagerComponent"
            IComponent_ProcessMessage = False
    End Select
End Function

' Методы обработки событий
Private Function HandleExecuteCommand(data As Variant) As Boolean
    ' Проверка формата данных
    If TypeOf data Is Object  Then
        ' Если данные содержат команду напрямую
        If data.exists("Command") Then
            Dim command As Object
            Set command = data("Command")
            
            ' Установка команды в инвокер
            this.CommandInvoker.SetCommand command
            
            ' Выполнение команды
            HandleExecuteCommand = this.CommandInvoker.ExecuteCommand()
            Exit Function
        End If
    ElseIf IsObject(data) Then
        ' Если данные - это команда напрямую
        ' Установка команды в инвокер
        this.CommandInvoker.SetCommand data
        
        ' Выполнение команды
        HandleExecuteCommand = this.CommandInvoker.ExecuteCommand()
        Exit Function
    End If
    
    ' Если формат данных неверный
    this.logger.LogWarning "Неверный формат данных для выполнения команды", "CommandManagerComponent"
    HandleExecuteCommand = False
End Function

Private Function HandleUndoCommand(data As Variant) As Boolean
    ' Проверка, что инвокер имеет команды в истории
    If this.CommandInvoker.GetHistoryCount() = 0 Then
        this.logger.LogWarning "Нет команд для отмены", "CommandManagerComponent"
        this.LastErrorMessage = "Нет команд для отмены"
        HandleUndoCommand = False
        Exit Function
    End If
    
    ' Отмена последней команды
    Dim success As Boolean
    success = this.CommandInvoker.UndoLastCommand()
    
    ' Если команда отменена успешно, уведомляем через медиатор
    If success Then
        ' Получаем команду из истории (если возможно)
        Dim command As Object
        On Error Resume Next
        Set command = this.CommandInvoker.GetCommand
        On Error GoTo 0
        
        ' Создание словаря с информацией для передачи компонентам
        Dim resultData As Object
        Set resultData = CreateObject("Scripting.Dictionary")
        resultData.Add "Command", command
        resultData.Add "Success", True
        
        ' Отправка уведомления через медиатор
        this.mediator.NotifyCommandUndone command, True
    Else
        ' Если команда не отменена, сохраняем сообщение об ошибке
        this.LastErrorMessage = this.CommandInvoker.LastError
        
        ' Получаем команду из истории (если возможно)
        Dim errorCommand As Object
        On Error Resume Next
        Set errorCommand = this.CommandInvoker.GetCommand
        On Error GoTo 0
        
        ' Создание словаря с информацией для передачи компонентам
        Dim errorData As Object
        Set errorData = CreateObject("Scripting.Dictionary")
        errorData.Add "Command", errorCommand
        errorData.Add "Success", False
        errorData.Add "ErrorMessage", this.LastErrorMessage
        
        ' Отправка уведомления через медиатор
        this.mediator.NotifyCommandUndone errorCommand, False
        
        ' Создание и отправка информации об ошибке
        Dim errorInfo As New errorInfo
        errorInfo.Number = vbObjectError + 5001
        errorInfo.description = this.LastErrorMessage
        errorInfo.source = "CommandManagerComponent"
        errorInfo.context = "Отмена команды"
        errorInfo.commandName = IIf(errorCommand Is Nothing, "Unknown", TypeName(errorCommand))
        errorInfo.severity = 3 ' Средняя серьезность
        
        this.mediator.NotifyErrorOccurred errorInfo
    End If
    
    HandleUndoCommand = success
End Function

Private Function HandleErrorOccurred(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is errorInfo Then
        this.logger.LogWarning "Неверный формат данных для события ErrorOccurred", "CommandManagerComponent"
        HandleErrorOccurred = False
        Exit Function
    End If
    
    ' Получение данных об ошибке
    Dim errorInfo As errorInfo
    Set errorInfo = data
    
    ' Логирование события
    this.logger.LogError "Получено уведомление об ошибке: " & errorInfo.description, "CommandManagerComponent"
    
    ' Сохранение сообщения об ошибке
    this.LastErrorMessage = errorInfo.description
    
    HandleErrorOccurred = True
End Function

' Публичные методы
Public Function ExecuteCommand(command As Object) As Boolean
    ' Проверка наличия инвокера
    If this.CommandInvoker Is Nothing Then
        this.logger.LogError "CommandInvoker не инициализирован", "CommandManagerComponent"
        ExecuteCommand = False
        Exit Function
    End If

    ' Проверка наличия медиатора
    If this.mediator Is Nothing Then
        ' Если медиатор не установлен, выполняем команду напрямую через инвокер
        On Error Resume Next
        this.CommandInvoker.SetCommand command
        
        If err.Number <> 0 Then
            this.logger.LogError "Ошибка при установке команды: " & err.description, "CommandManagerComponent"
            ExecuteCommand = False
            Exit Function
        End If
        
        ExecuteCommand = this.CommandInvoker.ExecuteCommand()
        
        If err.Number <> 0 Then
            this.logger.LogError "Ошибка при выполнении команды: " & err.description, "CommandManagerComponent"
            ExecuteCommand = False
        End If
        On Error GoTo 0
    Else
        ' Если медиатор установлен, обрабатываем команду через HandleExecuteCommand
        ' чтобы избежать бесконечной рекурсии
        ExecuteCommand = HandleExecuteCommand(command)
    End If
End Function

Public Function UndoLastCommand() As Boolean
    ' Проверка наличия медиатора
    If this.mediator Is Nothing Then
        ' Если медиатор не установлен, отменяем команду напрямую через инвокер
        UndoLastCommand = this.CommandInvoker.UndoLastCommand()
    Else
        ' Если медиатор установлен, отправляем запрос на отмену команды
        UndoLastCommand = this.mediator.SendMessage("CommandManager", "UndoCommand", Nothing)
    End If
End Function

Public Property Get LastError() As String
    LastError = this.LastErrorMessage
End Property

Public Property Get HistoryCount() As Long
    HistoryCount = this.CommandInvoker.GetHistoryCount()
End Property

Public Sub ClearCommandHistory()
    this.CommandInvoker.ClearHistory
    this.logger.LogInfo "История команд очищена", "CommandManagerComponent"
End Sub

'-------------------------------------------
' Component: MediatorSystemInitializer
'-------------------------------------------
' Модуль MediatorSystemInitializer
' Инициализация системы с использованием Mediator
Option Explicit

' Инициализация системы с использованием медиатора
Public Sub InitializeSystem()
    ' Получаем логгер для вывода информации
    Dim logger As Object
    Set logger = GetLogger
    
    logger.LogInfo "Начало инициализации системы с использованием Mediator", "MediatorSystemInitializer"
    
    ' Получаем медиатор
    Dim mediator As CommandMediator
    Set mediator = GetCommandMediator
    
    ' Инициализируем медиатор
    mediator.Initialize
    
    ' Создаем компоненты системы
    Dim loggerComponent As New loggerComponent
    Dim errorHandlerComponent As New errorHandlerComponent
    Dim commandManagerComponent As New commandManagerComponent
    
    ' Регистрируем компоненты в медиаторе
    mediator.RegisterComponent loggerComponent
    mediator.RegisterComponent errorHandlerComponent
    mediator.RegisterComponent commandManagerComponent
    
    ' Инициализация системы обработки ошибок
    InitializeErrorSystem
    
    ' Отправка тестового сообщения через медиатор
    mediator.BroadcastMessage "SystemInitialized", Now
    
    logger.LogInfo "Система успешно инициализирована", "MediatorSystemInitializer"
End Sub

' Тестирование системы с использованием медиатора
Public Sub TestMediatorSystem()
    ' Получаем логгер для вывода информации
    Dim logger As Object
    Set logger = GetLogger
    
    logger.LogInfo "Начало тестирования системы с использованием Mediator", "MediatorSystemInitializer"
    
    ' Получаем медиатор
    Dim mediator As CommandMediator
    Set mediator = GetCommandMediator
    
    ' Проверяем, что медиатор инициализирован
    If Not mediator.IsInitialized Then
        InitializeSystem
    End If
    
    ' Создаем тестовую команду
    Dim testCommand As New SetCellValueCommand
    testCommand.Initialize "Sheet1", "C1", "Тестовое значение через Mediator"
    
    ' Получаем компонент управления командами
    Dim commandManager As commandManagerComponent
    For Each key In mediator.Components.Keys()
        If key = "CommandManager" Then
            Set commandManager = mediator.Components(key)
            Exit For
        End If
    Next key
    
    If commandManager Is Nothing Then
        logger.LogError "Компонент CommandManager не найден", "MediatorSystemInitializer"
        Exit Sub
    End If
    
    ' Выполняем команду через компонент управления командами
    logger.LogInfo "Выполнение команды через Mediator", "MediatorSystemInitializer"
    
    Dim success As Boolean
    success = commandManager.ExecuteCommand(testCommand)
    
    If success Then
        logger.LogInfo "Команда успешно выполнена", "MediatorSystemInitializer"
    Else
        logger.LogError "Ошибка выполнения команды: " & commandManager.LastError, "MediatorSystemInitializer"
    End If
    
    ' Отменяем команду через компонент управления командами
    logger.LogInfo "Отмена команды через Mediator", "MediatorSystemInitializer"
    
    success = commandManager.UndoLastCommand()
    
    If success Then
        logger.LogInfo "Команда успешно отменена", "MediatorSystemInitializer"
    Else
        logger.LogError "Ошибка отмены команды: " & commandManager.LastError, "MediatorSystemInitializer"
    End If
    
    ' Генерируем тестовую ошибку
    logger.LogInfo "Генерация тестовой ошибки", "MediatorSystemInitializer"
    
    Dim errorInfo As New errorInfo
    errorInfo.Number = vbObjectError + 9999
    errorInfo.description = "Тестовая ошибка для проверки работы Mediator"
    errorInfo.source = "MediatorSystemInitializer"
    errorInfo.context = "Тестирование Mediator"
    errorInfo.severity = 2 ' Низкая серьезность
    
    ' Получаем компонент обработки ошибок
    Dim errorHandler As errorHandlerComponent
    For Each key In mediator.Components.Keys()
        If key = "ErrorHandler" Then
            Set errorHandler = mediator.Components(key)
            Exit For
        End If
    Next key
    
    If errorHandler Is Nothing Then
        logger.LogError "Компонент ErrorHandler не найден", "MediatorSystemInitializer"
        Exit Sub
    End If
    
    ' Обрабатываем ошибку через компонент обработки ошибок
    errorHandler.HandleError errorInfo
    
    logger.LogInfo "Тестирование системы с использованием Mediator успешно завершено", "MediatorSystemInitializer"
End Sub

' Добавление публичных свойств IsInitialized в CommandMediator класс
Sub UpdateCommandMediatorClass()
    ' Получаем медиатор
    Dim mediator As CommandMediator
    Set mediator = GetCommandMediator
    
    ' Проверяем наличие свойства IsInitialized
    Dim IsInitialized As Boolean
    On Error Resume Next
    IsInitialized = mediator.IsInitialized
    On Error GoTo 0
    
    If err.Number <> 0 Then
        MsgBox "Необходимо добавить свойство IsInitialized в класс CommandMediator", vbInformation, "Обновление класса"
    End If
End Sub

' Вспомогательная функция для ожидания
Sub WaitForSeconds(seconds As Integer)
    Application.Wait Now + TimeValue("00:00:" & seconds)
End Sub

'-------------------------------------------
' Component: TestComponent
'-------------------------------------------
' Класс TestComponent
' Тестовый компонент для демонстрации работы с медиатором
Option Explicit

Implements IComponent

' Внутренняя структура для хранения данных
Private Type TTestComponent
    mediator As CommandMediator
    logger As Object
    MessageCount As Long
End Type

Private this As TTestComponent

' Инициализация
Private Sub Class_Initialize()
    Set this.logger = GetLogger
    this.MessageCount = 0
End Sub

' Реализация интерфейса IComponent
Private Sub IComponent_SetMediator(mediator As CommandMediator)
    Set this.mediator = mediator
End Sub

Private Function IComponent_GetComponentID() As String
    IComponent_GetComponentID = "TestComponent"
End Function

Private Function IComponent_ProcessMessage(messageType As String, data As Variant) As Boolean
    ' Проверка наличия медиатора
    If this.mediator Is Nothing Then
        this.logger.LogError "Медиатор не установлен", "TestComponent"
        IComponent_ProcessMessage = False
        Exit Function
    End If
    
    ' Увеличиваем счетчик сообщений
    this.MessageCount = this.MessageCount + 1
    
    ' Обработка различных типов сообщений
    Select Case messageType
        Case "TestMessage"
            ' Обработка тестового сообщения
            IComponent_ProcessMessage = HandleTestMessage(data)
            
        Case "SystemStatus"
            ' Обработка статуса системы
            IComponent_ProcessMessage = HandleSystemStatus(data)
            
        Case "SystemInitialized"
            ' Обработка инициализации системы
            IComponent_ProcessMessage = HandleSystemInitialized(data)
            
        Case Else
            ' Неизвестный тип сообщения
            this.logger.LogDebug "Получено неизвестное сообщение: " & messageType, "TestComponent"
            IComponent_ProcessMessage = False
    End Select
End Function

' Методы обработки сообщений
Private Function HandleTestMessage(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is Object  Then
        this.logger.LogWarning "Неверный формат данных для события TestMessage", "TestComponent"
        HandleTestMessage = False
        Exit Function
    End If
    
    ' Проверка наличия необходимых ключей
    If Not data.exists("Message") Then
        this.logger.LogWarning "Отсутствуют необходимые данные для события TestMessage", "TestComponent"
        HandleTestMessage = False
        Exit Function
    End If
    
    ' Получение данных
    Dim message As String
    message = data("Message")
    
    ' Логирование события
    this.logger.LogInfo "Получено тестовое сообщение: " & message, "TestComponent"
    
    ' Отправка ответа через медиатор
    Dim responseData As Object
    Set responseData = CreateObject("Scripting.Dictionary")
    responseData.Add "Response", "Сообщение получено и обработано!"
    responseData.Add "MessageCount", this.MessageCount
    responseData.Add "Source", "TestComponent"
    
    If data.exists("Source") Then
        this.mediator.SendMessage data("Source"), "TestResponse", responseData
    End If
    
    HandleTestMessage = True
End Function

Private Function HandleSystemStatus(data As Variant) As Boolean
    ' Проверка формата данных
    If Not TypeOf data Is Object  Then
        this.logger.LogWarning "Неверный формат данных для события SystemStatus", "TestComponent"
        HandleSystemStatus = False
        Exit Function
    End If
    
    ' Проверка наличия необходимых ключей
    If Not data.exists("Status") Then
        this.logger.LogWarning "Отсутствуют необходимые данные для события SystemStatus", "TestComponent"
        HandleSystemStatus = False
        Exit Function
    End If
    
    ' Получение данных
    Dim status As String
    status = data("Status")
    
    ' Логирование события
    this.logger.LogInfo "Получен статус системы: " & status, "TestComponent"
    
    HandleSystemStatus = True
End Function

Private Function HandleSystemInitialized(data As Variant) As Boolean
    ' Логирование события
    this.logger.LogInfo "Система инициализирована", "TestComponent"
    
    HandleSystemInitialized = True
End Function

' Публичные методы
Public Property Get MessageCount() As Long
    MessageCount = this.MessageCount
End Property

'-------------------------------------------
' Component: Module1
'-------------------------------------------
' Модуль MediatorTestModule
' Демонстрация работы паттерна Mediator и его интеграции в систему
Option Explicit

' Главная процедура тестирования системы с паттерном Mediator
Public Sub TestMediatorIntegration()
    ' Настраиваем логирование для тестирования
    InitializeLogger
    
    ' Получаем логгер для вывода информации
    Dim logger As Object
    Set logger = GetLogger
    
    ' Логируем начало тестирования
    logger.LogInfo "=== НАЧАЛО ТЕСТИРОВАНИЯ ПАТТЕРНА MEDIATOR ===", "MediatorTestModule"
    
    ' Инициализируем систему с использованием медиатора
    InitializeSystemWithMediator
    
    ' Последовательно выполняем тесты
    TestCommandExecution
    TestCommandUndo
    TestErrorHandling
    TestComponentCommunication
    TestDynamicComponentAddition
    
    ' Логируем завершение тестирования
    logger.LogInfo "=== ТЕСТИРОВАНИЕ ПАТТЕРНА MEDIATOR УСПЕШНО ЗАВЕРШЕНО ===", "MediatorTestModule"
    
    ' Выводим результаты в MsgBox
    MsgBox "Тестирование системы с использованием паттерна Mediator успешно завершено! " & _
           "Проверьте журнал для получения подробной информации.", _
           vbInformation, "Тестирование паттерна Mediator"
End Sub

' Инициализация логгера
Private Sub InitializeLogger()
    Dim logger As Object
    Set logger = GetLogger
    
    ' Настройка файла лога в папке документов пользователя
    Dim logPath As String
    logPath = Environ("USERPROFILE") & "\Documents\CommandPatternLogs\"
    
    ' Создаем папку, если она не существует
    On Error Resume Next
    If Dir(logPath, vbDirectory) = "" Then
        MkDir logPath
    End If
    On Error GoTo 0
    
    ' Устанавливаем файл лога
    logger.LogFile = logPath & "MediatorSystem_" & Format(Now, "yyyymmdd_hhmmss") & ".log"
    
    ' Включаем все уровни логирования для тестирования
    logger.MinimumLevel = LogLevel.LogDebug
    
    ' Очищаем историю логов
    logger.ClearHistory
    
    ' Логируем информацию о начале тестирования
    logger.LogInfo "Логгер инициализирован. Файл лога: " & logger.LogFile, "MediatorTestModule"
End Sub

' Инициализация системы с использованием медиатора
Private Sub InitializeSystemWithMediator()
    Dim logger As Object
    Set logger = GetLogger
    
    logger.LogInfo "Начало инициализации системы с использованием Mediator", "MediatorTestModule"
    
    ' Получаем медиатор
    Dim mediator As CommandMediator
    Set mediator = GetCommandMediator
    
    If mediator Is Nothing Then
        logger.LogError "Не удалось получить экземпляр CommandMediator", "MediatorTestModule"
        Exit Sub
    End If
    
    ' Инициализируем медиатор
    mediator.Initialize
    
    ' Создаем компоненты системы
    Dim loggerComponent As New loggerComponent
    Dim errorHandlerComponent As New errorHandlerComponent
    Dim commandManagerComponent As New commandManagerComponent
    
    If loggerComponent Is Nothing Or errorHandlerComponent Is Nothing Or commandManagerComponent Is Nothing Then
        logger.LogError "Не удалось создать компоненты системы", "MediatorTestModule"
        Exit Sub
    End If
    
    ' Регистрируем компоненты в медиаторе
    logger.LogInfo "Регистрация компонентов в медиаторе:", "MediatorTestModule"
    
    logger.LogInfo "1. Регистрация компонента Logger", "MediatorTestModule"
    mediator.RegisterComponent loggerComponent
    
    logger.LogInfo "2. Регистрация компонента ErrorHandler", "MediatorTestModule"
    mediator.RegisterComponent errorHandlerComponent
    
    logger.LogInfo "3. Регистрация компонента CommandManager", "MediatorTestModule"
    mediator.RegisterComponent commandManagerComponent
    
    ' Инициализация системы обработки ошибок
    logger.LogInfo "Инициализация системы обработки ошибок", "MediatorTestModule"
    InitializeErrorSystem
    
    ' Отправка тестового сообщения через медиатор
    logger.LogInfo "Отправка тестового широковещательного сообщения", "MediatorTestModule"
    
    ' Создаем объект Dictionary для передачи информации
    On Error Resume Next
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")
    
    If err.Number <> 0 Then
        logger.LogError "Ошибка при создании Dictionary: " & err.description, "MediatorTestModule"
        Exit Sub
    End If
    
    data.Add "Message", "Система успешно инициализирована"
    data.Add "Timestamp", Now
    
    Dim componentsNotified As Long
    componentsNotified = mediator.BroadcastMessage("SystemInitialized", data)
    
    If err.Number <> 0 Then
        logger.LogError "Ошибка при отправке сообщения: " & err.description, "MediatorTestModule"
    Else
        logger.LogInfo "Сообщение получено " & componentsNotified & " компонентами", "MediatorTestModule"
    End If
    On Error GoTo 0
    
    logger.LogInfo "Система успешно инициализирована с использованием Mediator", "MediatorTestModule"
End Sub

' Тестирование выполнения команды через медиатор
Private Sub TestCommandExecution()
    Dim logger As Object
    Set logger = GetLogger
    
    logger.LogInfo "=== ТЕСТ 1: Выполнение команды через Mediator ===", "MediatorTestModule"
    
    ' Получаем медиатор
    Dim mediator As CommandMediator
    Set mediator = GetCommandMediator
    
    ' Создаем тестовую команду
    Dim testCommand As New SetCellValueCommand
    testCommand.Initialize "Sheet1", "D1", "Значение через Mediator"
    
    ' Получаем компонент управления командами
    Dim commandManager As Object
    Set commandManager = GetComponentByID(mediator, "CommandManager")
    
    If commandManager Is Nothing Then
        logger.LogError "Компонент CommandManager не найден", "MediatorTestModule"
        Exit Sub
    End If
    
   ' Выполняем команду через медиатор
logger.LogInfo "Выполнение команды установки значения в ячейку D1", "MediatorTestModule"

Dim success As Boolean
success = mediator.ExecuteCommand(testCommand)

' Проверяем результат
If success Then
    logger.LogInfo "Команда успешно выполнена через Mediator", "MediatorTestModule"
    
    ' Проверяем, что значение действительно установлено
    Dim actualValue As String
    actualValue = ThisWorkbook.Worksheets("Sheet1").Range("D1").value
    
    logger.LogInfo "Значение в ячейке D1: " & actualValue, "MediatorTestModule"
    
    If actualValue = "Значение через Mediator" Then
        logger.LogInfo "Значение установлено корректно", "MediatorTestModule"
    Else
        logger.LogError "Значение установлено некорректно", "MediatorTestModule"
    End If
Else
    logger.LogError "Ошибка выполнения команды", "MediatorTestModule"
End If
    
    logger.LogInfo "Тест выполнения команды через Mediator завершен", "MediatorTestModule"
End Sub

' Тестирование отмены команды через медиатор
Private Sub TestCommandUndo()
    Dim logger As Object
    Set logger = GetLogger
    
    logger.LogInfo "=== ТЕСТ 2: Отмена команды через Mediator ===", "MediatorTestModule"
    
    ' Получаем медиатор
    Dim mediator As CommandMediator
    Set mediator = GetCommandMediator
    
    ' Создаем тестовую команду
    Dim testCommand As New SetCellValueCommand
    testCommand.Initialize "Sheet1", "E1", "Значение для отмены через Mediator"
    
    ' Получаем компонент управления командами
    Dim commandManager As Object
    Set commandManager = GetComponentByID(mediator, "CommandManager")
    
    If commandManager Is Nothing Then
        logger.LogError "Компонент CommandManager не найден", "MediatorTestModule"
        Exit Sub
    End If
    
    ' Запоминаем текущее значение
    Dim originalValue As Variant
    originalValue = ThisWorkbook.Worksheets("Sheet1").Range("E1").value
    
    logger.LogInfo "Исходное значение в ячейке E1: " & IIf(IsEmpty(originalValue), "[EMPTY]", originalValue), "MediatorTestModule"
    
    ' Выполняем команду через компонент управления командами
    logger.LogInfo "Выполнение команды установки значения в ячейку E1", "MediatorTestModule"
    
    Dim success As Boolean
    success = commandManager.ExecuteCommand(testCommand)
    
    If success Then
        logger.LogInfo "Команда успешно выполнена", "MediatorTestModule"
        
        ' Проверяем, что значение действительно установлено
        Dim newValue As String
        newValue = ThisWorkbook.Worksheets("Sheet1").Range("E1").value
        
        logger.LogInfo "Новое значение в ячейке E1: " & newValue, "MediatorTestModule"
        
        ' Отменяем команду через компонент управления командами
        logger.LogInfo "Отмена команды через Mediator", "MediatorTestModule"
        
        success = commandManager.UndoLastCommand()
        
        If success Then
            logger.LogInfo "Команда успешно отменена через Mediator", "MediatorTestModule"
            
            ' Проверяем, что значение восстановлено
            Dim restoredValue As Variant
            restoredValue = ThisWorkbook.Worksheets("Sheet1").Range("E1").value
            
            logger.LogInfo "Восстановленное значение в ячейке E1: " & IIf(IsEmpty(restoredValue), "[EMPTY]", restoredValue), "MediatorTestModule"
            
            ' Сравниваем с исходным значением (учитывая возможность Empty)
            If IsEmpty(originalValue) And IsEmpty(restoredValue) Then
                logger.LogInfo "Значение успешно восстановлено (пустое)", "MediatorTestModule"
            ElseIf originalValue = restoredValue Then
                logger.LogInfo "Значение успешно восстановлено", "MediatorTestModule"
            Else
                logger.LogError "Значение восстановлено некорректно", "MediatorTestModule"
            End If
        Else
            logger.LogError "Ошибка отмены команды: " & commandManager.LastError, "MediatorTestModule"
        End If
    Else
        logger.LogError "Ошибка выполнения команды: " & commandManager.LastError, "MediatorTestModule"
    End If
    
    logger.LogInfo "Тест отмены команды через Mediator завершен", "MediatorTestModule"
End Sub

' Тестирование обработки ошибок через медиатор
Private Sub TestErrorHandling()
    Dim logger As Object
    Set logger = GetLogger
    
    logger.LogInfo "=== ТЕСТ 3: Обработка ошибок через Mediator ===", "MediatorTestModule"
    
    ' Получаем медиатор
    Dim mediator As CommandMediator
    Set mediator = GetCommandMediator
    
    ' Получаем компонент обработки ошибок
    Dim errorHandler As Object
    Set errorHandler = GetComponentByID(mediator, "ErrorHandler")
    
    If errorHandler Is Nothing Then
        logger.LogError "Компонент ErrorHandler не найден", "MediatorTestModule"
        Exit Sub
    End If
    
    ' Создаем тестовую ошибку низкой серьезности
    logger.LogInfo "Генерация тестовой ошибки низкой серьезности", "MediatorTestModule"
    
    Dim lowErrorInfo As New errorInfo
    lowErrorInfo.Number = vbObjectError + 8001
    lowErrorInfo.description = "Тестовая ошибка низкой серьезности"
    lowErrorInfo.source = "MediatorTestModule"
    lowErrorInfo.context = "Тестирование обработки ошибок"
    lowErrorInfo.severity = 2 ' Низкая серьезность
    
    ' Обрабатываем ошибку через компонент обработки ошибок
    errorHandler.HandleError lowErrorInfo
    
    ' Создаем тестовую ошибку высокой серьезности
    logger.LogInfo "Генерация тестовой ошибки высокой серьезности", "MediatorTestModule"
    
    Dim highErrorInfo As New errorInfo
    highErrorInfo.Number = vbObjectError + 8002
    highErrorInfo.description = "Тестовая ошибка высокой серьезности"
    highErrorInfo.source = "MediatorTestModule"
    highErrorInfo.context = "Тестирование обработки ошибок"
    highErrorInfo.severity = 4 ' Высокая серьезность
    
    ' Обрабатываем ошибку через компонент обработки ошибок
    errorHandler.HandleError highErrorInfo
    
    ' Создаем тестовую ошибку валидации
    logger.LogInfo "Генерация тестовой ошибки валидации", "MediatorTestModule"
    
    Dim validationErrorInfo As New errorInfo
    validationErrorInfo.Number = vbObjectError + 10050
    validationErrorInfo.description = "Недопустимый формат данных"
    validationErrorInfo.source = "ValidationManager"
    validationErrorInfo.context = "Валидация входных данных"
    validationErrorInfo.severity = 3 ' Средняя серьезность
    
    ' Обрабатываем ошибку через компонент обработки ошибок
    errorHandler.HandleError validationErrorInfo
    
    logger.LogInfo "Тест обработки ошибок через Mediator завершен", "MediatorTestModule"
End Sub

' Тестирование коммуникации между компонентами через медиатор
Private Sub TestComponentCommunication()
    Dim logger As Object
    Set logger = GetLogger
    
    logger.LogInfo "=== ТЕСТ 4: Коммуникация между компонентами через Mediator ===", "MediatorTestModule"
    
    ' Получаем медиатор
    Dim mediator As CommandMediator
    Set mediator = GetCommandMediator
    
    ' Получаем компоненты
    Dim loggerComponent As Object
    Dim errorHandler As Object
    Dim commandManager As Object
    
    Set loggerComponent = GetComponentByID(mediator, "Logger")
    Set errorHandler = GetComponentByID(mediator, "ErrorHandler")
    Set commandManager = GetComponentByID(mediator, "CommandManager")
    
    If loggerComponent Is Nothing Or errorHandler Is Nothing Or commandManager Is Nothing Then
        logger.LogError "Не удалось получить все необходимые компоненты", "MediatorTestModule"
        Exit Sub
    End If
    
    ' Демонстрация взаимодействия компонентов через медиатор
    logger.LogInfo "Демонстрация взаимодействия компонентов через Mediator", "MediatorTestModule"
    
    ' Создаем событие для отправки всем компонентам
    Dim eventData As Object
    Set eventData = CreateObject("Scripting.Dictionary")
    eventData.Add "EventType", "SystemStatus"
    eventData.Add "Status", "Running"
    eventData.Add "Timestamp", Now
    eventData.Add "Source", "MediatorTestModule"
    
    ' Отправляем событие всем компонентам через медиатор
    logger.LogInfo "Отправка события 'SystemStatus' всем компонентам", "MediatorTestModule"
    
    Dim componentsNotified As Long
    componentsNotified = mediator.BroadcastMessage("SystemStatus", eventData)
    
    logger.LogInfo "Событие получено " & componentsNotified & " компонентами", "MediatorTestModule"
    
    ' Отправка целевого сообщения конкретному компоненту
    logger.LogInfo "Отправка целевого сообщения компоненту CommandManager", "MediatorTestModule"
    
    Dim targetData As Object
    Set targetData = CreateObject("Scripting.Dictionary")
    targetData.Add "Action", "ClearHistory"
    targetData.Add "Source", "MediatorTestModule"
    
    Dim messageReceived As Boolean
    messageReceived = mediator.SendMessage("CommandManager", "CommandAction", targetData)
    
    If messageReceived Then
        logger.LogInfo "Сообщение успешно обработано компонентом CommandManager", "MediatorTestModule"
    Else
        logger.LogWarning "Сообщение не обработано компонентом CommandManager", "MediatorTestModule"
    End If
    
    logger.LogInfo "Тест коммуникации между компонентами через Mediator завершен", "MediatorTestModule"
End Sub

' Тестирование динамического добавления компонентов
Private Sub TestDynamicComponentAddition()
    Dim logger As Object
    Set logger = GetLogger
    
    logger.LogInfo "=== ТЕСТ 5: Динамическое добавление компонентов ===", "MediatorTestModule"
    
    ' Получаем медиатор
    Dim mediator As CommandMediator
    Set mediator = GetCommandMediator
    
    ' Создаем новый компонент для тестирования
    logger.LogInfo "Создание нового тестового компонента", "MediatorTestModule"
    
    Dim testComponent As New testComponent
    
    ' Регистрируем компонент в медиаторе
    logger.LogInfo "Регистрация нового компонента в медиаторе", "MediatorTestModule"
    mediator.RegisterComponent testComponent
    
    ' Проверяем, что компонент зарегистрирован
    Dim registeredComponent As Object
    Set registeredComponent = GetComponentByID(mediator, "TestComponent")
    
    If Not registeredComponent Is Nothing Then
        logger.LogInfo "Компонент TestComponent успешно зарегистрирован", "MediatorTestModule"
        
        ' Отправляем сообщение новому компоненту
        logger.LogInfo "Отправка сообщения новому компоненту", "MediatorTestModule"
        
        Dim testData As Object
        Set testData = CreateObject("Scripting.Dictionary")
        testData.Add "Message", "Hello TestComponent!"
        testData.Add "Source", "MediatorTestModule"
        
        Dim messageReceived As Boolean
        messageReceived = mediator.SendMessage("TestComponent", "TestMessage", testData)
        
        If messageReceived Then
            logger.LogInfo "Сообщение успешно обработано компонентом TestComponent", "MediatorTestModule"
        Else
            logger.LogWarning "Сообщение не обработано компонентом TestComponent", "MediatorTestModule"
        End If
        
        ' Удаляем компонент из медиатора
        logger.LogInfo "Удаление компонента из медиатора", "MediatorTestModule"
        mediator.UnregisterComponent "TestComponent"
        
        ' Проверяем, что компонент удален
        Set registeredComponent = GetComponentByID(mediator, "TestComponent")
        
        If registeredComponent Is Nothing Then
            logger.LogInfo "Компонент TestComponent успешно удален", "MediatorTestModule"
        Else
            logger.LogWarning "Компонент TestComponent не был удален", "MediatorTestModule"
        End If
    Else
        logger.LogError "Не удалось зарегистрировать компонент TestComponent", "MediatorTestModule"
    End If
    
    logger.LogInfo "Тест динамического добавления компонентов завершен", "MediatorTestModule"
End Sub

' Вспомогательная функция для получения компонента по идентификатору
Private Function GetComponentByID(mediator As CommandMediator, componentID As String) As Object
    If mediator Is Nothing Then
        Debug.Print "GetComponentByID: mediator is Nothing"
        Set GetComponentByID = Nothing
        Exit Function
    End If
    
    If mediator.Components Is Nothing Then
        Debug.Print "GetComponentByID: mediator.Components is Nothing"
        Set GetComponentByID = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    If mediator.Components.exists(componentID) Then
        Set GetComponentByID = mediator.Components(componentID)
    Else
        Set GetComponentByID = Nothing
    End If
    
    If err.Number <> 0 Then
        Debug.Print "GetComponentByID error: " & err.description
        Set GetComponentByID = Nothing
    End If
    On Error GoTo 0
End Function

' Класс TestComponent для демонстрации динамического добавления компонентов
' (Требуется создание отдельного файла класса TestComponent.cls)
' @Code
'
' ' Класс TestComponent
' ' Тестовый компонент для демонстрации работы с медиатором
' Option Explicit
'
' Implements IComponent
'
' ' Внутренняя структура для хранения данных
' Private Type TTestComponent
'     Mediator As CommandMediator
'     Logger As Object
'     MessageCount As Long
' End Type
'
' Private this As TTestComponent
'
' ' Инициализация
' Private Sub Class_Initialize()
'     Set this.Logger = GetLogger
'     this.MessageCount = 0
' End Sub
'
' ' Реализация интерфейса IComponent
' Private Sub IComponent_SetMediator(mediator As CommandMediator)
'     Set this.Mediator = mediator
' End Sub
'
' Private Function IComponent_GetComponentID() As String
'     IComponent_GetComponentID = "TestComponent"
' End Function
'
' Private Function IComponent_ProcessMessage(messageType As String, data As Variant) As Boolean
'     ' Проверка наличия медиатора
'     If this.Mediator Is Nothing Then
'         this.Logger.LogError "Медиатор не установлен", "TestComponent"
'         IComponent_ProcessMessage = False
'         Exit Function
'     End If
'
'     ' Увеличиваем счетчик сообщений
'     this.MessageCount = this.MessageCount + 1
'
'     ' Обработка различных типов сообщений
'     Select Case messageType
'         Case "TestMessage"
'             ' Обработка тестового сообщения
'             IComponent_ProcessMessage = HandleTestMessage(data)
'
'         Case "SystemStatus"
'             ' Обработка статуса системы
'             IComponent_ProcessMessage = HandleSystemStatus(data)
'
'         Case "SystemInitialized"
'             ' Обработка инициализации системы
'             IComponent_ProcessMessage = HandleSystemInitialized(data)
'
'         Case Else
'             ' Неизвестный тип сообщения
'             this.Logger.LogDebug "Получено неизвестное сообщение: " & messageType, "TestComponent"
'             IComponent_ProcessMessage = False
'     End Select
' End Function
'
' ' Методы обработки сообщений
' Private Function HandleTestMessage(data As Variant) As Boolean
'     ' Проверка формата данных
'     If Not TypeOf data Is Object Then
'         this.Logger.LogWarning "Неверный формат данных для события TestMessage", "TestComponent"
'         HandleTestMessage = False
'         Exit Function
'     End If
'
'     ' Проверка наличия необходимых ключей
'     If Not data.Exists("Message") Then
'         this.Logger.LogWarning "Отсутствуют необходимые данные для события TestMessage", "TestComponent"
'         HandleTestMessage = False
'         Exit Function
'     End If
'
'     ' Получение данных
'     Dim message As String
'     message = data("Message")
'
'     ' Логирование события
'     this.Logger.LogInfo "Получено тестовое сообщение: " & message, "TestComponent"
'
'     ' Отправка ответа через медиатор
'     Dim responseData As Object
'     Set responseData = CreateObject("Scripting.Dictionary")
'     responseData.Add "Response", "Сообщение получено и обработано!"
'     responseData.Add "MessageCount", this.MessageCount
'     responseData.Add "Source", "TestComponent"
'
'     If data.Exists("Source") Then
'         this.Mediator.SendMessage data("Source"), "TestResponse", responseData
'     End If
'
'     HandleTestMessage = True
' End Function
'
' Private Function HandleSystemStatus(data As Variant) As Boolean
'     ' Проверка формата данных
'     If Not TypeOf data Is Object Then
'         this.Logger.LogWarning "Неверный формат данных для события SystemStatus", "TestComponent"
'         HandleSystemStatus = False
'         Exit Function
'     End If
'
'     ' Проверка наличия необходимых ключей
'     If Not data.Exists("Status") Then
'         this.Logger.LogWarning "Отсутствуют необходимые данные для события SystemStatus", "TestComponent"
'         HandleSystemStatus = False
'         Exit Function
'     End If
'
'     ' Получение данных
'     Dim status As String
'     status = data("Status")
'
'     ' Логирование события
'     this.Logger.LogInfo "Получен статус системы: " & status, "TestComponent"
'
'     HandleSystemStatus = True
' End Function
'
' Private Function HandleSystemInitialized(data As Variant) As Boolean
'     ' Логирование события
'     this.Logger.LogInfo "Система инициализирована", "TestComponent"
'
'     HandleSystemInitialized = True
' End Function
'
' ' Публичные методы
' Public Property Get MessageCount() As Long
'     MessageCount = this.MessageCount
' End Property
'
' ' @EndCode

