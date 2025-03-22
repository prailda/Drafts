Attribute VB_Name = "PatternDemoAdvanced"
' Создать новый модуль PatternDemoAdvanced.bas
' Демонстрация интеграции паттернов
Option Explicit

' Заголовок для сообщений
Private Const DEMO_TITLE As String = "Интеграция паттернов проектирования"

' Запуск полной демонстрации всех паттернов
Public Sub RunAdvancedPatternDemo()
    ' Инициализация всех компонентов
    InitializeAllPools
    
    ' Приветственное сообщение
    MsgBox "Добро пожаловать в демонстрацию интеграции паттернов проектирования!" & vbCrLf & vbCrLf & _
           "Будут продемонстрированы:" & vbCrLf & _
           "1. Command и Object Pool - базовые паттерны" & vbCrLf & _
           "2. Decorator - для расширения поведения команд" & vbCrLf & _
           "3. Factory - для создания объектов команд" & vbCrLf & _
           "4. Builder - для пошагового создания макрокоманд" & vbCrLf & vbCrLf & _
           "Нажмите OK для начала демонстрации.", _
           vbInformation, DEMO_TITLE

    ' 1. Демонстрация фабрики команд
    DemoCommandFactory
    
    ' 2. Демонстрация декораторов
    DemoDecorators
    
    ' 3. Демонстрация строителя макрокоманд
    DemoCommandBuilder
    
    ' 4. Демонстрация комбинирования всех паттернов
    DemoCombinedPatterns
    
    ' Завершение
    MsgBox "Демонстрация завершена! Все паттерны были показаны в действии." & vbCrLf & _
           "Проверьте окно Immediate (Ctrl+G) для просмотра подробных результатов.", _
           vbInformation, DEMO_TITLE
           
    ' Очистка ресурсов
    ReleaseAllPools
End Sub

' Демонстрация использования фабрики команд
Private Sub DemoCommandFactory()
    Debug.Print "=== ДЕМОНСТРАЦИЯ ФАБРИКИ КОМАНД ==="
    
    ' Получаем фабрику команд
    Dim factory As CommandFactory
    Set factory = GetCommandFactory()
    
    ' Получаем инвокер
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    ' Создаем команды через фабрику
    Dim logCmd As ICommand
    Set logCmd = factory.CreateLogInfoCommand("Команда создана через фабрику", "Factory Demo")
    
    ' Выполняем команду
    invoker.ExecuteCommand logCmd
    
    ' Создаем команду добавления записи
    Dim addCmd As ICommand
    Set addCmd = factory.CreateAddRecordCommand("Фабричные данные", "FactoryTable")
    
    ' Выполняем команду
    invoker.ExecuteCommand addCmd
    
    MsgBox "Демонстрация фабрики команд завершена!" & vbCrLf & _
           "Команды были созданы и выполнены через фабрику.", _
           vbInformation, DEMO_TITLE & " - Фабрика"
    
    Debug.Print "=== ЗАВЕРШЕНИЕ ДЕМОНСТРАЦИИ ФАБРИКИ ==="
End Sub

' Демонстрация использования декораторов
Private Sub DemoDecorators()
    Debug.Print "=== ДЕМОНСТРАЦИЯ ДЕКОРАТОРОВ ==="
    
    ' Получаем фабрику и инвокер
    Dim factory As CommandFactory
    Set factory = GetCommandFactory()
    
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    ' Создаем базовую команду
    Dim baseCmd As ICommand
    Set baseCmd = factory.CreateLogInfoCommand("Базовая команда логирования", "Decorator Demo")
    
    ' Создаем декорированную версию команды
    Dim decoratedCmd As ICommand
    Set decoratedCmd = factory.CreateLoggingDecorator(baseCmd, "DEBUG")
    
    ' Выполняем декорированную команду
    MsgBox "Сейчас будет выполнена декорированная команда логирования." & vbCrLf & _
           "Обратите внимание на дополнительные сообщения в логе.", _
           vbInformation, DEMO_TITLE & " - Декораторы"
    
    invoker.ExecuteCommand decoratedCmd
    
    ' Создаем команду с данными и декорируем ее
    Dim dataCmd As ICommand
    Set dataCmd = factory.CreateAddRecordCommand("Декорированные данные", "DecoratorTable")
    
    Dim decoratedDataCmd As ICommand
    Set decoratedDataCmd = factory.CreateLoggingDecorator(dataCmd, "INFO")
    
    ' Выполняем декорированную команду с данными
    invoker.ExecuteCommand decoratedDataCmd
    
    MsgBox "Демонстрация декораторов завершена!" & vbCrLf & _
           "Команды были обернуты декораторами, добавляющими логирование.", _
           vbInformation, DEMO_TITLE & " - Декораторы"
    
    Debug.Print "=== ЗАВЕРШЕНИЕ ДЕМОНСТРАЦИИ ДЕКОРАТОРОВ ==="
End Sub

' Демонстрация использования строителя макрокоманд
Private Sub DemoCommandBuilder()
    Debug.Print "=== ДЕМОНСТРАЦИЯ СТРОИТЕЛЯ МАКРОКОМАНД ==="
    
    ' Получаем инвокер
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    ' Создаем строитель макрокоманд
    Dim builder As New MacroCommandBuilder
    
    ' Строим макрокоманду
    MsgBox "Сейчас будет создана и выполнена макрокоманда с помощью строителя.", _
           vbInformation, DEMO_TITLE & " - Строитель"
    
    Dim macro As MacroCommand
    Set macro = builder.WithName("Процесс обработки данных") _
                       .AddLogCommand("Начало процесса обработки", "Builder Demo") _
                       .AddRecordOperation("Новая запись от строителя", "BuilderTable") _
                       .AddShowMessageOperation("Запись успешно добавлена!", "Результат операции") _
                       .AddLogCommand("Завершение процесса обработки", "Builder Demo") _
                       .Build()
    
    ' Выполняем макрокоманду
    invoker.ExecuteCommand macro
    
    ' Показываем информацию о созданной макрокоманде
' Должно быть:
' Простейшее решение - просто убрать ссылку на имя:
MsgBox "Макрокоманда успешно создана и выполнена!" & vbCrLf & _
       "Количество команд в макрокоманде: " & macro.CommandCount, _
       vbInformation, DEMO_TITLE & " - Строитель"
    
    Debug.Print "=== ЗАВЕРШЕНИЕ ДЕМОНСТРАЦИИ СТРОИТЕЛЯ ==="
End Sub

' Демонстрация комбинирования всех паттернов вместе
Private Sub DemoCombinedPatterns()
    Debug.Print "=== ДЕМОНСТРАЦИЯ КОМБИНАЦИИ ВСЕХ ПАТТЕРНОВ ==="
    
    ' Получаем фабрику, инвокер и пулы
    Dim factory As CommandFactory
    Set factory = GetCommandFactory()
    
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    Dim logPool As CommandPool
    Set logPool = GetLogCommandPool()
    
    ' 1. Получаем команду из пула (Object Pool)
    Dim logCmd As ICommand
    Set logCmd = logPool.GetCommand("LogInfoCommand")
    
    ' 2. Настраиваем команду
    Dim logInfoCmd As LogInfoCommand
    Set logInfoCmd = logCmd
    logInfoCmd.Initialize "Комбинированная демонстрация паттернов", "Combined Demo"
    
    ' 3. Создаем декоратор с помощью фабрики (Factory + Decorator)
    Dim decoratedLogCmd As ICommand
    Set decoratedLogCmd = factory.CreateLoggingDecorator(logCmd)
    
    ' 4. Создаем строитель макрокоманд (Builder)
    Dim builder As New MacroCommandBuilder
    
    ' 5. Строим сложную макрокоманду
    MsgBox "Сейчас будет создана и выполнена сложная макрокоманда," & vbCrLf & _
           "использующая все изученные паттерны проектирования.", _
           vbInformation, DEMO_TITLE & " - Комбинация паттернов"
    
    ' Создаем команду для добавления записи с помощью фабрики
    Dim addCmd As ICommand
    Set addCmd = factory.CreateAddRecordCommand("Интегрированные данные", "MasterTable")
    
    ' Создаем декорированную версию команды добавления
    Dim decoratedAddCmd As ICommand
    Set decoratedAddCmd = factory.CreateLoggingDecorator(addCmd)
    
    ' Строим макрокоманду, включая обычные и декорированные команды
    Dim macro As MacroCommand
    Set macro = builder.WithName("Комплексная операция") _
                       .AddCommand(decoratedLogCmd) _
                       .AddShowMessageOperation("Начинаем комплексную операцию", "Процесс") _
                       .AddCommand(decoratedAddCmd) _
                       .AddLogCommand("Проверка результатов операции", "Combined Demo") _
                       .AddShowMessageOperation("Комплексная операция успешно завершена!", "Результат") _
                       .Build()
    
    ' Выполняем макрокоманду
    invoker.ExecuteCommand macro
    
    ' 6. Возвращаем команду логирования в пул (Object Pool)
    logPool.ReturnCommand logCmd
    
    ' Отображаем статистику пулов
    Debug.Print GetPoolsStatistics()
    
    MsgBox "Демонстрация комбинации паттернов завершена!" & vbCrLf & vbCrLf & _
           "Использованные паттерны:" & vbCrLf & _
           "• Command - для инкапсуляции операций" & vbCrLf & _
           "• Object Pool - для управления ресурсами команд" & vbCrLf & _
           "• Decorator - для расширения функциональности команд" & vbCrLf & _
           "• Factory - для создания и инициализации команд" & vbCrLf & _
           "• Builder - для конструирования сложных макрокоманд", _
           vbInformation, DEMO_TITLE & " - Итоги"
    
    Debug.Print "=== ЗАВЕРШЕНИЕ КОМБИНИРОВАННОЙ ДЕМОНСТРАЦИИ ==="
End Sub

