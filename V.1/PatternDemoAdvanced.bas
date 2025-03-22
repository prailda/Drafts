Attribute VB_Name = "PatternDemoAdvanced"
' ������� ����� ������ PatternDemoAdvanced.bas
' ������������ ���������� ���������
Option Explicit

' ��������� ��� ���������
Private Const DEMO_TITLE As String = "���������� ��������� ��������������"

' ������ ������ ������������ ���� ���������
Public Sub RunAdvancedPatternDemo()
    ' ������������� ���� �����������
    InitializeAllPools
    
    ' �������������� ���������
    MsgBox "����� ���������� � ������������ ���������� ��������� ��������������!" & vbCrLf & vbCrLf & _
           "����� ������������������:" & vbCrLf & _
           "1. Command � Object Pool - ������� ��������" & vbCrLf & _
           "2. Decorator - ��� ���������� ��������� ������" & vbCrLf & _
           "3. Factory - ��� �������� �������� ������" & vbCrLf & _
           "4. Builder - ��� ���������� �������� �����������" & vbCrLf & vbCrLf & _
           "������� OK ��� ������ ������������.", _
           vbInformation, DEMO_TITLE

    ' 1. ������������ ������� ������
    DemoCommandFactory
    
    ' 2. ������������ �����������
    DemoDecorators
    
    ' 3. ������������ ��������� �����������
    DemoCommandBuilder
    
    ' 4. ������������ �������������� ���� ���������
    DemoCombinedPatterns
    
    ' ����������
    MsgBox "������������ ���������! ��� �������� ���� �������� � ��������." & vbCrLf & _
           "��������� ���� Immediate (Ctrl+G) ��� ��������� ��������� �����������.", _
           vbInformation, DEMO_TITLE
           
    ' ������� ��������
    ReleaseAllPools
End Sub

' ������������ ������������� ������� ������
Private Sub DemoCommandFactory()
    Debug.Print "=== ������������ ������� ������ ==="
    
    ' �������� ������� ������
    Dim factory As CommandFactory
    Set factory = GetCommandFactory()
    
    ' �������� �������
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    ' ������� ������� ����� �������
    Dim logCmd As ICommand
    Set logCmd = factory.CreateLogInfoCommand("������� ������� ����� �������", "Factory Demo")
    
    ' ��������� �������
    invoker.ExecuteCommand logCmd
    
    ' ������� ������� ���������� ������
    Dim addCmd As ICommand
    Set addCmd = factory.CreateAddRecordCommand("��������� ������", "FactoryTable")
    
    ' ��������� �������
    invoker.ExecuteCommand addCmd
    
    MsgBox "������������ ������� ������ ���������!" & vbCrLf & _
           "������� ���� ������� � ��������� ����� �������.", _
           vbInformation, DEMO_TITLE & " - �������"
    
    Debug.Print "=== ���������� ������������ ������� ==="
End Sub

' ������������ ������������� �����������
Private Sub DemoDecorators()
    Debug.Print "=== ������������ ����������� ==="
    
    ' �������� ������� � �������
    Dim factory As CommandFactory
    Set factory = GetCommandFactory()
    
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    ' ������� ������� �������
    Dim baseCmd As ICommand
    Set baseCmd = factory.CreateLogInfoCommand("������� ������� �����������", "Decorator Demo")
    
    ' ������� �������������� ������ �������
    Dim decoratedCmd As ICommand
    Set decoratedCmd = factory.CreateLoggingDecorator(baseCmd, "DEBUG")
    
    ' ��������� �������������� �������
    MsgBox "������ ����� ��������� �������������� ������� �����������." & vbCrLf & _
           "�������� �������� �� �������������� ��������� � ����.", _
           vbInformation, DEMO_TITLE & " - ����������"
    
    invoker.ExecuteCommand decoratedCmd
    
    ' ������� ������� � ������� � ���������� ��
    Dim dataCmd As ICommand
    Set dataCmd = factory.CreateAddRecordCommand("�������������� ������", "DecoratorTable")
    
    Dim decoratedDataCmd As ICommand
    Set decoratedDataCmd = factory.CreateLoggingDecorator(dataCmd, "INFO")
    
    ' ��������� �������������� ������� � �������
    invoker.ExecuteCommand decoratedDataCmd
    
    MsgBox "������������ ����������� ���������!" & vbCrLf & _
           "������� ���� �������� ������������, ������������ �����������.", _
           vbInformation, DEMO_TITLE & " - ����������"
    
    Debug.Print "=== ���������� ������������ ����������� ==="
End Sub

' ������������ ������������� ��������� �����������
Private Sub DemoCommandBuilder()
    Debug.Print "=== ������������ ��������� ����������� ==="
    
    ' �������� �������
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    ' ������� ��������� �����������
    Dim builder As New MacroCommandBuilder
    
    ' ������ ������������
    MsgBox "������ ����� ������� � ��������� ������������ � ������� ���������.", _
           vbInformation, DEMO_TITLE & " - ���������"
    
    Dim macro As MacroCommand
    Set macro = builder.WithName("������� ��������� ������") _
                       .AddLogCommand("������ �������� ���������", "Builder Demo") _
                       .AddRecordOperation("����� ������ �� ���������", "BuilderTable") _
                       .AddShowMessageOperation("������ ������� ���������!", "��������� ��������") _
                       .AddLogCommand("���������� �������� ���������", "Builder Demo") _
                       .Build()
    
    ' ��������� ������������
    invoker.ExecuteCommand macro
    
    ' ���������� ���������� � ��������� ������������
' ������ ����:
' ���������� ������� - ������ ������ ������ �� ���:
MsgBox "������������ ������� ������� � ���������!" & vbCrLf & _
       "���������� ������ � ������������: " & macro.CommandCount, _
       vbInformation, DEMO_TITLE & " - ���������"
    
    Debug.Print "=== ���������� ������������ ��������� ==="
End Sub

' ������������ �������������� ���� ��������� ������
Private Sub DemoCombinedPatterns()
    Debug.Print "=== ������������ ���������� ���� ��������� ==="
    
    ' �������� �������, ������� � ����
    Dim factory As CommandFactory
    Set factory = GetCommandFactory()
    
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    Dim logPool As CommandPool
    Set logPool = GetLogCommandPool()
    
    ' 1. �������� ������� �� ���� (Object Pool)
    Dim logCmd As ICommand
    Set logCmd = logPool.GetCommand("LogInfoCommand")
    
    ' 2. ����������� �������
    Dim logInfoCmd As LogInfoCommand
    Set logInfoCmd = logCmd
    logInfoCmd.Initialize "��������������� ������������ ���������", "Combined Demo"
    
    ' 3. ������� ��������� � ������� ������� (Factory + Decorator)
    Dim decoratedLogCmd As ICommand
    Set decoratedLogCmd = factory.CreateLoggingDecorator(logCmd)
    
    ' 4. ������� ��������� ����������� (Builder)
    Dim builder As New MacroCommandBuilder
    
    ' 5. ������ ������� ������������
    MsgBox "������ ����� ������� � ��������� ������� ������������," & vbCrLf & _
           "������������ ��� ��������� �������� ��������������.", _
           vbInformation, DEMO_TITLE & " - ���������� ���������"
    
    ' ������� ������� ��� ���������� ������ � ������� �������
    Dim addCmd As ICommand
    Set addCmd = factory.CreateAddRecordCommand("��������������� ������", "MasterTable")
    
    ' ������� �������������� ������ ������� ����������
    Dim decoratedAddCmd As ICommand
    Set decoratedAddCmd = factory.CreateLoggingDecorator(addCmd)
    
    ' ������ ������������, ������� ������� � �������������� �������
    Dim macro As MacroCommand
    Set macro = builder.WithName("����������� ��������") _
                       .AddCommand(decoratedLogCmd) _
                       .AddShowMessageOperation("�������� ����������� ��������", "�������") _
                       .AddCommand(decoratedAddCmd) _
                       .AddLogCommand("�������� ����������� ��������", "Combined Demo") _
                       .AddShowMessageOperation("����������� �������� ������� ���������!", "���������") _
                       .Build()
    
    ' ��������� ������������
    invoker.ExecuteCommand macro
    
    ' 6. ���������� ������� ����������� � ��� (Object Pool)
    logPool.ReturnCommand logCmd
    
    ' ���������� ���������� �����
    Debug.Print GetPoolsStatistics()
    
    MsgBox "������������ ���������� ��������� ���������!" & vbCrLf & vbCrLf & _
           "�������������� ��������:" & vbCrLf & _
           "� Command - ��� ������������ ��������" & vbCrLf & _
           "� Object Pool - ��� ���������� ��������� ������" & vbCrLf & _
           "� Decorator - ��� ���������� ���������������� ������" & vbCrLf & _
           "� Factory - ��� �������� � ������������� ������" & vbCrLf & _
           "� Builder - ��� ��������������� ������� �����������", _
           vbInformation, DEMO_TITLE & " - �����"
    
    Debug.Print "=== ���������� ��������������� ������������ ==="
End Sub

