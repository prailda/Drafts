VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' �����: CommandFactory.cls
' ��������: ������� ��� �������� ������ ��������� �����
Option Explicit

' �������� ������� �� ����� ����
Public Function CreateCommand(ByVal commandType As String) As ICommand
    On Error Resume Next
    
    Select Case commandType
        ' ������� �����������
        Case "LogInfoCommand"
            Dim logInfoCmd As New LogInfoCommand
            Set CreateCommand = logInfoCmd
            
        Case "LogErrorCommand"
            Dim logErrorCmd As New LogErrorCommand
            Set CreateCommand = logErrorCmd
            
        ' ������� ��� ������ � �������
        Case "AddRecordCommand"
            Dim addCmd As New AddRecordCommand
            Set CreateCommand = addCmd
            
        Case "DeleteRecordCommand"
            Dim delCmd As New DeleteRecordCommand
            Set CreateCommand = delCmd
            
        ' ������� ��� ����������
        Case "ShowMessageCommand"
            Dim msgCmd As New ShowMessageCommand
            Set CreateCommand = msgCmd
            
        Case "RequestInputCommand"
            Dim inputCmd As New RequestInputCommand
            Set CreateCommand = inputCmd
            
        ' ������������ � ����������
        Case "MacroCommand"
            Dim macroCmd As New MacroCommand
            Set CreateCommand = macroCmd
            
        Case "LoggingDecorator"
            Dim logDec As New LoggingDecorator
            Set CreateCommand = logDec
            
        Case Else
            Debug.Print "ERROR: Unknown command type: " & commandType
            Set CreateCommand = Nothing
    End Select
    
    ' �������� �� ������ ��������
    If Err.Number <> 0 Then
        Debug.Print "Error creating command " & commandType & ": " & Err.Description
        Err.Clear
        Set CreateCommand = Nothing
    End If
    
    On Error GoTo 0
End Function

' �������� ������� � �������������
Public Function CreateLogInfoCommand(ByVal message As String, Optional ByVal source As String = "") As ICommand
    Dim cmd As LogInfoCommand
    Set cmd = CreateCommand("LogInfoCommand")
    
    If Not cmd Is Nothing Then
        cmd.Initialize message, source
    End If
    
    Set CreateLogInfoCommand = cmd
End Function

' �������� ������� ���������� ������ � �������������
Public Function CreateAddRecordCommand(ByVal data As String, ByVal tableName As String) As ICommand
    Dim cmd As AddRecordCommand
    Set cmd = CreateCommand("AddRecordCommand")
    
    If Not cmd Is Nothing Then
        cmd.Initialize data, tableName
    End If
    
    Set CreateAddRecordCommand = cmd
End Function

' �������� �������������� ������� �����������
Public Function CreateLoggingDecorator(ByVal command As ICommand, _
                                     Optional ByVal logLevel As String = "INFO") As ICommand
    Dim decorator As LoggingDecorator
    Set decorator = CreateCommand("LoggingDecorator")
    
    If Not decorator Is Nothing Then
        decorator.Initialize command, logLevel
    End If
    
    Set CreateLoggingDecorator = decorator
End Function

' �������� ������������ � ������� ������
Public Function CreateMacroCommand(ByVal commands As Collection, _
                                 Optional ByVal Name As String = "MacroCommand") As ICommand
    Dim macroCmd As MacroCommand
    Set macroCmd = CreateCommand("MacroCommand")
    
    If Not macroCmd Is Nothing Then
        macroCmd.Name = Name
        
        Dim cmd As ICommand
        For Each cmd In commands
            macroCmd.AddCommand cmd
        Next cmd
    End If
    
    Set CreateMacroCommand = macroCmd
End Function

