VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacroCommandBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' �����: MacroCommandBuilder.cls
' ��������: ��������� ��� �������� �����������
Option Explicit

' ���������� ������������
Private m_MacroCommand As MacroCommand

' �������������
Private Sub Class_Initialize()
    Reset
End Sub

' ����� ���������
Public Sub Reset()
    Set m_MacroCommand = New MacroCommand
End Sub

' ��������� ����� ������������
Public Function WithName(ByVal Name As String) As MacroCommandBuilder
    m_MacroCommand.Name = Name
    Set WithName = Me
End Function

' ���������� ������� �����������
Public Function AddLogCommand(ByVal message As String, Optional ByVal source As String = "") As MacroCommandBuilder
    Dim cmd As New LogInfoCommand
    cmd.Initialize message, source
    m_MacroCommand.AddCommand cmd
    Set AddLogCommand = Me
End Function

' ���������� ������� ����������� ������
Public Function AddErrorLogCommand(ByVal message As String, Optional ByVal source As String = "", _
                                  Optional ByVal errorCode As Long = 0) As MacroCommandBuilder
    Dim cmd As New LogErrorCommand
    cmd.Initialize message, source, errorCode
    m_MacroCommand.AddCommand cmd
    Set AddErrorLogCommand = Me
End Function

' ���������� ������� ���������� ������
Public Function AddRecordOperation(ByVal data As String, ByVal tableName As String) As MacroCommandBuilder
    Dim cmd As New AddRecordCommand
    cmd.Initialize data, tableName
    m_MacroCommand.AddCommand cmd
    Set AddRecordOperation = Me
End Function

' ���������� ������� �������� ������
Public Function AddDeleteOperation(ByVal recordID As String, ByVal tableName As String) As MacroCommandBuilder
    Dim cmd As New DeleteRecordCommand
    cmd.Initialize recordID, tableName
    m_MacroCommand.AddCommand cmd
    Set AddDeleteOperation = Me
End Function

' ���������� ������� ����������� ���������
Public Function AddShowMessageOperation(ByVal message As String, Optional ByVal title As String = "����������") As MacroCommandBuilder
    Dim cmd As New ShowMessageCommand
    cmd.Initialize message, title
    m_MacroCommand.AddCommand cmd
    Set AddShowMessageOperation = Me
End Function

' ���������� ������������ �������
Public Function AddCommand(ByVal command As ICommand) As MacroCommandBuilder
    m_MacroCommand.AddCommand command
    Set AddCommand = Me
End Function

' �������� �������������� ������� �����������
Public Function AddLoggingDecorator(ByVal command As ICommand) As MacroCommandBuilder
    Dim decorator As New LoggingDecorator
    decorator.Initialize command
    m_MacroCommand.AddCommand decorator
    Set AddLoggingDecorator = Me
End Function

' ���������� ���������� � ��������� ����������
Public Function Build() As MacroCommand
    Dim result As MacroCommand
    Set result = m_MacroCommand
    Reset ' ����� ��� ���������� ����������
    Set Build = result
End Function

