VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandInvoker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �����: CommandInvoker.cls
' ��������: ������� ������ - �����, ����������� ���������� ������
Option Explicit

' ������� ����������� ������ ��� ������
Private m_CommandHistory As Collection

' �������������
Private Sub Class_Initialize()
    Set m_CommandHistory = New Collection
End Sub

' �������
Private Sub Class_Terminate()
    Set m_CommandHistory = Nothing
End Sub

' ���������� �������
Public Function ExecuteCommand(ByVal command As ICommand) As Boolean
    On Error GoTo ErrorHandler
    
    ' ���������� �������
    command.Execute
    
    ' ���� ������� ������������ ������, ��������� � �������
    If command.SupportsUndo Then
        m_CommandHistory.Add command
    End If
    
    ExecuteCommand = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in ExecuteCommand: " & Err.Description
    ExecuteCommand = False
End Function

' ������ ��������� �������
Public Function UndoLastCommand() As Boolean
    On Error GoTo ErrorHandler
    
    ' ��������� ������� ������ � �������
    If m_CommandHistory.Count = 0 Then
        UndoLastCommand = False
        Exit Function
    End If
    
    ' �������� ��������� �������
    Dim command As ICommand
    Set command = m_CommandHistory(m_CommandHistory.Count)
    
    ' �������� ����������
    command.Undo
    
    ' ������� �� �������
    m_CommandHistory.Remove m_CommandHistory.Count
    
    UndoLastCommand = True
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in UndoLastCommand: " & Err.Description
    UndoLastCommand = False
End Function

' �������� ����������� ������
Public Property Get CanUndo() As Boolean
    CanUndo = (m_CommandHistory.Count > 0)
End Property

' ���������� ������ � �������
Public Property Get CommandHistoryCount() As Long
    CommandHistoryCount = m_CommandHistory.Count
End Property

' ������� ������� ������
Public Sub ClearHistory()
    Set m_CommandHistory = New Collection
End Sub
