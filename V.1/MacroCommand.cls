VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MacroCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' �����: MacroCommand.cls
' ��������: ���������� ������������ (����������� ������� Command)
Option Explicit

Implements ICommand
Implements PoolableObject

' ��������� ������ � ������������
Private m_Commands As Collection
Private m_Name As String
Private m_ID As String
Private m_IsActive As Boolean

' �������������
Private Sub Class_Initialize()
    Set m_Commands = New Collection
    m_Name = "MacroCommand"
    m_ID = "MACRO_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
    m_IsActive = True
End Sub

' �������
Private Sub Class_Terminate()
    Set m_Commands = Nothing
End Sub

' ��������� ����� ������������
Public Property Let Name(ByVal value As String)
    m_Name = value
End Property

' ��������� ����� ������������
Public Property Get Name() As String
    Name = m_Name
End Property

' ���������� ������� � ������������
Public Sub AddCommand(ByVal command As ICommand)
    m_Commands.Add command
End Sub

' �������� ���� ������
Public Sub ClearCommands()
    Set m_Commands = New Collection
End Sub

' ��������� ���������� ������
Public Property Get CommandCount() As Long
    CommandCount = m_Commands.Count
End Property

' ���������� ICommand.Execute
Private Sub ICommand_Execute()
    If Not m_IsActive Then Exit Sub
    
    ' ���������� ���� ������ ���������������
    Dim cmd As ICommand
    Dim i As Long
    
    For i = 1 To m_Commands.Count
        Set cmd = m_Commands(i)
        cmd.Execute
    Next i
End Sub

' ���������� ICommand.Undo
Private Sub ICommand_Undo()
    If Not m_IsActive Then Exit Sub
    
    ' ������ ���� ������ � �������� �������
    Dim cmd As ICommand
    Dim i As Long
    
    For i = m_Commands.Count To 1 Step -1
        Set cmd = m_Commands(i)
        If cmd.SupportsUndo Then
            cmd.Undo
        End If
    Next i
End Sub

' ���������� ICommand.Name
Private Property Get ICommand_Name() As String
    ICommand_Name = m_Name
End Property

' ���������� ICommand.SupportsUndo
Private Property Get ICommand_SupportsUndo() As Boolean
    ' ������������ ������������ ������, ���� ���� �� ���� �� ������ ������������
    Dim cmd As ICommand
    Dim i As Long
    
    For i = 1 To m_Commands.Count
        Set cmd = m_Commands(i)
        If cmd.SupportsUndo Then
            ICommand_SupportsUndo = True
            Exit Property
        End If
    Next i
    
    ICommand_SupportsUndo = False
End Property

' ���������� PoolableObject.Reset
Private Sub PoolableObject_Reset()
    Set m_Commands = New Collection
    m_Name = "MacroCommand"
    m_IsActive = True
End Sub

' ���������� PoolableObject.Activate
Private Sub PoolableObject_Activate()
    m_IsActive = True
End Sub

' ���������� PoolableObject.Deactivate
Private Sub PoolableObject_Deactivate()
    m_IsActive = False
End Sub

' ���������� PoolableObject.ObjectID
Private Property Get PoolableObject_ObjectID() As String
    PoolableObject_ObjectID = m_ID
End Property
