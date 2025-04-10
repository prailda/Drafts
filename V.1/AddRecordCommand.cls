VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddRecordCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �����: AddRecordCommand.cls
' ��������: ������� ��� ���������� ������ � ���� ������
Option Explicit

Implements ICommand
Implements PoolableObject

' ��������� �������
Private m_Data As String
Private m_TableName As String
Private m_RecordID As String ' ��� ������
Private m_ID As String
Private m_IsActive As Boolean

' �������������
Private Sub Class_Initialize()
    m_ID = "ADDREC_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
    m_IsActive = True
End Sub

' ��������� �������
Public Sub Initialize(ByVal data As String, ByVal tableName As String)
    m_Data = data
    m_TableName = tableName
    m_RecordID = "REC_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
End Sub

' ICommand.Execute
Private Sub ICommand_Execute()
    If Not m_IsActive Then Exit Sub
    
    ' ���������� ������ (���������)
    Debug.Print "Adding record '" & m_Data & "' to table '" & m_TableName & "' with ID: " & m_RecordID
    
    ' � �������� ������� ����� ��� �� ��� ���������� � ��
End Sub

' ICommand.Undo
Private Sub ICommand_Undo()
    If Not m_IsActive Then Exit Sub
    
    ' ������ ���������� (��������� ��������)
    Debug.Print "Undoing add record: Deleting record '" & m_RecordID & "' from table '" & m_TableName & "'"
    
    ' � �������� ������� ����� ��� �� ��� �������� �� ��
End Sub

' ICommand.Name
Private Property Get ICommand_Name() As String
    ICommand_Name = "AddRecordCommand"
End Property

' ICommand.SupportsUndo
Private Property Get ICommand_SupportsUndo() As Boolean
    ICommand_SupportsUndo = True ' ������������ ������
End Property

' PoolableObject.Reset
Private Sub PoolableObject_Reset()
    m_Data = ""
    m_TableName = ""
    m_RecordID = ""
    m_IsActive = True
End Sub

' PoolableObject.Activate
Private Sub PoolableObject_Activate()
    m_IsActive = True
End Sub

' PoolableObject.Deactivate
Private Sub PoolableObject_Deactivate()
    m_IsActive = False
End Sub

' PoolableObject.ObjectID
Private Property Get PoolableObject_ObjectID() As String
    PoolableObject_ObjectID = m_ID
End Property
