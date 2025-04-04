VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DeleteRecordCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �����: DeleteRecordCommand.cls
' ��������: ������� ��� �������� ������ �� ���� ������
Option Explicit

Implements ICommand
Implements PoolableObject

' ��������� �������
Private m_RecordID As String
Private m_TableName As String
Private m_BackupData As String ' ��� ������ ��������
Private m_ID As String
Private m_IsActive As Boolean

' �������������
Private Sub Class_Initialize()
    m_ID = "DELREC_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
    m_IsActive = True
End Sub

' ��������� �������
Public Sub Initialize(ByVal recordID As String, ByVal tableName As String)
    m_RecordID = recordID
    m_TableName = tableName
    ' � �������� ���������� ����� ����� ��������� ������ ������ ��� ������
    m_BackupData = "BACKUP_DATA_" & recordID
End Sub

' ICommand.Execute
Private Sub ICommand_Execute()
    If Not m_IsActive Then Exit Sub
    
    ' �������� ������ (���������)
    Debug.Print "Deleting record '" & m_RecordID & "' from table '" & m_TableName & "'"
    
    ' � �������� ������� ����� ��� �� ��� �������� �� ��
End Sub

' ICommand.Undo
Private Sub ICommand_Undo()
    If Not m_IsActive Then Exit Sub
    
    ' ������ �������� (�������������� �� ��������� �����)
    Debug.Print "Undoing delete record: Restoring record '" & m_RecordID & "' to table '" & m_TableName & "'"
    Debug.Print "Restored data: " & m_BackupData
    
    ' � �������� ������� ����� ��� �� ��� �������������� ������
End Sub

' ICommand.Name
Private Property Get ICommand_Name() As String
    ICommand_Name = "DeleteRecordCommand"
End Property

' ICommand.SupportsUndo
Private Property Get ICommand_SupportsUndo() As Boolean
    ICommand_SupportsUndo = True ' ������������ ������
End Property

' PoolableObject.Reset
Private Sub PoolableObject_Reset()
    m_RecordID = ""
    m_TableName = ""
    m_BackupData = ""
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
