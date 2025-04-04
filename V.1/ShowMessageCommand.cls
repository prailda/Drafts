VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShowMessageCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �����: ShowMessageCommand.cls
' ��������: ������� ��� ����������� ��������� ������������
Option Explicit

Implements ICommand
Implements PoolableObject

' ��������� �������
Private m_Message As String
Private m_Title As String
Private m_ID As String
Private m_IsActive As Boolean

' �������������
Private Sub Class_Initialize()
    m_ID = "SHOWMSG_" & Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd() * 10000)
    m_IsActive = True
End Sub

' ��������� �������
Public Sub Initialize(ByVal message As String, Optional ByVal title As String = "����������")
    m_Message = message
    m_Title = title
End Sub

' ICommand.Execute
Private Sub ICommand_Execute()
    If Not m_IsActive Then Exit Sub
    
    ' ����������� ���������
    MsgBox m_Message, vbInformation, m_Title
    
    ' ������ � ���
    Debug.Print "Displayed message: " & m_Title & " - " & m_Message
End Sub

' ICommand.Undo - ����������� ��������� �� ������������ ������
Private Sub ICommand_Undo()
    Debug.Print "Cannot undo showing message to user"
End Sub

' ICommand.Name
Private Property Get ICommand_Name() As String
    ICommand_Name = "ShowMessageCommand"
End Property

' ICommand.SupportsUndo
Private Property Get ICommand_SupportsUndo() As Boolean
    ICommand_SupportsUndo = False ' �� ������������ ������
End Property

' PoolableObject.Reset
Private Sub PoolableObject_Reset()
    m_Message = ""
    m_Title = "����������"
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
