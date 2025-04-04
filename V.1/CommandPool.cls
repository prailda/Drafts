VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CommandPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �����: CommandPool.cls
' ��������: ��� ��� ���������� ���������-���������
Option Explicit

' ��� ��������� ��������
Private m_AvailableCommands As Collection
' ��� ������������ ��������
Private m_InUseCommands As Collection
Private m_Initialized As Boolean ' �������� � ������ private ����������

' ��������� ����
Private m_PoolName As String
Private m_MaxPoolSize As Long
Private m_CommandLifetime As Long
Private m_CommandFactory As CommandFactory

' �������������
Private Sub Class_Initialize()
    Set m_AvailableCommands = New Collection
    Set m_InUseCommands = New Collection
    m_MaxPoolSize = 10 ' �� ���������
    m_CommandLifetime = 30 ' 30 ����� �� ���������
    m_Initialized = False
End Sub

' ������� ��������
Private Sub Class_Terminate()
    ClearAllObjects
    Set m_AvailableCommands = Nothing
    Set m_InUseCommands = Nothing
End Sub

' ��������� ���������� ������� ������
Public Function GetCommandFactory() As CommandFactory
    ' ��������, ��� ��� ���������� ����������������
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    If m_CommandFactory Is Nothing Then
        Set m_CommandFactory = New CommandFactory
    End If
    
    Set GetCommandFactory = m_CommandFactory
End Function

' �������� ����� InitializeAllPools
Public Sub InitializeAllPools()
    If m_ModuleInitialized Then Exit Sub
    
    ' �������������� ���� ������
    Set m_LogCommandPool = New CommandPool
    m_LogCommandPool.Initialize "LogCommands", 10, 30
    
    Set m_DataCommandPool = New CommandPool
    m_DataCommandPool.Initialize "DataCommands", 15, 30
    
    Set m_UICommandPool = New CommandPool
    m_UICommandPool.Initialize "UICommands", 8, 30
    
    ' �������������� ������� � �������
    Set m_CommandInvoker = New CommandInvoker
    Set m_CommandFactory = New CommandFactory
    
    m_ModuleInitialized = True
    
    ' ������� ���������� �� �������������
    Debug.Print "=== ��� ���������� ������� ���������������� ==="
End Sub

' � ������ Initialize ��������:
Public Sub Initialize(ByVal PoolName As String, Optional ByVal MaxPoolSize As Long = 10, _
                     Optional ByVal CommandLifetime As Long = 30)
    m_PoolName = PoolName
    m_MaxPoolSize = MaxPoolSize
    m_CommandLifetime = CommandLifetime
    m_Initialized = True ' ������������� ���� �������������
End Sub

' ��������� ����� ��������
Public Property Get IsInitialized() As Boolean
    IsInitialized = m_Initialized
End Property

' ��������� ������� �� ���� �� ����� �������
Public Function GetCommand(ByVal commandName As String) As ICommand
    If Not m_Initialized Then
        Debug.Print "ERROR: Command pool not initialized"
        Set GetCommand = Nothing
        Exit Function
    End If
    
    ' ���������, ���� �� ��������� ������� ������� ���� � ����
    Dim i As Long
    For i = 1 To m_AvailableCommands.Count
        Dim cmd As ICommand
        Set cmd = m_AvailableCommands(i)
        
        If cmd.Name = commandName Then
            ' ����� ������� ������� ����
            m_AvailableCommands.Remove i
            
            ' ���������� �������
            If TypeOf cmd Is PoolableObject Then
                Dim poolableCmd As PoolableObject
                Set poolableCmd = cmd
                poolableCmd.Activate
            End If
            
            ' ��������� � ������ ������������
            Dim key As String
            key = "CMD_" & commandName & "_" & Format(Now, "yyyymmddhhnnss") & "_" & m_InUseCommands.Count + 1
            m_InUseCommands.Add cmd, key
            
            Set GetCommand = cmd
            Exit Function
        End If
    Next i
    
    ' �������� ����� ������� - ������������ ������

    Set cmd = CreateCommand(commandName)
    
    ' ���� ������� �������, ��������� � ������������
    If Not cmd Is Nothing Then
        Dim newKey As String
        newKey = "CMD_" & commandName & "_" & Format(Now, "yyyymmddhhnnss") & "_" & m_InUseCommands.Count + 1
        m_InUseCommands.Add cmd, newKey
    End If
    
    Set GetCommand = cmd
End Function

' �������� ����� ������� �� �����
' ������������ ������ ������ ��� �������� ������
' � ������ CommandPool.cls - ����������� ������ �������� ������

Private Function CreateCommand(ByVal commandName As String) As ICommand
    On Error Resume Next ' ��������� ��������� ������
    
    ' ��������� ����� ��� �������� ������ ������ �����
    Select Case commandName
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
            ' ��������� ����� �������� �������� ���� �������
            Dim inputCmd As New RequestInputCommand
            ' ���������, ��� ������ ������
            If Err.Number <> 0 Then
                Debug.Print "Error creating RequestInputCommand: " & Err.Description
                Err.Clear
                Set CreateCommand = Nothing
                Exit Function
            End If
            Set CreateCommand = inputCmd
            
        Case Else
            Debug.Print "ERROR: Unknown command type: " & commandName
            Set CreateCommand = Nothing
    End Select
    
    ' ���������, ��� ������� ������� �������
    If Err.Number <> 0 Then
        Debug.Print "Error creating command " & commandName & ": " & Err.Description
        Err.Clear
        Set CreateCommand = Nothing
    End If
    
    On Error GoTo 0
End Function

' ������� ������� � ���
Public Sub ReturnCommand(ByRef command As ICommand)
    If command Is Nothing Then Exit Sub
    
    ' ���� ������� � ������ ������������
    Dim i As Long
    Dim found As Boolean
    found = False
    
    For i = 1 To m_InUseCommands.Count
        If m_InUseCommands(i) Is command Then
            found = True
            Exit For
        End If
    Next i
    
    ' ���� ������� ������� � ������ ������������
    If found Then
        ' ������� �� ������ ������������
        m_InUseCommands.Remove i
        
        ' ������������ � ���������� �������
        If TypeOf command Is PoolableObject Then
            Dim poolableCmd As PoolableObject
            Set poolableCmd = command
            poolableCmd.Deactivate
            poolableCmd.Reset
        End If
        
        ' ���������, �� �������� �� ������������ ������ ����
        If m_AvailableCommands.Count < m_MaxPoolSize Then
            ' ��������� ������� � ��� ���������
            m_AvailableCommands.Add command
        End If
    End If
    
    ' ������� ������ �� �������
    Set command = Nothing
End Sub

' ������� ���� ������ � �����
' �����: CommandPool.cls
' ����� �������, ������� ����� ���������

Public Sub ClearAllObjects()
    On Error Resume Next ' ��������� ��������� ������
    
    Dim i As Long
    
    ' ������� ��������� ������
    For i = m_AvailableCommands.Count To 1 Step -1
        Set m_AvailableCommands(i) = Nothing
        m_AvailableCommands.Remove i
    Next i
    
    ' ������� ������������ ������
    For i = m_InUseCommands.Count To 1 Step -1
        Set m_InUseCommands(i) = Nothing
        m_InUseCommands.Remove i
    Next i
    
    On Error GoTo 0
End Sub

' ��������� ���������� ����
Public Function GetStatistics() As String
    Dim stats As String
    stats = "===== " & m_PoolName & " Pool Statistics =====" & vbCrLf
    stats = stats & "Available commands: " & m_AvailableCommands.Count & vbCrLf
    stats = stats & "In-use commands: " & m_InUseCommands.Count & vbCrLf
    stats = stats & "Max pool size: " & m_MaxPoolSize & vbCrLf
    stats = stats & "Command lifetime: " & m_CommandLifetime & " minutes" & vbCrLf
    stats = stats & "=================================="
    
    GetStatistics = stats
End Function

' �����: CommandPool.cls
' ���������� ��������� ������� ��� ��������� ����������

' ���������� ��������� �������� � ����
Public Property Get AvailableObjectCount() As Long
    AvailableObjectCount = m_AvailableCommands.Count
End Property

' ���������� ������������ ��������
Public Property Get InUseObjectCount() As Long
    InUseObjectCount = m_InUseCommands.Count
End Property

' ������������ ������ ����
Public Property Get MaxPoolSize() As Long
    MaxPoolSize = m_MaxPoolSize
End Property

' ����� ����� �������
Public Property Get CommandLifetime() As Long
    CommandLifetime = m_CommandLifetime
End Property

' ��� ����
Public Property Get PoolName() As String
    PoolName = m_PoolName
End Property
