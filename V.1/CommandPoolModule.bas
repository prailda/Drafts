Attribute VB_Name = "CommandPoolModule"
'' ������: CommandPoolModule.bas
'' ��������: ����� ������ ��� ���������� �������� Singleton ��� ����� ������
'Option Explicit
'
'' ���������� ����� ������
'Private m_LogCommandPool As CommandPool
'Private m_DataCommandPool As CommandPool
'Private m_UICommandPool As CommandPool
'
'' ��������� �������� ������
'Private m_CommandInvoker As CommandInvoker
'
'' ��������� ���������� ���� ������ ��� �����������
'Public Function GetLogCommandPool() As CommandPool
'    If m_LogCommandPool Is Nothing Then
'        Set m_LogCommandPool = New CommandPool
'        m_LogCommandPool.Initialize "LogCommands", 10, 30 ' ����. 10 ������, 30 ����� �����
'    End If
'    Set GetLogCommandPool = m_LogCommandPool
'End Function
'
'' ��������� ���������� ���� ������ ��� ������ � �������
'Public Function GetDataCommandPool() As CommandPool
'    If m_DataCommandPool Is Nothing Then
'        Set m_DataCommandPool = New CommandPool
'        m_DataCommandPool.Initialize "DataCommands", 15, 30 ' ����. 15 ������, 30 ����� �����
'    End If
'    Set GetDataCommandPool = m_DataCommandPool
'End Function
'
'' ��������� ���������� ���� ������ ��� ����������
'Public Function GetUICommandPool() As CommandPool
'    If m_UICommandPool Is Nothing Then
'        Set m_UICommandPool = New CommandPool
'        m_UICommandPool.Initialize "UICommands", 8, 30 ' ����. 8 ������, 30 ����� �����
'    End If
'    Set GetUICommandPool = m_UICommandPool
'End Function
'
'' ��������� ���������� �������� ������
'Public Function GetCommandInvoker() As CommandInvoker
'    If m_CommandInvoker Is Nothing Then
'        Set m_CommandInvoker = New CommandInvoker
'    End If
'    Set GetCommandInvoker = m_CommandInvoker
'End Function
'
'' ��������� ���������� ���� ����� ������
'Public Function GetPoolsStatistics() As String
'    Dim stats As String
'    stats = "======= Command Pool Statistics =======" & vbCrLf & vbCrLf
'
'    If Not m_LogCommandPool Is Nothing Then
'        stats = stats & m_LogCommandPool.GetStatistics() & vbCrLf & vbCrLf
'    End If
'
'    If Not m_DataCommandPool Is Nothing Then
'        stats = stats & m_DataCommandPool.GetStatistics() & vbCrLf & vbCrLf
'    End If
'
'    If Not m_UICommandPool Is Nothing Then
'        stats = stats & m_UICommandPool.GetStatistics() & vbCrLf & vbCrLf
'    End If
'
'    stats = stats & "======================================"
'
'    GetPoolsStatistics = stats
'End Function
'' �������� ���� ����� � ������ CommandPoolModule.bas
'Public Function CreateRequestInputCommand(ByVal prompt As String, _
'                                         Optional ByVal title As String = "���� ������", _
'                                         Optional ByVal defaultValue As String = "") As RequestInputCommand
'    Dim cmd As New RequestInputCommand
'    cmd.Initialize prompt, title, defaultValue
'    Set CreateRequestInputCommand = cmd
'End Function
'
'Public Sub ReleaseAllPools()
'    On Error Resume Next ' ��������� ��������� ������
'
'    If Not m_LogCommandPool Is Nothing Then
'        ' ���������, ��� ������ ��������������� ���������
'        If m_LogCommandPool.IsInitialized Then
'            m_LogCommandPool.ClearAllObjects
'        End If
'        Set m_LogCommandPool = Nothing
'    End If
'
'    If Not m_DataCommandPool Is Nothing Then
'        If m_DataCommandPool.IsInitialized Then
'            m_DataCommandPool.ClearAllObjects
'        End If
'        Set m_DataCommandPool = Nothing
'    End If
'
'    If Not m_UICommandPool Is Nothing Then
'        If m_UICommandPool.IsInitialized Then
'            m_UICommandPool.ClearAllObjects
'        End If
'        Set m_UICommandPool = Nothing
'    End If
'
'    Set m_CommandInvoker = Nothing
'
'    ' ���������, ���� �� ������
'    If Err.Number <> 0 Then
'        Debug.Print "Error in ReleaseAllPools: " & Err.Description
'        Err.Clear
'    End If
'
'    On Error GoTo 0
'End Sub
'


' ������: CommandPoolModule.bas
' ��������: ����� ������ ��� ���������� �������� Singleton ��� ����� ������
Option Explicit

' ���������� ����� ������
Private m_LogCommandPool As CommandPool
Private m_DataCommandPool As CommandPool
Private m_UICommandPool As CommandPool
Private m_CommandFactory As CommandFactory


' ��������� �������� ������
Private m_CommandInvoker As CommandInvoker

' ���� ������������� ������
Private m_ModuleInitialized As Boolean

' ������������� ���� �����
Public Sub InitializeAllPools()
    If m_ModuleInitialized Then Exit Sub
    
    ' �������������� ���� ������
    Set m_LogCommandPool = New CommandPool
    m_LogCommandPool.Initialize "LogCommands", 10, 30
    
    Set m_DataCommandPool = New CommandPool
    m_DataCommandPool.Initialize "DataCommands", 15, 30
    
    Set m_UICommandPool = New CommandPool
    m_UICommandPool.Initialize "UICommands", 8, 30
    
    ' �������������� �������
    Set m_CommandInvoker = New CommandInvoker
    
    m_ModuleInitialized = True
    
    ' ������� ���������� �� �������������
    Debug.Print "=== ��� ���� ������ ���������������� ==="
End Sub

' ��������� ���������� ���� ������ ��� �����������
Public Function GetLogCommandPool() As CommandPool
    ' ��������, ��� ��� ���� ����������������
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    Set GetLogCommandPool = m_LogCommandPool
End Function

' ��������� ���������� ���� ������ ��� ������ � �������
Public Function GetDataCommandPool() As CommandPool
    ' ��������, ��� ��� ���� ����������������
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    Set GetDataCommandPool = m_DataCommandPool
End Function

' ��������� ���������� ���� ������ ��� ����������
Public Function GetUICommandPool() As CommandPool
    ' ��������, ��� ��� ���� ����������������
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    Set GetUICommandPool = m_UICommandPool
End Function

' ��������� ���������� �������� ������
Public Function GetCommandInvoker() As CommandInvoker
    ' ��������, ��� ��� ���� ����������������
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    Set GetCommandInvoker = m_CommandInvoker
End Function

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




' ��������� ���������� ���� ����� ������
Public Function GetPoolsStatistics() As String
    ' ��������, ��� ��� ���� ����������������
    If Not m_ModuleInitialized Then
        InitializeAllPools
    End If
    
    ' �������� ���������� � ����� � ��������
    Dim stats As String
    stats = "======= Command Pool Statistics =======" & vbCrLf & vbCrLf
    
    ' ��������� ���������� � ���� ������ �����������
    stats = stats & "--- Log Commands Pool ---" & vbCrLf
    stats = stats & "Available commands: " & m_LogCommandPool.AvailableObjectCount & vbCrLf
    stats = stats & "In-use commands: " & m_LogCommandPool.InUseObjectCount & vbCrLf
    stats = stats & "Max pool size: " & m_LogCommandPool.MaxPoolSize & vbCrLf & vbCrLf
    
    ' ��������� ���������� � ���� ������ ��� ������ � �������
    stats = stats & "--- Data Commands Pool ---" & vbCrLf
    stats = stats & "Available commands: " & m_DataCommandPool.AvailableObjectCount & vbCrLf
    stats = stats & "In-use commands: " & m_DataCommandPool.InUseObjectCount & vbCrLf
    stats = stats & "Max pool size: " & m_DataCommandPool.MaxPoolSize & vbCrLf & vbCrLf
    
    ' ��������� ���������� � ���� ������ ��� ����������
    stats = stats & "--- UI Commands Pool ---" & vbCrLf
    stats = stats & "Available commands: " & m_UICommandPool.AvailableObjectCount & vbCrLf
    stats = stats & "In-use commands: " & m_UICommandPool.InUseObjectCount & vbCrLf
    stats = stats & "Max pool size: " & m_UICommandPool.MaxPoolSize & vbCrLf & vbCrLf
    
    ' ��������� ���������� � �������� � �������
    stats = stats & "--- Command History ---" & vbCrLf
    stats = stats & "Commands in history: " & m_CommandInvoker.CommandHistoryCount & vbCrLf
    stats = stats & "Can undo operations: " & IIf(m_CommandInvoker.CanUndo, "Yes", "No") & vbCrLf & vbCrLf
    
    stats = stats & "======================================"
    
    GetPoolsStatistics = stats
End Function

' ������� ���� ����� ������
Public Sub ReleaseAllPools()
    On Error Resume Next
    
    If Not m_ModuleInitialized Then Exit Sub
    
    ' ��������� ���������� ����� ������������� ��������
    Debug.Print "=== ���������� ����� ������������� �������� ==="
    Debug.Print GetPoolsStatistics()
    
    If Not m_LogCommandPool Is Nothing Then
        m_LogCommandPool.ClearAllObjects
        Set m_LogCommandPool = Nothing
    End If
    
    If Not m_DataCommandPool Is Nothing Then
        m_DataCommandPool.ClearAllObjects
        Set m_DataCommandPool = Nothing
    End If
    
    If Not m_UICommandPool Is Nothing Then
        m_UICommandPool.ClearAllObjects
        Set m_UICommandPool = Nothing
    End If
    
    Set m_CommandInvoker = Nothing
    
    m_ModuleInitialized = False
    
    Debug.Print "=== ��� ������� ����� ����������� ==="
    
    If Err.Number <> 0 Then
        Debug.Print "������ ��� ������������ ��������: " & Err.Description
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub


' �������� � ����� ������� ������
Public Sub InitBeforeDemo()
    On Error Resume Next
    ' ����������� ��� ������� ����� �������������
    ReleaseAllPools
    
    ' ������������� �������������� ����
    InitializeAllPools
    
    ' ������� ����������
    Debug.Print "=== ���� ���������������� � ������ � ������������ ==="
    Debug.Print GetPoolsStatistics()
    
    MsgBox "��� ���� ������ ���������������� � ������ � ������������!", _
           vbInformation, "����������"
End Sub

