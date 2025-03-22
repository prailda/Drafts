Attribute VB_Name = "TestCommandPool"
' ������: CommandPoolDemo.bas
' ���������� ������������ ��������� Command � Object Pool
Option Explicit

' ��������� ��� ����������� ����������
Private Const DEMO_TITLE As String = "������������ ��������� Command � Object Pool"

' ������ ������ ������������
'Public Sub RunFullDemo()
'    ' ����������
'    ShowWelcomeMessage
'
'    ' �������� ������������
'    DemoCommandExecution
'    DemoCommandHistory
'    DemoObjectPoolStatistics
'
'    ' ����������
'    ShowFinalMessage
'End Sub

' ������������ ���������� ������
Private Sub DemoCommandExecution()
    ' �������� ������� ������
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    ' ���������� �������������� ���������
    MsgBox "������ ����� ������������������ ���������� ������." & vbCrLf & _
           "�� ������� ������ �� ���� � ��������� � ���������� ������.", _
           vbInformation, DEMO_TITLE & " - ��� 1"
    
    ' 1. ������ ����������������� �����
    Dim requestInput As New RequestInputCommand
    requestInput.Initialize "������� ���� ��� ��� ������������:", "���� ������", "������������"
    
    ' ����������� � ��������� � ���������
    Dim inputCmd As ICommand
    Set inputCmd = requestInput
    invoker.ExecuteCommand inputCmd
    
    ' �������� ���������
    Dim userName As String
    userName = requestInput.result
    
    ' 2. ������� ��������� ���������
    MsgBox "�� �����: " & userName & vbCrLf & _
           "��� ������ ���� �������� � ������� ������� RequestInputCommand � ���������.", _
           vbInformation, DEMO_TITLE & " - ��� 2"
    
    ' 3. ������������ ���������� ������
    MsgBox "������ ����� ��������� ������� AddRecordCommand, ������� ������� ������ � �������." & vbCrLf & _
           "��� ������� ������������ ������ ��������.", _
           vbInformation, DEMO_TITLE & " - ��� 3"
    
    Dim addCmd As New AddRecordCommand
    addCmd.Initialize userName, "UsersTable"
    
    ' ��������� �������
    invoker.ExecuteCommand addCmd
    
    ' 4. ������������ �������� ������
    MsgBox "������ ����� ��������� ������� DeleteRecordCommand, ������� ������ �������� ������." & vbCrLf & _
           "��� ������� ����� ������������ ������ ��������.", _
           vbInformation, DEMO_TITLE & " - ��� 4"
    
    Dim delCmd As New DeleteRecordCommand
    delCmd.Initialize "DEMO_ID_123", "DemoTable"
    
    ' ��������� �������
    invoker.ExecuteCommand delCmd
    
    ' 5. ������������ ����������� ������
    MsgBox "������ ����� ��������� ������� LogErrorCommand, ������� ������� ���������� �� ������." & vbCrLf & _
           "��������� ����� ����� ������� � ���� Immediate (Ctrl+G).", _
           vbInformation, DEMO_TITLE & " - ��� 5"
    
    Dim errorCmd As New LogErrorCommand
    errorCmd.Initialize "���������������� ������", "DemoModule", 1234, "��� �� ��������� ������, � ������������"
    
    ' ��������� �������
    invoker.ExecuteCommand errorCmd
End Sub

' ������������ ������� ������
Private Sub DemoCommandHistory()
    ' �������� ������� ������
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    
    ' ���������� ���������� �� ������� ������
    MsgBox "� ������� ������ ������ ��������� " & invoker.CommandHistoryCount & " ������(�), �������������� ������." & vbCrLf & _
           "������ ����� ������������������ ������ ��������� �������.", _
           vbInformation, DEMO_TITLE & " - ������� ������"
    
    ' �������� ��������� �������, ���� ���� ����� �����������
    If invoker.CanUndo Then
        invoker.UndoLastCommand
        
        ' ���������� � ����������� ������
        MsgBox "������� ���� ������� ��������!" & vbCrLf & _
               "������ � ������� " & invoker.CommandHistoryCount & " ������(�).", _
               vbInformation, DEMO_TITLE & " - ��������� ������"
    Else
        MsgBox "� ������� ��� ������, �������������� ������.", _
               vbExclamation, DEMO_TITLE & " - ������� ������"
    End If
End Sub

'' ������������ ���������� ���� ��������
'Private Sub DemoObjectPoolStatistics()
'    ' �������� ���������� �����
'    Dim stats As String
'    stats = GetPoolsStatistics()
'
'    ' ���������� ����������
'    MsgBox "���������� ����� ��������:" & vbCrLf & vbCrLf & _
'           stats & vbCrLf & vbCrLf & _
'           "��� ���������� ���������� ������������� �������� Object Pool ��� ���������� ���������.", _
'           vbInformation, DEMO_TITLE & " - ���������� �����"
'
'    ' ������� ����� � ���� Immediate ��� �������� �����
'    Debug.Print "===== ���������� ����� �������� ====="
'    Debug.Print stats
'    Debug.Print "===================================="
'End Sub

' �������������� ���������
Private Sub ShowWelcomeMessage()
    MsgBox "����� ���������� � ������������ ��������� Command � Object Pool!" & vbCrLf & vbCrLf & _
           "��� ������������ �������:" & vbCrLf & _
           "1. �������� � ���������� ������" & vbCrLf & _
           "2. ������ ������ ����� �������" & vbCrLf & _
           "3. ���������� ��������� ����� ���" & vbCrLf & vbCrLf & _
           "������� OK, ����� ������ ������������.", _
           vbInformation, DEMO_TITLE
End Sub

' ��������� ���������
Private Sub ShowFinalMessage()
    ' ����������� �������
    On Error Resume Next
    ReleaseAllPools
    If Err.Number <> 0 Then
        Debug.Print "������ ��� ������������ ��������: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    ' ������� ��������� ���������
    MsgBox "������������ ��������� Command � Object Pool ���������!" & vbCrLf & vbCrLf & _
           "��� ���� ������������������:" & vbCrLf & _
           "� ������� Command ��� ������������ ��������" & vbCrLf & _
           "� ������ �������� ����� ������� ������" & vbCrLf & _
           "� ������� Object Pool ��� ������������ ���������� ���������" & vbCrLf & vbCrLf & _
           "���������� ���������� ����� �������� � ���� Immediate (Ctrl+G).", _
           vbInformation, DEMO_TITLE & " - ����������"
End Sub

' ������� ���� �� ����������������� �������
Public Sub QuickTest()
    ' ������� �������� �������
    Dim testCmd As New LogInfoCommand
    testCmd.Initialize "�������� ������� ���������", "QuickTest"
    
    ' �������� ������� � ��������� �������
    Dim invoker As CommandInvoker
    Set invoker = GetCommandInvoker()
    invoker.ExecuteCommand testCmd
    
    ' ������� ���������
    MsgBox "���� �������� �������! ��������� ���� Immediate (Ctrl+G).", _
           vbInformation, "������� ����"
End Sub

' �������� � ������ CommandPoolDemo.bas
Public Sub PreloadPoolsDemo()
    ' �������������� ��� ����
    InitializeAllPools
    
    ' ������������ ������ � ����
    MsgBox "������ ����� ��������� ������������ ������ � ����." & vbCrLf & _
           "��� �������� ������������������ ��������� ������������� ��������.", _
           vbInformation, DEMO_TITLE & " - ������������"
    
    ' ��������� ��������� ������ � ��� �����������
    Dim i As Integer
    For i = 1 To 3
        ' ������� ������� �����������
        Dim logCmd As New LogInfoCommand
        logCmd.Initialize "��������������� ������� #" & i, "PreloadDemo"
        
        ' ��������� �������
        Dim invoker As CommandInvoker
        Set invoker = GetCommandInvoker()
        invoker.ExecuteCommand logCmd
        
        ' ���������� � ���
        GetLogCommandPool().ReturnCommand logCmd
    Next i
    
    ' ��������� ������� � ��� ������
    Dim dataCmd As New AddRecordCommand
    dataCmd.Initialize "Preloaded Data", "TestTable"
    invoker.ExecuteCommand dataCmd
    GetDataCommandPool().ReturnCommand dataCmd
    
    ' ���������� ���������� ����� ������������
    MsgBox "������������ ���������. ������ � ����� ���� ������� ��� ���������� �������������." & vbCrLf & vbCrLf & _
           "���������� �����:" & vbCrLf & vbCrLf & _
           GetPoolsStatistics(), _
           vbInformation, DEMO_TITLE & " - ���������� ������������"
End Sub

' ����������� ������ DemoObjectPoolStatistics
Private Sub DemoObjectPoolStatistics()
    ' ������������� ����������� ��������� ������, ����� ���������, ��� ���������� ����� �� ������
    Dim logCmd As ICommand
    Set logCmd = GetLogCommandPool().GetCommand("LogInfoCommand")
    Dim dataCmd As ICommand
    Set dataCmd = GetDataCommandPool().GetCommand("AddRecordCommand")
    Dim uiCmd As ICommand
    Set uiCmd = GetUICommandPool().GetCommand("ShowMessageCommand")
    
    ' �� ��������� �������, ������ ���������� �� � ���
    GetLogCommandPool().ReturnCommand logCmd
    GetDataCommandPool().ReturnCommand dataCmd
    GetUICommandPool().ReturnCommand uiCmd
    
    ' �������� ���������� �����
    Dim stats As String
    stats = GetPoolsStatistics()
    
    ' ���������, ��� ���������� �� ������
    If Len(stats) < 100 Then
        ' �������������� ��������� ����������
        stats = "======= Command Pool Statistics =======" & vbCrLf & vbCrLf
        stats = stats & "--- Log Commands Pool ---" & vbCrLf
        stats = stats & "Available commands: " & GetLogCommandPool().AvailableObjectCount & vbCrLf
        stats = stats & "In-use commands: " & GetLogCommandPool().InUseObjectCount & vbCrLf
        stats = stats & "Max pool size: " & GetLogCommandPool().MaxPoolSize & vbCrLf & vbCrLf
        
        stats = stats & "--- Data Commands Pool ---" & vbCrLf
        stats = stats & "Available commands: " & GetDataCommandPool().AvailableObjectCount & vbCrLf
        stats = stats & "In-use commands: " & GetDataCommandPool().InUseObjectCount & vbCrLf
        stats = stats & "Max pool size: " & GetDataCommandPool().MaxPoolSize & vbCrLf & vbCrLf
        
        stats = stats & "--- UI Commands Pool ---" & vbCrLf
        stats = stats & "Available commands: " & GetUICommandPool().AvailableObjectCount & vbCrLf
        stats = stats & "In-use commands: " & GetUICommandPool().InUseObjectCount & vbCrLf
        stats = stats & "Max pool size: " & GetUICommandPool().MaxPoolSize & vbCrLf & vbCrLf
        
        stats = stats & "--- Command History ---" & vbCrLf
        stats = stats & "Commands in history: " & GetCommandInvoker().CommandHistoryCount & vbCrLf
        stats = stats & "Can undo operations: " & IIf(GetCommandInvoker().CanUndo, "Yes", "No") & vbCrLf & vbCrLf
        
        stats = stats & "======================================"
    End If
    
    ' ���������� ����������
    MsgBox "���������� ����� ��������:" & vbCrLf & vbCrLf & _
           stats & vbCrLf & vbCrLf & _
           "��� ���������� ���������� ������������� �������� Object Pool ��� ���������� ���������." & vbCrLf & _
           "�������� �������� �� ����������� ��������� � ������������ ��������.", _
           vbInformation, DEMO_TITLE & " - ���������� �����"
    
    ' ������� ����� � ���� Immediate ��� �������� �����
    Debug.Print "===== ���������� ����� �������� ====="
    Debug.Print stats
    Debug.Print "===================================="
End Sub

' ���������� ������ RunFullDemo
Public Sub RunFullDemo()
    ' ����������
    ShowWelcomeMessage
    
    ' ������������ ����� ��� ������������ ���������� �������������
    PreloadPoolsDemo
    
    ' �������� ������������
    DemoCommandExecution
    DemoCommandHistory
    DemoObjectPoolStatistics
    
    ' ����������
    ShowFinalMessage
End Sub

