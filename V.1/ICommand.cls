VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' �����: ICommand.cls
' ��������: ��������� ��� ���� ������
Option Explicit

' ����� ���������� �������
Public Sub Execute()
    ' ����� ���������� � ���������� ��������
End Sub

' ����� ������ ������� (���� ��������������)
Public Sub Undo()
    ' ����� ���������� � ���������� ��������
End Sub

' �������� �������
Public Property Get Name() As String
    ' ����� ���������� � ���������� ��������
End Property

' ������������ �� ������� ������
Public Property Get SupportsUndo() As Boolean
    ' ����� ���������� � ���������� ��������
End Property
