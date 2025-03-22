VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Класс: ICommand.cls
' Описание: Интерфейс для всех команд
Option Explicit

' Метод выполнения команды
Public Sub Execute()
    ' Будет реализован в конкретных командах
End Sub

' Метод отмены команды (если поддерживается)
Public Sub Undo()
    ' Будет реализован в конкретных командах
End Sub

' Название команды
Public Property Get Name() As String
    ' Будет реализован в конкретных командах
End Property

' Поддерживает ли команда отмену
Public Property Get SupportsUndo() As Boolean
    ' Будет реализован в конкретных командах
End Property
