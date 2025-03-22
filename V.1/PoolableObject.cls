VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PoolableObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Класс: PoolableObject.cls
' Описание: Интерфейс для объектов, которые могут быть использованы в пуле
Option Explicit

' Метод для сброса состояния объекта перед повторным использованием
Public Sub Reset()
    ' Будет реализован в конкретных классах
End Sub

' Метод для активации объекта после извлечения из пула
Public Sub Activate()
    ' Будет реализован в конкретных классах
End Sub

' Метод для деактивации объекта перед возвратом в пул
Public Sub Deactivate()
    ' Будет реализован в конкретных классах
End Sub

' Идентификатор объекта (для отслеживания)
Public Property Get ObjectID() As String
    ' Будет реализован в конкретных классах
End Property
