VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ќтранжирƒиапазон"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public массивЌаимен
Public массивѕозиц
Public место As Worksheet

Public Sub ѕерем олонок()
' *************************************************************************
' ћетод дл€ растановки столбцов из массива - "массивЌаимен" в позиции из
' массива - "массивѕозиц"
' *************************************************************************
    Dim номѕосл ол As Integer ' номер последнего столбца обрабатываемой таблицы
    Dim заголовоки As Range, заголовок As Range
    Dim i As Integer ' сетчик цикла по массивЌаимен
    Dim j As Integer ' сетчик цикла по массивЌаимен
    Dim n As Integer ' номер колонки
    
    номѕосл ол = место.Cells(1, Columns.Count).End(xlToLeft).Column
    Set заголовоки = место.Range(Cells(1, 1), Cells(1, номѕосл ол))
    
    For i = 0 To UBound(массивЌаимен)
    ' -------------------------------------------------------------
    ' цикл дл€ переноса столбцов последовательно в конец таблицы
    ' -------------------------------------------------------------
        For Each заголовок In заголовоки
            n = заголовок.Column
            If заголовок.Value = массивЌаимен(i) Then
                Columns(n).Cut
                Columns(номѕосл ол + 1).Insert
                Exit For
            End If
        Next заголовок
    Next i
        
    For j = 0 To UBound(массивЌаимен)
    ' -------------------------------------------------------------------
    ' цикл дл€ переноса столбцов в позиции согласно массива - массивѕозиц
    ' -------------------------------------------------------------------
        For Each заголовок In заголовоки
            n = заголовок.Column
            If заголовок.Value = массивЌаимен(j) Then
                If n <> массивѕозиц(j) Then
                    Columns(n).Cut
                    Columns(массивѕозиц(j)).Insert
                End If
                Exit For
            End If
        Next заголовок
    Next j
    
End Sub
