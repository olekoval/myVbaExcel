VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "����������������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ������������
Public �����������
Public ����� As Worksheet

Public Sub ������������()
' *************************************************************************
' ����� ��� ���������� �������� �� ������� - "������������" � ������� ��
' ������� - "�����������"
' *************************************************************************
    Dim ���������� As Integer ' ����� ���������� ������� �������������� �������
    Dim ���������� As Range, ��������� As Range
    Dim i As Integer ' ������ ����� �� ������������
    Dim j As Integer ' ������ ����� �� ������������
    Dim n As Integer ' ����� �������
    
    ���������� = �����.Cells(1, Columns.Count).End(xlToLeft).Column
    Set ���������� = �����.Range(Cells(1, 1), Cells(1, ����������))
    
    For i = 0 To UBound(������������)
    ' -------------------------------------------------------------
    ' ���� ��� �������� �������� ��������������� � ����� �������
    ' -------------------------------------------------------------
        For Each ��������� In ����������
            n = ���������.Column
            If ���������.Value = ������������(i) Then
                Columns(n).Cut
                Columns(���������� + 1).Insert
                Exit For
            End If
        Next ���������
    Next i
        
    For j = 0 To UBound(������������)
    ' -------------------------------------------------------------------
    ' ���� ��� �������� �������� � ������� �������� ������� - �����������
    ' -------------------------------------------------------------------
        For Each ��������� In ����������
            n = ���������.Column
            If ���������.Value = ������������(j) Then
                If n <> �����������(j) Then
                    Columns(n).Cut
                    Columns(�����������(j)).Insert
                End If
                Exit For
            End If
        Next ���������
    Next j
    
End Sub
