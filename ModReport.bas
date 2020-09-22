Attribute VB_Name = "ModReport"
Option Explicit

Private Type MyData
    FunctionName As String
    Calls As Long
    TotalTime As Double
End Type


Private Type MyDataCalls
    FunctionName As String
    TimeSpent As Double
End Type
Public Data(1 To 10) As MyData
Public DataCalls(1 To 1000) As MyDataCalls
Public DataIndex As Long

Public Sub CallLog(FuncName As String, Status As Boolean)
Static Tim As Double
Static Func As String
If DataIndex >= 1000 Then Exit Sub
    If Status Then
        Tim = GetTickCount
        DataIndex = DataIndex + 1
        Func = FuncName
        DataCalls(DataIndex).FunctionName = Func
    Else
        If Func = FuncName Then
            DataCalls(DataIndex).TimeSpent = GetTickCount - Tim
        End If
    End If
End Sub
Public Sub ShowLog()
Dim i As Long
Dim d As Long

Dim T As Single
Dim C As Long
Dim datastr As String
Dim datamax As Long
Dim FunctionNameFound As Boolean

    For i = 1 To DataIndex
        For d = 1 To datamax
            FunctionNameFound = False
            If Data(d).FunctionName = DataCalls(i).FunctionName Then
                Data(d).Calls = Data(d).Calls + 1
                Data(d).TotalTime = Data(d).TotalTime + DataCalls(i).TimeSpent
                FunctionNameFound = True
                Exit For
            End If
        Next
        If FunctionNameFound = False Then
            datamax = datamax + 1
            Data(datamax).FunctionName = DataCalls(i).FunctionName
            Data(datamax).Calls = 1
            Data(datamax).TotalTime = DataCalls(i).TimeSpent
        End If
    Next
    For d = 1 To datamax
        datastr = datastr & d & vbTab & Data(d).FunctionName & Space(12) & vbTab & Data(d).Calls & vbTab & Data(d).TotalTime & vbTab & Format(Data(d).Calls / (Data(d).TotalTime + 0.01) * 1000, "###.##0") & " Fps" & vbNewLine
        
        If Data(d).FunctionName = "TransformationMatrix" Or Data(d).FunctionName = "ProcessZorder" Or Data(d).FunctionName = "RefreshDraw3d" Then
       
        T = T + Data(d).TotalTime
         C = Data(d).Calls

        End If
    Next
If T > 0 Then
frmReport.Text1 = datastr & vbNewLine & "Screen Refesh Rate " & Format(C / (T) * 1000, "###.##0") & " Fps"
Else
frmReport.Text1 = datastr & vbNewLine
End If
frmReport.Show 1
DataIndex = 0
End Sub

