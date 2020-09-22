Attribute VB_Name = "ModMainCode"
Option Explicit

Public Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function GetInputState Lib "user32" () As Long

Public GL_DemoRunning As Boolean       'DemoMode
Public XEnambled As Boolean
Public YEnambled As Boolean
Public ZEnambled As Boolean

Public GL_Select As Boolean
Public GL_Move As Boolean
Public GL_Rotate As Boolean


Public CurrentLibraryId As Integer
Public CurrentOPFFile As String
Public CurrentOPFFileDirty As Boolean


Public GL_LibraryPath As String
Public GL_AppPath As String

Public Sub Main()
    XEnambled = True
    YEnambled = True
    GL_Move = True
    MY3D_Initialize
    CurrentOPFFileDirty = False
    GL_AppPath = App.Path & "\"
    Load frmmain
    frmmain.Show
End Sub

Public Sub CloseProgram()
    GL_DemoRunning = False
    MY3D_ShutDown
      If CurrentOPFFileDirty Then SaveOPFFile CurrentOPFFile
    End
End Sub
    
Public Sub UpdateStatus(pno As Integer, str As String)
    frmmain.Sbar.Panels(pno).Text = str
End Sub

Public Sub ResetView()
    With GL3D_CAMTransformation
        .YRot = 0
        .XRot = 0
        .ZRot = 0
        .XOff = 0
        .YOff = 0
        .ZOff = 0
        .ScaleAll = 1
    End With
    
       With GL3D_WorldTransformation
        .YRot = 0
        .XRot = 0
        .ZRot = 0
        .XOff = 0
        .YOff = 0
        .ZOff = 0
        .ScaleAll = 1
    End With
End Sub

'#############################################
'#  Demo Loop Code
'#############################################
Public Sub DoDemoLoop()
Static rot As Single
Dim Tim As Double
Dim counter As Long
Dim NextUpdateTime As Double
Const OneSec = 1000
Const StepAngle = 2
    UpdateStatus 1, "Demo Mode :"
    Tim = GetTickCount
    NextUpdateTime = Tim + OneSec
    Do While GL_DemoRunning
        counter = counter + 1
        rot = rot + StepAngle
        If rot >= 360 Then rot = 0
        With GL3D_CAMTransformation
            If XEnambled Then .XRot = rot
            If YEnambled Then .YRot = rot
            If ZEnambled Then .ZRot = rot
        End With
        If NextUpdateTime < GetTickCount Then
            UpdateStatus 2, "F/S :" & Format(counter, "###,#0")
            Tim = GetTickCount
            NextUpdateTime = Tim + OneSec
            counter = 1
        End If
        MY3D_DrawView
        DoEvents
    Loop
    UpdateStatus 1, "Ready :"
    UpdateStatus 2, ""
End Sub

'other file stuff
Public Function StripPath(spfile As String) As String
Dim i As Integer
    For i = 0 To Len(spfile)
        If Left(Right(spfile, i), 1) = "\" Then
            StripPath = Left(spfile, Len(spfile) - i + 1)
            Exit For
        End If
    Next
End Function

Public Function Stripfile(spPath As String) As String
Dim i As Integer
    For i = 0 To Len(spPath)
        If Left(Right(spPath, i), 1) = "\" Then
            Stripfile = Right(spPath, i - 1)
            Exit For
        End If
    Next
End Function

