Attribute VB_Name = "ModOPFCode"
Option Explicit

Public Type MyPoints
    X As Single                         'Real Data
    Y As Single
    Z As Single
    WX As Single                        'World Data
    WY As Single
    WZ As Single
End Type

Public Type MyFaceIndex
   A As Integer
   B As Integer
   C As Integer
   Colour As Long
End Type

Public Type MyObject
   PartID As String
   Name As String
   Colour As Long
   XOff As Single
   YOff As Single
   ZOff As Single
   XRot As Single
   YRot As Single
   ZRot As Single
   Points() As MyPoints
   PointsMax As Integer
   FaceIndex() As MyFaceIndex
   FaceMax As Integer
End Type

Public Object() As MyObject             'The merged Object Array
Public ObjectMax As Integer

Public Sub GetFiles(FileList() As String, ThisPath As String, Filter As String)
Dim MyName As String
Dim FileCounter As Long
On Error Resume Next
    FileCounter = 0
    MyName = Dir(ThisPath & Filter)    ' Retrieve the first entry.
    Do While MyName <> ""   ' Start the loop.
        FileCounter = FileCounter + 1
        MyName = Dir
    Loop
    ReDim FileList(FileCounter) As String
    FileCounter = 0
    MyName = Dir(ThisPath & Filter)    ' Retrieve the first entry.
    Do While MyName <> ""   ' Start the loop.
        FileCounter = FileCounter + 1
        FileList(FileCounter) = MyName
        MyName = Dir   ' Get next entry.Loop
    Loop
End Sub


Public Function LoadObjectTree() As Boolean
Dim mNode As Node ' Module-level variable for Nodes
Dim pno As Long
Dim A As Long
Dim Index As Long
Dim pIndex As Long
Dim Files() As String
frmmain.TreeView1.Nodes.Clear
GetFiles Files(), GL_LibraryPath, "*.OPF"
Set mNode = frmmain.TreeView1.Nodes.Add(, tvwFirst, "0K", "Library", "Closed")
mNode.Tag = "ROOT"
For A = 1 To UBound(Files)
Index = Index + 1
    Set mNode = frmmain.TreeView1.Nodes.Add(1, tvwChild, A & "K", Files(A), "Closed")
    mNode.Tag = "OPF"
   Next
   frmmain.TreeView1.Nodes(1).Expanded = True
   frmmain.TreeView1.Nodes(1).Sorted = True
   Set frmmain.TreeView1.SelectedItem = frmmain.TreeView1.Nodes.Item(1)
End Function


'###########################################################
'#
'#  BEA 3D Engine Load Data types from File
'#
'###########################################################

Public Function LoadObjectFile(objname As String) As Boolean
Dim Return_String As String
Dim FileNumber As Long
Dim Col As Integer
Dim ObjectNo As Long
Dim PointNo As Long
Dim PointsMax As Integer
Dim FaceNo As Long
Dim FaceTotalCounter As Long
On Error GoTo err1
  If CurrentOPFFileDirty Then SaveOPFFile CurrentOPFFile
SelectedID = 0
CallLog "LoadObjectFile", True
ObjectNo = 1
ReDim Object(1 To 1) As MyObject
    FileNumber = FreeFile
    Open objname For Input As #FileNumber
        Do While Not EOF(FileNumber)
            Line Input #FileNumber, Return_String
            If InStr(1, UCase(Return_String), "3DOBJECT", vbBinaryCompare) = 1 Then
                
                With Object(ObjectNo)
                    For Col = 1 To 2 '  fields
                        Input #FileNumber, Return_String
                        Select Case Col
                           Case 1
                             .Colour = 255 ' = Val(return_string)
                           Case 2
                              .Name = Return_String
                        End Select
                    Next
                End With
            End If
         
            If InStr(1, UCase(Return_String), "3DPOINTS", vbTextCompare) > 0 Then
                PointNo = 0
                Do While Not EOF(FileNumber)
                    PointNo = PointNo + 1 'Add Another point for this objectPionts
                    ReDim Preserve Object(ObjectNo).Points(1 To PointNo) As MyPoints
                    For Col = 1 To 3 '  fields
                    Input #FileNumber, Return_String
                    If InStr(1, UCase(Return_String), "END3DPOINTS", vbTextCompare) > 0 Then
                        ReDim Preserve Object(ObjectNo).Points(1 To PointNo - 1) As MyPoints
                        GoTo PointsFinished
                    End If
                    Select Case Col
                       Case 1
                          Object(ObjectNo).Points(PointNo).X = Val(Return_String)
                       Case 2
                          Object(ObjectNo).Points(PointNo).Y = Val(Return_String)
                       Case 3
                          Object(ObjectNo).Points(PointNo).Z = Val(Return_String)
                    End Select
                Next
                
                Object(ObjectNo).PointsMax = PointNo

                Loop
            End If
         
PointsFinished:
         
            If InStr(1, UCase(Return_String), "3DFACES", vbTextCompare) > 0 Then
                FaceNo = 0
                Do While Not EOF(FileNumber)
                    FaceNo = FaceNo + 1 'Add another face for this object
                    FaceTotalCounter = FaceTotalCounter + 1
                    ReDim Preserve Object(ObjectNo).FaceIndex(1 To FaceNo) As MyFaceIndex
                    For Col = 1 To 4 '  fields
                        Input #FileNumber, Return_String
                        If InStr(1, UCase(Return_String), "END3DFACES", vbTextCompare) > 0 Then
                            ReDim Preserve Object(ObjectNo).FaceIndex(1 To FaceNo - 1) As MyFaceIndex
                            GoTo Finished
                        End If
                        Select Case Col
                              Case 1
                                  Object(ObjectNo).FaceIndex(FaceNo).A = Val(Return_String)
                              Case 2
                                 Object(ObjectNo).FaceIndex(FaceNo).B = Val(Return_String)
                              Case 3
                                 Object(ObjectNo).FaceIndex(FaceNo).C = Val(Return_String)
                               Case 4
                                 Object(ObjectNo).FaceIndex(FaceNo).Colour = Val(Return_String)
                            
                        End Select
                    Next
                    Object(ObjectNo).FaceMax = FaceNo
                Loop
            End If
Finished:
            If InStr(1, UCase(Return_String), "END3DOBJECT", vbBinaryCompare) > 0 Then
                Exit Do
            End If
        Loop
    Close #FileNumber
    UpdateStatus 3, "Loading Library Objects :" & ObjectNo ': DoEvents
    ObjectMax = ObjectNo
    LoadObjectFile = True
     ObjectMax = 1
    ReDim VisiblePoly(1 To Object(ObjectNo).FaceMax) As MyVisiblePoly
    CurrentLibraryId = ObjectNo
    GL3D_FileLoaded = True
    ResetView
    UpdateStatus 3, "Setting ...  :" & Object(ObjectNo).Name
    MY3D_SetObjectLibToWorld
    MY3D_DrawView
    UpdateStatus 3, "File loaded :" & Object(ObjectNo).Name
    frmmain.LabP = "Points : " & Object(ObjectNo).PointsMax
    frmmain.LabPoly = "Polys : " & Object(ObjectNo).FaceMax
CallLog "LoadObjectFile", False

Exit Function
err1:
     MsgBox "Loading Library Error " & vbCr & objname & vbCr & Err.Description
    Close #FileNumber
    LoadObjectFile = False
End Function


'###########################################################
'#
'#  BEA 3D Engine Save Data types from File
'#
'###########################################################
Public Function SaveOPFFile(Save3DFile As String) As Boolean
Dim Return_String As String
Dim FileNumber As Long
Dim Col As Integer
Dim ObjectNo As Long
Dim PointNo As Long
Dim FaceNo As Long
Dim FaceTotalCounter As Long
On Error GoTo err1

CallLog "SaveOPFFile", True

    FileNumber = FreeFile
    Open Save3DFile For Output As #FileNumber
         ObjectNo = 1
            With Object(ObjectNo)
                Print #FileNumber, "------------"
                Print #FileNumber, ""
                Print #FileNumber, "3DOBJECT"
                Print #FileNumber, ObjectNo & "," & .Name
            
                Print #FileNumber, "3DPOINTS"
                For PointNo = 1 To .PointsMax
                    With .Points(PointNo)
                        Print #FileNumber, .X & "," & .Y & "," & .Z
                    End With
                Next
                Print #FileNumber, "END3DPOINTS"
 
                Print #FileNumber, "3DFACES"
                For FaceNo = 1 To .FaceMax
                    With .FaceIndex(FaceNo)
                        Print #FileNumber, .A & "," & .B & "," & .C & "," & .Colour
                    End With
                Next
                Print #FileNumber, "END3DFACES"
                
                Print #FileNumber, "END3DOBJECT"
            End With
    Close #FileNumber
    SaveOPFFile = True
      CurrentOPFFileDirty = False
CallLog "SaveOPFFile", False

Exit Function
err1:
    MsgBox "Saveing Library Error " & vbCr & Save3DFile & vbCr & Err.Description
    Close #FileNumber
    SaveOPFFile = False
End Function
