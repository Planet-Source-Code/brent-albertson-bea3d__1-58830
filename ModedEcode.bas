Attribute VB_Name = "ModedEcode"
Option Explicit
'################################################
'   BEA3DEngine
'   2002
'################################################

''API Text constants
Private Const LF_FACESIZE = 32
Private Const LOGPIXELSY = 90

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lsngStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lsngPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE - 1) As Byte
End Type
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal intIndex As Long) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long

'API Stuff
Private Const SRCCOPY = &HCC0020

'use by the Polygon type api functions
Public Type POINTAPI
        X As Long
        Y As Long
End Type

'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Long) As Integer

Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoints As Any, ByVal nCount As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

'BACKGROUND pICTURE
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'CreatePen(nPenStyle)
Private Const PS_SOLID = 0
Private Const PS_DASH = 1                    '  -------
Private Const PS_DOT = 2                     '  .......
Private Const PS_DASHDOT = 3                 '  _._._._
Private Const PS_DASHDOTDOT = 4              '  _.._.._
Private Const PS_NULL = 5
Private Const PS_INSIDEFRAME = 6

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long


Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

'lookups for Sin And Cos may Help Speed
Private Sine(0 To 360) As Single
Private CoSine(0 To 360) As Single

Public Const Pi = 3.14159265358979
Public Const CRan = (Pi / 180)
Public Const HalfPi = (Pi / 2)

' 3d to 2d useing this the Perspective
Private VP_X As Single
Private VP_Y As Single
Private VP_Z As Single

'transformed  matrix
Private TS(1 To 4, 1 To 4) As Single            'scale
Private TR(1 To 4, 1 To 4) As Single            'rotation
Private WorldT(1 To 4, 1 To 4) As Single        'world view
Private CAMT(1 To 4, 1 To 4) As Single          'Cam view
Private ViewT(1 To 4, 1 To 4) As Single         'View view
'Size of World
Private MyViewWidth As Single
Private MyViewHeight As Single

'Centre of World
Public MyViewXoff As Single
Public MyViewYoff As Single
Public MyViewZoff As Single


Public Type MyVisiblePoly
   OnwersName As String
   OwnerID As Integer
   FaceID As Integer
   VisblePF(2) As POINTAPI
   Colour As Long
   Zorder As Single                     'Each Visible Faces is Zorderd
   ZOrderIndex As Integer
   ZOrderProcessed As Boolean
End Type
Public VisiblePoly() As MyVisiblePoly   'Visible Faces Array
Public VisiblePolyMax As Integer

'Picture Box for View rendering
Private DstPicture  As PictureBox


Private Type TransformationMatrix
   XOff As Single
   YOff As Single
   ZOff As Single
   XRot As Single
   YRot As Single
   ZRot As Single
   ScaleAll As Single
End Type
Private ObjectTransformation As TransformationMatrix    'useby MY3D_SetToWorld
Public SelectedID As Integer
Public SelectedObject As Integer
Public SelectedFace As Integer

Public SelectedFaceColour As Long
Public SelectedObjectcolour As Long


'control of engine stuff
Public GL3D_WorldTransformation As TransformationMatrix
Public GL3D_CAMTransformation As TransformationMatrix
Public GL3D_ScaleFactor As Single
Public GL3D_RotationFactor As Single
Public GL3D_Render As String
Public GL3D_FileLoaded As Boolean
Public GL3D_BackColour As Long




Private hMemDC As Long
Private hMemBitmap As Long
Private hOldMemBitmap As Long
Private hErrCode As Long
Private mlfFont As LOGFONT

'################################################
'   BEA3DEngine
'   Call to Initialize the 3d Engine
'################################################
Public Sub MY3D_Initialize()
Dim i As Long

CallLog "MY3D_Initialize", True

    For i = 0 To 360
        Sine(i) = Sin(i * CRan)
        CoSine(i) = Cos(i * CRan)
    Next
    Initializemartix TS
    Initializemartix TR
    Initializemartix WorldT
    Initializemartix CAMT
    Initializemartix ViewT
    GL3D_Render = "Hidden Line"
    GL3D_FileLoaded = False
    GL3D_ScaleFactor = 0.1
    GL3D_RotationFactor = 0.5
    GL3D_WorldTransformation.ScaleAll = 1
    GL3D_CAMTransformation.ScaleAll = 1
    GL3D_BackColour = &HE0E0E0 'light grey
    
CallLog "MY3D_Initialize", False
    
End Sub

Private Sub Initializemartix(T() As Single)
   T(1, 1) = 1
   T(1, 2) = 0
   T(1, 3) = 0
   T(1, 4) = 1
   T(2, 1) = 0
   T(2, 2) = 1
   T(2, 3) = 0
   T(2, 4) = 1
   T(3, 1) = 0
   T(3, 2) = 0
   T(3, 3) = 1
   T(3, 4) = 1
   T(4, 1) = 0
   T(4, 2) = 0
   T(4, 3) = 0
   T(4, 4) = 1
End Sub
Public Sub MY3D_ShutDown()
    GL3D_FileLoaded = False
    DeleteDC hMemDC
    DeleteObject hMemBitmap
End Sub
'################################################
'   BEA3DEngine
'   Call when Picture Box Objects Change Size
'################################################
'Call this First or when the Picture boxs change size
Public Sub MY3D_Setup(pict As PictureBox, VPoffset As Single)
Dim bytBuf() As Byte
Dim intI As Integer
CallLog "MY3D_Setup", True


    Set DstPicture = pict
    With DstPicture
        .ScaleMode = 3
        .AutoRedraw = False
        .Visible = True
        .FillStyle = vbSolid 'vbTransparent
        .DrawMode = vbCopyPen
        .DrawStyle = vbSolid
        MyViewWidth = .ScaleWidth
        MyViewHeight = .ScaleHeight
    End With
    
    MyViewXoff = MyViewWidth / 2
    MyViewYoff = MyViewHeight / 2
    MyViewZoff = 1 'VPoffset
    VP_X = MyViewXoff
    VP_Y = MyViewYoff
    VP_Z = VPoffset
    
    hMemDC = CreateCompatibleDC(DstPicture.hdc)
    hMemBitmap = CreateCompatibleBitmap(DstPicture.hdc, MyViewWidth, MyViewHeight)
    hOldMemBitmap = SelectObject(hMemDC, hMemBitmap)

    'Prepare font name, decoding from Unicode
    With mlfFont
        bytBuf = StrConv("Courier New" & Chr$(0), vbFromUnicode)
        For intI = 0 To UBound(bytBuf)
            .lfFaceName(intI) = bytBuf(intI)
        Next intI
        .lfHeight = 10 * GetDeviceCaps(hMemDC, LOGPIXELSY) \ 72
        .lfEscapement = CLng(0 * 10#)
        .lfOrientation = mlfFont.lfEscapement
        .lfItalic = 0 'Set Italic or not
        .lfUnderline = 0 'Set Underline or not
        .lsngStrikeOut = 0 'Set Strikethrough or not
        .lfWeight = 500 'Set Bold or not (use font's weight ie 400 to 500)
    End With
      
CallLog "MY3D_Setup", False
    
End Sub


'################################################
'   BEA3DEngine
'   Main Redraw Function
'################################################
Public Sub MY3D_DrawView()
    If GL3D_FileLoaded Then
       SETTransformation WorldT(), GL3D_WorldTransformation
       SETTransformation CAMT(), GL3D_CAMTransformation
       MatrixMatrixMult ViewT(), CAMT(), WorldT()
       TransformationMatrix ViewT()
       ProcessZorder
       RefreshDraw3d
    End If
End Sub


Private Sub SETTransformation(TempT() As Single, Transformation As TransformationMatrix)
Dim SinX As Single
Dim SinY As Single
Dim SinZ As Single
Dim CosY As Single
Dim CosX As Single
Dim CosZ As Single

'CallLog "SETTransformation", True

    With Transformation
       SinX = Sine(.XRot)
       SinY = Sine(.YRot)
       SinZ = Sine(.ZRot)
       CosX = CoSine(.XRot)
       CosY = CoSine(.YRot)
       CosZ = CoSine(.ZRot)
       'Set The position Transformation
       TS(1, 4) = .XOff
       TS(2, 4) = .YOff
       TS(3, 4) = .ZOff
       'Set The Scale Transformation
       TS(1, 1) = .ScaleAll   'Note: you could scale each axis
       TS(2, 2) = .ScaleAll
       TS(3, 3) = .ScaleAll
       TS(4, 4) = 1
    End With
    'Set The Rotation Transformation
    TR(1, 1) = CosZ * CosY
    TR(2, 1) = CosZ * -SinY * -SinX + SinZ * CosX
    TR(3, 1) = CosZ * -SinY * CosX + SinZ * SinX
    TR(1, 2) = -SinZ * CosY
    TR(2, 2) = -SinZ * -SinY * -SinX + CosZ * CosX
    TR(3, 2) = -SinZ * -SinY * CosX + CosZ * SinX
    TR(1, 3) = SinY
    TR(2, 3) = CosY * -SinX
    TR(3, 3) = CosY * CosX
   
    MatrixMatrixMult TempT(), TS(), TR()
'CallLog "SETTransformation", False

End Sub



'Matrix Multply
Private Sub MatrixMatrixMult(r() As Single, A() As Single, B() As Single)
Dim i As Byte
Dim j As Byte
Dim k As Byte
Dim value As Single
    For i = 1 To 4
        For j = 1 To 4
            value = 0
            For k = 1 To 4
                value = value + A(i, k) * B(k, j)
            Next k
            r(i, j) = value
        Next j
    Next i
End Sub

'Transform all 3dpoints by the Main View Matrix
Private Sub TransformationMatrix(TM() As Single)
Dim O As Integer
Dim f As Integer
Dim NearDis As Single
Dim X As Single
Dim Y As Single
Dim Z As Single
Dim VAX As Single
Dim VAY As Single
Dim VAZ As Single
Dim VBX As Single
Dim VBY As Single
Dim VBZ As Single
Dim VCX As Single
Dim VCY As Single
Dim VCZ As Single
Dim Zorder As Single

CallLog "TransformationMatrix", True
    TM(1, 4) = TM(1, 4) + MyViewXoff
    TM(2, 4) = TM(2, 4) + MyViewYoff '+ MyViewYoff 'only for car demo
    TM(3, 4) = TM(3, 4) + MyViewZoff

    VisiblePolyMax = 0
    For O = 1 To ObjectMax
        With Object(O)
            For f = 1 To .FaceMax
                'Transformation
                With .Points(.FaceIndex(f).A)
                    X = .WX
                    Y = .WY
                    Z = .WZ
                End With
                VAX = X * TM(1, 1) + Y * TM(1, 2) + Z * TM(1, 3) + TM(1, 4)
                VAY = X * TM(2, 1) + Y * TM(2, 2) + Z * TM(2, 3) + TM(2, 4)
                VAZ = X * TM(3, 1) + Y * TM(3, 2) + Z * TM(3, 3) + TM(3, 4)
                
                With .Points(.FaceIndex(f).B)
                    X = .WX
                    Y = .WY
                    Z = .WZ
                End With
                VBX = X * TM(1, 1) + Y * TM(1, 2) + Z * TM(1, 3) + TM(1, 4)
                VBY = X * TM(2, 1) + Y * TM(2, 2) + Z * TM(2, 3) + TM(2, 4)
                VBZ = X * TM(3, 1) + Y * TM(3, 2) + Z * TM(3, 3) + TM(3, 4)
                
                With .Points(.FaceIndex(f).C)
                    X = .WX
                    Y = .WY
                    Z = .WZ
                End With
                VCX = X * TM(1, 1) + Y * TM(1, 2) + Z * TM(1, 3) + TM(1, 4)
                VCY = X * TM(2, 1) + Y * TM(2, 2) + Z * TM(2, 3) + TM(2, 4)
                VCZ = X * TM(3, 1) + Y * TM(3, 2) + Z * TM(3, 3) + TM(3, 4)
                              
                'Propestive
                NearDis = ((VP_Z - VAZ) / VP_Z) ^ 2
                VAX = VP_X - (NearDis * (VP_X - VAX))
                VAY = VP_Y - (NearDis * (VP_Y - VAY))
                NearDis = ((VP_Z - VBZ) / VP_Z) ^ 2
                VBX = VP_X - (NearDis * (VP_X - VBX))
                VBY = VP_Y - (NearDis * (VP_Y - VBY))
                NearDis = ((VP_Z - VCZ) / VP_Z) ^ 2
                VCX = VP_X - (NearDis * (VP_X - VCX))
                VCY = VP_Y - (NearDis * (VP_Y - VCY))
                
                'Cull
                If (VCX - VAX) * (VBY - VAY) > (VBX - VAX) * (VCY - VAY) Then GoTo NextFace

                'clip vertical
                If MyViewHeight < VAY And MyViewHeight < VBY And MyViewHeight < VCY Then GoTo NextFace
                If 0 > VAY And 0 > VBY And 0 > VCY Then GoTo NextFace
                'Clip horzital
                If MyViewWidth < VAX And MyViewWidth < VBX And MyViewWidth < VCX Then GoTo NextFace
                If 0 > VAX And 0 > VBX And 0 > VCX Then GoTo NextFace
                
                'Set Zorder
                Zorder = (VCZ + VBZ + VAZ) / 3
                If Zorder > VP_Z Then GoTo NextFace

                'Face Visible
                VisiblePolyMax = VisiblePolyMax + 1
                With VisiblePoly(VisiblePolyMax)
                'If O = 1 And f = 306 Then MsgBox Object(O).FaceIndex(f).Colour
                
               ' If SelectedObject = O Then
                 '   .Colour = SelectedObjectcolour
                 '   Else
                    
                    .Colour = Object(O).FaceIndex(f).Colour
                 '   End If
                    .OnwersName = Object(O).Name
                    .OwnerID = O
                    .FaceID = f
                    .VisblePF(0).X = VAX
                    .VisblePF(0).Y = VAY
                    .VisblePF(1).X = VBX
                    .VisblePF(1).Y = VBY
                    .VisblePF(2).X = VCX
                    .VisblePF(2).Y = VCY
                    .Zorder = Zorder
                    .ZOrderIndex = VisiblePolyMax
                    .ZOrderProcessed = False
                End With
NextFace:
            Next
    End With
   Next
   
CallLog "TransformationMatrix", False

End Sub

'Process Only Visible Faces For Zorder
Private Sub ProcessZorder()
Dim V As Integer
Dim Z As Integer
Dim TempLow As Single
Dim TempIndex As Integer   ' Find the Zorder to render

CallLog "ProcessZorder", True

    For V = 1 To VisiblePolyMax
        TempLow = -9999
        For Z = 1 To VisiblePolyMax
            If VisiblePoly(Z).ZOrderProcessed = False Then
                If TempLow < VisiblePoly(Z).Zorder Then
                    TempLow = VisiblePoly(Z).Zorder
                    TempIndex = Z
                End If
            End If
        Next
        VisiblePoly(V).ZOrderIndex = TempIndex
        VisiblePoly(TempIndex).ZOrderProcessed = True
    Next
    
CallLog "ProcessZorder", False
           
End Sub

'Draw the objects to the BackGround the flip to the Front Picture Box
Private Sub RefreshDraw3d()
Dim V As Integer     'Visible Object
Dim hBrush As Long
Dim hOldBrush As Long
Dim hPen As Long
Dim hOldPen As Long
Dim lngFont As Long
Dim lngOldFont As Long
Dim StrText As String
Dim Line_P As POINTAPI
Dim highlight As Long
Dim BrightNess As Long
Dim Colour As Long
Dim RetVal As Long
Dim RC As Byte
Dim BC As Byte
Dim GC As Byte
On Error GoTo err1
CallLog "RefreshDraw3d", True

    'CLS
    hBrush = CreateSolidBrush(GL3D_BackColour)
    hOldBrush = SelectObject(hMemDC, hBrush)
    Rectangle hMemDC, 0, 0, MyViewWidth, MyViewHeight
    SelectObject hMemDC, hOldBrush
    DeleteObject hBrush
      
    'High light Near Faces
    highlight = (VisiblePolyMax / 128) + 1

    'Make Pen
    hPen = CreatePen(PS_NULL, 1, 0)
    hOldPen = SelectObject(hMemDC, hPen)
    'Draw Polys
    For V = 1 To VisiblePolyMax
        With VisiblePoly(VisiblePoly(V).ZOrderIndex) '       NextToRender
            'Make Colour
            BrightNess = (V / highlight)
            RC = (.Colour And &HFF) / 2
            GC = ((.Colour \ &H100) And &HFF) / 2
            BC = ((.Colour \ &H10000) And &HFF) / 2
            Colour = RGB(RC + BrightNess, GC + BrightNess, BC + BrightNess)
            
            'Make Colour Bush
            hBrush = CreateSolidBrush(Colour)
            hOldBrush = SelectObject(hMemDC, hBrush)
            
            'Draw Face
            Polygon hMemDC, .VisblePF(0), 3
            
            'Delete Bush
            SelectObject hMemDC, hOldBrush
            DeleteObject hBrush
         End With
    Next
    SelectObject hMemDC, hOldPen
    DeleteObject hPen
    
   If SelectedID <> 0 Then
        'Make Pen
    hPen = CreatePen(PS_SOLID, 1, vbWhite)
    hOldPen = SelectObject(hMemDC, hPen)
   With VisiblePoly(SelectedID)
        'Make Colour
         '   BrightNess = (V / highlight)
          '  RC = (.Colour And &HFF) / 2
          '  GC = ((.Colour \ &H100) And &HFF) / 2
           ' BC = ((.Colour \ &H10000) And &HFF) / 2
           ' Colour = RGB(RC + BrightNess, GC + BrightNess, BC + BrightNess)
            'Make Colour Bush
            hBrush = CreateSolidBrush(.Colour)
            hOldBrush = SelectObject(hMemDC, hBrush)
            'Draw Face
            Polygon hMemDC, .VisblePF(0), 3
            'Delete Bush
            SelectObject hMemDC, hOldBrush
            DeleteObject hBrush
         End With
    
    SelectObject hMemDC, hOldPen
    DeleteObject hPen
    End If
    
    
    
    hBrush = CreateSolidBrush(vbWhite)
    hOldBrush = SelectObject(hMemDC, hBrush)
    Rectangle hMemDC, 0, MyViewHeight - 17, MyViewWidth, MyViewHeight
    SelectObject hMemDC, hOldBrush
    DeleteObject hBrush

     'Draw Text Box
    hPen = CreatePen(PS_SOLID, 1, vbBlack)
    hOldPen = SelectObject(hMemDC, hPen)
    MoveToEx hMemDC, 1, MyViewHeight - 18, Line_P
    LineTo hMemDC, MyViewWidth, MyViewHeight - 18
    SelectObject hMemDC, hOldPen
    DeleteObject hPen

'    'Build temporary new font and output the string
    lngFont = CreateFontIndirect(mlfFont)
    lngOldFont = SelectObject(hMemDC, lngFont)
    With GL3D_WorldTransformation
        StrText = " Off Set : " & Format(.XOff, " 0#.#0 ")
        TextOut hMemDC, 2, MyViewHeight - 15, StrText, Len(StrText)
        StrText = Format(.YOff, " 0#.#0 ")
        TextOut hMemDC, 118, MyViewHeight - 15, StrText, Len(StrText)
        StrText = Format(.ZOff, " 0#.#0 ")
        TextOut hMemDC, 162, MyViewHeight - 15, StrText, Len(StrText)
    End With
    
    With GL3D_CAMTransformation
        StrText = " Rotation : " & Format(.XRot, " 00# ")
        TextOut hMemDC, 220, MyViewHeight - 15, StrText, Len(StrText)
        StrText = Format(.YRot, " 00# ")
        TextOut hMemDC, 342, MyViewHeight - 15, StrText, Len(StrText)
        StrText = Format(.ZRot, " 00# ")
        TextOut hMemDC, 392, MyViewHeight - 15, StrText, Len(StrText)
        StrText = " Scale : " & Format(.ScaleAll, " 0#.#0 ")
        TextOut hMemDC, 450, MyViewHeight - 15, StrText, Len(StrText)
    End With
    lngFont = SelectObject(hMemDC, lngOldFont)
    DeleteObject lngFont
    
   
    'Draw Cross
    hPen = CreatePen(PS_DOT, 1, vbYellow)
    hOldPen = SelectObject(hMemDC, hPen)
    MoveToEx hMemDC, MyViewXoff, MyViewYoff - 25, Line_P
    LineTo hMemDC, MyViewXoff, MyViewYoff + 25
    MoveToEx hMemDC, MyViewXoff - 25, MyViewYoff, Line_P
    LineTo hMemDC, MyViewXoff + 25, MyViewYoff
    SelectObject hMemDC, hOldPen
    DeleteObject hPen
    
    'Copy background picture(hMemDC) to the forground
    RetVal = BitBlt(DstPicture.hdc, 0, 0, MyViewWidth, MyViewHeight, hMemDC, 0, 0, SRCCOPY)
   'Call BitBlt(frmmain.PicView.hdc, 1, 1, 10, 10, hMemDC, 1, 1, SRCCOPY)
 If RetVal = 0 Then MsgBox GetLastError()
  'Debug.Print RetVal

    
CallLog "RefreshDraw3d", False
Exit Sub
err1:
MsgBox Err.Description
End Sub


' check All Visible Polys
Public Function MY3D_CheckForLocation(Xpos As Single, Ypos As Single) As Boolean
Dim RgnHandel As Long
Dim V As Long
    MY3D_CheckForLocation = False
    SelectedObject = 0
    SelectedFace = 0
    For V = 1 To VisiblePolyMax
        With VisiblePoly(V) '       NextToRender
            RgnHandel = CreatePolygonRgn(.VisblePF(0), 3, 1)
            If PtInRegion(RgnHandel, Xpos, Ypos) Then GoTo LocationFound
            DeleteObject (RgnHandel)
        End With
    Next
Exit Function
LocationFound:
    MY3D_CheckForLocation = True
    SelectedID = V
    SelectedObject = VisiblePoly(V).OwnerID
    SelectedFace = VisiblePoly(V).FaceID
    SelectedFaceColour = VisiblePoly(V).Colour
    
    DeleteObject (RgnHandel)
Exit Function
err1:
    MY3D_CheckForLocation = False
End Function

Public Function MY3D_ColourSelectedFace(Colour As Long) As Boolean
On Error GoTo err1
    Object(SelectedObject).FaceIndex(SelectedFace).Colour = Colour
    MY3D_ColourSelectedFace = True
    CurrentOPFFileDirty = True
Exit Function
err1:
MY3D_ColourSelectedFace = False
End Function

Public Function MY3D_ColourSelectedObject(Colour As Long) As Boolean
On Error GoTo err1
MY3D_ColourSelectedObject = True
    SelectedObjectcolour = Colour
Exit Function
err1:
MY3D_ColourSelectedObject = False
End Function


Public Sub MY3D_SetObjectLibToWorld()
Dim O As Integer
Dim p As Integer
Dim nx As Single
Dim ny As Single
Dim nz As Single
ReDim T(1 To 4, 1 To 4) As Single
    Initializemartix T
    With Object(1)
        ObjectTransformation.XOff = .XOff '+ MyViewXoff
        ObjectTransformation.YOff = .YOff '+ MyViewYoff
        ObjectTransformation.ZOff = .ZOff '+ MyViewZoff
        ObjectTransformation.XRot = .XRot
        ObjectTransformation.YRot = .YRot
        ObjectTransformation.ZRot = .ZRot
        ObjectTransformation.ScaleAll = 1
        SETTransformation T(), ObjectTransformation
         For p = 1 To UBound(.Points)
            With .Points(p)
                  nx = .X  '+ Offset to Object Centre in the X Axis
                  ny = .Y '+ Offset to Object Centre in the y Axis
                  nz = .Z  '+ Offset to Object Centre in the z Axis
                  .WX = (nx * T(1, 1)) + (ny * T(1, 2)) + (nz * T(1, 3)) + T(1, 4)
                  .WY = (nx * T(2, 1)) + (ny * T(2, 2)) + (nz * T(2, 3)) + T(2, 4)
                  .WZ = (nx * T(3, 1)) + (ny * T(3, 2)) + (nz * T(3, 3)) + T(3, 4)
              End With
         Next
     End With
End Sub

