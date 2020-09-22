VERSION 5.00
Begin VB.Form frmPal 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   1110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   74
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmPal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long

Dim MySelColour As Long

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MySelColour = Me.Point(X, Y)
    Me.Hide
End Sub

Private Sub Form_Paint()
Dim row As Integer
Dim Column As Integer
Dim BoxHeight As Integer
Dim BoxWidth As Integer
Dim colIndex As Long
Dim HBursh As Long
Dim HOldBush As Long
Dim DUMMY As Long
    BoxHeight = ScaleHeight \ 5
    BoxWidth = ScaleWidth \ 4
    For colIndex = 1 To 20
        row = colIndex \ 5 + 1
        Column = colIndex Mod 4 + 1
        HBursh = CreateSolidBrush(&H1000000 Or colIndex)
        DUMMY = SelectObject(hdc, HBursh)
        DUMMY = Rectangle(hdc, (Column - 1) * BoxWidth, (row - 1) * BoxHeight, Column * BoxWidth, row * BoxHeight)
        DUMMY = SelectObject(hdc, GetStockObject(4))
        DUMMY = DeleteObject(HBursh)
    Next
End Sub

Public Property Get SelColour() As Long
    SelColour = MySelColour
End Property

Public Property Let SelColour(ByVal vNewValue As Long)
    SelColour = vNewValue
End Property
