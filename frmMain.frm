VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "BEA 3D Engine Demo"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11385
   ClipControls    =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   759
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1260
      Width           =   2355
      Begin VB.Frame Frame2 
         Caption         =   "Information"
         Height          =   1635
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   4560
         Width           =   2235
         Begin VB.Label LabfaceId 
            Caption         =   "Face ID:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1260
            Width           =   1395
         End
         Begin VB.Label LabFaceCol 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1560
            TabIndex        =   10
            Top             =   1260
            Width           =   555
         End
         Begin VB.Label LabObject 
            Caption         =   "Object :"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   1875
         End
         Begin VB.Label LabP 
            Caption         =   "Pionts :"
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   1755
         End
         Begin VB.Label LabPoly 
            Caption         =   "Polys"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   660
            Width           =   1875
         End
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3315
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   5847
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   295
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         SingleSel       =   -1  'True
         ImageList       =   "ImageListTree"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6735
      Left            =   60
      TabIndex        =   3
      Top             =   900
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   11880
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Object Library"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar Sbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7680
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14446
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Select"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":075E
            Key             =   "ZoomOut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A7A
            Key             =   "ZoomIn"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D96
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11EA
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12FE
            Key             =   "Properties"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   1535
      ButtonWidth     =   1429
      ButtonHeight    =   1376
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            Key             =   "Properties"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset"
            Key             =   "Reset"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Demo"
            Key             =   "Demo"
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Select"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Move"
            Key             =   "Move"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rotate"
            Key             =   "Rotate"
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ZoomOut"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ZoomIn"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "X"
            Key             =   "X"
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Y"
            Key             =   "Y"
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Z"
            Key             =   "Z"
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicView 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   2955
      Left            =   2820
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   371
      TabIndex        =   0
      Top             =   1140
      Width           =   5595
   End
   Begin MSComctlLib.ImageList ImageListTree 
      Left            =   2640
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1412
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1866
            Key             =   "Part"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CBA
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":210E
            Key             =   "Object"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'#############################################
'#  Form Varibles
'#############################################



'#############################################
'#  Form Code
'#############################################
Private Sub Form_Load()
   LoadObjectTree
   UpdateStatus 1, "Ready :"
End Sub

Private Sub Form_Resize()
Dim T As Single
Dim H As Single
    If Me.WindowState <> vbMinimized Then
        T = Toolbar1.Height + 4
        H = Me.ScaleHeight - (Sbar.Height + 4 + T)
        PicView.Move 174, T, Me.ScaleWidth - 176, H
        TabStrip1.Height = H
        Frame1(0).Height = H - 35
        TreeView1.Height = (H - 240) * 15
        Frame2(0).Top = (H - 220) * 15
        MY3D_Setup PicView, 4000
        MY3D_DrawView
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    CloseProgram
End Sub

Private Sub LabFaceCol_Click()
    frmPal.Left = Me.Left + LabFaceCol.Left + LabFaceCol.Width + 200
    frmPal.Top = Me.Top + Frame2(0).Top + TabStrip1.Top + Toolbar1.Height + LabFaceCol.Top + 1600
    frmPal.Show 1
    SelectedFaceColour = frmPal.SelColour
    LabFaceCol.BackColor = SelectedFaceColour
    MY3D_ColourSelectedFace SelectedFaceColour
    MY3D_DrawView
End Sub

'#############################################
'#  Menu Code
'#############################################


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub PicView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If GL_Select Then
        If MY3D_CheckForLocation(X, Y) Then
            LabObject = "Object ID: " & SelectedObject
            LabfaceId = "Face ID: " & SelectedFace
            LabFaceCol.BackColor = SelectedFaceColour
            MY3D_DrawView
        End If
    End If
End Sub

'#############################################
'#  Picture View Code
'#############################################
Private Sub PicView_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static LastMousePositionX As Single
Static LastMousePositionY As Single
    If Button <> 0 Then
        Select Case Button
            Case 1 'orbit
                If GL_Rotate Then DOMouse "Rotate", X - LastMousePositionX, Y - LastMousePositionY
                If GL_Move Then DOMouse "Move", X - LastMousePositionX, Y - LastMousePositionY
            Case 2 'dolly
                DOMouse "Zoom", X - LastMousePositionX, Y - LastMousePositionY
          End Select
    End If
    LastMousePositionX = X
    LastMousePositionY = Y
End Sub

Private Sub DOMouse(Action As String, xdif As Single, ydiff As Single)
    Select Case Action
        Case "Move"
        With GL3D_WorldTransformation
            If XEnambled Then .XOff = .XOff + xdif / GL3D_CAMTransformation.ScaleAll '* GL_TranslationFactor
            If YEnambled Then .YOff = .YOff + ydiff / GL3D_CAMTransformation.ScaleAll
            If ZEnambled Then .ZOff = .ZOff + ydiff / GL3D_CAMTransformation.ScaleAll
          End With
        Case "Rotate"
            With GL3D_CAMTransformation
                If XEnambled Then .XRot = .XRot - xdif * GL3D_RotationFactor
                If YEnambled Then .YRot = .YRot - ydiff * GL3D_RotationFactor
                If ZEnambled Then .ZRot = .ZRot - ydiff * GL3D_RotationFactor
                .XRot = (.XRot + 360) Mod (360)
                .YRot = (.YRot + 360) Mod (360)
                .ZRot = (.ZRot + 360) Mod (360)
            End With
        Case "Zoom"
              GL3D_CAMTransformation.ScaleAll = GL3D_CAMTransformation.ScaleAll + ydiff * GL3D_ScaleFactor
    End Select
    MY3D_DrawView
End Sub

Private Sub PicView_Paint()
  MY3D_DrawView
End Sub

'#############################################
'#  Main Tool Bar code
'#############################################
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key

        Case "Save"
              SaveOPFFile CurrentOPFFile
        Case "Properties"
            ShowLog
        Case "Reset"
            ResetView
            MY3D_DrawView
        Case "Select"
            GL_Select = True
            GL_Move = False
            GL_Rotate = False
            SelectedID = 0
         
         Case "Demo"
            GL_DemoRunning = IIf(Button.value = tbrPressed, True, False)
            If GL_DemoRunning Then DoDemoLoop
         Case "Move"
            GL_Move = True
            GL_Select = False
            GL_Rotate = False
            SelectedID = 0
            
         Case "Rotate"
            GL_Rotate = True
            GL_Select = False
            GL_Move = False
            SelectedID = 0
        Case "ZoomIn"
            GL3D_CAMTransformation.ScaleAll = GL3D_CAMTransformation.ScaleAll + (2 * GL3D_ScaleFactor)
            MY3D_DrawView
        Case "ZoomOut"
            GL3D_CAMTransformation.ScaleAll = GL3D_CAMTransformation.ScaleAll + (-2 * GL3D_ScaleFactor)
            MY3D_DrawView
        Case "X"
            XEnambled = IIf(Button.value = tbrPressed, True, False)
        Case "Y"
            YEnambled = IIf(Button.value = tbrPressed, True, False)
            If YEnambled Then
                Toolbar1.Buttons("Z").value = tbrUnpressed
                ZEnambled = False
            End If
        Case "Z"
            ZEnambled = IIf(Button.value = tbrPressed, True, False)
            If ZEnambled Then
                Toolbar1.Buttons("Y").value = tbrUnpressed
                YEnambled = False
            End If
    End Select
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    If LoadObjectFile(GL_LibraryPath & Node.Text) Then
        CurrentOPFFile = GL_LibraryPath & Node.Text
    End If
End Sub
