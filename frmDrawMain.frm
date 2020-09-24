VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDrawMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zac's Paint Program"
   ClientHeight    =   8055
   ClientLeft      =   4065
   ClientTop       =   1890
   ClientWidth     =   9345
   FillStyle       =   0  'Solid
   Icon            =   "frmDrawMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmDrawMain.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   623
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Zac"
      DialogTitle     =   "Saving..."
      FileName        =   "New Picture"
      Filter          =   "Image Files (*.Zac)"
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   0
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "Zac"
      DialogTitle     =   "Opening..."
      Filter          =   "Image Files (*.Zac)"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save As..."
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuToolBar 
      Caption         =   "&Tool Bar"
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmDrawMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExtFloodFill Lib "Gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Function Distance(X1 As Single, Y1 As Single, x2 As Single, y2 As Single) As Single
Distance = Sqr((X1 - x2) ^ 2 + (Y1 - y2) ^ 2)
End Function
Private Sub cmdReady_Click()
FrameHowTo.Visible = False
cmdReady.Visible = False
End Sub
Public Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static X1 As Single, Y1 As Single
Static fill
Static p
If Button = vbLeftButton Then
If Sing Then
    If Not Drwing Then
        Drwing = True
        PSet (X, Y), colr
    Else
       Line -(X, Y), colr
        Drwing = False
    End If
End If
End If
If Button = vbLeftButton Then
If Cntin Then
    If Not Drwing Then
        Drwing = True
        PSet (X, Y), colr
    Else
        Line -(X, Y), colr
    End If
End If
End If
If Fre And Button = vbLeftButton Then
    Drwing = True
    PSet (X, Y), colr
End If
If Button = vbLeftButton Then
If Rectangle Then
    If p = 0 Then
        PSet (X, Y), colr
        X1 = X: Y1 = Y
        p = 1
    ElseIf p = 1 Then
        Line -(X1, Y), colr
        Line -(X, Y), colr
        Line -(X, Y1), colr
        Line -(X1, Y1), colr
        p = 0
    End If
End If
End If
If Button = vbLeftButton Then
If Triangle Then
    If p = 0 Then
        PSet (X, Y), colr
        X1 = X: Y1 = Y
        p = 1
    ElseIf p = 1 Then
        Line -(X, Y), colr
        p = 2
    ElseIf p = 2 Then
        Line -(X, Y), colr
        Line -(X1, Y1), colr
        p = 0
    End If
End If
End If
If Button = vbLeftButton Then
If Center Then
FillStyle = 1
    If p = 0 Then
        PSet (X, Y), colr
        X1 = X: Y1 = Y
        p = 1
    ElseIf p = 1 Then
        Circle (X1, Y1), Distance(X, Y, X1, Y1), colr
        p = 0
    End If
End If
End If
If Button = vbLeftButton Then
If Diameter Then
FillStyle = 1
    If p = 0 Then
        PSet (X, Y), colr
        X1 = X: Y1 = Y
        p = 1
    ElseIf p = 1 Then
       Circle ((X1 + X) / 2, (Y1 + Y) / 2), Distance(X, Y, X1, Y1) / 2, colr
       p = 0
    End If
End If
End If
If E And Button = vbLeftButton Then
    Drwing = True
    PSet (X, Y), QBColor(15)
End If
If Button = vbLeftButton Then
If Ebox Then
    If p = 0 Then
        PSet (X, Y), colr
        X1 = X: Y1 = Y
        p = 1
    ElseIf p = 1 Then
        Line -(X1, Y), colr
        Line -(X, Y), colr
        Line -(X, Y1), colr
        Line -(X1, Y1), colr
        Line (X1, Y1)-(X, Y), QBColor(15), BF
        p = 0
    End If
End If
End If
If Button = vbRightButton Then
    frmDrawMain.FillColor = FrmDrawing.dlgColor.Color
    frmDrawMain.FillStyle = 0
    ExtFloodFill frmDrawMain.hdc, X, Y, frmDrawMain.Point(X, Y), 1
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Fre And Drwing Then
    frmDrawMain.Line -(X, Y), colr
End If
If E And Drwing Then
    frmDrawMain.Line -(X, Y), QBColor(15)
End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Fre Then Drwing = False
If E Then Drwing = False
End Sub
Private Sub mnuhelp_Click()
frmhelp.Visible = True
End Sub
Private Sub mnuHide_Click()
FrmDrawing.Visible = False
End Sub
Private Sub mnuShow_Click()
FrmDrawing.Visible = True
End Sub
Private Sub mnuNew_Click()
frmDrawMain.Line (0, 0)-(9999, 9999), QBColor(15), BF
frmDrawMain.Caption = "                                                  Zac's Paint Program!-New Project"
End Sub
Private Sub mnuOpen_Click()
On Error Resume Next
dlgOpen.ShowOpen
frmDrawMain.Picture = LoadPicture(dlgOpen.FileName)
frmDrawMain.Caption = "                                                      Zac's Paint Program!" & " " & dlgOpen.FileName
End Sub
Private Sub mnuSaveAs_Click()
On Error Resume Next
dlgSave.ShowSave
SavePicture frmDrawMain.Image, dlgSave.FileName
frmDrawMain.Caption = "                                                      Zac's Paint Program!" & " " & dlgSave.FileName
End Sub
