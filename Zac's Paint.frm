VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmDrawing 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ToolBar"
   ClientHeight    =   7785
   ClientLeft      =   780
   ClientTop       =   1890
   ClientWidth     =   3210
   FillStyle       =   0  'Solid
   ForeColor       =   &H00404040&
   Icon            =   "Zac's Paint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   519
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   214
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar ToolbarColorChoice 
      Height          =   8730
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   15399
      ButtonWidth     =   1984
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Black"
            Key             =   "C0"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Blue"
            Key             =   "C1"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Green"
            Key             =   "C2"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cyan"
            Key             =   "C3"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Red"
            Key             =   "C4"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Magenta"
            Key             =   "C5"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Yellow"
            Key             =   "C6"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Gray"
            Key             =   "C8"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Light Blue "
            Key             =   "C9"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Light Green"
            Key             =   "C10"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Light  Cyan"
            Key             =   "C11"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Light Red"
            Key             =   "C12"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Light Magenta"
            Key             =   "C13"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Light Yellow"
            Key             =   "C14"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture6 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         ScaleHeight     =   195
         ScaleWidth      =   75
         TabIndex        =   54
         Top             =   7080
         Width           =   135
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FF00FF&
         Height          =   135
         Left            =   0
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   53
         Top             =   6600
         Width           =   135
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   52
         Top             =   6000
         Width           =   135
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   51
         Top             =   5520
         Width           =   135
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H0000FF00&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   50
         Top             =   4920
         Width           =   135
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   49
         Top             =   4440
         Width           =   135
      End
      Begin VB.PictureBox Gray 
         BackColor       =   &H00808080&
         Height          =   375
         Left            =   120
         ScaleHeight     =   508.846
         ScaleMode       =   0  'User
         ScaleWidth      =   75
         TabIndex        =   48
         Top             =   3840
         Width           =   135
      End
      Begin VB.PictureBox Yellow 
         BackColor       =   &H00008080&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   47
         Top             =   3360
         Width           =   135
      End
      Begin VB.PictureBox Magenta 
         BackColor       =   &H00800080&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   46
         Top             =   2760
         Width           =   135
      End
      Begin VB.PictureBox Red 
         BackColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   45
         Top             =   2280
         Width           =   135
      End
      Begin VB.PictureBox Cyan 
         BackColor       =   &H00808000&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   44
         Top             =   1680
         Width           =   135
      End
      Begin VB.PictureBox Green 
         BackColor       =   &H00008000&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   43
         Top             =   1200
         Width           =   135
      End
      Begin VB.PictureBox Blue 
         BackColor       =   &H00800000&
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   42
         Top             =   600
         Width           =   135
      End
      Begin VB.PictureBox Black 
         BackColor       =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   75
         TabIndex        =   41
         Top             =   120
         Width           =   135
      End
   End
   Begin ComctlLib.Toolbar ToolbarDrawWidth 
      Height          =   5055
      Left            =   0
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   8916
      ButtonWidth     =   9472
      ButtonHeight    =   8758
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "1"
            Key             =   "D1"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "2"
            Key             =   "D2"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "3"
            Key             =   "D3"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "5"
            Key             =   "D4"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Custom"
            Key             =   "D5"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ComctlLib.Toolbar ToolbarErase 
      Height          =   630
      Left            =   0
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   1111
      ButtonWidth     =   1111
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "1"
            Key             =   "C1"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "2"
            Key             =   "C2"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "3"
            Key             =   "C3"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "5"
            Key             =   "C4"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Custom"
            Key             =   "C5"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdFillColor 
      Caption         =   "Set Fill Col&or"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   5640
      Width           =   1815
   End
   Begin VB.HScrollBar HScrollRed 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      SmallChange     =   5
      TabIndex        =   32
      Top             =   840
      Width           =   1815
   End
   Begin VB.HScrollBar HScrollGreen 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      SmallChange     =   5
      TabIndex        =   31
      Top             =   1320
      Width           =   1815
   End
   Begin VB.HScrollBar HScrollBlue 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      SmallChange     =   5
      TabIndex        =   30
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Frame FrameLines 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lines"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   26
      Top             =   2880
      Width           =   2055
      Begin VB.CommandButton cmdLines 
         Caption         =   "&Single Line (Point-Point)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "Click to draw single lines"
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdLines 
         Caption         =   "&Continuous Line (Point-Point-Point)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Click to draw continuous lines"
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdLines 
         Caption         =   "&Free Hand Drawing"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "click to draw free-handed"
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame FramePolygons 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Polygons"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   16
      Left            =   0
      TabIndex        =   21
      Top             =   4080
      Width           =   2055
      Begin VB.CommandButton cmdPolygons 
         Caption         =   "&Rectangle"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Click to draw rectangles"
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdPolygons 
         Caption         =   "&Triangle"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdPolygons 
         Caption         =   "Circle - Center, &Radius"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Click to draw circles"
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdPolygons 
         Caption         =   "Circle - &Diameter"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Click to draw circles"
         Top             =   960
         Width           =   1815
      End
   End
   Begin VB.Frame FrameDrawWidth 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Draw Width"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   19
      Top             =   6600
      Width           =   2055
      Begin VB.CommandButton cmdDrawWidth 
         Caption         =   "&Draw Width"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame FrameColor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Draw Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   17
      Top             =   6000
      Width           =   2055
      Begin VB.CommandButton cmdColorSelect 
         Caption         =   "Set Draw &Color"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame FrameFillColor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fill Color"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   16
      Top             =   5400
      Width           =   2055
   End
   Begin VB.PictureBox PictureCCSave1 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   2040
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
      Begin VB.CommandButton Use1 
         Caption         =   "Use"
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.PictureBox PictureCCSave2 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   2040
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   12
      Top             =   3840
      Width           =   1215
      Begin VB.CommandButton Use2 
         Caption         =   "Use"
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.PictureBox PictureCCSave3 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   2040
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
      Begin VB.CommandButton Use3 
         Caption         =   "Use"
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.PictureBox PictureCCSave4 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   2040
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
      Begin VB.CommandButton Use4 
         Caption         =   "Use"
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.PictureBox PictureCCSave0 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   2040
      ScaleHeight     =   1275
      ScaleWidth      =   1155
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
      Begin VB.CommandButton Use0 
         Caption         =   "Use"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
   End
   Begin VB.Timer Timer2 
      Left            =   3240
      Top             =   7320
   End
   Begin VB.Frame FrameEraser 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eraser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   2055
      Begin VB.CommandButton cmdLines 
         Caption         =   "Rec&tangle"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdLines 
         Caption         =   "Fr&ee Hand"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Click to erase"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame FrameCustomColor 
      BackColor       =   &H00808080&
      Caption         =   "Custom Color"
      Height          =   2535
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton cmdSetCustomColor 
         Caption         =   "&Set Draw Color"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Click to use the custom color you made"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Shape shpCustomColorPreview 
         BackStyle       =   1  'Opaque
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   240
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   840
         TabIndex        =   37
         Top             =   600
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   720
         TabIndex        =   36
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   840
         TabIndex        =   35
         Top             =   1560
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Zac's Paint.frx":17B2
      Height          =   1215
      Left            =   2040
      TabIndex        =   38
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "FrmDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function Done()
Dim X
For X = FrmDrawing.Timer2.Interval = 2 To 200 Step 2
    frmDrawMain.Width = frmDrawMain.Width - 60
    frmDrawMain.Height = frmDrawMain.Height - 60
    frmDrawMain.Top = frmDrawMain.Top + 100
    frmDrawMain.Left = frmDrawMain.Left + 10
Next X
If X >= 199 Then
    End
End If
End Function
Private Sub cmdColorSelect_Click()
ToolbarColorChoice.Visible = True
ToolbarDrawWidth.Visible = False
ToolbarErase.Visible = False
End Sub
Private Sub cmdDrawWidth_Click()
ToolbarDrawWidth.Visible = True
ToolbarColorChoice.Visible = False
ToolbarErase.Visible = False
End Sub
Private Sub cmdFillColor_Click()
dlgColor.ShowColor
End Sub
Private Sub cmdLines_Click(Index As Integer)
Drwing = False
Sing = False
Cntin = False
Fre = False
Triangle = False
Rectangle = False
Center = False
Diameter = False
E = False
Ebox = False
ToolbarDrawWidth.Visible = False
ToolbarColorChoice.Visible = False
Select Case Index
    Case 0
        Sing = True
    Case 1
        Cntin = True
    Case 2
        Fre = True
    Case 3
        ToolbarErase.Visible = True
        E = True
    Case 4
        Ebox = True
End Select
End Sub
Private Sub cmdpolygons_Click(Index As Integer)
Drwing = False
Sing = False
Cntin = False
Fre = False
Triangle = False
Rectangle = False
Center = False
Diameter = False
E = False
Ebox = False
Select Case Index
    Case 0
        Rectangle = True
    Case 1
        Triangle = True
    Case 2
        Center = True
    Case 3
        Diameter = True
    End Select
End Sub
Private Sub cmdSetCustomColor_Click()
colr = RGB(HScrollRed.Value, HScrollGreen.Value, HScrollBlue.Value)
FillColor = RGB(HScrollRed.Value, HScrollGreen.Value, HScrollBlue.Value)
End Sub
Private Sub Form_Load()
frmDrawMain.Visible = True
End Sub
Private Sub Form_Resize()
FrameEraser.Height = FrmDrawing.Height
End Sub
Private Sub HScrollBlue_Change()
shpCustomColorPreview.FillColor = RGB(HScrollRed.Value, HScrollGreen.Value, HScrollBlue.Value)
End Sub
Private Sub HScrollGreen_Change()
shpCustomColorPreview.FillColor = RGB(HScrollRed.Value, HScrollGreen.Value, HScrollBlue.Value)
End Sub
Private Sub HScrollRed_Change()
shpCustomColorPreview.FillColor = RGB(HScrollRed.Value, HScrollGreen.Value, HScrollBlue.Value)
End Sub
Private Sub mnuExit_Click()
FrmDrawing.Visible = False
Done
End Sub
Private Sub PictureCCSave0_DblClick()
PictureCCSave0.BackColor = colr
End Sub
Private Sub PictureCCSave1_Click()
PictureCCSave1.BackColor = colr
End Sub
Private Sub PictureCCSave2_Click()
PictureCCSave2.BackColor = colr
End Sub
Private Sub PictureCCSave3_Click()
PictureCCSave3.BackColor = colr
End Sub
Private Sub PictureCCSave4_Click()
PictureCCSave4.BackColor = colr
End Sub
Private Sub ToolbarColorChoice_ButtonClick(ByVal Button As ComctlLib.Button)
FillColor = colr
Select Case Button.Key
    Case "C0"
        colr = QBColor(0)
    Case "C1"
        colr = QBColor(1)
    Case "C2"
        colr = QBColor(2)
    Case "C3"
        colr = QBColor(3)
    Case "C4"
        colr = QBColor(4)
    Case "C5"
        colr = QBColor(5)
    Case "C6"
        colr = QBColor(6)
    Case "C8"
        colr = QBColor(8)
    Case "C9"
        colr = QBColor(9)
    Case "C10"
        colr = QBColor(10)
    Case "C11"
        colr = QBColor(11)
    Case "C12"
        colr = QBColor(12)
    Case "C13"
        colr = QBColor(13)
    Case "C14"
        colr = QBColor(14)
End Select
ToolbarColorChoice.Visible = False
End Sub
Private Sub ToolbarDrawWidth_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "D1"
        frmDrawMain.DrawWidth = 1
    Case "D2"
        frmDrawMain.DrawWidth = 2
    Case "D3"
        frmDrawMain.DrawWidth = 3
    Case "D4"
        frmDrawMain.DrawWidth = 5
    Case "D5"
        On Error GoTo Handle
        Dw = InputBox("Enter a drawing width between 1-50", "Draw Width", 1)
        frmDrawMain.DrawWidth = Dw
        If Val(Dw) > 50 Then
            frmDrawMain.DrawWidth = 1
        End If
Handle:
End Select
ToolbarDrawWidth.Visible = False
End Sub
Private Sub ToolbarErase_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "C1"
        frmDrawMain.DrawWidth = 1
    Case "C2"
        frmDrawMain.DrawWidth = 2
    Case "C3"
        frmDrawMain.DrawWidth = 3
    Case "C4"
        frmDrawMain.DrawWidth = 5
    Case "C5"
        On Error GoTo Handle
        Dw = InputBox("Enter a drawing width between 1-50", "Draw Width", 1)
        frmDrawMain.DrawWidth = Dw
        If Val(Dw) > 50 Then
            frmDrawMain.DrawWidth = 1
        End If
Handle:
End Select
ToolbarErase.Visible = False
End Sub
Private Sub Use0_Click()
colr = PictureCCSave0.BackColor
End Sub
Private Sub Use1_Click()
colr = PictureCCSave1.BackColor
End Sub
Private Sub Use2_Click()
colr = PictureCCSave2.BackColor
End Sub
Private Sub Use3_Click()
colr = PictureCCSave3.BackColor
End Sub
Private Sub Use4_Click()
colr = PictureCCSave4.BackColor
End Sub
