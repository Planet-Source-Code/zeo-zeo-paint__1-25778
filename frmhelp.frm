VERSION 5.00
Begin VB.Form frmhelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "help"
   ClientHeight    =   4785
   ClientLeft      =   225
   ClientTop       =   1800
   ClientWidth     =   3675
   Icon            =   "frmhelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   3675
   Begin VB.Frame FrameHowTo 
      Caption         =   "How To Use Zac's Paint!"
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3615
      Begin VB.CommandButton cmdReady 
         Caption         =   "Ready to draw"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Made By: Zac Orns"
         Height          =   255
         Left            =   720
         TabIndex        =   9
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Start 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmhelp.frx":164A
         ForeColor       =   &H0000FF00&
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3285
      End
      Begin VB.Label Start 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "To use the lines click on the button, once clicked you will draw that type of line"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   3315
      End
      Begin VB.Label Start 
         BackColor       =   &H00808000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "To use the polygons click on the one you want to use, once clicked you will draw that type of polygon."
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   3330
      End
      Begin VB.Label Start 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "To fill a shape just right click inside the shape. To change the fill color click fill color."
         ForeColor       =   &H00FF80FF&
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   3315
      End
      Begin VB.Label Start 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click on color then choose"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   2640
         Width           =   1950
      End
      Begin VB.Label Start 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click on draw width to change the size"
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   3
         Top             =   2880
         Width           =   2790
      End
      Begin VB.Label Start 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click on erase then choose a side and erase."
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   2
         Top             =   3120
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReady_Click()
frmhelp.Visible = False
End Sub
