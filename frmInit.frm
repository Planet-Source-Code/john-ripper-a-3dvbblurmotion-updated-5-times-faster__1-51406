VERSION 5.00
Begin VB.Form frmInit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Blur Motion in Pure VB Source"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFullScreen 
      Caption         =   "FullScreen Mode [ESC or Click to finish]"
      Height          =   255
      Left            =   30
      TabIndex        =   9
      Top             =   2130
      Width           =   4005
   End
   Begin VB.Frame Frame1 
      Caption         =   "Blur Palette"
      Height          =   1125
      Left            =   2160
      TabIndex        =   5
      Top             =   900
      Width           =   1845
      Begin VB.OptionButton OptPalette 
         Caption         =   "OnlyBlue31"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   750
         Width           =   1575
      End
      Begin VB.OptionButton OptPalette 
         Caption         =   "Cooler96"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   510
         Width           =   1575
      End
      Begin VB.OptionButton OptPalette 
         Caption         =   "Fire126"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1740
      Top             =   510
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "(Direct X 7.0 screen mode initialization)"
      Height          =   225
      Left            =   330
      TabIndex        =   10
      Top             =   2400
      Width           =   3705
   End
   Begin VB.Label Label5 
      Caption         =   "Tip: Compile this source and execute the EXE file for better performance!"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   2190
      TabIndex        =   4
      Top             =   30
      Width           =   1845
   End
   Begin VB.Label Label3 
      Caption         =   "Use Shift and Z, X, Y to slow down speed rotation"
      Height          =   225
      Left            =   30
      TabIndex        =   3
      Top             =   3060
      Width           =   4005
   End
   Begin VB.Label Label2 
      Caption         =   "Use Z, X, Y to increase speed rotation"
      Height          =   225
      Left            =   30
      TabIndex        =   2
      Top             =   2790
      Width           =   4005
   End
   Begin VB.Label Label1 
      Caption         =   "Select the Mesh:"
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   1755
   End
End
Attribute VB_Name = "frmInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub File1_Click()
    MeshName = File1.FileName
    If chkFullScreen.Value = vbChecked Then
        FullScreen = True
    Else
        FullScreen = False
    End If
    Timer1.Enabled = True
    
End Sub

Private Sub Form_Load()
    File1.Path = App.Path & "\Meshes"
    PaletteName = "Fire126"
End Sub

Private Sub OptPalette_Click(Index As Integer)
    PaletteName = OptPalette(Index).Caption
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    frmMain.Show vbModal
    Unload Me
End Sub
