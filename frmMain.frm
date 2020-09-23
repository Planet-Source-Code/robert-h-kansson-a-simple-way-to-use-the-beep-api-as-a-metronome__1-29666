VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Metronome"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider sTempo 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Min             =   100
      Max             =   1000
      SelStart        =   100
      TickStyle       =   3
      Value           =   100
   End
   Begin VB.Timer T1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3240
      Top             =   240
   End
   Begin VB.CheckBox chkStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin MSComctlLib.Slider sTone 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Min             =   100
      Max             =   1000
      SelStart        =   100
      TickStyle       =   3
      Value           =   100
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Created by Robert@Hakansson.nu 2001 "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Tone:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Tempo:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Shape S1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   1920
      Shape           =   3  'Circle
      Top             =   240
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   2055
      Left            =   120
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program is created by Robert Håkansson 2001
'Contact: robert@kltdata.se



Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Dim Tempo As Integer
Dim Tone As Integer



Private Sub chkStart_Click()
    If chkStart.Caption = "Start" Then
        T1.Enabled = True
        chkStart.Caption = "Stop"
    Else
        T1.Enabled = False
        chkStart.Caption = "Start"
        
    End If
End Sub

Private Sub Form_Load()
    Call Start
    sTone = 500
    sTempo = 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Beep 1000, 150
    Beep 200, 150
    Beep 400, 150
    Beep 600, 150
    Beep 500, 150
End Sub

Private Sub sTempo_Change()
    Tempo = sTempo.Value
End Sub

Private Sub sTempo_Scroll()
    Tempo = sTempo.Value
End Sub

Private Sub sTone_Change()
    Tone = sTone.Value
End Sub

Private Sub sTone_Scroll()
    Tone = sTone.Value
End Sub

Private Sub T1_Timer()
    T1.Interval = Tempo
    S1.BackColor = vbBlack
    S1.Refresh
    Beep Tone, 50
    S1.BackColor = vbGreen
    S1.Refresh
End Sub


Private Sub Start()
    Beep 500, 150
    Beep 600, 150
    Beep 400, 150
    Beep 200, 150
    Beep 1000, 200
End Sub
