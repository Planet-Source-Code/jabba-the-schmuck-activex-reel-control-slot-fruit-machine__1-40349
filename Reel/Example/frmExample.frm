VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkHS 
      Caption         =   "Horizontal Spin"
      Height          =   375
      Left            =   480
      TabIndex        =   31
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CheckBox chkA 
      Caption         =   "Align on Stop"
      Height          =   270
      Left            =   5760
      TabIndex        =   23
      Top             =   5160
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00E0E0E0&
      Caption         =   "STOP"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Spin Type"
      Height          =   4215
      Left            =   5760
      TabIndex        =   7
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton cmdSpeed 
         Caption         =   "Speed = 25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CheckBox chkEST 
         Caption         =   "Enable Spin Type"
         Height          =   270
         Left            =   360
         TabIndex        =   33
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txtStopPos 
         Height          =   390
         Left            =   1320
         TabIndex        =   27
         Text            =   "0"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtMinCycles 
         Height          =   390
         Left            =   1320
         TabIndex        =   26
         Text            =   "0"
         Top             =   2640
         Width           =   735
      End
      Begin VB.OptionButton optST 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   2040
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton optST 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   1200
         TabIndex        =   19
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton optST 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton optST 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   17
         Top             =   360
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton optST 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton optST 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "If Spin Type is enabled, these 3 options will be ignored!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   240
         TabIndex        =   32
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "0 is random / default"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   720
         TabIndex        =   28
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Stop on reel position:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Minimum full cycles for spin:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   24
         Top             =   2520
         Width           =   975
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   2760
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3015
      Begin VB.CheckBox chkNF 
         Caption         =   "Natural Finish"
         Height          =   270
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkEH 
         Caption         =   "Enable Hold"
         Height          =   270
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CommandButton cmdReels 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   495
      End
      Begin VB.CommandButton cmdReels 
         Caption         =   "5"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Number of Reels"
         Height          =   495
         Left            =   960
         TabIndex        =   21
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   2880
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lblReels 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   2880
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   2415
      Begin VB.CheckBox chkDirection 
         Caption         =   "Up"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1560
         TabIndex        =   30
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton cmdNudge 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nudge"
         Height          =   405
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdSpin 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Spin"
         Height          =   1005
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cmbReel 
         Height          =   390
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "CLICK ON A REEL !"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   480
      TabIndex        =   35
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   9000
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label lblReelPos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   4
      Left            =   5160
      TabIndex        =   4
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblReelPos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   3
      Left            =   4680
      TabIndex        =   3
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblReelPos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   2
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblReelPos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblReelPos 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   0
      Top             =   3840
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9000
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line lblPointer 
      Index           =   0
      X1              =   3480
      X2              =   3480
      Y1              =   4200
      Y2              =   4440
   End
   Begin VB.Line lblPointer 
      Index           =   4
      X1              =   5400
      X2              =   5400
      Y1              =   4200
      Y2              =   4440
   End
   Begin VB.Line lblPointer 
      Index           =   3
      X1              =   4920
      X2              =   4920
      Y1              =   4200
      Y2              =   4440
   End
   Begin VB.Line lblPointer 
      Index           =   2
      X1              =   4440
      X2              =   4440
      Y1              =   4200
      Y2              =   4440
   End
   Begin VB.Line lblPointer 
      Index           =   1
      X1              =   3960
      X2              =   3960
      Y1              =   4200
      Y2              =   4440
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEH_Click()

    ucReel1.EnableHold = Not ucReel1.EnableHold

End Sub

Private Sub chkNF_Click()

    ucReel1.NaturalFinish = Not ucReel1.NaturalFinish
    
End Sub

Private Sub chkHS_Click()

    ucReel1.HorizontalSpin = Not ucReel1.HorizontalSpin

End Sub

Private Sub cmdNudge_Click()

    ucReel1.Nudge Val(cmbReel.Text), chkDirection.Value

End Sub

Private Sub cmdReels_Click(Index As Integer)

    If Index = 0 Then
        lblReels.Caption = Val(lblReels.Caption) + 1
        If lblReels.Caption = "5" Then
            cmdReels(0).Enabled = False
        ElseIf lblReels.Caption = "2" Then
            cmdReels(1).Enabled = True
        End If
        ucReel1.Reels = Val(lblReels.Caption)
    Else
        lblReels.Caption = Val(lblReels.Caption) - 1
        If lblReels.Caption = "1" Then
            cmdReels(1).Enabled = False
        ElseIf lblReels.Caption = "4" Then
            cmdReels(0).Enabled = True
        End If
         ucReel1.Reels = Val(lblReels.Caption)
    End If

End Sub

Private Sub cmdSpeed_Click()

    If cmdSpeed.Caption = "Speed = 25" Then
        cmdSpeed.Caption = "Speed = 35"
    ElseIf cmdSpeed.Caption = "Speed = 35" Then
        cmdSpeed.Caption = "Speed = 15"
    ElseIf cmdSpeed.Caption = "Speed = 15" Then
        cmdSpeed.Caption = "Speed = 25"
    End If

End Sub

Private Sub cmdSpin_Click()

Dim i As Long
    If chkEST.Value = 1 Then
        For i = 0 To optST.UBound
            If optST(i).Value = True Then
                ucReel1.Spin Val(cmbReel.Text), i + 1
                Exit For
            End If
        Next i
    Else
        ucReel1.Spin Val(cmbReel.Text), 3, Val(txtStopPos.Text), , Val(txtMinCycles.Text), Val(Right$(cmdSpeed.Caption, 2))
    End If

End Sub

Private Sub cmdStop_Click()

    ucReel1.StopSpinning (chkA.Value)

End Sub

Private Sub Form_Load()

Dim i As Byte
    For i = 0 To ucReel1.Reels
        cmbReel.AddItem i
    Next i

    With ucReel1
        .NaturalFinish = chkNF.Value
        .EnableHold = chkEH.Value
        .HorizontalSpin = chkHS.Value
    End With

End Sub

Private Sub ucReel1_OnReelStop(ByVal whatReel As Integer)

    lblReelPos(whatReel).Caption = ucReel1.ReelPosition(whatReel)

End Sub
