VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AUTO DRAWING"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5685
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5685
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   510
      ScaleWidth      =   3855
      TabIndex        =   2
      Top             =   5445
      Width           =   3885
      Begin VB.Label lblSelectedColour 
         BackColor       =   &H0000FFFF&
         Height          =   420
         Left            =   45
         TabIndex        =   3
         Top             =   45
         Width           =   405
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2415
      Top             =   6015
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&START"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3960
      TabIndex        =   1
      Top             =   5505
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C00000&
      Height          =   5430
      Left            =   15
      ScaleHeight     =   5430
      ScaleWidth      =   5700
      TabIndex        =   0
      Top             =   -15
      Width           =   5700
      Begin VB.Label lblAxis 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4380
         TabIndex        =   4
         Top             =   105
         Width           =   45
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FSO As FileSystemObject
Private TS As TextStream
Private StrPoints As String, strX As String, strY As String
Private dblSelectedColour As Double
Private Sub cmdStart_Click()
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Set FSO = New FileSystemObject
    Set TS = FSO.OpenTextFile(App.Path & "\Points.pan", ForReading)
    Picture1.DrawWidth = 2
    dblSelectedColour = vbYellow
End Sub
Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblSelectedColour.BackColor = Picture2.Point(X, Y)
    dblSelectedColour = lblSelectedColour.BackColor
End Sub
Private Sub Timer1_Timer()
    On Error GoTo Finish
    StrPoints = TS.ReadLine
    strX = Mid(StrPoints, 1, InStr(1, StrPoints, ":") - 1)
    strY = Mid(StrPoints, InStr(1, StrPoints, ":") + 1)
    lblAxis.Caption = "X=" & strX & " : Y=" & strY
    Picture1.PSet (strX, strY), dblSelectedColour
Finish:
If Err.Number = 62 Then
    If MsgBox("Do You Want To Continue ?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
        Set TS = FSO.OpenTextFile(App.Path & "\Points.pan", ForReading)
        Picture1.Cls
    Else
        MsgBox "Please...Post Your Vote", vbInformation, "Bye"
        Timer1.Enabled = False
        End
        Exit Sub
    End If
End If
End Sub
