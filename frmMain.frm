VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transparent Form Maker"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExample 
      Caption         =   "View Example Form"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoadPicture 
      Caption         =   "Load Picture"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   3000
      Width           =   975
   End
   Begin VB.PictureBox pctTransparentColor 
      Height          =   375
      Left            =   1560
      ScaleHeight     =   315
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdShowForm 
      Caption         =   "Show The Form"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   4200
      ScaleHeight     =   3315
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.PictureBox pctOpen 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   360
         ScaleHeight     =   2415
         ScaleWidth      =   3615
         TabIndex        =   6
         Top             =   480
         Width           =   3615
      End
   End
   Begin MSComDlg.CommonDialog CmnDialog 
      Left            =   3600
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   $"frmMain.frx":0000
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Transparent Color:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strPicFile As String
Dim strLastUsedPath As String


Private Sub cmdExample_Click()
Dim fExample As frmExample

Set fExample = New frmExample

fExample.Show
Set fExample = Nothing

End Sub

Private Sub cmdLoadPicture_Click()
On Error GoTo ErrorHandle
Dim proj As VBProject

Set proj = objVBE.ActiveVBProject



With CmnDialog
    .DialogTitle = "Select Picture"
    .Filter = "All Picture Files|*.bmp;*.jpg;*.gif"
    If proj.FileName <> "" Then
        .InitDir = GetPathToFile(proj.FileName)
    ElseIf strLastUsedPath <> "" Then
        .InitDir = strLastUsedPath
    Else
        .InitDir = objVBE.LastUsedPath
    End If
    .FileName = ""
    .ShowOpen
    strLastUsedPath = GetPathToFile(.FileName)
    strPicFile = .FileName
End With

LoadPictureBox strPicFile, pctOpen, Picture1
cmdShowForm.Enabled = True

Exit Sub
ErrorHandle:
End Sub

Private Sub cmdShowForm_Click()
Dim fRegion As frmRegion
Dim fSave As frmSave

Me.MousePointer = vbHourglass
Me.Enabled = False

Set fRegion = New frmRegion
Load fRegion
fRegion.SetPicture strPicFile, pctTransparentColor.BackColor
Me.MousePointer = vbDefault
Me.Visible = False
fRegion.Show

Set fSave = New frmSave
Load fSave

fSave.SetPictureFile strPicFile

fSave.Left = fRegion.Left - fSave.Width
fSave.Top = fRegion.Top - fSave.Height
If fSave.Left < 0 Then
    fSave.Left = 1000
End If
If fSave.Top < 0 Then
    fSave.Top = 1000
End If

fSave.Show
Set fSave = Nothing
Set fRegion = Nothing

End Sub

Private Sub pctOpen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
pctTransparentColor.BackColor = pctOpen.Point(X, Y)
End Sub

