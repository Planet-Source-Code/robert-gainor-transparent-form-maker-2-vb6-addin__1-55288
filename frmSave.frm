VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Transparent Form"
   ClientHeight    =   1860
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4920
   Begin VB.CommandButton CancelButton 
      Caption         =   "Close"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Save"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Click on the Save button to save your new Transparent form. Click Cancel to return to the main program window."
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim strPicFile As String

Private Sub CancelButton_Click()
Unload Me

End Sub

Private Sub Form_Load()
StayOnTop Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim frm As Form
For Each frm In Forms
    If frm.Name = "frmRegion" Then
        Unload frm
        Set frm = Nothing
    ElseIf frm.Name = "frmMain" Then
        frm.Enabled = True
        frm.Visible = True
        SetFormOnTop frm
    End If
    
Next frm

End Sub

Private Sub OKButton_Click()
    Me.MousePointer = vbHourglass
    
    Dim fRegion As Form
        
    For Each fRegion In Forms
        If fRegion.Name = "frmRegion" Then
            If fRegion.SaveForm(strPicFile) = True Then
                Label1.Caption = "Your new form has been saved." & vbCrLf & "Click the Close Button to go back to the main program window"
                OKButton.Visible = False
                
            End If
            Exit For
        End If
    Next fRegion
    Me.MousePointer = vbDefault

    
End Sub

Public Sub SetPictureFile(strFile As String)
    strPicFile = strFile
End Sub

