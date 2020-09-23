Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public objVBE As VBIDE.VBE
Private Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2





Public Sub CenterForm(frmCenter As Form)
frmCenter.Left = (Screen.Width - frmCenter.Width) / 2
frmCenter.Top = (Screen.Height - frmCenter.Height) / 2
End Sub


Public Function GetFileName(ByVal strPath As String) As String
Dim lPosition As Long

lPosition = InStrRev(strPath, "\", Len(strPath))
GetFileName = Right(strPath, Len(strPath) - lPosition)

End Function

Public Function RemoveFileExt(ByVal strFile As String) As String
Dim lPosition As Long
lPosition = InStrRev(strFile, ".", Len(strFile))
RemoveFileExt = Left(strFile, lPosition - 1)
End Function

Public Function GetPathToFile(ByVal strPath As String) As String
Dim lPosition As Long

lPosition = InStrRev(strPath, "\", Len(strPath))
GetPathToFile = Left(strPath, lPosition)


End Function

Public Sub SetFormOnTop(ByRef frm As Form)

Dim lHandle As Long
'First we're going to retrieve the handle of this window
    ' "ThunderRT5Form" is the classname of a VB-window
    lHandle = FindWindow("ThunderRT5Form", frm.Caption)
    'Set this window to the foreground
    lHandle = SetForegroundWindow(lHandle)

End Sub

Public Sub StayOnTop(frm As Form)
  SetWindowPos frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
