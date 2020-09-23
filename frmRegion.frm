VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRegion 
   BorderStyle     =   0  'None
   ClientHeight    =   5325
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRegion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long

Dim bytRegion() As Byte
Dim nBytes As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


Private Sub MakeRegion(ByVal TrnsColor As Long)
Me.BorderStyle = 0
Dim ScaleSize As Long
Dim Width, Height As Long
Dim rgnMain As Long
Dim X, Y As Long
Dim rgnPixel As Long
Dim RGBColor As Long
Dim dcMain As Long
Dim bmpMain As Long

'save the current scalemode of the form
ScaleSize = Me.ScaleMode

'set the form scalemode to pixels
Me.ScaleMode = 3

'get the size of the picture
Width = Me.ScaleX(Me.Picture.Width, vbHimetric, vbPixels)
Height = Me.ScaleY(Me.Picture.Height, vbHimetric, vbPixels)

'size the form to the size of the picture
Me.Width = Width * Screen.TwipsPerPixelX
Me.Height = Height * Screen.TwipsPerPixelY

'create the transparent region
rgnMain = CreateRectRgn(0, 0, Width, Height)

'create a device context
dcMain = CreateCompatibleDC(Me.hDC)

'save the handle to the form picture
bmpMain = SelectObject(dcMain, Me.Picture.Handle)

'check each pixel of the picture
'if the pixel matches the transparent color then
'remove it from the rgnMain
For Y = 0 To Height
    For X = 0 To Width
        RGBColor = GetPixel(dcMain, X, Y)
        If RGBColor = TrnsColor Then
            rgnPixel = CreateRectRgn(X, Y, X + 1, Y + 1)
            CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR
            DeleteObject rgnPixel
        End If
    Next X
Next Y

'delete the dcmain and the bmpMain
SelectObject dcMain, bmpMain
DeleteDC dcMain
DeleteObject bmpMain

'Save the region data for later use in creating the form's code
'and set the window region to rgnmain
If rgnMain <> 0 Then
 nBytes = GetRegionData(rgnMain, 0, ByVal 0&)
    If nBytes > 0 Then
        ReDim bytRegion(0 To nBytes - 1)
        nBytes = GetRegionData(rgnMain, nBytes, bytRegion(0))
    End If
    SetWindowRgn Me.hWnd, rgnMain, True
    CenterForm Me
End If
'set the scalemode of the form back to the original
Me.ScaleMode = ScaleSize
'for debug purposes
Debug.Print Me.Width
Debug.Print Me.Height

End Sub

Public Sub SetPicture(FileName As String, ClrTransparent As Long)

Set Me.Picture = LoadPicture(FileName)
MakeRegion ClrTransparent


End Sub

Private Sub Form_DblClick()
Dim frm As Form
For Each frm In Forms
    If frm.Name = "frmSave" Then
        Unload frm
    End If
Next frm
        
Unload Me

End Sub

Public Function SaveForm(strPicFile As String) As Boolean
'when this add-in is run from the VB IDE there is a bug that will not allow
'the new form to load a picture
'when you compile this into a dll file the loadpicture function works fine
'Check out this webpage for a better explaination

'http://groups.google.com/groups?hl=en&lr=&ie=UTF-8&threadm=uHdERNtG%24GA.176%40cppssbbsa02.microsoft.com&rnum=1&prev=/groups%3Fq%3DVBComponent%2BProperty%2BObject%2BPicture%26hl%3Den%26lr%3D%26ie%3DUTF-8%26selm%3DuHdERNtG%2524GA.176%2540cppssbbsa02.microsoft.com%26rnum%3D1

On Error Resume Next
Dim fWaiting As frmWaiting
Dim i As Long
Dim LineCount As Long
Dim CurrentProc As Long
Dim CurrentByte As Long
Dim strName As String
Dim cmponent As VBComponent
Dim proj As VBProject
Dim strCode As String

Me.MousePointer = vbHourglass



'get the name of the new form
strName = InputBox("Enter a name for your new form", "SSE Transparent Form Creator")

'check the name of the new form and for any spaces in the name of the form
If strName = "" Then
    Me.MousePointer = vbDefault
    SaveForm = False
    Exit Function
ElseIf InStr(1, strName, " ") Then
    MsgBox "The Name of the form cannot have any spaces", vbCritical, "SSE Transparent Form Maker"
    Me.MousePointer = False
    SaveForm = False
    Exit Function
End If

'Show frmWaiting while the transparent form is being built
Set fWaiting = New frmWaiting
fWaiting.Show

'get the current vb project
Set proj = objVBE.ActiveVBProject

'create a new form component in the project
Set cmponent = proj.VBComponents.Add(vbext_ct_VBForm)

'set the form component's name
cmponent.Name = strName

'set the component's properties
With cmponent
    .Properties("BorderStyle") = 0
    .Properties("ControlBox") = False
    .Properties("ShowInTaskbar") = False
    .Properties("Caption") = False
    .Properties("Height") = Me.Height
    .Properties("Width") = Me.Width
    Set .Properties("Picture").Object = LoadPicture(strPicFile)
End With

'used for debug purposes
Debug.Print "Region Height " & Me.Height
Debug.Print "Region Width " & Me.Width

Debug.Print "Form Height " & cmponent.Properties("Height")
Debug.Print "Form Width " & cmponent.Properties("Width")

'create the code for the new form
strCode = strCode & "'API Declares" & vbCrLf & _
                    "Private Declare Function ExtCreateRegion Lib ""gdi32"" (lpXform As Any, ByVal nCount As Long, lpRgnData As Any) As Long" & vbCrLf & _
                    "Private Declare Function SetWindowRgn Lib ""user32"" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long" & vbCrLf & _
                    "Private Declare Function SendMessage Lib ""user32"" Alias ""SendMessageA"" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long" & vbCrLf & _
                    "Private Declare Function ReleaseCapture Lib ""user32"" () As Long" & vbCrLf & _
                    vbCrLf & _
                    "Dim bytRegion(" & nBytes - 1 & ") As Byte" & vbCrLf & _
                    "Dim nBytes As Long" & vbCrLf & _
                    vbCrLf & _
                    "Private Sub Form_Load()" & vbCrLf & _
                    "    Dim rgnMain as Long" & vbCrLf & _
                    vbCrLf & _
                    "    nBytes = " & nBytes & vbCrLf & _
                    vbCrLf & _
                    "    LoadBytes1" & vbCrLf & _
                    vbCrLf & _
                    "    rgnMain = ExtCreateRegion(ByVal 0&, nBytes, bytRegion(0))" & vbCrLf & _
                    "    SetWindowRgn Me.hwnd, rgnMain, True" & vbCrLf & _
                    vbCrLf & _
                    "End Sub" & vbCrLf & _
                    vbCrLf

'This is written in such a way as to not exceed the 64kbyte limit on a procedure
'CurrentProc is the current procedure being written
'CurrentByte is the current byte that is being written from nbytes array

CurrentProc = 1
CurrentByte = 0
Do Until CurrentByte = nBytes - 1
    
        strCode = strCode & "Private Sub LoadBytes" & CurrentProc & "()" & vbCrLf & _
                            vbCrLf
        LineCount = 1
        Do Until LineCount = 3000 Or CurrentByte = nBytes - 1
            If bytRegion(CurrentByte) <> 0 Then
                strCode = strCode & "bytregion(" & CurrentByte & ")= " & bytRegion(CurrentByte) & vbCrLf
                LineCount = LineCount + 1
            End If
            CurrentByte = CurrentByte + 1
            
        Loop
        CurrentProc = CurrentProc + 1
        If CurrentByte < nBytes - 1 Then
            strCode = strCode & "LoadBytes" & CurrentProc & vbCrLf
        End If
        strCode = strCode & "End Sub" & vbCrLf
        
Loop

strCode = strCode & vbCrLf & _
          "Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)" & vbCrLf & _
          "'Next two lines enable window drag from anywhere on form.  Remove them" & vbCrLf & _
          "'to allow window drag from title bar only." & vbCrLf & _
          vbCrLf & _
          "ReleaseCapture" & vbCrLf & _
          "SendMessage Me.hWnd, &HA1, 2, 0&" & vbCrLf & _
          "End Sub" & vbCrLf & vbCrLf & _
          "Private Sub Form_DblClick()" & vbCrLf & vbCrLf & _
          "Unload Me" & vbCrLf & vbCrLf & _
          "End Sub"

'add the code to the new form component
cmponent.CodeModule.AddFromString (strCode)

'show the new form component in the VB IDE
cmponent.CodeModule.CodePane.Show

'show the new form component's designer window
cmponent.DesignerWindow.Visible = True

'cleanup objects used
Set cmponent = Nothing
Set proj = Nothing

'unload the waiting form to show that the process is finished
Unload fWaiting
Set fWaiting = Nothing

'find the frmSave form and set it to the forground
Dim frm As Form
For Each frm In Forms
    If frm.Name = "frmSave" Then
        SetFormOnTop frm
    End If
Next frm

'set the mouse pointer back to the arrow pointer
Me.MousePointer = vbDefault

'return true
SaveForm = True

Exit Function
ErrorHandle:
MsgBox Err.Number & " " & Err.Description
Me.MousePointer = vbDefault
If Not fWaiting Is Nothing Then
    Unload fWaiting
End If
SaveForm = False
End Function

Private Sub Form_Load()
StayOnTop Me

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
