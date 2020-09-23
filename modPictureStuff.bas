Attribute VB_Name = "modPictureStuff"
Option Explicit



Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As RasterOpConstants) As Long

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
    ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Type Rect
    Width As Long
    Height As Long
End Type




Public Sub LoadPictureBox(FileName As String, pctOpen As PictureBox, pctSize As PictureBox)
Dim dblRatio As Double
Dim tempPic As StdPicture  'picture we're going to load
'Dim hmWidth As Long, hmHeight As Long 'size of the picture, converted to pixels
Dim picDC As Long 'the temp DC to select the picture into
Dim rctSize As Rect


pctOpen.Picture = LoadPicture

Set tempPic = LoadPicture(FileName)

'set the initial rectangle size
rctSize.Width = pctOpen.ScaleX(tempPic.Width, vbHimetric, vbTwips)
rctSize.Height = pctOpen.ScaleY(tempPic.Height, vbHimetric, vbTwips)

'Get the ratio of the picture
Select Case rctSize.Width
    Case Is >= rctSize.Height
        dblRatio = rctSize.Width / rctSize.Height
    Case Is < rctSize.Height
        dblRatio = rctSize.Height / rctSize.Width
        
End Select

'set the rectangle size
Select Case rctSize.Width
    Case Is >= rctSize.Height
        rctSize.Width = pctSize.ScaleWidth
        rctSize.Height = rctSize.Width / dblRatio
        If rctSize.Height > pctSize.ScaleHeight Then
            rctSize.Height = pctSize.ScaleHeight
            rctSize.Width = rctSize.Height * dblRatio
        End If
    Case Is < rctSize.Height
        rctSize.Height = pctSize.Height
        rctSize.Width = rctSize.Height / dblRatio
        If rctSize.Width > pctSize.ScaleHeight Then
            rctSize.Width = pctSize.ScaleWidth
            rctSize.Height = rctSize.Width * dblRatio
        End If
End Select


'set the size of pctopen to the right size and position
pctOpen.Width = rctSize.Width
pctOpen.Height = rctSize.Height

pctOpen.Left = (pctSize.ScaleWidth - pctOpen.Width) / 2
pctOpen.Top = (pctSize.ScaleHeight - pctOpen.Height) / 2

'set the mode
SetStretchBltMode pctOpen.hDC, 4

'create the backbuffer
picDC = CreateCompatibleDC(pctOpen.hDC)

'select the bitmap onto the temp DC
DeleteObject SelectObject(picDC, tempPic.Handle)
StretchBlt pctOpen.hDC, 0, 0, pctOpen.Width, pctOpen.Height, picDC, 0, 0, pctOpen.ScaleX(tempPic.Width, vbHimetric, vbTwips), pctOpen.ScaleY(tempPic.Height, vbHimetric, vbTwips), vbSrcCopy
pctOpen.Refresh
'clean up
DeleteDC picDC

End Sub
