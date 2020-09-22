Attribute VB_Name = "DirectX"
Option Explicit
'------------------------------------------------------------------------------------------------------
'These functions are used to copy a non bitmap onto a DirectDraw Surface
Private Const SRCCOPY = &HCC0020
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'------------------------------------------------------------------------------------------------------
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public DX As New DirectX7

Public DDraw As DirectDraw7

Public PrimarySurface As DirectDrawSurface7
Public PrimarySurfaceDescription As DDSURFACEDESC2

Public BackBuffer As DirectDrawSurface7
Public BackBufferDescription As DDSURFACEDESC2

Private DXhWnd As Long
Public SRECT As RECT
Private Const PI As Double = 3.14159265358979 'The definition of PI
Public Function InitiateDX(DXhWnd As Long) As Boolean
On Error GoTo An_Error

Set DDraw = DX.DirectDrawCreate("")
DDraw.SetCooperativeLevel DXhWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE Or DDSCL_ALLOWMODEX
DDraw.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT

SRECT.Right = 640
SRECT.Bottom = 480

PrimarySurfaceDescription.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
PrimarySurfaceDescription.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
PrimarySurfaceDescription.lBackBufferCount = 1
Set PrimarySurface = DDraw.CreateSurface(PrimarySurfaceDescription)

Dim Caps As DDSCAPS2
Caps.lCaps = DDSCAPS_BACKBUFFER
Set BackBuffer = PrimarySurface.GetAttachedSurface(Caps)
BackBuffer.GetSurfaceDesc BackBufferDescription

ShowCursor (0)
InitiateDX = True

An_Error:
InitiateDX = False
End Function
Public Function EndDX()
DDraw.RestoreDisplayMode
DDraw.SetCooperativeLevel DXhWnd, DDSCL_NORMAL
Set BackBuffer = Nothing
Set PrimarySurface = Nothing
Set DDraw = Nothing
Set DX = Nothing
ShowCursor (1)
End Function

Function Blt()
On Error GoTo Err
DoEvents
LostSurfaces
PrimarySurface.Flip Nothing, DDFLIP_WAIT
Err:
End Function
Public Function Clr()
On Error GoTo Err
LostSurfaces
BackBuffer.BltColorFill SRECT, 0
Err:
End Function
Public Function PText(Text As String, TX As Long, TY As Long, Optional TColor As Long = -1)
On Error GoTo Err
If TColor <> -1 Then BackBuffer.SetForeColor TColor
LostSurfaces
BackBuffer.DrawText TX, TY, Text, False
Err:
End Function
Public Function CreateSurfaceFromFile(ByVal FileName As String) As DirectDrawSurface7
Dim Picture As StdPicture
Dim Width As Long, Height As Long
Dim Surface As DirectDrawSurface7
Dim SurfaceDesc As DDSURFACEDESC2
Dim hdcPicture As Long, hdcSurface As Long
Dim ddtrans As DDCOLORKEY
  
Set Picture = LoadPicture(FileName)

Width = CLng((Picture.Width * 0.001) * 567 / Screen.TwipsPerPixelX)
Height = CLng((Picture.Height * 0.001) * 567 / Screen.TwipsPerPixelY)
    
With SurfaceDesc
    .lFlags = DDSD_CAPS
    .lFlags = .lFlags Or DDSD_WIDTH Or DDSD_HEIGHT
    .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    .lWidth = Width
    .lHeight = Height
End With

Set Surface = DDraw.CreateSurface(SurfaceDesc)
hdcPicture = CreateCompatibleDC(0)
SelectObject hdcPicture, Picture.Handle
hdcSurface = Surface.GetDC

StretchBlt hdcSurface, 0, 0, SurfaceDesc.lWidth, SurfaceDesc.lHeight, hdcPicture, 0, 0, Width, Height, SRCCOPY

Surface.ReleaseDC hdcSurface

ddtrans.low = vbMagenta
ddtrans.high = ddtrans.low
Surface.SetColorKey DDCKEY_SRCBLT, ddtrans
 
DeleteDC hdcPicture
Set Picture = Nothing
Set CreateSurfaceFromFile = Surface
Set Surface = Nothing
End Function
Public Function BltSurface(DstSurf As DirectDrawSurface7, SrcSurf As DirectDrawSurface7, Optional SX As Long, Optional SY As Long)
Dim SrcDsc As DDSURFACEDESC2, SrcRct As RECT
On Error GoTo Err
SrcSurf.GetSurfaceDesc SrcDsc
SrcRct.Right = SrcDsc.lWidth
SrcRct.Bottom = SrcDsc.lHeight
LostSurfaces
DstSurf.BltFast SX, SY, SrcSurf, SrcRct, DDBLTFAST_WAIT
Err:
End Function
Public Function BltRotate(DstSurf As DirectDrawSurface7, SrcSurf As DirectDrawSurface7, SAngle As Integer, Optional XDest As Long = 0, Optional Ydest As Long = 0, Optional Transparency As Long = -1)
Dim SrcSD As DDSURFACEDESC2, rEmpty As RECT
Dim ScreenSD As DDSURFACEDESC2, ScreenW As Long, ScreenH As Long
Dim Rad As Single, SinA As Single, CosA As Single
Dim SrcWidth As Long, SrcHeight As Long, Hyp As Long
Dim XP As Long, YP As Long, GXP As Long, GYP As Long
Dim LngColor As Long

On Error GoTo Err

Rad = SAngle * PI / 180
SinA = Sin(Rad)
CosA = Cos(Rad)

SrcSurf.GetSurfaceDesc SrcSD
BackBuffer.GetSurfaceDesc ScreenSD

SrcWidth = SrcSD.lWidth
SrcHeight = SrcSD.lHeight
ScreenW = ScreenSD.lWidth
ScreenH = ScreenSD.lHeight
Hyp = Sqr(SrcWidth ^ 2 + SrcHeight ^ 2)

LostSurfaces

DstSurf.Lock rEmpty, SrcSD, DDLOCK_WAIT, 0
SrcSurf.Lock rEmpty, SrcSD, DDLOCK_WAIT, 0

For XP = -Hyp / 2 To Hyp / 2
    For YP = -Hyp / 2 To Hyp / 2
        GXP = SrcWidth / 2 + (XP * CosA + YP * SinA)
        GYP = SrcHeight / 2 - (XP * SinA - YP * CosA)
        If GXP > 0 And GYP > 0 And _
        GXP < SrcWidth And GYP < SrcHeight And _
        XDest + XP > 0 And Ydest + YP > 0 And _
        XDest + XP < 640 And Ydest + YP < 480 Then
            LngColor = SrcSurf.GetLockedPixel(GXP, GYP)
            If LngColor <> Transparency Then
                DstSurf.SetLockedPixel XDest + XP, Ydest + YP, LngColor
            End If
        End If
    Next YP
Next XP

SrcSurf.Unlock rEmpty
DstSurf.Unlock rEmpty
Err:
End Function
Public Function ExclusiveMode() As Boolean
Dim lTestExMode As Long

'Test if we're still in exclusive mode
lTestExMode = DDraw.TestCooperativeLevel

If (lTestExMode = DD_OK) Then
    ExclusiveMode = True
Else
    ExclusiveMode = False
End If
End Function
Public Function LostSurfaces() As Boolean
'Check if we should reload our bitmaps or not
LostSurfaces = False
Do Until ExclusiveMode
    DoEvents
    LostSurfaces = True
Loop
    
'Lost bitmaps, restore the surfaces and return 'true'
DoEvents

If LostSurfaces Then
    DDraw.RestoreAllSurfaces
End If
    
End Function
