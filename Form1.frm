VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6705
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   180
      Picture         =   "Form1.frx":CF26
      ScaleHeight     =   900
      ScaleWidth      =   2190
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2250
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Demonstration of how to transparently put a gif/bmp
'on top of another image using Device Context (DC) and
'TransparentBlt API

Private Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

'drawing API
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Private Sub Form_Resize()
    Me.Cls
    Dim lWidth As Long
    Dim lHeight As Long
    lWidth = (Me.ScaleWidth / 2) - (Me.Picture1.ScaleWidth / 2)
    lHeight = Me.ScaleHeight / 2 - (Me.Picture1.ScaleHeight / 2)
    'draw the image centre of the DC / 15 (15 = Twips per pixel)
    DrawTransPicture Me.Picture1.Picture, lWidth / 15, lHeight / 15, RGB(255, 0, 255)
End Sub

Private Sub DrawTransPicture(img As StdPicture, ImageX As Long, ImageY As Long, ImgTransColour As Long)
    Dim hbmDc As Long
    Dim hBmp As Long
    Dim hBmpOld As Long
    Dim bmp As BITMAP
    'if the picture is a bitmap...
    If img.Type = vbPicTypeBitmap Then
        hBmp = img.Handle
        'create a memory device context
        hbmDc = CreateCompatibleDC(0&)
        If hbmDc <> 0 Then
            'select the bitmap into the context
            hBmpOld = SelectObject(hbmDc, hBmp)
            'retrieve information for the
            'specified graphics object
            If GetObject(hBmp, Len(bmp), bmp) <> 0 Then
                'draw the bitmap with the
                'specified transparency colour
                Call TransparentBlt(Me.hdc, ImageX, ImageY, bmp.bmWidth, bmp.bmHeight, hbmDc, 0, 0, bmp.bmWidth, bmp.bmHeight, ImgTransColour)
            End If  'GetObject
            Call SelectObject(hbmDc, hBmpOld)
            DeleteObject hBmpOld
            DeleteDC hbmDc
        End If  'hbmDc
    ElseIf img.Type = vbPicTypeIcon Then
        'if the picture is an icon
        Call Me.PaintPicture(img, ImageX, ImageY)
    End If
End Sub
