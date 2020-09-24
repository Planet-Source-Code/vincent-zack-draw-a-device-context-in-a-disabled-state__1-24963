<div align="center">

## Draw A Device Context In A Disabled State


</div>

### Description



Cut and Paste the code below into a new project.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vincent Zack](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vincent-zack.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vincent-zack-draw-a-device-context-in-a-disabled-state__1-24963/archive/master.zip)

### API Declarations

```
Private Const SRCCOPY = &HCC0020
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
```


### Source Code

```
Sub DisableHDC(SourceDC As Long, SourceWidth As Long, SourceHeight As Long)
Const BLACK = 0
Const DARKGREY = &H808080
Const WHITE = &HFFFFFF
Dim i As Long
Dim j As Long
Dim PixelColor As Long
Dim BackgroundColor As Long
Dim MemoryDC As Long
Dim MemoryBitmap As Long
Dim OldBitmap As Long
Dim BooleanArray() As Boolean
ReDim BooleanArray(SourceWidth, SourceHeight)
MemoryDC = CreateCompatibleDC(SourceDC)
MemoryBitmap = CreateCompatibleBitmap(SourceDC, SourceWidth, SourceHeight)
OldBitmap = SelectObject(MemoryDC, MemoryBitmap)
BitBlt MemoryDC, 0, 0, SourceWidth, SourceHeight, SourceDC, 0, 0, SRCCOPY
BackgroundColor = GetBkColor(SourceDC)
' Scan Pixels and if the pixel is black
' it is flagged as true and saved in BooleanArray(x,y)
' then colored dark grey (disabled color)
For i = 0 To SourceWidth
  For j = 0 To SourceHeight
    PixelColor = GetPixel(MemoryDC, i, j)
    If PixelColor <> BackgroundColor Then ' skip background color pixels
      If PixelColor = BLACK Then
        BooleanArray(i, j) = True
        SetPixel MemoryDC, i, j, DARKGREY
      Else
        SetPixel MemoryDC, i, j, BackgroundColor
      End If
    End If
  Next
Next
' For each Black pixel, draw a white shadow 1 pixel down and
' 1 pixel to the right to create a shadow effect
For i = 0 To SourceWidth - 1
  For j = 0 To SourceHeight - 1
    If BooleanArray(i, j) = True Then
      If BooleanArray(i + 1, j + 1) = False Then
      SetPixel MemoryDC, i + 1, j + 1, WHITE
      End If
    End If
  Next
Next
BitBlt SourceDC, 0, 0, SourceWidth, SourceHeight, MemoryDC, 0, 0, SRCCOPY
SelectObject MemoryDC, OldBitmap
DeleteObject MemoryBitmap
DeleteDC MemoryDC
End Sub
Private Sub Form_Load()
Me.Picture = Me.Icon
End Sub
' Hold down mouse button to disable
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const PICSIZE = 32
Me.Picture = Me.Icon
Me.AutoRedraw = True
Me.ScaleMode = vbPixels
DisableHDC Me.hdc, PICSIZE, PICSIZE
Me.Refresh
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Picture = Me.Icon
End Sub
```

