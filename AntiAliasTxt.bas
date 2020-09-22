Attribute VB_Name = "AntiAliasText"
' Code originally by Roger Johansson.

Private Declare Sub CopyMemoryLong Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub MemCpy Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Long, ByRef Source As Long, ByVal Length As Long)

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function GetTabbedTextExtent Lib "user32" Alias "GetTabbedTextExtentA" (ByVal hdc As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long) As Long

Private Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hbrush As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long

Private Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long


Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type


Private Type POINTAPI
        x As Long
        y As Long
End Type


Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type

' Uncomment if you move this module into a different project
'Private Type DWORD
'    low As Integer
'    high As Integer
'End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Function DrawAntiAliasedText(hdc As Long, Text As String, xpos As Long, ypos As Long, color As Long, opacity As Double, fontname As String, fontsize As Long) As DWORD
  Dim size                                  As DWORD
  Dim ret                                   As Long
  Dim ndc                                   As Long
  Dim nbmp                                  As Long
  Dim hjunk
  Dim font                                  As LOGFONT
  Dim hfont                                 As Long
  Dim hfont2                                As Long
  Dim pixels()                              As RGBQUAD
  Dim npixels()                             As RGBQUAD
  Dim bgpixels()                            As RGBQUAD
  Dim rgbcol(3)                             As Byte
  Dim x, y, yy
  Dim bminfo                                As BITMAPINFO
  Dim tmp                                   As Double
  Dim alpha                                 As Double
  Dim RetValue                             As DWORD
  Dim i As Long, j As Long, k As Long
  Dim hObj As Long
    With font
        .lfHeight = -(fontsize * 20) / Screen.TwipsPerPixelY ' set font size
        .lfFaceName = fontname & Chr(0) 'apply font name
        .lfWeight = 0   'this is how bold the font is .. apply a in param if you want
    End With
    
    '-----------------------------------------
    'create a dc for our backbuffer
    ndc = CreateCompatibleDC(hdc)
    'create a bitmap for our backbuffer
    nbmp = CreateCompatibleBitmap(hdc, 1, 1) 'make a temp bitmap so we can get the size of the text
    'attach our bitmap to our backbuffer
    hjunk = SelectObject(ndc, nbmp)
    'apply the font to our backbuffer
    hfont = CreateFontIndirect(font)
    hfont2 = SelectObject(ndc, hfont)
    
    'get size of the text we want to draw
    ret = GetTabbedTextExtent(ndc, Text, Len(Text), 0, 0)
    
    'delete our temp bmp
    hObj = SelectObject(ndc, hfont2)
    i = DeleteObject(hObj)
    hObj = SelectObject(ndc, hjunk)
    i = DeleteObject(hObj)
    i = DeleteDC(ndc)

    'this part was only to measure the size of the text
    '----------------------------------------
    'now lets draw the text...
    
    
    'split our color value to a byte array
    'this is my own invention ... pretty nice (?)
    CopyMemoryLong VarPtr(rgbcol(0)), VarPtr(color), 4
    'split the return value from gettextextent into two integers
    CopyMemoryLong VarPtr(size), VarPtr(ret), 4
    ' And copy to function value so that the text extents can be utilised.
    CopyMemoryLong VarPtr(RetValue), VarPtr(ret), 4
    DrawAntiAliasedText = RetValue
    
    ypos = ypos - size.high / 2
    'create a dc for our backbuffer
    ndc = CreateCompatibleDC(hdc)
    'create a bitmap for our backbuffer
    nbmp = CreateCompatibleBitmap(hdc, size.low, size.high)
    'attach our bitmap to our backbuffer
    hjunk = SelectObject(ndc, nbmp)
    'apply the font to our backbuffer
    hfont = CreateFontIndirect(font)
    hfont2 = SelectObject(ndc, hfont)
    'set black background coloy
    SetBkColor ndc, 0
    'set white forecolor
    SetTextColor ndc, vbWhite
    'write the text to our backbuffer
    TabbedTextOut ndc, 0, 0, Text, Len(Text), 0, 0, 0
    'resize the arrays to the same size as the bbuffer
    ReDim pixels(size.low - 1, size.high - 1)
    ReDim npixels(size.low - 1, size.high - 1)
    ReDim bgpixels(size.low - 1, size.high - 1)
    
    'set the bitmap info (so we can get the gfx data in and out of our arrays
    With bminfo.bmiHeader
        .biSize = Len(bminfo.bmiHeader)
        .biWidth = size.low
        .biHeight = size.high
        .biPlanes = 1
        .biBitCount = 32
    End With
    'store the drawn text in our "pixels" array
    GetDIBits ndc, nbmp, 0, size.high, pixels(0, 0), bminfo, 1
    'get the bg graphics into our "bgpixels" array
    BitBlt ndc, 0, 0, size.low, size.high, hdc, xpos, ypos, vbSrcCopy
    GetDIBits ndc, nbmp, 0, size.high, bgpixels(0, 0), bminfo, 1
    yy = Int(size.high / 2)
    
    ' npixels = bgpixels ' cannot do this in VB5
    For i = LBound(bgpixels, 1) To UBound(bgpixels, 1)
      For j = LBound(bgpixels, 2) To UBound(bgpixels, 2)
        npixels(i, j) = bgpixels(i, j)
      Next j
    Next i
    'npixels = bgpixels
    'copyMemoryLong VarPtr(npixels(LBound(npixels, 1), LBound(npixels, 2))), VarPtr(bgpixels(LBound(bgpixels, 1), LBound(bgpixels, 2))), (UBound(bgpixels, 1) - LBound(bgpixels)) * (UBound(bgpixels, 2) - LBound(bgpixels, 2))
    
    For x = 0 To size.low - 2 Step 2
        For y = 0 To size.high - 2 Step 2
            'alpha is the average of the color of 2*2 pixels /255
            'now we have a value between 0 and 1
            '0 is transparent
            '1 is soild white
            'now multiply alpha with the opacity factor
            'ie if opacity is 0.5 ...  aplha will be max 0.5
            'since we draw our text with white . we only need to check the strength of one color (in this case blue)
            'coz red and green will always be the same as the blue
            alpha = (((0 + (pixels(x + 0, y + 0).rgbBlue) + (pixels(x + 1, y + 0).rgbBlue) + (pixels(x + 0, y + 1).rgbBlue) + (pixels(x + 1, y + 1).rgbBlue)) / 4) / 255) * opacity
            'alpha is now the opacity factor 0-1
            'calculate amount of blue to apply
            'and how much of the background that is going to be seen
            tmp = (alpha * rgbcol(2)) + bgpixels(x / 2, y / 2).rgbBlue * (1 - alpha)
            'never go higher than 255
            If tmp > 255 Then tmp = 255
            'store the result at x/2 and y/2 (the new picture is only 0.5 times as high and wide
            npixels(x / 2, y / 2).rgbBlue = tmp
            'calculate amount of red to apply
            'and how much of the background that is going to be seen
            tmp = (alpha * rgbcol(0)) + bgpixels(x / 2, y / 2).rgbRed * (1 - alpha)
            'never go higher than 255
            If tmp > 255 Then tmp = 255
            npixels(x / 2, y / 2).rgbRed = tmp
            'calculate amount of green to apply
            'and how much of the background that is going to be seen
            tmp = (alpha * rgbcol(1)) + bgpixels(x / 2, y / 2).rgbGreen * (1 - alpha)
            'never go higher than 255
            If tmp > 255 Then tmp = 255
            npixels(x / 2, y / 2).rgbGreen = tmp
        Next
    Next
    'apply the new picture to our bbuffer-dc
    SetDIBits ndc, nbmp, 0, size.high, npixels(0, 0), bminfo, 1
    'blit our bbuffer-dc to the screen
    BitBlt hdc, xpos, ypos, size.low, size.high, ndc, 0, 0, vbSrcCopy
    'clean up
    hObj = SelectObject(ndc, hfont2)
    i = DeleteObject(hObj)
    hObj = SelectObject(ndc, hjunk)
    i = DeleteObject(hObj)
    i = DeleteDC(ndc)
End Function

