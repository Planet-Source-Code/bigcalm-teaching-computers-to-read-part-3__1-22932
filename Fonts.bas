Attribute VB_Name = "Fonts"
Option Explicit

' Original code by Tom Walker.  tomwalker@hotmail.com
' API callback enumeration for Fonts.

Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
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
        lfFaceName(LF_FACESIZE) As Byte
End Type
Private Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type
Private Const NTM_REGULAR = &H40&
Private Const NTM_BOLD = &H20&
Private Const NTM_ITALIC = &H1&
Private Const TMPF_FIXED_PITCH = &H1
Private Const TMPF_VECTOR = &H2
Private Const TMPF_DEVICE = &H8
Private Const TMPF_TRUETYPE = &H4
Private Const ELF_VERSION = 0
Private Const ELF_CULTURE_LATIN = 0
Private Const RASTER_FONTTYPE = &H1
Private Const DEVICE_FONTTYPE = &H2
Private Const TRUETYPE_FONTTYPE = &H4
Private Declare Function EnumFontFamilies Lib "gdi32" Alias "EnumFontFamiliesA" (ByVal hdc As Long, ByVal lpszFamily As String, ByVal lpEnumFontFamProc As Long, LParam As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private JustTextualFonts As Boolean

Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, ByVal FontType As Long, LParam As ListBox) As Long
Dim FaceName As String
Dim FullName As String
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    If JustTextualFonts = True And lpNTM.tmCharSet > 0 Then
    Else
      If Len(FaceName) > 0 Then
        LParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
      End If
    End If
    EnumFontFamProc = 1
End Function

Sub FillListWithFonts(LB As ListBox)
Dim hdc As Long
    LB.Clear
    hdc = GetDC(LB.hwnd)
    EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, LB
    ReleaseDC LB.hwnd, hdc
End Sub

Sub FillComboWithFonts(LB As ComboBox, Optional TextFontsOnly As Boolean = False)
Dim hdc As Long
    JustTextualFonts = TextFontsOnly
    LB.Clear
    hdc = GetDC(LB.hwnd)
    EnumFontFamilies hdc, vbNullString, AddressOf EnumFontFamProc, LB
    ReleaseDC LB.hwnd, hdc
End Sub


