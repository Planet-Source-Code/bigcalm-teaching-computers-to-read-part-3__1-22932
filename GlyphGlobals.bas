Attribute VB_Name = "GlyphGlobals"
Option Explicit
' All declares and variables that are visible throughout the
' whole project are placed here.

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
      ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
      ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
      ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
  ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
  ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
  ByVal ySrc As Long, ByVal nSrcWidth As Long, _
  ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, _
      ByVal nStretchMode As Long) As Long
      
Public Type DWORD
    low As Integer
    high As Integer
End Type

Public Enum StretchBltModes
  BLACKONWHITE = 1
  WHITEONBLACK = 2
  COLORONCOLOR = 3
  HALFTONE = 4
  MAXSTRETCHBLTMODE = 4
End Enum
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const cMaxPath = &H104
Public Enum EShellShowConstants
    essSW_HIDE = 0
    essSW_MAXIMIZE = 3
    essSW_MINIMIZE = 6
    essSW_SHOWMAXIMIZED = 3
    essSW_SHOWMINIMIZED = 2
    essSW_SHOWNORMAL = 1
    essSW_SHOWNOACTIVATE = 4
    essSW_SHOWNA = 8
    essSW_SHOWMINNOACTIVE = 7
    essSW_SHOWDEFAULT = 10
    essSW_RESTORE = 9
    essSW_SHOW = 5
End Enum
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5        ' access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                ' file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                ' path not found
Private Const SE_ERR_OOM = 8                ' out of memory
Private Const SE_ERR_SHARE = 26

Global Const CB_ERR = -1
Global Const CB_FINDSTRING = &H14C
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public TrainingFonts(0 To 50) As String
Public TestingFonts(0 To 5) As String
Public TestingGlyphs() As String

Public Sub PopulateTrainingFonts()
  TrainingFonts(0) = "Abadi MT Condensed"
  TrainingFonts(1) = "Amazone BT"
  TrainingFonts(2) = "Ambient"
  TrainingFonts(3) = "Arial"
  TrainingFonts(4) = "Arial Black"
  TrainingFonts(5) = "Arial Narrow"
  TrainingFonts(6) = "Arial Rounded MT Bold"
  TrainingFonts(7) = "Arnold Boecklin"
  TrainingFonts(8) = "Balmoral LET"
  TrainingFonts(9) = "Book Antiqua"
  TrainingFonts(10) = "Bookman Old Style"
  TrainingFonts(11) = "Britannic Bold"
  TrainingFonts(12) = "Brush Script MT"
  TrainingFonts(13) = "Cancellaresca Script LET"
  TrainingFonts(14) = "Caslon 3"
  TrainingFonts(15) = "Century Gothic"
  TrainingFonts(16) = "Century Schoolbook"
  TrainingFonts(17) = "Comic Sans MS"
  TrainingFonts(18) = "Courier New"
  TrainingFonts(19) = "Dom Casual"
  TrainingFonts(20) = "Fette Fraktur"
  TrainingFonts(21) = "Fixedsys"
  TrainingFonts(22) = "Footlight MT Light"
  TrainingFonts(23) = "Franklin Cond. Gothic"
  TrainingFonts(24) = "Franklin Extra Cond. Gothic"
  TrainingFonts(25) = "Gando BT"
  TrainingFonts(26) = "Garamond"
  TrainingFonts(27) = "GeoSlab703 Md BT"
  TrainingFonts(28) = "GF Gesetz"
  TrainingFonts(29) = "Goudy Old Style"
  TrainingFonts(30) = "Haettenschweiler"
  TrainingFonts(31) = "Heritage"
  TrainingFonts(32) = "Impact"
  TrainingFonts(33) = "Japan"
  TrainingFonts(34) = "Kis BT"
  TrainingFonts(35) = "LeiScriptSSk"
  TrainingFonts(36) = "Liberty"
  TrainingFonts(37) = "Malibu LET"
  TrainingFonts(38) = "Mural Script"
  TrainingFonts(39) = "Nora Casual"
  TrainingFonts(40) = "Oak"
  TrainingFonts(41) = "Old Towne No.536"
  TrainingFonts(42) = "Optimum"
  TrainingFonts(43) = "Park Avenue"
  TrainingFonts(44) = "Playbill"
  TrainingFonts(45) = "Romeo"
  TrainingFonts(46) = "Serpentine"
  TrainingFonts(47) = "Small Fonts"
  TrainingFonts(48) = "System"
  TrainingFonts(49) = "Tahoma"
'  TrainingFonts(50) = "Tango BT"
'  TrainingFonts(51) = "Times New Roman"
'  TrainingFonts(52) = "Trebuchet MS"
'  TrainingFonts(53) = "Van Dijk Bold LET"
'  TrainingFonts(54) = "Verdana"
'  TrainingFonts(55) = "Windsor"

  
  ' Now set up testing fonts - these will not be trained, but will be tested
  TestingFonts(0) = "Tango BT"
  TestingFonts(1) = "Times New Roman"
  TestingFonts(2) = "Trebuchet MS"
  TestingFonts(3) = "Van Dijk Bold LET"
  TestingFonts(4) = "Verdana"
  TestingFonts(5) = "Windsor"
'  TestingFonts(6) = "Parisian BT"
  
  ReDim TestingGlyphs(0 To 9)
  TestingGlyphs(0) = "0"
  TestingGlyphs(1) = "1"
  TestingGlyphs(2) = "2"
  TestingGlyphs(3) = "3"
  TestingGlyphs(4) = "4"
  TestingGlyphs(5) = "5"
  TestingGlyphs(6) = "6"
  TestingGlyphs(7) = "7"
  TestingGlyphs(8) = "8"
  TestingGlyphs(9) = "9"
End Sub

Public Function ShellEx( _
        ByVal sFIle As String, _
        Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
        Optional ByVal sParameters As String = "", _
        Optional ByVal sDefaultDir As String = "", _
        Optional sOperation As String = "open", _
        Optional Owner As Long = 0 _
    ) As Boolean
Dim lR As Long
Dim lErr As Long, sErr As String
    If (InStr(UCase$(sFIle), ".EXE") <> 0) Then
        eShowCmd = 0
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then
        lR = ShellExecuteForExplore(Owner, sOperation, sFIle, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFIle, sParameters, sDefaultDir, eShowCmd)
    End If
    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR
        Select Case lR
        Case 0
            lErr = 7: sErr = "Out of memory"
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "File not found"
        Case ERROR_PATH_NOT_FOUND
            lErr = 76: sErr = "Path not found"
        Case ERROR_BAD_FORMAT
            sErr = "The executable file is invalid or corrupt"
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "Path/file access error"
        Case SE_ERR_ASSOCINCOMPLETE
            sErr = "This file type does not have a valid file association."
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "The file could not be opened because the target application is busy. Please try again in a moment."
        Case SE_ERR_DDEFAIL
            lErr = 285: sErr = "The file could not be opened because the DDE transaction failed. Please try again in a moment."
        Case SE_ERR_DDETIMEOUT
            lErr = 286: sErr = "The file could not be opened due to time out. Please try again in a moment."
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "The specified dynamic-link library was not found."
        Case SE_ERR_FNF
            lErr = 53: sErr = "File not found"
        Case SE_ERR_NOASSOC
            sErr = "No application is associated with this file type."
        Case SE_ERR_OOM
            lErr = 7: sErr = "Out of memory"
        Case SE_ERR_PNF
            lErr = 76: sErr = "Path not found"
        Case SE_ERR_SHARE
            lErr = 75: sErr = "A sharing violation occurred."
        Case Else
            sErr = "An error occurred occurred whilst trying to open or print the selected file."
        End Select
                
        Err.Raise lErr, , App.EXEName & ".GShell: " & sErr
        ShellEx = False
    End If

End Function


