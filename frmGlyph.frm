VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGlyph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Neural Net - Teaching Computers To Read Part 3 (Numbers)"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Glyph Manipulation"
      Height          =   4695
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   6135
      Begin VB.CommandButton Command3 
         Caption         =   "Huh?"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear Pic"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Test Glyph"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton cmdBattery 
         Caption         =   "Battery"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Output"
         Height          =   855
         Left            =   120
         TabIndex        =   19
         Top             =   3720
         Width           =   5895
         Begin VB.CheckBox chkCorrect 
            Enabled         =   0   'False
            Height          =   255
            Left            =   5520
            TabIndex        =   31
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox txtOutputPercentage 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   495
         End
         Begin VB.Label lblTotPercent 
            Height          =   375
            Left            =   5280
            TabIndex        =   29
            Top             =   360
            Width           =   495
         End
         Begin VB.Label lblOutput 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.PictureBox picGlyphBits 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   2775
         Left            =   1200
         ScaleHeight     =   181
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   317
         TabIndex        =   18
         Top             =   960
         Width           =   4815
      End
      Begin VB.PictureBox picFlipG 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   5280
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   45
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.PictureBox picGlyph 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1440
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblLearningRate 
         Height          =   255
         Left            =   2160
         TabIndex        =   30
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblSSE 
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblTotalTrainCycles 
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Press a key to test output of Neural Net."
         Height          =   1095
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Net-Input Glyph"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Stretched Glyph"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Font"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblTrainingFont 
         Height          =   495
         Left            =   3360
         TabIndex        =   12
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Commands"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   6135
      Begin VB.CommandButton btCreate 
         Caption         =   "Create Net"
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton btSave 
         Caption         =   "Save Net"
         Height          =   495
         Left            =   3720
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton btRestore 
         Caption         =   "Load Net"
         Height          =   495
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton btKill 
         Caption         =   "Destroy Net"
         Height          =   495
         Left            =   4920
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton btTrain 
         Caption         =   "Train Net"
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   6840
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7110
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCycles 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6840
      Width           =   1455
   End
End
Attribute VB_Name = "frmGlyph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Program to recognise whether a glyph is "0" - "9"
' This will be expanded _if_it_works_

' Coding by Jonathan Daniel. bigcalm@hotmail.com

' The first rule of programming is to steal all the code you can,
' so you don't have to write it yourself (credits in full here of course)
' So therefore..
' I have nicked a ridiculous amount of code off other people to do this
' The credits are as follows:

' Ulli (umgedv@aol.com): All Neural Net classes and initialisation functions. mNet.bas, cDendrite.cls, cLayer.cls, cNet.cls, cNeuron.cls
' Tom Sawyer (tomsawyer@hotmail.com): Font API enumeration. Fonts.bas
' Roger Johansson: Anti-aliased text.  AntiAliasTxt.bas

' I have changed some of this code to fit my needs.

'----------------------------------------------------------------------------------------------------------------
' Please see module NotesOnCode for a list of testing, changes & history
#Const OutputStyle = 1 ' Ulli requested that Sum of output boxes should equal
      ' 100%.  Setting this to 1 should do this. (I don't like it myself so it's a compiler
      ' constant)
#Const UseOptimisedNet = 1

#If UseOptimisedNet = 1 Then
Private WithEvents oNet As NetFast
Attribute oNet.VB_VarHelpID = -1
#Else
Private WithEvents oNet As cNet
Attribute oNet.VB_VarHelpID = -1
#End If
Private TrainingIterations As Long
Private Const CharacterWidth As Long = 30
Private Const CharacterHeight As Long = 20
Private Const AcceptableShortFall = 0.3
Private Const AcceptableOpposableError = 0.3
Private Running As Boolean
Private Stopping As Boolean

Private Sub FlipAndShift()
Dim i As Long
Dim j As Long
Dim k As Long
Dim NotBlank As Boolean
Dim LastNonBlankRow As Long
Dim FirstNonBlankCol As Long

  Set picFlipG = LoadPicture
  BitBlt picFlipG.hdc, 0, 0, CharacterWidth, CharacterHeight, _
    picGlyph.hdc, 0, 0, vbSrcCopy
  Exit Sub

End Sub


Private Sub btTrain_Click()
Dim i As Long, j As Long, k As Long, l As Long, m As Long
Dim TrainingSet(0 To ((CharacterWidth + 1) * (CharacterHeight + 1))) As Long
Dim ExpectedResultSet() As Double
Dim TrainFont As String
Dim NumberOfCyclesStr As String

  If oNet Is Nothing Then
    MsgBox "Create net first!", vbExclamation
    Exit Sub
  End If
  
  If Running = True Then
    Stopping = True
    Exit Sub
  Else
    NumberOfCyclesStr = InputBox("Please enter the number of training cycles", , 1000)
    If IsNumeric(NumberOfCyclesStr) Then
      TrainingIterations = NumberOfCyclesStr
    Else
      MsgBox "You must enter a numeric number of cycles!", vbExclamation
      Exit Sub
    End If
    Stopping = False
    Running = True
  End If
  
  ' Initialise Training Fonts
  PopulateTrainingFonts
  ReDim ExpectedResultSet(LBound(TestingGlyphs) To UBound(TestingGlyphs))
    
  ' Set up progress bars
  ProgressBar1.Visible = True
  ProgressBar1.Min = 0
  ProgressBar1.Max = 100
  ProgressBar1.Value = 0
  ProgressBar2.Visible = True
  ProgressBar2.Min = 0
  ProgressBar2.Max = TrainingIterations
  ProgressBar2.Value = 0
  StatusBar1.Panels(1).Text = "Training"
  lblCycles.Caption = "Cycles: 0" & " / " & TrainingIterations
  btTrain.Caption = "Stop Training"
  lblTrainingFont.Caption = "Training Font: " & ""
     
  With oNet
      ' Initialise training
      Me.MousePointer = vbArrowHourglass
      i = 1
      Do While i < TrainingIterations
        For m = LBound(TrainingFonts) To UBound(TrainingFonts)
          TrainFont = TrainingFonts(m)
          lblTrainingFont.Caption = "Training Font: " & TrainFont
          For l = LBound(TestingGlyphs) To UBound(TestingGlyphs)
            PrepareGlyph TestingGlyphs(l), TrainFont
            ' Copy glyph into training set array.
            Erase TrainingSet
            For j = 0 To CharacterWidth
              For k = 0 To CharacterHeight
                TrainingSet((j * CharacterHeight) + k) = 1 - ((picGlyph.Point(j, k) And &HFF&) / 255)
              Next k
            Next j
            For j = LBound(ExpectedResultSet) To UBound(ExpectedResultSet)
              If j = l Then
                ExpectedResultSet(j) = 1
              Else
                ExpectedResultSet(j) = 0
              End If
            Next j
            ' Train
            .Train TrainingSet, ExpectedResultSet
            
            ShowOutput l
            
            lblSSE.Caption = .AverageSquaredError
            lblLearningRate.Caption = .LearningCoefficient
            DoEvents
            If Stopping = True Then
              Exit For
            End If
          Next l
  
          DoEvents
          If Stopping = True Then
            Exit Do
          End If
          ProgressBar2.Value = i
          lblCycles.Caption = "Cycles: " & i & " / " & TrainingIterations
          lblTotalTrainCycles.Caption = oNet.TrainingCycles
          i = i + 1
        Next m
      Loop
    End With
    
   ' Tidy up.
   btTrain.Caption = "Train Net"
   picGlyph.Visible = True
   ProgressBar2.Visible = False
   lblTrainingFont.Caption = ""
   WipePanelsAndProgress
End Sub

Private Function GetTrueExtents(TextExtents As DWORD) As RECT
Dim i As Long
Dim j As Long
Dim Colour As Long
Dim RealTextExtent As RECT

  With RealTextExtent
    .Left = -1
    .Top = -1
    .Right = -1
    .Bottom = -1
  End With
  ' Top
  For i = 0 To TextExtents.high
    For j = 0 To TextExtents.low
      Colour = GetPixel(picGlyphBits.hdc, j, i)
      If Colour <> &HFFFFFF And Colour <> -1 Then
        RealTextExtent.Top = i
        Exit For
      End If
    Next j
    If RealTextExtent.Top <> -1 Then
      Exit For
    End If
  Next i
  If RealTextExtent.Top = -1 Then
    GetTrueExtents = RealTextExtent
    Exit Function
  End If
  ' Left
  For i = 0 To TextExtents.low
    For j = RealTextExtent.Top To TextExtents.high
      Colour = GetPixel(picGlyphBits.hdc, i, j)
      If Colour <> &HFFFFFF And Colour <> -1 Then
        RealTextExtent.Left = i
        Exit For
      End If
    Next j
    If RealTextExtent.Left <> -1 Then
      Exit For
    End If
  Next i
  If RealTextExtent.Left = -1 Then
    GetTrueExtents = RealTextExtent
    Exit Function
  End If

  ' Right
  For i = TextExtents.low To RealTextExtent.Left Step -1
    For j = RealTextExtent.Top To TextExtents.high
      Colour = GetPixel(picGlyphBits.hdc, i, j)
      If Colour <> -1 And Colour <> &HFFFFFF Then
        RealTextExtent.Right = i
        Exit For
      End If
    Next j
    If RealTextExtent.Right <> -1 Then
      Exit For
    End If
  Next i
  If RealTextExtent.Right = -1 Then
    GetTrueExtents = RealTextExtent
    Exit Function
  End If

  ' Bottom
  For i = TextExtents.high To RealTextExtent.Top Step -1
    For j = RealTextExtent.Left To TextExtents.low
      Colour = GetPixel(picGlyphBits.hdc, j, i)
      If Colour <> -1 And Colour <> &HFFFFFF Then
        RealTextExtent.Bottom = i
        Exit For
      End If
    Next j
    If RealTextExtent.Bottom <> -1 Then
      Exit For
    End If
  Next i
  If RealTextExtent.Bottom = -1 Then
    GetTrueExtents = RealTextExtent
    Exit Function
  End If

  GetTrueExtents = RealTextExtent

End Function

Private Sub PrepareGlyph(pChar As String, pFont As String)
Dim TextExtents As DWORD
Dim i As Long, j As Long
Dim RealTextExtent As RECT
Dim Colour As Long
Dim StartTime As Long

  ' Test for stretchblt.
  Set picGlyph = LoadPicture
  Set picGlyphBits = LoadPicture
  TextExtents = AntiAliasText.DrawAntiAliasedText(picGlyphBits.hdc, pChar, 0, 0, &H0, 1, pFont, 144)

  picGlyphBits.Refresh
  
  ' Get real text extents (perhaps move this to AAtxt module?)
  ' This is a crock, but it works.
  ' Init
  RealTextExtent = GetTrueExtents(TextExtents)
  
  If RealTextExtent.Left = -1 Or RealTextExtent.Right = -1 Or RealTextExtent.Bottom = -1 Or RealTextExtent.Top = -1 Then
    ' must be a spacer character
    Exit Sub
  End If
  
  SetStretchBltMode picGlyph.hdc, StretchBltModes.HALFTONE ' Dont really know whether this does anything. Shrug.  It doesn't hurt.
  SetStretchBltMode picGlyphBits.hdc, StretchBltModes.HALFTONE
  StretchBlt picGlyph.hdc, 0, 0, CharacterWidth, CharacterHeight, _
          picGlyphBits.hdc, RealTextExtent.Left, RealTextExtent.Top, _
          (RealTextExtent.Bottom - RealTextExtent.Top) * (CharacterWidth / CharacterHeight), _
          RealTextExtent.Bottom - RealTextExtent.Top, vbSrcCopy
  
  ' Draw lines to indicate where the text extents are
  picGlyphBits.Line (TextExtents.low \ 2, 0)-(TextExtents.low \ 2, TextExtents.high \ 2), vbBlue
  picGlyphBits.Line (0, TextExtents.high \ 2)-(TextExtents.low \ 2, TextExtents.high \ 2), vbRed

  ' Draw lines to indicate real text extent.
  picGlyphBits.Line (RealTextExtent.Left, RealTextExtent.Top)-(RealTextExtent.Right, RealTextExtent.Top), vbGreen
  picGlyphBits.Line (RealTextExtent.Left, RealTextExtent.Top)-(RealTextExtent.Left, RealTextExtent.Bottom), vbYellow
  picGlyphBits.Line (RealTextExtent.Right, RealTextExtent.Top)-(RealTextExtent.Right, RealTextExtent.Bottom), vbMagenta
  picGlyphBits.Line (RealTextExtent.Left, RealTextExtent.Bottom)-(RealTextExtent.Right, RealTextExtent.Bottom), vbCyan
    
  picGlyph.Refresh
  picGlyph.font = Combo1.Text
  picGlyph.font.size = 12
  FlipAndShift
End Sub

' This tests the net to see how well it's learnt.
Private Sub cmdBattery_Click()
Dim i As Long, j As Long, k As Long, l As Long
Dim TrainFont As String
Dim InputSet(0 To ((CharacterWidth + 1) * (CharacterHeight + 1))) As Long
Dim RightResult As Double
Dim WrongResult As Double
Dim RightTotal As Double
Dim WrongTotal As Double
Dim tmpResult As Double
Dim ResultMessage As String
Dim Successes As Long
Dim Failures As Long
Dim tmpString As String

  If oNet Is Nothing Then
    MsgBox "Create and train the net first!", vbExclamation
    Exit Sub
  End If
  
  If oNet.TrainingCycles = 0 Then
    MsgBox "Train net first!", vbExclamation
    Exit Sub
  End If

  If Running = True Then
    Exit Sub
  Else
    Running = True
    Stopping = False
  End If
  
  ' Initialise Training Fonts
  PopulateTrainingFonts
    
  ' Set up progress bars
  ProgressBar1.Visible = True
  ProgressBar1.Min = 0
  ProgressBar1.Max = 100
  ProgressBar1.Value = 0
  ProgressBar2.Visible = True
  ProgressBar2.Min = 0
  ProgressBar2.Max = UBound(TrainingFonts) - LBound(TrainingFonts) + 1
  ProgressBar2.Value = 0
  StatusBar1.Panels(1).Text = "Testing Known Fonts"
  lblCycles.Caption = "Cycles: 0" & " / " & UBound(TrainingFonts) - LBound(TrainingFonts) + 1
  lblTrainingFont.Caption = "Test Font: " & ""
  For i = LBound(TrainingFonts) To UBound(TrainingFonts)
    For l = LBound(TestingGlyphs) To UBound(TestingGlyphs)
      TrainFont = TrainingFonts(i)
      lblTrainingFont.Caption = "Training Font: " & TrainFont
      PrepareGlyph TestingGlyphs(l), TrainFont
      Erase InputSet
      For j = 0 To CharacterWidth
        For k = 0 To CharacterHeight
          InputSet((j * CharacterHeight) + k) = 1 - ((picGlyph.Point(j, k) And &HFF&) / 255)
        Next k
      Next j
      oNet.SetInput InputSet
      oNet.ProcessOutput
      RightResult = oNet.OutputLayer(l + 1)
      'txtOutputPercentage(l).Text = Format$(RightResult * 100, "0.00") & "%"
      WrongResult = 0
      For j = LBound(TestingGlyphs) To UBound(TestingGlyphs)
        If j <> l Then
          tmpResult = oNet.OutputLayer(j + 1)
          If tmpResult > WrongResult Then
            WrongResult = tmpResult
          End If
          'txtOutputPercentage(j).Text = Format$(tmpResult * 100, "0.00") & "%"
        End If
      Next j
      ShowOutput l
      RightTotal = RightTotal + RightResult
      WrongTotal = WrongTotal + WrongResult
      If RightResult >= 1 - AcceptableShortFall And WrongResult <= 0 + AcceptableOpposableError Then
        Successes = Successes + 1
      Else
        tmpString = ""
        For j = LBound(TestingGlyphs) To UBound(TestingGlyphs)
          tmpString = tmpString & "   " & txtOutputPercentage(j).Text
        Next j
        Debug.Print "Failure on font: " & TrainFont & ". Character: " & TestingGlyphs(l) & vbCrLf & _
          "Output: " & tmpString
        Failures = Failures + 1
      End If
      DoEvents
      If Stopping = True Then
        Exit For
      End If
    Next l
    ProgressBar2.Value = i + 1
    lblCycles.Caption = "Cycles: " & i + 1 & " / " & UBound(TrainingFonts) - LBound(TrainingFonts) + 1
    If Stopping = True Then
      Exit For
    End If
  Next i
  If Stopping = True Then
    Running = False
    picGlyph.Visible = True
    ProgressBar2.Visible = False
    lblTrainingFont.Caption = ""
    WipePanelsAndProgress
    Exit Sub
  End If
  
  ResultMessage = "Trained Fonts" & vbCrLf & _
                             "--------------------" & vbCrLf & _
                             "Successes: " & Successes & vbCrLf & _
                             "Failures: " & Failures & vbCrLf & _
                             "Average Correctness: " & Format$(RightTotal / ((UBound(TrainingFonts) - LBound(TrainingFonts) + 1) * (UBound(TestingGlyphs) - LBound(TestingGlyphs) + 1)) * 100, "0.00") & "%" & vbCrLf & _
                             "Average Incorrectness: " & Format$(WrongTotal / ((UBound(TrainingFonts) - LBound(TrainingFonts) + 1) * (UBound(TestingGlyphs) - LBound(TestingGlyphs) + 1)) * 100, "0.00") & "%" & vbCrLf
                             
  ' Reinit variables for second run on untrained fonts:
  Successes = 0
  Failures = 0
  RightTotal = 0
  WrongTotal = 0
  
  ' Untrained
  StatusBar1.Panels(1).Text = "Testing Unknown Fonts"
  ProgressBar2.Max = UBound(TestingFonts)
  ProgressBar2.Value = 0
  lblCycles.Caption = "Cycles: 0" & " / " & UBound(TestingFonts) - LBound(TestingFonts) + 1
  lblTrainingFont.Caption = "Test Font: " & ""
  For i = LBound(TestingFonts) To UBound(TestingFonts)
    For l = LBound(TestingGlyphs) To UBound(TestingGlyphs)
      TrainFont = TestingFonts(i)
      lblTrainingFont.Caption = "Testing Font: " & TrainFont
      PrepareGlyph TestingGlyphs(l), TrainFont
      Erase InputSet
      For j = 0 To CharacterWidth
        For k = 0 To CharacterHeight
          InputSet((j * CharacterHeight) + k) = 1 - ((picGlyph.Point(j, k) And &HFF&) / 255)
        Next k
      Next j
      oNet.SetInput InputSet
      oNet.ProcessOutput
      RightResult = oNet.OutputLayer(l + 1)
      'txtOutputPercentage(l).Text = Format$(RightResult * 100, "0.00") & "%"
      WrongResult = 0
      ShowOutput l
      For j = LBound(TestingGlyphs) To UBound(TestingGlyphs)
        If j <> l Then
          tmpResult = oNet.OutputLayer(j + 1)
          If tmpResult > WrongResult Then
            WrongResult = tmpResult
          End If
          'txtOutputPercentage(j).Text = Format$(tmpResult * 100, "0.00") & "%"
        End If
      Next j
      RightTotal = RightTotal + RightResult
      WrongTotal = WrongTotal + WrongResult
      If RightResult >= 1 - AcceptableShortFall And WrongResult <= 0 + AcceptableOpposableError Then
        Successes = Successes + 1
      Else
        tmpString = ""
        For j = LBound(TestingGlyphs) To UBound(TestingGlyphs)
          tmpString = tmpString & "   " & txtOutputPercentage(j).Text
        Next j
        Debug.Print "Failure on font: " & TrainFont & ". Character: " & TestingGlyphs(l) & vbCrLf & _
          "Output: " & tmpString
        Failures = Failures + 1
      End If
      DoEvents
      If Stopping = True Then
        Exit For
      End If
    Next l
    ProgressBar2.Value = i
    lblCycles.Caption = "Cycles: " & i + 1 & " / " & UBound(TestingFonts) - LBound(TestingFonts) + 1
    If Stopping = True Then
      Exit For
    End If
  Next i
  
  If Stopping = True Then
    Running = False
    picGlyph.Visible = True
    ProgressBar2.Visible = False
    lblTrainingFont.Caption = ""
    WipePanelsAndProgress
    Exit Sub
  End If
  
  ResultMessage = ResultMessage & vbCrLf & _
                            "UnTrained Fonts" & vbCrLf & _
                             "--------------------" & vbCrLf & _
                             "Successes: " & Successes & vbCrLf & _
                             "Failures: " & Failures & vbCrLf & _
                             "Average Correctness: " & Format$(RightTotal / ((UBound(TestingFonts) - LBound(TestingFonts) + 1) * (UBound(TestingGlyphs) - LBound(TestingGlyphs) + 1)) * 100, "0.00") & "%" & vbCrLf & _
                             "Average Incorrectness: " & Format$(WrongTotal / ((UBound(TestingFonts) - LBound(TestingFonts) + 1) * (UBound(TestingGlyphs) - LBound(TestingGlyphs) + 1)) * 100, "0.00") & "%" & vbCrLf

  Debug.Print ResultMessage ' so I can copy/paste it to "NotesOnCode" module!
  MsgBox ResultMessage, vbInformation
  
  Running = False
  picGlyph.Visible = True
  ProgressBar2.Visible = False
  lblTrainingFont.Caption = ""
  WipePanelsAndProgress
End Sub

Private Sub Command1_Click()
  ' intended to test net output.
Dim InputSet(0 To ((CharacterWidth + 1) * (CharacterHeight + 1))) As Long
Dim i As Long, j As Long, k As Long, l As Long
Dim RightResult As Double
Dim WrongResult As Double
Dim tmpResult As Double
Dim RealTextExtent As RECT
Dim MaxExtents As DWORD
  
  If Running = True Then
    Exit Sub
  End If
  MaxExtents.low = picGlyphBits.ScaleHeight
  MaxExtents.high = picGlyphBits.ScaleWidth
  RealTextExtent = GetTrueExtents(MaxExtents)
  If RealTextExtent.Left = -1 Or RealTextExtent.Right = -1 Or RealTextExtent.Bottom = -1 Or RealTextExtent.Top = -1 Then
    ' must be a spacer character
    Exit Sub
  End If
  
  SetStretchBltMode picGlyph.hdc, StretchBltModes.HALFTONE ' Dont really know whether this does anything. Shrug.  It doesn't hurt.
  SetStretchBltMode picGlyphBits.hdc, StretchBltModes.HALFTONE
  StretchBlt picGlyph.hdc, 0, 0, CharacterWidth, CharacterHeight, _
          picGlyphBits.hdc, RealTextExtent.Left, RealTextExtent.Top, _
          (RealTextExtent.Bottom - RealTextExtent.Top) * (CharacterWidth / CharacterHeight), _
          RealTextExtent.Bottom - RealTextExtent.Top, vbSrcCopy

  ' Draw lines to indicate real text extent.
  picGlyphBits.Line (RealTextExtent.Left, RealTextExtent.Top)-(RealTextExtent.Right, RealTextExtent.Top), vbGreen
  picGlyphBits.Line (RealTextExtent.Left, RealTextExtent.Top)-(RealTextExtent.Left, RealTextExtent.Bottom), vbYellow
  picGlyphBits.Line (RealTextExtent.Right, RealTextExtent.Top)-(RealTextExtent.Right, RealTextExtent.Bottom), vbMagenta
  picGlyphBits.Line (RealTextExtent.Left, RealTextExtent.Bottom)-(RealTextExtent.Right, RealTextExtent.Bottom), vbCyan
    
  picGlyphBits.Refresh
  FlipAndShift
  DoEvents
  
  If Not oNet Is Nothing Then
    ' We need to check the output from the NN
    ' Copy glyph into training set array.
    For j = 0 To CharacterWidth
      If j = CharacterWidth Then
        Exit For
      End If
      For k = 0 To CharacterHeight
        If k = CharacterHeight Then
          Exit For
        End If
        InputSet((j * CharacterHeight) + k) = 1 - ((picGlyph.Point(j, k) And &HFF&) / 255)
      Next k
    Next j
    oNet.SetInput InputSet
    oNet.ProcessOutput
    ' Is the character we're testing for one of the test glyphs?
    l = -1
    ShowOutput l
    If l > -1 Then
      RightResult = oNet.OutputLayer(l + 1)
      'txtOutputPercentage(l).Text = Format$(RightResult * 100, "0.00") & "%"
    End If
    WrongResult = 0
    For j = LBound(TestingGlyphs) To UBound(TestingGlyphs)
      If j <> l Then
        tmpResult = oNet.OutputLayer(j + 1)
        If tmpResult > WrongResult Then
          WrongResult = tmpResult
        End If
        'txtOutputPercentage(j).Text = Format$(tmpResult * 100, "0.00") & "%"
      End If
    Next j
  End If
  
End Sub

Private Sub Command2_Click()
  Set picGlyphBits = LoadPicture
End Sub

Private Sub Command3_Click()
  ShellEx App.Path & "\Readme.txt"
End Sub

' Output now shown as true percentage, as requested by Ulli.
Private Sub ShowOutput(Optional CorrectResult As Long = -1)
Dim j As Long
Dim TotalOutputAmount As Double
Dim Cap As Double
Dim Highest As Double
Dim HighPtr As Long

  #If OutputStyle = 1 Then
    chkCorrect.Value = 0
    For j = LBound(TestingGlyphs) To UBound(TestingGlyphs)
      TotalOutputAmount = TotalOutputAmount + oNet.OutputLayer(j + 1)
    Next j
    Cap = 0
    Highest = 0
    For j = LBound(TestingGlyphs) To UBound(TestingGlyphs)
      If oNet.OutputLayer(j + 1) > Highest Then
        Highest = oNet.OutputLayer(j + 1)
        HighPtr = j
      End If
      txtOutputPercentage(j).Text = Format$(oNet.OutputLayer(j + 1) / TotalOutputAmount * 100, "0.00") & "%"
      txtOutputPercentage(j).BackColor = vbWindowBackground
      Cap = Cap + (oNet.OutputLayer(j + 1) / TotalOutputAmount * 100)
    Next j
    If HighPtr = CorrectResult Then
      chkCorrect.Value = 1
    End If
    txtOutputPercentage(HighPtr).BackColor = vbCyan
    lblTotPercent.Caption = Cap
  #Else
    chkCorrect.Value = 0
    Cap = 0
    Highest = 0
    For j = LBound(TestingGlyphs) To UBound(TestingGlyphs)
      If oNet.OutputLayer(j + 1) > Highest Then
        Highest = oNet.OutputLayer(j + 1)
        HighPtr = j
      End If
      txtOutputPercentage(j).Text = Format$(oNet.OutputLayer(j + 1) * 100, "0.00") & "%"
      txtOutputPercentage(j).BackColor = vbWindowBackground
      Cap = Cap + oNet.OutputLayer(j + 1) * 100
    Next j
    If HighPtr = CorrectResult Then
      chkCorrect.Value = 1
    End If
    txtOutputPercentage(HighPtr).BackColor = vbCyan
    lblTotPercent.Caption = Cap
  #End If
End Sub

Private Sub Command4_Click()
  oNet.Jitter
End Sub

' This tests the trained net with a single char
Private Sub Form_KeyPress(KeyAscii As Integer)
Dim InputSet(0 To ((CharacterWidth + 1) * (CharacterHeight + 1))) As Long
Dim i As Long, j As Long, k As Long, l As Long
Dim RightResult As Double
Dim WrongResult As Double
Dim tmpResult As Double

  If Running = True Then
    Exit Sub
  End If
  PrepareGlyph Chr(KeyAscii), Combo1.Text

  If Not oNet Is Nothing Then
    ' We need to check the output from the NN
    ' Copy glyph into training set array.
    For j = 0 To CharacterWidth
      If j = CharacterWidth Then
        Exit For
      End If
      For k = 0 To CharacterHeight
        If k = CharacterHeight Then
          Exit For
        End If
        InputSet((j * CharacterHeight) + k) = 1 - ((picGlyph.Point(j, k) And &HFF&) / 255)
      Next k
    Next j
    oNet.SetInput InputSet
    oNet.ProcessOutput
    ' Is the character we're testing for one of the test glyphs?
    l = -1
    For j = LBound(TestingGlyphs) To UBound(TestingGlyphs)
      If Chr(KeyAscii) = TestingGlyphs(j) Then
        l = j
        Exit For
      End If
    Next j
    ShowOutput j
    If l > -1 Then
      RightResult = oNet.OutputLayer(l + 1)
      'txtOutputPercentage(l).Text = Format$(RightResult * 100, "0.00") & "%"
    End If
    WrongResult = 0
    For j = LBound(TestingGlyphs) To UBound(TestingGlyphs)
      If j <> l Then
        tmpResult = oNet.OutputLayer(j + 1)
        If tmpResult > WrongResult Then
          WrongResult = tmpResult
        End If
        'txtOutputPercentage(j).Text = Format$(tmpResult * 100, "0.00") & "%"
      End If
    Next j
  End If
  KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    Stopping = True
  End If
  KeyCode = 0
End Sub

Private Sub Form_Load()
  ' Populate Combo with fonts.
  FillComboWithFonts Combo1, True
  'SortComboBox Combo1
  Combo1.ListIndex = 0
  ' Other initialisations
  TrainingIterations = 1000
  SetUpOutputButtons
  Randomize
End Sub

Private Sub SetUpOutputButtons()
Dim i As Long
  PopulateTrainingFonts
  For i = LBound(TestingGlyphs) To UBound(TestingGlyphs)
    If i > 0 Then
      Load lblOutput(i)
      lblOutput(i).Left = lblOutput(i - 1).Left + lblOutput(i - 1).Width
      Load txtOutputPercentage(i)
      txtOutputPercentage(i).Left = lblOutput(i - 1).Left + lblOutput(i - 1).Width
    End If
    lblOutput(i).Caption = "  " & TestingGlyphs(i)
    txtOutputPercentage(i).Text = ""
    lblOutput(i).Visible = True
    txtOutputPercentage(i).Visible = True
  Next
End Sub

' Create Network
Private Sub btCreate_Click()

  If Running = True Then
    Stopping = True
    Exit Sub
  Else
    Stopping = False
    Running = True
  End If
  
  ProgressBar1.Min = 0
  ProgressBar1.Max = 100
  ProgressBar1.Value = 0
  ProgressBar1.Visible = True
  StatusBar1.Panels(1).Text = "Creating Net"
  Me.MousePointer = vbArrowHourglass
  
  ' Create uninitialised Neural Net
  If oNet Is Nothing Then
    Set oNet = New NetFast
  Else
    oNet.DestroyNicely
  End If
  oNet.CreateNet Nothing, 651, 96, 144, 10
                      ' Input layer = 4 neurons
                      ' 1st Hidden layer = 3 neurons
                      ' Output layer = 2 neurons
  oNet.LearningCoefficient = 0.9
  oNet.LearningRateIncrease = 1.05
  oNet.LearningRateDecrease = 0.97
  oNet.AnnealingEpoch = (UBound(TrainingFonts) - LBound(TrainingFonts) + 1) * (UBound(TestingGlyphs) - LBound(TestingGlyphs) + 1) * 2
  
  ' Tidy up
  picGlyph.Visible = True
  WipePanelsAndProgress

End Sub

Private Sub btSave_Click()
    If Not oNet Is Nothing Then
      With CommonDialog1
        .CancelError = True
        .DefaultExt = "NNT"
        .DialogTitle = "Save Neural Net As"
        .filename = ""
        .Filter = "Neural Net Files (*.NNT)|*.NNT|Neural Net (Version 1.2) File|*.NNT|All Files (*.*)|*.*"
        .FilterIndex = 1
        .InitDir = App.Path
        .Flags = cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
        On Error Resume Next
        .ShowSave
        If Err Then
          Err.Clear
          Exit Sub
        End If
        On Error GoTo 0
      End With
      Me.MousePointer = vbArrowHourglass
      ProgressBar1.Min = 0
      ProgressBar1.Max = 100
      ProgressBar1.Visible = True
      StatusBar1.Panels(1).Text = "Saving Net"
      If CommonDialog1.FilterIndex = 2 Then
        oNet.SaveNet CommonDialog1.filename, "1.2"
      Else
        oNet.SaveNet CommonDialog1.filename
      End If
      WipePanelsAndProgress
    Else
      MsgBox "No Net to save", vbExclamation
    End If
End Sub

Private Sub btKill_Click()

    If oNet Is Nothing Then
      Exit Sub
    End If
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    ProgressBar1.Visible = True
    StatusBar1.Panels(1).Text = "Destroying Net"
    oNet.DestroyNicely
    Set oNet = Nothing
    WipePanelsAndProgress
End Sub

Private Sub WipePanelsAndProgress()
    ProgressBar1.Visible = False
    StatusBar1.Panels(1).Text = "Ready"
    StatusBar1.Panels(2).Text = ""
    Running = False
    Stopping = False
    lblCycles.Caption = ""
    Me.MousePointer = vbDefault
End Sub


Private Sub btRestore_Click()
    With CommonDialog1
      .CancelError = True
      .DefaultExt = "NNT"
      .DialogTitle = "Load Neural Net"
      .filename = ""
      .Filter = "Neural Net Files (*.NNT)|*.NNT|All Files(*.*)|*.*"
      .FilterIndex = 1
      .InitDir = App.Path
      .Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNFileMustExist
      On Error Resume Next
      .ShowOpen
      If Err Then
        Err.Clear
        Exit Sub
      End If
      On Error GoTo 0
    End With


    Me.MousePointer = vbArrowHourglass
    ProgressBar1.Min = 0
    ProgressBar1.Max = 100
    ProgressBar1.Visible = True
    StatusBar1.Panels(1).Text = "Loading Net"
    If Not oNet Is Nothing Then
      oNet.DestroyNicely
    Else
      Set oNet = New NetFast
    End If
    oNet.LoadNet CommonDialog1.filename
    lblTotalTrainCycles.Caption = oNet.TrainingCycles
    WipePanelsAndProgress
End Sub

Private Sub Form_Resize()
Dim i As Long
Dim Width As Single
  ' resize status panels
  Width = 0
  For i = 1 To StatusBar1.Panels.Count - 1
    Width = Width + StatusBar1.Panels(i).Width
  Next i
  If Me.WindowState = vbMinimized Then
  Else
    StatusBar1.Panels(StatusBar1.Panels.Count).Width = Me.Width - Width
  End If
End Sub

Private Sub oNet_InfoMessage(vTag As Variant, Info As String)
  StatusBar1.Panels(2).Text = Info
  DoEvents
End Sub


Private Sub oNet_Progress(vTag As Variant, Percentage As Single)
  ProgressBar1.Value = Percentage
  DoEvents
End Sub

Private Sub picGlyphBits_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = vbLeftButton Then
    picGlyphBits.DrawWidth = 10
    picGlyphBits.PSet (x, y), vbBlack
    picGlyphBits.DrawWidth = 1
  End If
End Sub

' Dead code below (i.e. No longer used):
'Private Sub SelectGlyph(pCharacter As String, pFont As String)
'  Set picGlyph = LoadPicture
'  picGlyph.CurrentX = 0
'  picGlyph.CurrentY = 0
'  On Error Resume Next
'  If pFont = "" Then
'    picGlyph.font = "Arial"
'  Else
'    picGlyph.font = pFont
'  End If
'  If Err Then
'    Err.Clear
'    On Error GoTo 0
'    picGlyph.font = "Arial"  ' Defaults to this if the font isn't installed locally.
'  End If
'  picGlyph.fontsize = 12
'  picGlyph.Print pCharacter
'End Sub
'
'Private Sub SelectBigGlyph(pCharacter As String, pFont As String)
'  Set picGlyphBits = LoadPicture
'  picGlyphBits.CurrentX = 10
'  picGlyphBits.CurrentY = 10
'  On Error Resume Next
'  If pFont = "" Then
'    picGlyphBits.font = "Arial"
'  Else
'    picGlyphBits.font = pFont
'  End If
'  If Err Then
'    Err.Clear
'    On Error GoTo 0
'    picGlyphBits.font = "Arial"  ' Defaults to this if the font isn't installed locally.
'  End If
'  picGlyphBits.fontsize = 120
'  picGlyphBits.Print pCharacter
'End Sub
'
'Private Sub ShrinkAndAntiAlias(pFont As String)
'Dim i As Long, j As Long, k As Long, l As Long
'Dim TotalBlackPixels As Long
'Dim Avg As Single
'Dim Calc As Long
'
'  Set picGlyph = LoadPicture
'  picGlyph.font = pFont
'  picGlyph.CurrentX = 0
'  picGlyph.CurrentY = 0
'  picGlyph.font.size = 12
'
'  For i = 0 To picGlyphBits.ScaleWidth - 1 Step 10
'    For j = 0 To picGlyphBits.ScaleHeight - 1 Step 10
'      TotalBlackPixels = 0
'      For k = 0 To 9
'        For l = 0 To 9
'          If picGlyphBits.Point(i + k, j + l) = &H0 Then
'            TotalBlackPixels = TotalBlackPixels + 1
'          End If
'        Next l
'      Next k
'      Avg = TotalBlackPixels / 100
'      Calc = (100 - TotalBlackPixels) * 2.55  ' To get colour intensity
'      picGlyph.PSet (i / 10, j / 10), RGB(Calc, Calc, Calc)
'    Next j
'  Next i
'
'End Sub
'
'Private Sub Pixelise(pTextHeight As Long, pTextWidth As Long)
'Dim i As Long
'Dim j As Long
'
'  Set picGlyphBits = LoadPicture
'  For i = 0 To pTextWidth
'    For j = 0 To pTextHeight
'      picGlyphBits.Line (i * 10, j * 10)-((i * 10) + 8, (j * 10) + 8), picGlyph.Point(i, j), BF
'    Next j
'  Next i
'  ' Draw blue lines in picture
'  For i = 0 To pTextWidth + 1
'    picGlyphBits.Line (-1 + (i * 10), 0)-(-1 + (i * 10), picGlyphBits.ScaleHeight), vbBlue
'  Next
'  ' Draw Red lines
'  For j = 0 To pTextHeight + 1
'    picGlyphBits.Line (0, -1 + (j * 10))-(picGlyphBits.ScaleWidth, -1 + (j * 10)), vbRed
'  Next
'End Sub
'
'Private Sub Pixelise2(pTextHeight As Long, pTextWidth As Long)
'Dim i As Long
'Dim j As Long
'
'  Set picGlyphBits = LoadPicture
'  For i = 0 To pTextWidth
'    For j = 0 To pTextHeight
'      picGlyphBits.Line (i * 10, j * 10)-((i * 10) + 8, (j * 10) + 8), picFlipG.Point(i, j), BF
'    Next j
'  Next i
'  ' Draw blue lines in picture
'  For i = 0 To pTextWidth + 1
'    picGlyphBits.Line (-1 + (i * 10), 0)-(-1 + (i * 10), picGlyphBits.ScaleHeight), vbBlue
'  Next
'  ' Draw Red lines
'  For j = 0 To pTextHeight + 1
'    picGlyphBits.Line (0, -1 + (j * 10))-(picGlyphBits.ScaleWidth, -1 + (j * 10)), vbRed
'  Next
'End Sub

'Private Sub FlipAndShift()
'Dim i As Long
'Dim j As Long
'Dim k As Long
'Dim NotBlank As Boolean
'Dim LastNonBlankRow As Long
'Dim FirstNonBlankCol As Long
'  ' The purpose of this function is to take the glyph drawn in
'  ' picGlyph and find the last row where pixels are used.
'  ' Then, copy an upside down image into picFlipG
'
'  ' No longer necessary - just blit one to the other
'  Set picFlipG = LoadPicture
'  BitBlt picFlipG.hdc, 0, 0, CharacterWidth, CharacterHeight, _
'    picGlyph.hdc, 0, 0, vbSrcCopy
'  Exit Sub
'
''  ' Blank picFlipG
''  Set picFlipG = LoadPicture
''
''  ' Find last empty row
''  LastNonBlankRow = picGlyph.ScaleHeight - 1
''
''  For i = picGlyph.ScaleHeight - 1 To 0 Step -1
''    NotBlank = False
''    For j = 0 To picGlyph.ScaleWidth - 1
''      If picGlyph.Point(j, i) <> &HFFFFFF Then
''        NotBlank = True
''        Exit For
''      End If
''    Next j
''    If NotBlank = True Then
''      Exit For
''    End If
''    LastNonBlankRow = i
''  Next i
''  LastNonBlankRow = LastNonBlankRow - 1
''
''  ' Find first column
''  FirstNonBlankCol = 0
''  For i = 0 To picGlyph.ScaleWidth - 1
''    NotBlank = False
''    For j = 0 To picGlyph.ScaleHeight - 1
''      If picGlyph.Point(i, j) <> &HFFFFFF Then
''        NotBlank = True
''        Exit For
''      End If
''    Next j
''    If NotBlank = True Then
''      Exit For
''    End If
''    FirstNonBlankCol = FirstNonBlankCol + 1
''  Next i
''  If FirstNonBlankCol >= picGlyph.ScaleWidth - 1 Then
''    FirstNonBlankCol = 0
''  End If
''
''  ' Now loads and loads and loads of blitting.
''  j = 0
''  For i = LastNonBlankRow To 0 Step -1
''    k = BitBlt(picFlipG.hdc, 0, j, picGlyph.ScaleWidth - FirstNonBlankCol, 1, picGlyph.hdc, FirstNonBlankCol, i, vbSrcCopy)
''    j = j + 1
''  Next i
'End Sub


