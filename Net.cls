VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Layer"
Attribute VB_Ext_KEY = "Member0" ,"Layer"
Option Explicit
Option Base 0
Option Compare Text

'----------------------------------------------------------------------------------------------------------------
' Neural Net Code and classes originally by Ulli (umgedv@aol.com)
' Conversion to ActiveX Object Library and other improvements
' by Jonathan Daniel (bigcalm@hotmail.com)
'
' This code is shareware.  Please credit the authors if you use this code.
'
' Bugs, fixes, improvements, and suggestions to bigcalm@hotmail.com
'
'----------------------------------------------------------------------------------------------------------------

' This class is the only public create-able one in this Project.
' Contains methods/properties/events for the Net as a whole.

' Constants for load/save.
Private Const OCXNAME As String = "Perceptron"
Private Const VERSION As String = "1.2"

' Private collections & Temporary variables
Private oLayer      As cLayer
Private oNeuron     As cNeuron
Private oDendrite   As cDendrite
Private Layers As Collection

' Modular Property variables
Private mvarCreated As Boolean
Private mTag As Variant
Private mTrainingCycles As Long
Private mLearningCoefficient As Double
Private mSumSquaredError As Double
Private mRunning As Boolean ' Because we're allowing the caller to utilise
  ' DoEvents when we Raise an event, we need a Boolean value to stop
  ' something being called twice (or say Destroy being called while we're training).
  ' Also allow owner to read this property if they want.
Private mStopping As Boolean ' Set this flag to abort current procedure.
      ' Set by calling StopWorking.  Not available as a property (yet).

' Events
' These three events will be raised so the controlling application has information
' on what's going on.  It'll also give the controlling app a chance to do things
' like DoEvents for example.
' These are useful for progress bars/panels/etc.  Just define your net "WithEvents"
' to use these.
Public Event InfoMessage(vTag As Variant, Info As String)
Public Event Progress(vTag As Variant, Percentage As Single)
'
' Example use for Progress...
' Private WithEvents MyNet as New cNet
' Private Sub Form_Click()
'   If MyNet.Running = True Then
'     Exit Sub
'   End If
'   ProgressBar1.Visible = True
'   ProgressBar1.Min = 0
'   ProgressBar1.Max = 0
'   ProgressBar1.Value = 0
'   MyNet.CreateNet(Array(60,8,2))
'   ProgressBar1.Visible = False
' End Sub
' Private Sub MyNet_Progress(vTag as Variant, Percentage as Single)
'   ProgressBar1.Value = Percentage
'   DoEvents
' End Sub


' Property Let/Get/Sets.
Public Property Get Created() As Boolean
  Created = mvarCreated
End Property
Friend Property Let Created(ByVal vData As Boolean)
  mvarCreated = vData
End Property
Public Property Get TrainingCycles() As Long
  TrainingCycles = mTrainingCycles
End Property
Friend Property Let TrainingCycles(ByVal vData As Long)
  mTrainingCycles = vData
End Property
Public Property Let LearningCoefficient(ByVal vData As Double)
  mLearningCoefficient = vData
End Property
Public Property Get LearningCoefficient() As Double
  LearningCoefficient = mLearningCoefficient
End Property
Public Property Get SumSquaredError() As Double
  SumSquaredError = mSumSquaredError
End Property
Friend Property Let SumSquaredError(ByVal vData As Double)
  mSumSquaredError = vData
End Property
Public Property Get AverageSquaredError() As Double
  If mvarCreated = True Then
    AverageSquaredError = mSumSquaredError / Layers(Layers.Count).NeuronCount
  Else
    AverageSquaredError = 0
  End If
End Property

Public Property Get Tag() As Variant
  If IsObject(mTag) Or mTag Is Nothing Then
    Set Tag = mTag
  Else
    Tag = mTag
  End If
End Property
Public Property Set Tag(ByVal vData As Variant)
  Set mTag = vData
End Property
Public Property Let Tag(ByVal vData As Variant)
  mTag = vData
End Property

Public Property Get Running() As Boolean
  Running = mRunning
End Property
Private Property Let Running(ByVal vData As Boolean)
  mRunning = vData
End Property
Public Property Get TotalNeuronCount() As Long
  TotalNeuronCount = 0
  For Each oLayer In Layers
    TotalNeuronCount = TotalNeuronCount + oLayer.NeuronCount
  Next
End Property


' Layers collection properties
Public Sub Add(Layer As cLayer)
    Layers.Add Layer
End Sub
Public Property Get Layer(Index As Long) As cLayer
    Set Layer = Layers(Index)
End Property
Public Property Get LayerCount() As Long
    LayerCount = Layers.Count
End Property
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = Layers.[_NewEnum]
End Property

' Class Events
Private Sub Class_Initialize()
    Set Layers = New Collection
    Set mTag = Nothing
    mLearningCoefficient = 1.5
    mSumSquaredError = 0
    mvarCreated = False
End Sub
Private Sub Class_Terminate()
    mStopping = True
    Set Layers = Nothing
    mvarCreated = False
    If IsObject(mTag) Then
      Set mTag = Nothing
    End If
End Sub

' Neural Net functions
Public Sub Train(Data As Variant, Target As Variant)
  If mvarCreated = True Then
    If mRunning = True Then
      Exit Sub
    End If
    mRunning = True
    mStopping = False

    SetInput Data
    If mStopping = False Then
      CalculateOutput
      If mStopping = False Then
        AdjustWeights Target
      End If
    End If
    
  Else
    Err.Raise vbObjectError + 1, "Perceptron", "You must initialise your net before you attempt to train it"
    mStopping = True
  End If
  If mStopping = False Then
    mTrainingCycles = mTrainingCycles + 1
  End If
  mRunning = False
End Sub

' Adds random noise to every dendrite.  Used to avoid hill-descent problems.
Public Sub Jitter(Optional MaxVariance As Double = 0.05)
Dim WorkToDo As Long
Dim Percentage As Long
Dim WorkDone As Long
Dim i As Long, j As Long, k As Long
Dim Variance As Double

  If mRunning = True Then
    Exit Sub
  End If
  mRunning = True
  mStopping = False

  ' Work out how much work we have to do first.
  WorkToDo = TotalNeuronCount
  Percentage = WorkToDo / 100
  WorkDone = 0
  
  ' Jiggle all layers
  i = 0
  For Each oLayer In Layers
    RaiseEvent InfoMessage(mTag, "Jiggling Layer " & i + 1)
    For Each oNeuron In oLayer
      For Each oDendrite In oNeuron
        Variance = (MaxVariance * Rnd * 2) - MaxVariance
        oDendrite.TransferWeight = oDendrite.TransferWeight + Variance
        If oDendrite.TransferWeight < -1 Then
          oDendrite.TransferWeight = -1
        End If
        If oDendrite.TransferWeight > 1 Then
          oDendrite.TransferWeight = 1
        End If
      Next oDendrite
      WorkDone = WorkDone + 1
      If WorkDone Mod Percentage = 0 Then
        RaiseEvent Progress(mTag, (WorkDone / WorkToDo) * 100)
      End If
      If mStopping = True Then
        Exit For
      End If
    Next oNeuron
    If mStopping = True Then
      Exit For
    End If
    i = i + 1
  Next oLayer
  mRunning = False
End Sub

' dont like having this public, but can't see another way around it at the mo.
Public Sub SetInput(Data As Variant)
Dim i As Long
    i = 0
    RaiseEvent InfoMessage(mTag, "Initialising Input Layer")
    If mStopping = True Then
      Exit Sub
    End If
    For Each oNeuron In Layer(1)
        oNeuron.Value = CDbl(Data(i))
        i = i + 1
    Next oNeuron
End Sub

' Use this function to retrieve the values in the output layer after calling
' ProcessOutput or CalculateOutput
Public Property Get OutputLayer(Index As Long) As Double
  If mvarCreated = True Then
    OutputLayer = Layers(Layers.Count).Neuron(Index).Value
  Else
    Err.Raise vbObjectError + 2, "Perceptron", "You must initialise your net before you attempt to retrieve output from it"
    OutputLayer = 0
  End If
End Property

' This can be publicly called - the function that actually does the work is a hidden
' procedure called CalculateOutput (below).
Public Sub ProcessOutput()
  If mvarCreated = False Then
    Err.Raise vbObjectError + 2, "Perceptron", "You must initialise your net before you attempt to retrieve output from it"
    Exit Sub
  End If
  If mRunning = True Then
    Exit Sub
  End If
  mRunning = True
  mStopping = False
  CalculateOutput
  mRunning = False
End Sub

Friend Sub CalculateOutput()
Dim i As Long
Dim TotalToDo As Long
Dim Percentage As Long
Dim WorkDone As Long
  
  ' Calculate work to be done
  TotalToDo = TotalNeuronCount
  Percentage = TotalNeuronCount \ 100
  WorkDone = 0
  i = 1
  
  ' Fire all neurons
  For Each oLayer In Layers
    RaiseEvent InfoMessage(mTag, "Calculating Output For Layer " & i)
      For Each oNeuron In oLayer
          oNeuron.Activate
          WorkDone = WorkDone + 1
          If WorkDone Mod Percentage = 0 Then
            RaiseEvent Progress(mTag, (WorkDone / TotalToDo) * 100)
          End If
          If mStopping = True Then
            Exit For
          End If
      Next oNeuron
      If mStopping = True Then
        Exit For
      End If
      i = i + 1
  Next oLayer
End Sub

Private Sub AdjustWeights(Target As Variant)
Dim i As Long
Dim Percentage As Long
Dim WorkDone As Long
Dim TotalToDo As Long
Dim SSE As Double

    i = 0
    RaiseEvent InfoMessage(mTag, "Training the Net")
    
    'output layer
    SumSquaredError = 0
    For Each oNeuron In Layers(Layers.Count)
        With oNeuron
            .Error = CDbl(Target(i)) - .Value 'raw error
            SSE = SSE + (.Error * .Error)
        End With
        If mStopping = True Then
          Exit For
        End If
        i = i + 1
    Next oNeuron
    SumSquaredError = SSE
    
    If mStopping = True Then
      Exit Sub
    End If
    
    ' Calculate work to be done
    TotalToDo = TotalNeuronCount
    Percentage = TotalNeuronCount \ 100
    WorkDone = 0
    
    'hidden layers
    For i = Layers.Count To 2 Step -1
        RaiseEvent InfoMessage(mTag, "Running Back Propogation on Layer " & i)
        For Each oNeuron In Layers(i)
            oNeuron.BackPropagate
            WorkDone = WorkDone + 1
            If WorkDone Mod Percentage = 0 Then
              RaiseEvent Progress(mTag, (WorkDone / TotalToDo) * 100)
            End If
            If mStopping = True Then
              Exit For
            End If
        Next oNeuron
        If mStopping = True Then
          Exit For
        End If
    Next i
    
    If mStopping = True Then
      Exit Sub
    End If
    
    ' Calculate work to be done
    Percentage = TotalNeuronCount \ 100
    TotalToDo = TotalNeuronCount
    WorkDone = 0
    i = 1
    
    ' Update weights
    For Each oLayer In Layers
        RaiseEvent InfoMessage(mTag, "Updating Weights in Layer " & i)
        For Each oNeuron In oLayer
            oNeuron.UpdateWeights
            WorkDone = WorkDone + 1
            If WorkDone Mod Percentage = 0 Then
              RaiseEvent Progress(mTag, (WorkDone / TotalToDo) * 100)
            End If
            If mStopping = True Then
              Exit For
            End If
        Next oNeuron
        i = i + 1
        If mStopping = True Then
          Exit For
        End If
    Next oLayer
End Sub

' Create/Load/Save code
' I may change this to using a ParamArray instead - more convenient.
Public Sub CreateNet(Struc As Variant, Optional pTag As Variant = Nothing)
Dim oLayer         As cLayer
Dim oPreviousLayer As cLayer
Dim oNeuron        As cNeuron
Dim oDendrite      As cDendrite
Dim ll             As Long
Dim n              As Long
Dim nn             As Long
Dim d              As Long
Dim dd             As Long
Dim x              As Double
Dim l As Long
Dim TotalToDo As Long
Dim TotalDone As Long
Dim PercentageHit As Long
Dim i As Long

    If mRunning = True Then
      Exit Sub
    End If
    mRunning = True
    mStopping = False
    ' Initialise
    ' Need to destroy first
    Set Layers = Nothing
    Set Layers = New Collection
    mvarCreated = False
    Set Tag = pTag
    mTrainingCycles = 0
    mSumSquaredError = 0
    
    ' Calculate work to be done
    TotalToDo = 0
    For i = 0 To UBound(Struc)
      TotalToDo = TotalToDo + Struc(i)
    Next i
    PercentageHit = TotalToDo / 100
    TotalDone = 0
    
    ' Create layers.
    Set oPreviousLayer = Nothing
    With Me
        For l = 0 To UBound(Struc)
            RaiseEvent InfoMessage(mTag, "Creating Layer " & l + 1)
            Set oLayer = New cLayer
            With oLayer
                Set .PreviousLayer = oPreviousLayer
                For n = 1 To Struc(l)
                    Set oNeuron = New cNeuron

                    With oNeuron
                        .LearningCoefficient = mLearningCoefficient '1.5
                        .BaseValue = GetRand
                        If l Then 'not first layer
                            For d = 1 To Struc(l - 1)
                                Set oDendrite = New cDendrite
                                With oDendrite
                                    .TransferWeight = GetRand
                                    Set .ConnectedTo = oLayer.PreviousLayer.Neuron(d)
                                End With
                                .Add oDendrite
                            Next d
                        End If
                    End With
                    .Add oNeuron
                    TotalDone = TotalDone + 1
                    If TotalDone Mod PercentageHit = 0 Then
                      RaiseEvent Progress(mTag, (TotalDone / TotalToDo) * 100)
                    End If
                    If mStopping = True Then
                      Exit For
                    End If
                Next n
            End With
            If mStopping = True Then
              Exit For
            End If
            .Add oLayer
            Set oPreviousLayer = oLayer
        Next l
    End With
    If mStopping = True Then
    Else
      mvarCreated = True
    End If
    mRunning = False
End Sub

' Filename is a variant so you can either pass a file number (long) or a file name (string)
Public Function SaveNet(Filename As Variant, Optional SaveVersion As String = VERSION) As Boolean
Dim FileNumber As Long
Dim oLayer         As cLayer
Dim oPreviousLayer As cLayer
Dim oNeuron        As cNeuron
Dim oDendrite      As cDendrite
Dim ll             As Long
Dim Percentage As Long
Dim WorkDone As Long
Dim TotalToDo As Long
Dim i As Long
Dim SaveHeader As String
Dim strVersion As String

  If mRunning = True Then
    Exit Function
  End If
  mRunning = True
  mStopping = False
  
  If mvarCreated = False Then
      Err.Raise vbObjectError + 3, "Perceptron", "There is no Net to save."
  Else
      ' Version 1.2 save format
      On Error GoTo ErrHandler
      If VarType(Filename) = vbString Then
        FileNumber = FreeFile
        Open Filename For Binary Access Write As #FileNumber
      Else
        FileNumber = Filename
      End If
      On Error GoTo 0
      
      ' Work out how much there is to do
      RaiseEvent InfoMessage(mTag, "Saving Neural Net")
      Percentage = TotalNeuronCount
      WorkDone = 0
      TotalToDo = TotalNeuronCount
      i = 1
      SaveHeader = OCXNAME
      Put #FileNumber, , SaveHeader
      strVersion = VERSION
      Put #FileNumber, , strVersion ' Write header
      Put #FileNumber, , Me.LayerCount ' Write Number of layers
      Put #FileNumber, , Me.TotalNeuronCount ' Write Total number of neurons
      Put #FileNumber, , Me.TrainingCycles ' Write Total amount trained
      For Each oLayer In Me
          RaiseEvent InfoMessage(mTag, "Saving Layer " & i)
          Put #FileNumber, , oLayer.NeuronCount
          For Each oNeuron In oLayer
              Put #FileNumber, , oNeuron.BaseValue
              Put #FileNumber, , oNeuron.DendriteCount
              For Each oDendrite In oNeuron
                  Put #FileNumber, , oDendrite.TransferWeight
              Next oDendrite
              WorkDone = WorkDone + 1
              If WorkDone Mod Percentage = 0 Then
                RaiseEvent Progress(mTag, (WorkDone / TotalToDo) * 100)
              End If
              If mStopping = True Then
                Exit For
              End If
          Next oNeuron
          If mStopping = True Then
            Exit For
          End If
          i = i + 1
      Next oLayer
      Close #FileNumber
    End If
    
    mRunning = False
    If mStopping = True Then
      SaveNet = False ' aborted save.
    Else
      SaveNet = True
    End If
    Exit Function
ErrHandler:
  MsgBox "Failed to save net: " & Err.Description
  Err.Clear
  SaveNet = False
  mRunning = False
End Function

' Filename is variant so you can either pass a file number or a file name
Public Function LoadNet(Filename As Variant, Optional pTag As Variant = Nothing, Optional pVersion As String = VERSION) As Boolean
Dim oLayer         As cLayer
Dim oPreviousLayer As cLayer
Dim oNeuron        As cNeuron
Dim oDendrite      As cDendrite
Dim FileNumber As Long
Dim ll             As Long
Dim nn             As Long
Dim n As Long
Dim l As Long
Dim dd             As Long
Dim d As Long
Dim x As Double
Dim Percentage As Long
Dim WorkDone As Long
Dim i As Long
Dim WorkToDo As Long
Dim SaveHeader As String
Dim strVersion As String

    If mRunning = True Then
      Exit Function
    End If
    mRunning = True
    mStopping = False
 
    ' Destroy net before reloading
    Set Layers = Nothing
    Set Layers = New Collection
    mvarCreated = False
    
    Set mTag = pTag
    Set oPreviousLayer = Nothing
    
    RaiseEvent InfoMessage(mTag, "Loading Neural Net")
    
    With Me
        On Error GoTo ErrHandler
        If VarType(Filename) = vbString Then
          FileNumber = FreeFile
          Open Filename For Binary Access Read As #FileNumber
        Else
          FileNumber = Filename
        End If
        
        On Error GoTo 0
        SaveHeader = String(Len(OCXNAME), " ")
        strVersion = String(Len(VERSION), " ")
        Get #FileNumber, , SaveHeader
        Get #FileNumber, , strVersion
        If SaveHeader = OCXNAME And strVersion = VERSION Then
          ' Valid 1.2 file
        Else
          ' This is not a valid file.
          Err.Clear
          Close #FileNumber
          GoTo ErrHandler  ' apologies!!!!  This will be changed in future.
        End If
        Get #FileNumber, , ll ' Number of layers
        Get #FileNumber, , WorkToDo ' Total neuron count!
        Get #FileNumber, , mTrainingCycles ' Total training cycles

        ' Work out how much there is to do
        WorkDone = 0
        Percentage = WorkToDo \ 100
        
        For l = 1 To ll
            Set oLayer = New cLayer
            RaiseEvent InfoMessage(mTag, "Loading Layer " & l)
            With oLayer
                Set .PreviousLayer = oPreviousLayer
                Get #FileNumber, , nn
                For n = 1 To nn
                    Set oNeuron = New cNeuron
                    With oNeuron
                        Get #FileNumber, , x
                        .BaseValue = x
                        Get #FileNumber, , dd
                        For d = 1 To dd
                            Set oDendrite = New cDendrite
                            With oDendrite
                                Get #FileNumber, , x
                                .TransferWeight = x
                                Set .ConnectedTo = oPreviousLayer.Neuron(d)
                            End With
                            .Add oDendrite
                        Next d
                    End With
                    .Add oNeuron
                    WorkDone = WorkDone + 1
                    If WorkDone Mod Percentage = 0 Then
                      RaiseEvent Progress(mTag, (WorkDone / WorkToDo) * 100)
                    End If
                    If mStopping = True Then
                      Exit For
                    End If
                Next n
            End With
            If mStopping = True Then
              Exit For
            End If
            .Add oLayer
            Set oPreviousLayer = oLayer
        Next l
                  
    End With
    Close #FileNumber
    
    mSumSquaredError = 0
    mRunning = False
    If mStopping = True Then
      LoadNet = False ' aborted load
    Else
      mvarCreated = True
      LoadNet = True
    End If
    Exit Function
    
ErrHandler:
  If Err = False Then
    MsgBox "Failed to load net: File is not in the correct format"
  Else
    MsgBox "Failed to load net: " & Err.Description
  End If
  Err.Clear
  LoadNet = False
  mRunning = False
End Function

' This function will destroy the various layers, neurons and dendrites in a
' co-operative way.
' With a large net, using Set Net = Nothing will take ages - this function
' allows some kind of progress to show.  Uses raised events, so the controlling
' application can call DoEvents if they need to.
' Tested.
Public Sub DestroyNicely()
Dim WorkToDo As Long
Dim Percentage As Long
Dim WorkDone As Long
Dim i As Long, j As Long, k As Long

  If mRunning = True Then
    Exit Sub
  End If
  mRunning = True
  mStopping = False

  ' Work out how much work we have to do first.
  WorkToDo = TotalNeuronCount
  Percentage = WorkToDo / 1000
  WorkDone = 0
  
  ' Destroy all layers
  i = 0
  For Each oLayer In Layers
    RaiseEvent InfoMessage(mTag, "Destroying Layer " & i + 1)
    For Each oNeuron In oLayer
      oNeuron.Destroy
      Set oNeuron = Nothing
      WorkDone = WorkDone + 1
      If WorkDone Mod Percentage = 0 Then
        RaiseEvent Progress(mTag, (WorkDone / WorkToDo) * 100)
      End If
      If mStopping = True Then
        Exit For
      End If
    Next oNeuron
    Set oLayer = Nothing
    If mStopping = True Then
      Exit For
    End If
    i = i + 1
  Next oLayer
  mvarCreated = False
  mTrainingCycles = 0
  mRunning = False
End Sub

Public Sub StopWorking()
  mStopping = True
End Sub

' Miscellaneous functions
Private Function GetRand() As Double

    Do
        GetRand = 1 - Rnd - Rnd
    Loop Until Abs(GetRand) >= 0.1

End Function

