VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NetFast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
#Const ShowDebugTimes = 0

'----------------------------------------------------------------------------------------------------------------
' Neural Net Code and classes originally by Ulli (umgedv@aol.com)
' Conversion to ActiveX Object Library and other improvements
' by Jonathan Daniel (bigcalm@hotmail.com)
' This class is an array based version of a neural net.
'
' This code is shareware.  Please credit the authors if you use this code.
'
' Bugs, fixes, improvements, and suggestions to bigcalm@hotmail.com
'
'----------------------------------------------------------------------------------------------------------------
' Many thanks to Chikh for all his help.
' Some code was adapted from Jason Tiscione's java code.
' Urgh my brain hurts after reading the neural net FAQ at faqs.org


Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Type Synapse
  TransferWeight As Double
End Type

Private Type Neuron
  Value As Double
  Error As Double
  BaseValue As Double
End Type

Private Type Layer
  StartNeuron As Long
  EndNeuron As Long
  StartDendrite As Long
  EndDendrite As Long
  NeuronCount As Long
  DendriteCount As Long
End Type

Private StartTime As Long

' Constants for load/save.
Private Const OCXNAME As String = "Perceptron"
Private Const VERSION As String = "1.3"

' Main objects
Private Dendrites() As Synapse
Private Neurons() As Neuron
Private Layers() As Layer

' From cNet
Private mvarCreated As Boolean
Private mTag As Variant
Private mTrainingCycles As Long
Private mLearningCoefficient As Double
Private mLearningRateIncrease As Double
Private mLearningRateDecrease As Double
'Private mMaximumErrorRate As Double
Private mAnnealingEpoch As Long
Private mAnnealingSSE As Double
Private mLastAnnealingSSE As Double
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

' Property Let/Get/Sets. from cNet
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
Public Property Let LearningRateIncrease(ByVal vData As Double)
  mLearningRateIncrease = vData
End Property
Public Property Get LearningRateIncrease() As Double
  LearningRateIncrease = mLearningRateIncrease
End Property
Public Property Let LearningRateDecrease(ByVal vData As Double)
  mLearningRateDecrease = vData
End Property
Public Property Get LearningRateDecrease() As Double
  LearningRateDecrease = mLearningRateDecrease
End Property
'Public Property Let MaximumErrorRate(ByVal vData As Double)
'  mMaximumErrorRate = vData
'End Property
'Public Property Get MaximumErrorRate() As Double
'  MaximumErrorRate = mMaximumErrorRate
'End Property

Public Property Get SumSquaredError() As Double
  SumSquaredError = mSumSquaredError
End Property
Friend Property Let SumSquaredError(ByVal vData As Double)
  mSumSquaredError = vData
End Property

' This should properly be called "Mean Square Error"
Public Property Get AverageSquaredError() As Double
  If mvarCreated = True Then
    AverageSquaredError = mSumSquaredError / Layers(UBound(Layers)).NeuronCount
  Else
    AverageSquaredError = 0
  End If
End Property
Public Property Get AnnealingEpoch() As Long
  AnnealingEpoch = mAnnealingEpoch
End Property
Public Property Let AnnealingEpoch(ByVal vData As Long)
  mAnnealingEpoch = vData
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
  TotalNeuronCount = UBound(Neurons) - LBound(Neurons) + 1
End Property

Public Property Get OutputLayer(Index As Long) As Double
  If mvarCreated = True Then
    OutputLayer = Neurons(Layers(UBound(Layers)).StartNeuron + Index - 1).Value
  Else
    Err.Raise vbObjectError + 2, "Perceptron", "You must initialise your net before you attempt to retrieve output from it"
    OutputLayer = 0
  End If
End Property

Public Sub SetInput(Data As Variant)
Dim i As Long

    If mStopping = True Then
      Exit Sub
    End If
    
    ' error checking on passed Data here:
    
    RaiseEvent InfoMessage(mTag, "Initialising Input Layer")
    
    For i = Layers(LBound(Layers)).StartNeuron To Layers(LBound(Layers)).EndNeuron
        Neurons(i).Value = CDbl(Data(i))
    Next i
End Sub

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
  #If ShowDebugTimes = 1 Then
  StartTime = GetTickCount
  #End If
  mRunning = True
  mStopping = False
  CalculateOutput
  #If ShowDebugTimes = 1 Then
  Debug.Print "Time to process output: " & GetTickCount - StartTime & " ms"
  #End If
  mRunning = False
End Sub

Friend Sub CalculateOutput()
Dim i As Long, j As Long, k As Long
Dim TotalToDo As Long
Dim Percentage As Long
Dim WorkDone As Long
Dim PrevNLayerPtr As Long
  
  ' Calculate work to be done
  TotalToDo = UBound(Neurons) - Layers(LBound(Layers) + 1).StartNeuron + 1
  Percentage = TotalToDo / 100
  WorkDone = 0
  
  ' Apologies for the complexity of this - whenever I've seen nn code before, it's
  ' always too complex to understand.  Hence the original OO stuff adapted from
  ' Ulli.  However, I've gone back to complexity to save on speed. :-(
  
  ' Go through all the layers except the first one
  For i = LBound(Layers) + 1 To UBound(Layers)
    RaiseEvent InfoMessage(mTag, "Calculating Output For Layer " & i)
    ' Go through all neurons in this layer
    For j = Layers(i).StartNeuron To Layers(i).EndNeuron
      ' for each neuron, sum the total of it's inputs from the previous layer.
      Neurons(j).Value = 0
      PrevNLayerPtr = Layers(i - 1).StartNeuron
      For k = Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) To Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) + Layers(i - 1).NeuronCount
        Neurons(j).Value = Neurons(j).Value + (Neurons(PrevNLayerPtr).Value * Dendrites(k).TransferWeight)
        PrevNLayerPtr = PrevNLayerPtr + 1
      Next k
      ' sigmoid squash
      Neurons(j).Value = Squish(Neurons(j).Value + Neurons(j).BaseValue)
      WorkDone = WorkDone + 1
      If WorkDone Mod Percentage = 0 Then
        RaiseEvent Progress(mTag, (WorkDone / TotalToDo) * 100)
      End If
      If mStopping = True Then
        Exit For
      End If
    Next j
    If mStopping = True Then
      Exit For
    End If
  Next i
End Sub

Private Sub AdjustWeights(Target As Variant)
  Dim i As Long, j As Long, k As Long
Dim Percentage As Long
Dim WorkDone As Long
Dim TotalToDo As Long
Dim SSE As Double
Dim PrevNLayerPtr As Long

    j = 0
    RaiseEvent InfoMessage(mTag, "Training the Net")
    
    'calculation of raw error in output layer
    SSE = 0
    For i = Layers(UBound(Layers)).StartNeuron To Layers(UBound(Layers)).EndNeuron
      Neurons(i).Error = CDbl(Target(j)) - Neurons(i).Value
      SSE = SSE + (Neurons(i).Error * Neurons(i).Error)
      j = j + 1
    Next i
    
    ' Simulated annealing - adjustment of learning coefficient to match error value.
    If mAnnealingEpoch > 0 Then
        mAnnealingSSE = mAnnealingSSE + SSE
        If mTrainingCycles Mod mAnnealingEpoch = 0 And mTrainingCycles > 0 Then
            If mLastAnnealingSSE > 0 Then
                If mAnnealingSSE < mLastAnnealingSSE Then
                    mLearningCoefficient = mLearningCoefficient * mLearningRateDecrease
                Else
                    mLearningCoefficient = mLearningCoefficient * mLearningRateIncrease
                End If
            End If
            mLastAnnealingSSE = mAnnealingSSE
            mAnnealingSSE = 0
        End If
    End If
    
    mSumSquaredError = SSE
    
    If mStopping = True Then
      Exit Sub
    End If
    
    ' Calculate work to be done
    TotalToDo = UBound(Neurons) - Layers(LBound(Layers) + 1).StartNeuron + 1
    Percentage = TotalToDo \ 100
    WorkDone = 0
    
    'hidden layers
    For i = UBound(Layers) To LBound(Layers) + 1 Step -1
        RaiseEvent InfoMessage(mTag, "Running Back Propogation on Layer " & i)
        For j = Layers(i).StartNeuron To Layers(i).EndNeuron
          ' Back propagate.
          With Neurons(j)
            .Error = .Error * .Value * (1# - .Value) ' proportional error
          End With
          ' Now, update all connected neurons error appropriately:
          PrevNLayerPtr = Layers(i - 1).StartNeuron
          For k = Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) To Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) + Layers(i - 1).NeuronCount
            Neurons(PrevNLayerPtr).Error = Neurons(PrevNLayerPtr).Error + (Neurons(j).Error * Dendrites(k).TransferWeight)
            PrevNLayerPtr = PrevNLayerPtr + 1
          Next k
        Next j
        WorkDone = WorkDone + 1
        If WorkDone Mod Percentage = 0 Then
          RaiseEvent Progress(mTag, (WorkDone / TotalToDo) * 100)
        End If
        If mStopping = True Then
          Exit For
        End If
    Next i
        
    If mStopping = True Then
      Exit Sub
    End If
    
    ' Calculate work to be done
    TotalToDo = UBound(Neurons) - Layers(LBound(Layers) + 1).StartNeuron + 1
    Percentage = TotalToDo \ 100
    WorkDone = 0
    
    ' Update weights
    For i = UBound(Layers) To LBound(Layers) + 1 Step -1
          RaiseEvent InfoMessage(mTag, "Updating Weights in Layer " & i)
          For j = Layers(i).StartNeuron To Layers(i).EndNeuron
              ' update base value
              With Neurons(j)
                .BaseValue = .BaseValue + mLearningCoefficient * .Error
              End With
              ' update dendrite weights
              PrevNLayerPtr = Layers(i - 1).StartNeuron
              For k = Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) To Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) + Layers(i - 1).NeuronCount
                With Dendrites(k)
                  .TransferWeight = .TransferWeight + mLearningCoefficient * Neurons(PrevNLayerPtr).Value * Neurons(j).Error
                End With
                PrevNLayerPtr = PrevNLayerPtr + 1
              Next k
              ' reset neuron error for next training cycle.
              Neurons(j).Error = 0
              WorkDone = WorkDone + 1
              If WorkDone Mod Percentage = 0 Then
                RaiseEvent Progress(mTag, (WorkDone / TotalToDo) * 100)
              End If
              If mStopping = True Then
                Exit For
              End If
          Next j
          If mStopping = True Then
            Exit For
          End If
    Next i
    
End Sub

Public Sub Train(Data As Variant, Target As Variant)
  If mvarCreated = True Then
    If mRunning = True Then
      Exit Sub
    End If
    ' need some error checking for data and target arrays to see if they're valid.
    
    mRunning = True
    mStopping = False
    #If ShowDebugTimes = 1 Then
    StartTime = GetTickCount
    #End If

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
  #If ShowDebugTimes = 1 Then
  Debug.Print "Time for one training cycle to take place: " & GetTickCount - StartTime & " ms"
  #End If
  mRunning = False
End Sub

' Test Ok.
Public Sub CreateNet(pTag As Variant, ParamArray Struc() As Variant)
Dim TotalToDo As Long
Dim TotalDone As Long
Dim PercentageHit As Long
Dim i As Long, TotalNeurons As Long, TotalSynapses As Long

    If mRunning = True Then
      Exit Sub
    End If
    mRunning = True
    mStopping = False
    #If ShowDebugTimes = 1 Then
    StartTime = GetTickCount
    #End If
    ' Initialise
    ' Need to destroy first
    DestroyNicely
    Set Tag = pTag
    mTrainingCycles = 0
    mSumSquaredError = 0
        
    ' Redim all arrays.
    ' Layers
    ReDim Layers(LBound(Struc) To UBound(Struc)) ' maintains array base of ParamArray
    ' Neurons
    ' Total neurons - add up all the values in the Struc array.
    ' Total synapses = Layer1*Layer2 + Layer2*Layer3 + Layer3*Layer4 etc.
    TotalNeurons = 0
    TotalSynapses = 0
    RaiseEvent InfoMessage(mTag, "Initialising Neural arrays")
    For i = LBound(Struc) To UBound(Struc)
      Layers(i).StartNeuron = TotalNeurons
      TotalNeurons = TotalNeurons + Struc(i)
      Layers(i).EndNeuron = TotalNeurons - 1
      Layers(i).NeuronCount = Struc(i)
      If i <> LBound(Struc) Then
        Layers(i).StartDendrite = TotalSynapses
        TotalSynapses = TotalSynapses + (Struc(i) * Struc(i - 1))
        Layers(i).EndDendrite = TotalSynapses - 1
        Layers(i).DendriteCount = Struc(i) * Struc(i - 1)
      Else
        Layers(i).StartDendrite = 0
        Layers(i).EndDendrite = 0
        Layers(i).DendriteCount = 0
      End If
    Next i
    ReDim Neurons(0 To TotalNeurons)
    ReDim Dendrites(0 To TotalSynapses)
        
    ' Calculate work to be done
'    RaiseEvent InfoMessage(mTag, "Randomising Weights and Base Values")
'    TotalToDo = -LBound(Neurons) + UBound(Neurons) + 1 - LBound(Dendrites) + UBound(Dendrites) + 1
'    PercentageHit = TotalToDo / 100
'    TotalDone = 0
    
    NguyenWidrowInitialise
    
'    ' Set BaseValue for each neuron
'    For i = LBound(Neurons) To UBound(Neurons)
'      Neurons(i).BaseValue = GetRand
'      TotalDone = TotalDone + 1
'      If TotalDone Mod PercentageHit = 0 Then
'        RaiseEvent Progress(mTag, (TotalDone / TotalToDo) * 100)
'      End If
'    Next i
'    ' Set TransferWeight for each dendrite.
'    For i = LBound(Dendrites) To UBound(Dendrites)
'      Dendrites(i).TransferWeight = GetRand
'      TotalDone = TotalDone + 1
'      If TotalDone Mod PercentageHit = 0 Then
'        RaiseEvent Progress(mTag, (TotalDone / TotalToDo) * 100)
'      End If
'    Next i
    
'    For i = LBound(Layers) To UBound(Layers)
'      Debug.Print "Layer " & i & ": Neurons: " & Layers(i).StartNeuron; " - " & Layers(i).EndNeuron & "(" & Layers(i).NeuronCount & ")" & "      Dendrites: " & Layers(i).StartDendrite & " - " & Layers(i).EndDendrite & "(" & Layers(i).DendriteCount & ")"
'    Next i
    
    If mStopping = True Then
    Else
      mvarCreated = True
    End If
    #If ShowDebugTimes = 1 Then
    Debug.Print "Time to create net: " & GetTickCount - StartTime & " ms"
    #End If
    mRunning = False
End Sub

' This is a bit pointless now - just here to preserve compatibility with original cNet module.
Public Sub DestroyNicely()
  If mRunning = True Then
    Exit Sub
  End If
  #If ShowDebugTimes = 1 Then
  StartTime = GetTickCount
  #End If
  RaiseEvent InfoMessage(mTag, "Destroying Dendrites")
  Erase Dendrites
  RaiseEvent InfoMessage(mTag, "Destroying Neurons")
  Erase Neurons
  RaiseEvent InfoMessage(mTag, "Destroying Layers")
  Erase Layers
  mvarCreated = False
  #If ShowDebugTimes = 1 Then
  Debug.Print "Time to destroy net: " & GetTickCount - StartTime & " ms"
  #End If
End Sub

Private Sub Class_Initialize()
  Randomize
  mvarCreated = False
  mLearningRateIncrease = 1#
  mLearningRateDecrease = 1#
  mAnnealingEpoch = 0
  mAnnealingSSE = 0
  mLastAnnealingSSE = 0
End Sub

Private Sub Class_Terminate()
  Erase Dendrites
  Erase Neurons
  Erase Layers
End Sub

Public Sub Jitter(Optional MaxVariance As Double = 0.05)
Dim i As Long, j As Long, k As Long
Dim Variance As Double
Dim TotalToDo As Long
Dim TotalDone As Long
Dim PercentageHit As Long

  If mRunning = True Or mvarCreated = False Then
    Exit Sub
  End If
  mRunning = True
  mStopping = False
  #If ShowDebugTimes = 1 Then
  StartTime = GetTickCount
  #End If

  
  TotalToDo = UBound(Dendrites) - LBound(Dendrites) + 1
  PercentageHit = TotalToDo / 100
  TotalDone = 0
  
  RaiseEvent InfoMessage(mTag, "Adding random noise to the net's weights")
  For i = LBound(Dendrites) To UBound(Dendrites)
    Variance = (MaxVariance * Rnd * 2) - MaxVariance
    Dendrites(i).TransferWeight = Dendrites(i).TransferWeight + Variance
    If Dendrites(i).TransferWeight < -1 Then
      Dendrites(i).TransferWeight = -1
    End If
    If Dendrites(i).TransferWeight > 1 Then
      Dendrites(i).TransferWeight = 1
    End If
    TotalDone = TotalDone + 1
    If TotalDone Mod PercentageHit = 0 Then
      RaiseEvent Progress(mTag, (TotalDone / TotalToDo) * 100)
    End If
  Next i
  
  #If ShowDebugTimes = 1 Then
  Debug.Print "Time to jitter net: " & GetTickCount - StartTime & " ms"
  #End If
  
  mRunning = False
End Sub

Public Sub KickZeros(Optional Amount As Double = 0.05)
Dim i As Long
Dim tmpAmount As Double
  For i = LBound(Dendrites) To UBound(Dendrites)
    If Dendrites(i).TransferWeight <= 0.001 And Dendrites(i).TransferWeight >= -0.001 Then
      tmpAmount = Amount * 2 * Rnd - Amount
      Dendrites(i).TransferWeight = tmpAmount
    End If
  Next i
  For i = LBound(Neurons) To UBound(Neurons)
    If Neurons(i).BaseValue <= 0.001 And Neurons(i).BaseValue >= -0.001 Then
      tmpAmount = Amount * Rnd
      Dendrites(i).TransferWeight = tmpAmount
    End If
  Next i
End Sub

Public Sub StopWorking()
  mStopping = True
End Sub

' ok
Public Function SaveNet(Filename As Variant, Optional SaveVersion As String = VERSION) As Boolean
Dim FileNumber As Long
Dim ll             As Long
Dim Percentage As Long
Dim WorkDone As Long
Dim TotalToDo As Long
Dim i As Long, j As Long, k As Long
Dim SaveHeader As String
Dim strVersion As String

  If mRunning = True Then
    Exit Function
  End If
  mRunning = True
  mStopping = False
  #If ShowDebugTimes = 1 Then
  StartTime = GetTickCount
  #End If
  
  If mvarCreated = False Then
      Err.Raise vbObjectError + 3, "Perceptron", "There is no Net to save."
  Else
      ' Version 1.3 save format
      On Error GoTo ErrHandler
      If VarType(Filename) = vbString Then
        FileNumber = FreeFile
        Open Filename For Binary Access Write As #FileNumber
      Else
        FileNumber = Filename
      End If
      On Error GoTo 0
      
      If SaveVersion = "1.2" Then
        SaveNetOnePointTwo FileNumber
      Else
        
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
        Put #FileNumber, , CLng(UBound(Layers) - LBound(Layers) + 1) ' Write Number of layers
        Put #FileNumber, , Me.TotalNeuronCount ' Write Total number of neurons
        Put #FileNumber, , CLng(UBound(Dendrites) - LBound(Dendrites) + 1) ' Write total number of dendrites
        Put #FileNumber, , Me.TrainingCycles ' Write Total amount trained
        ' Write learning coefficient information
        Put #FileNumber, , mAnnealingEpoch
        Put #FileNumber, , mAnnealingSSE
        Put #FileNumber, , mLastAnnealingSSE
        Put #FileNumber, , mLearningRateDecrease
        Put #FileNumber, , mLearningRateIncrease
        Put #FileNumber, , mLearningCoefficient
        
        For i = LBound(Layers) To UBound(Layers)
            RaiseEvent InfoMessage(mTag, "Saving Layer " & i)
            Put #FileNumber, , Layers(i).NeuronCount
            For j = Layers(i).StartNeuron To Layers(i).EndNeuron
                Put #FileNumber, , Neurons(j).BaseValue
                If i = LBound(Layers) Then
                  Put #FileNumber, , 0&
                Else
                  Put #FileNumber, , CLng(Layers(i - 1).NeuronCount)
                  For k = Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) To Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) + Layers(i - 1).NeuronCount - 1
                    Put #FileNumber, , Dendrites(k).TransferWeight
                  Next k
                End If
                WorkDone = WorkDone + 1
                If WorkDone Mod Percentage = 0 Then
                  RaiseEvent Progress(mTag, (WorkDone / TotalToDo) * 100)
                End If
                If mStopping = True Then
                  Exit For
                End If
            Next j
            If mStopping = True Then
              Exit For
            End If
        Next i
        Close #FileNumber
      End If
    End If
    
    mRunning = False
    If mStopping = True Then
      SaveNet = False ' aborted save.
    Else
      SaveNet = True
    End If
  #If ShowDebugTimes = 1 Then
  Debug.Print "Time to save net: " & GetTickCount - StartTime & " ms"
  #End If

    Exit Function
ErrHandler:
  MsgBox "Failed to save net: " & Err.Description
  Err.Clear
  SaveNet = False
  mRunning = False
End Function

Private Function SaveNetOnePointTwo(FileNumber As Long) As Boolean
Dim ll             As Long
Dim Percentage As Long
Dim WorkDone As Long
Dim TotalToDo As Long
Dim i As Long, j As Long, k As Long
Dim SaveHeader As String
Dim strVersion As String

    ' Version 1.2 save format
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
    Put #FileNumber, , CLng(UBound(Layers) - LBound(Layers) + 1) ' Write Number of layers
    Put #FileNumber, , Me.TotalNeuronCount ' Write Total number of neurons
    Put #FileNumber, , Me.TrainingCycles ' Write Total amount trained
    For i = LBound(Layers) To UBound(Layers)
        RaiseEvent InfoMessage(mTag, "Saving Layer " & i)
        Put #FileNumber, , Layers(i).NeuronCount
        For j = Layers(i).StartNeuron To Layers(i).EndNeuron
            Put #FileNumber, , Neurons(j).BaseValue
            If i = LBound(Layers) Then
              Put #FileNumber, , 0&
            Else
              Put #FileNumber, , CLng(Layers(i - 1).NeuronCount)
              For k = Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) To Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) + Layers(i - 1).NeuronCount - 1
                Put #FileNumber, , Dendrites(k).TransferWeight
              Next k
            End If
            WorkDone = WorkDone + 1
            If WorkDone Mod Percentage = 0 Then
              RaiseEvent Progress(mTag, (WorkDone / TotalToDo) * 100)
            End If
            If mStopping = True Then
              Exit For
            End If
        Next j
        If mStopping = True Then
          Exit For
        End If
    Next i
    Close #FileNumber
End Function

' Ok.
Private Function LoadNetOnePointTwo(FileNumber As Long) As Boolean
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
Dim CurrentNeuron As Long
Dim CurrentDendrite As Long

    ' Unfortunately in 1.2 save format the total number of dendrites are not stored.  This is
    ' unfortunate as we'll have to continually redim the Dendrites() array.
    ' This code will be removed once the project has been moved on a little (perhaps after part 3).
    
    Get #FileNumber, , ll ' Number of layers
    Get #FileNumber, , WorkToDo ' Total neuron count
    Get #FileNumber, , mTrainingCycles ' Total training cycles
    
    Percentage = WorkToDo / 100
    WorkDone = 0
    
    ReDim Layers(0 To ll - 1)
    ReDim Neurons(0 To WorkToDo)
    ReDim Dendrites(0 To 0)
    CurrentNeuron = 0
    CurrentDendrite = 0

    For l = 1 To ll
        RaiseEvent InfoMessage(mTag, "Loading Layer " & l)
        ' Get total neurons in this layer and redim dendrites
        Get #FileNumber, , nn
        Layers(l - 1).NeuronCount = nn
        Layers(l - 1).StartNeuron = CurrentNeuron
        Layers(l - 1).StartDendrite = CurrentDendrite
        For n = 1 To nn
            Get #FileNumber, , x
            Neurons(CurrentNeuron).BaseValue = x
            ' Get number of dendrites connected to this neuron.
            Get #FileNumber, , dd
            If dd > 0 Then
                ReDim Preserve Dendrites(0 To UBound(Dendrites) + dd)
                For d = 1 To dd
                    Get #FileNumber, , x
                    Dendrites(CurrentDendrite).TransferWeight = x
                    CurrentDendrite = CurrentDendrite + 1
                Next d
            End If
            CurrentNeuron = CurrentNeuron + 1
            WorkDone = WorkDone + 1
            If WorkDone Mod Percentage = 0 Then
              RaiseEvent Progress(mTag, (WorkDone / WorkToDo) * 100)
            End If
            If mStopping = True Then
              Exit For
            End If
        Next n
        Layers(l - 1).EndNeuron = CurrentNeuron - 1
        If dd > 0 Then
          Layers(l - 1).EndDendrite = CurrentDendrite - 1
        Else
          Layers(l - 1).EndDendrite = Layers(l - 1).StartDendrite
        End If
        Layers(l - 1).DendriteCount = Layers(l - 1).EndDendrite - Layers(l - 1).StartDendrite + 1
        If mStopping = True Then
          Exit For
        End If
    Next l

    Close #FileNumber
'    For i = LBound(Layers) To UBound(Layers)
'      Debug.Print "Layer " & i & ": Neurons: " & Layers(i).StartNeuron; " - " & Layers(i).EndNeuron & "(" & Layers(i).NeuronCount & ")" & "      Dendrites: " & Layers(i).StartDendrite & " - " & Layers(i).EndDendrite & "(" & Layers(i).DendriteCount & ")"
'    Next i
    
    mSumSquaredError = 0
    ' Now to set up all variables that weren't included in 1.2
    mAnnealingEpoch = 0
    mAnnealingSSE = 0#
    mLastAnnealingSSE = 0#
    mLearningRateDecrease = 1#
    mLearningRateIncrease = 1#
End Function

' Ok.
Public Function LoadNet(Filename As Variant, Optional pTag As Variant = Nothing, Optional pVersion As String = VERSION) As Boolean
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
Dim TotalDendrites As Long
Dim CurrentNeuron As Long
Dim CurrentDendrite As Long

    If mRunning = True Then
      Exit Function
    End If
    mRunning = True
    mStopping = False
  #If ShowDebugTimes = 1 Then
  StartTime = GetTickCount
  #End If

 
    ' Destroy net before reloading
    DestroyNicely
    
    Set mTag = pTag
    
    RaiseEvent InfoMessage(mTag, "Loading Neural Net")
    
    ' Open the file.
    On Error GoTo ErrHandler
    If VarType(Filename) = vbString Then
      FileNumber = FreeFile
      Open Filename For Binary Access Read As #FileNumber
    Else
      FileNumber = Filename
    End If

    SaveHeader = String(Len(OCXNAME), " ")
    strVersion = String(Len(VERSION), " ")
    On Error GoTo 0
    Get #FileNumber, , SaveHeader
    Get #FileNumber, , strVersion
    If SaveHeader = OCXNAME And strVersion = "1.2" Then
      LoadNetOnePointTwo FileNumber
    Else
      If SaveHeader = OCXNAME And strVersion = VERSION Then
        ' Valid 1.2 file
      Else
        ' This is not a valid file.
        Err.Clear
        Close #FileNumber
        GoTo ErrHandler  ' apologies!!!!  This will be changed in future.
      End If
      ' Unfortunately in 1.2 save format the total number of dendrites are not stored.  This is
      ' unfortunate as we'll have to continually redim the Dendrites() array.
      ' This code will be removed once the project has been moved on a little (perhaps after part 3).
    
      Get #FileNumber, , ll ' Number of layers
      Get #FileNumber, , WorkToDo ' Total neuron count
      Get #FileNumber, , TotalDendrites ' total dendrite count
      Get #FileNumber, , mTrainingCycles ' Total training cycles
      ' Read learning coefficient information
      Get #FileNumber, , mAnnealingEpoch
      Get #FileNumber, , mAnnealingSSE
      Get #FileNumber, , mLastAnnealingSSE
      Get #FileNumber, , mLearningRateDecrease
      Get #FileNumber, , mLearningRateIncrease
      Get #FileNumber, , mLearningCoefficient
   
      Percentage = WorkToDo / 100
      WorkDone = 0
      
      ReDim Layers(0 To ll - 1)
      ReDim Neurons(0 To WorkToDo)
      ReDim Dendrites(0 To TotalDendrites - 1)
      CurrentNeuron = 0
      CurrentDendrite = 0

      For l = 1 To ll
          RaiseEvent InfoMessage(mTag, "Loading Layer " & l)
          ' Get total neurons in this layer and redim dendrites
          Get #FileNumber, , nn
          Layers(l - 1).NeuronCount = nn
          Layers(l - 1).StartNeuron = CurrentNeuron
          Layers(l - 1).StartDendrite = CurrentDendrite
          For n = 1 To nn
              Get #FileNumber, , x
              Neurons(CurrentNeuron).BaseValue = x
              ' Get number of dendrites connected to this neuron.
              Get #FileNumber, , dd
              If dd > 0 Then
                  For d = 1 To dd
                      Get #FileNumber, , x
                      Dendrites(CurrentDendrite).TransferWeight = x
                      CurrentDendrite = CurrentDendrite + 1
                  Next d
              End If
              CurrentNeuron = CurrentNeuron + 1
              WorkDone = WorkDone + 1
              If WorkDone Mod Percentage = 0 Then
                RaiseEvent Progress(mTag, (WorkDone / WorkToDo) * 100)
              End If
              If mStopping = True Then
                Exit For
              End If
          Next n
          Layers(l - 1).EndNeuron = CurrentNeuron - 1
          If dd > 0 Then
            Layers(l - 1).EndDendrite = CurrentDendrite - 1
          Else
            Layers(l - 1).EndDendrite = Layers(l - 1).StartDendrite
          End If
          Layers(l - 1).DendriteCount = Layers(l - 1).EndDendrite - Layers(l - 1).StartDendrite + 1
          If mStopping = True Then
            Exit For
          End If
      Next l
  
      Close #FileNumber
    End If
'    For i = LBound(Layers) To UBound(Layers)
'      Debug.Print "Layer " & i & ": Neurons: " & Layers(i).StartNeuron; " - " & Layers(i).EndNeuron & "(" & Layers(i).NeuronCount & ")" & "      Dendrites: " & Layers(i).StartDendrite & " - " & Layers(i).EndDendrite & "(" & Layers(i).DendriteCount & ")"
'    Next i
    
    mSumSquaredError = 0
    mRunning = False
    If mStopping = True Then
      LoadNet = False ' aborted load
    Else
      mvarCreated = True
      LoadNet = True
    End If
  #If ShowDebugTimes = 1 Then
  Debug.Print "Time to load net: " & GetTickCount - StartTime & " ms"
  #End If

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


' Miscellaneous functions
Private Function GetRand() As Double
    Do
        GetRand = 1 - Rnd - Rnd
    Loop Until Abs(GetRand) >= 0.1
End Function

Private Function Squish(InVal As Double) As Double
Dim tmpDouble As Double
  'Uses Sigmoid Function to squash InVal into interval 0..1
 ' Now changed so I can set compiler optimisation on floating point vars if I need to.
 
  tmpDouble = 1# + Exp(-InVal)
  If tmpDouble <> 0 Then
    Squish = 1# / tmpDouble
  Else
    Squish = 0
  End If
End Function

' Random initialiser
' This was adapted from Jason Tiscione's Java code for a backprop neural net.
Private Sub NguyenWidrowInitialise()
Dim i As Long, j As Long, k As Long
Dim PrevNLayerPtr As Long
Dim magw As Double
Dim acc As Double, acc2 As Double
Dim TotalToDo As Long, Percentage As Long, WorkDone As Long

  RaiseEvent InfoMessage(mTag, "Randomising Weights and Base Values")
  TotalToDo = UBound(Neurons) - Layers(LBound(Layers) + 1).StartNeuron + 1
  Percentage = TotalToDo / 100
  WorkDone = 0
  
  For i = LBound(Layers) + 1 To UBound(Layers)
    magw = 2.8 * (Layers(i).NeuronCount ^ (1 / Layers(i).DendriteCount))
    For j = Layers(i).StartNeuron To Layers(i).EndNeuron
      Neurons(j).BaseValue = (2 * Rnd) - 1
      acc = 0
      PrevNLayerPtr = Layers(i - 1).StartNeuron
      For k = Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) To Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) + Layers(i - 1).NeuronCount
        Dendrites(k).TransferWeight = (2 * Rnd) - 1
        acc = acc + (Dendrites(k).TransferWeight * Dendrites(k).TransferWeight)
        PrevNLayerPtr = PrevNLayerPtr + 1
      Next k
      ' Normalise the "neuron j" row vector created in previous loop (?)
      acc = Sqr(1 / acc)
      acc2 = 0
      For k = Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) To Layers(i).StartDendrite + ((j - Layers(i).StartNeuron) * Layers(i - 1).NeuronCount) + Layers(i - 1).NeuronCount
        Dendrites(k).TransferWeight = Dendrites(k).TransferWeight * (2 * acc)
        acc2 = acc2 + Dendrites(k).TransferWeight
      Next k
      ' Normalise "neuron j" bias vector with respect to weights (?)
      Neurons(j).BaseValue = Neurons(j).BaseValue - (acc2 / 2)
      If mStopping = True Then
        Exit For
      End If
      WorkDone = WorkDone + 1
      If WorkDone Mod Percentage = 0 Then
        RaiseEvent Progress(mTag, (WorkDone / TotalToDo) * 100)
      End If
    Next j
    If mStopping = True Then
      Exit For
    End If
  Next i
  
End Sub
