VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNeuron"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Dendrite"
Attribute VB_Ext_KEY = "Member0" ,"Dendrite"
Option Explicit
Option Base 0
Option Compare Text

Private Dendrites           As Collection
Private oDendrite           As cDendrite

Public Value                As Double
Public Error                As Double
Public LearningCoefficient  As Double
Public BaseValue            As Double

Public Sub Add(Dendrite As cDendrite)
  
    Dendrites.Add Dendrite

End Sub

Public Property Get Dendrite(Index As Long) As cDendrite

    Set Dendrite = Dendrites(Index)

End Property

Public Property Get DendriteCount() As Long

    DendriteCount = Dendrites.Count

End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Set NewEnum = Dendrites.[_NewEnum]

End Property

Private Sub Class_Initialize()

    Set Dendrites = New Collection
    
End Sub

Private Sub Class_Terminate()

    Set Dendrites = Nothing

End Sub

Public Sub Activate()
    
    If Dendrites.Count Then  'not Top Layer
        Value = 0
        For Each oDendrite In Dendrites
            With oDendrite
                Value = Value + .ConnectedTo.Value * .TransferWeight
            End With
        Next oDendrite
        Value = TransferFunction(Value + BaseValue)
    End If

End Sub

' This is a gentle destroy.  Called from cNet but not available to outside the DLL.
Friend Sub Destroy()
Dim i As Long
  For Each oDendrite In Dendrites
    ' Set to nothing
    Set oDendrite.ConnectedTo = Nothing
  Next oDendrite
  For i = 1 To Dendrites.Count
    Dendrites.Remove 1
  Next i
End Sub

Public Sub BackPropagate()

    Error = Error * Value * (1# - Value) 'proportional error
    For Each oDendrite In Dendrites
        With oDendrite.ConnectedTo 'neuron in previous layer
            .Error = .Error + Error * oDendrite.TransferWeight 'raw error
        End With
    Next oDendrite

End Sub

Public Sub UpdateWeights()

    If Dendrites.Count Then 'not top layer
        BaseValue = BaseValue + LearningCoefficient * Error
        For Each oDendrite In Dendrites
            With oDendrite
                .TransferWeight = .TransferWeight + LearningCoefficient * .ConnectedTo.Value * Error
            End With
        Next oDendrite
        Error = 0 'reset for next tarining cycle
    End If

End Sub

Private Function TransferFunction(InVal As Double) As Double

  'Uses Sigmoid Function to squash InVal into interval 0..1

    On Error Resume Next
      TransferFunction = 1# / (1# + Exp(-InVal))
      If Err Then
          TransferFunction = 0#
      End If
    On Error GoTo 0
        
End Function

