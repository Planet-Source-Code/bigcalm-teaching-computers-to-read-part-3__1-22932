VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"Neuron"
Attribute VB_Ext_KEY = "Member0" ,"Neuron"
Option Explicit
Option Base 0
Option Compare Text

' A layer is a collection of neurons.

Public PreviousLayer As cLayer

Private Neurons As Collection

Public Sub Add(Neuron As cNeuron)
  
    Neurons.Add Neuron

End Sub

Public Property Get Neuron(Index As Long) As cNeuron

    Set Neuron = Neurons(Index)

End Property

Public Property Get NeuronCount() As Long

    NeuronCount = Neurons.Count

End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Set NewEnum = Neurons.[_NewEnum]

End Property

Private Sub Class_Initialize()

    Set Neurons = New Collection

End Sub

Private Sub Class_Terminate()

    Set Neurons = Nothing

End Sub
