VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDeckHole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:44 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CDeckHole.cls
'
'   Description: The symbol consists of a cylinder and port
'
'   Change History:
'   dd.mmm.yyyy         who          change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CWDeckHole:" 'Used for error messages

Private Sub Class_Initialize()
''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim iOutput     As Double
    iOutput = 0
    
    Dim parScupperHeight As Double
    Dim parOuterDiameter As Double
    Dim parCylinderHeight As Double
    Dim parInsulationThickness As Double
    Dim parInnerDiameter As Double
    Dim dCylinderHeight As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parScupperHeight = arrayOfInputs(2)
    parCylinderHeight = arrayOfInputs(3)
    parInnerDiameter = arrayOfInputs(4)
    parOuterDiameter = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEndPoint As AutoMath.DPosition
   
    
    If (lPartDataBasis <> 501) Then
        Set oStPoint = New AutoMath.DPosition
        oStPoint.Set 0.7 * parScupperHeight, 0, 0
        Set oEndPoint = New AutoMath.DPosition
        oEndPoint.Set parScupperHeight, 0, 0

    ElseIf (lPartDataBasis = 501) Then
        Set oStPoint = New AutoMath.DPosition
        oStPoint.Set (parScupperHeight - parCylinderHeight), 0, 0
        Set oEndPoint = New AutoMath.DPosition
        oEndPoint.Set parScupperHeight, 0, 0

    End If
    parOuterDiameter = (parOuterDiameter * 1.1)
    Dim ObjHoleBody As Object
    Set ObjHoleBody = PlaceCylinder(m_OutputColl, oStPoint, oEndPoint, parOuterDiameter, True)
    
    '   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoleBody
    Set ObjHoleBody = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
            Err.HelpFile, Err.HelpContext
End Sub


 
