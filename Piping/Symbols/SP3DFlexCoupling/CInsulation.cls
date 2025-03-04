VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author       :  PK
'   Creation Date:  Friday, Oct 05 2007
'   Description  :  Straub Flex Type Coupling
'   Change History:
'   dd.mmm.yyyy          who                         change description
'   -----------          -----                       ------------------
'   04.Oct.2007           PK                        CR-126631:Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt                As PartFacelets.IJDPart
    Dim parFacetoFace            As Double
    Dim parDiameter              As Double
    Dim parPorttoPortOffset      As Double
    Dim parBoltEndDimension      As Double
    Dim parInsulationThickness   As Double
    Dim ioutput                  As Double
    
    Dim oInsulationBody     As Object
    Dim oStPoint            As AutoMath.DPosition
    Dim oEnpoint            As AutoMath.DPosition
    
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parDiameter = arrayOfInputs(3)
    parPorttoPortOffset = arrayOfInputs(4)
    parBoltEndDimension = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnpoint = New AutoMath.DPosition
    
    ioutput = 0
    oStPoint.Set -parFacetoFace / 2, 0, 0
    oEnpoint.Set parFacetoFace / 2, 0, 0
    
    Set oInsulationBody = PlaceCylinder(m_OutputColl, oStPoint, oEnpoint, parDiameter + 2 * parInsulationThickness, True)
    
    ioutput = ioutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(ioutput), oInsulationBody
    Set oInsulationBody = Nothing
    Set oStPoint = Nothing
    Set oEnpoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
