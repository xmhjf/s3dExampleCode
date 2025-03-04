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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         RUK
'   Creation Date:  Monday, Feb 04, 2008
'   Description:
'       This is Open/Blind spcer symbol. This is prepared based Appendex E-94 in Piping Design Document L57.
'       CR-134984  Provide symbol for open spacer and blind spacer set
'               Blind spacer and open spacer set, specified by spacer length(758)
'               Blind spacer and open spacer set, specified by handle length from flow centerline(759)
'               Blind spacer and open spacer set, specified by handle length from spacer outside diameter(760)
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.Feb.2008     RUK     CR-134984  Provide symbol for open spacer and blind spacer set
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim parInstallationState As Double
    Dim parFacetoFace As Double
    Dim parSpacerLength As Double
    Dim parSpacerOutsideDia As Double
    Dim parSpacerInsideDia As Double
    Dim parHandleLength As Double
    Dim parHandleWidth As Double
    Dim parHandleThickness As Double
    Dim parHandleRadius As Double
    Dim parHoleDiameter As Double
    Dim parInsulationThickness As Double
    
    Set oPartFclt = arrayOfInputs(1)
    parInstallationState = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parSpacerLength = arrayOfInputs(4)
    parSpacerOutsideDia = arrayOfInputs(5)
    parSpacerInsideDia = arrayOfInputs(6)
    parHandleLength = arrayOfInputs(7)
    parHandleWidth = arrayOfInputs(8)
    parHandleThickness = arrayOfInputs(9)
    parHandleRadius = arrayOfInputs(10)
    parHoleDiameter = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    
    'Origin is at the center of inlet and outlet nozzles
    Dim iOutput As Long
    iOutput = 0
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    
    'Retrieve the nozzle parameters
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        cptOffset, depth
    Dim dInsDia As Double
    dInsDia = parSpacerOutsideDia + 2 * parInsulationThickness
    
    'Create the Insulated Body
    Dim objInsBody As Object
    oStPoint.Set parFacetoFace / 2, 0, 0
    oEnPoint.Set -parFacetoFace / 2, 0, 0
    Set objInsBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBody
    Set objInsBody = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD
End Sub

