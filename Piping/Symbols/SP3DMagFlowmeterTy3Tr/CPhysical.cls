VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          svsmylav
'   Creation Date:  Tuesday, Sep 17 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMFM3,
'   PDMS Category Name=IFMY). The symbol consists of seven outputs, five physical and
'   two insulation outputs. Symbol consists of a base (cylinder), a connector(cylinder)
'   and a Transmitter(box).
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

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
    Dim ObjFlowmtrBase As Object
    Dim ObjConnector As Object
    Dim ObjTransmitter As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentOffset As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentOffset1 As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFace2toCenter = arrayOfInputs(3) 'D3
    parFlowDiameter = arrayOfInputs(4)  'D4
    parInstrumentOffset = arrayOfInputs(5)  'D5
    parInstrumentLength = arrayOfInputs(6)  'D6
    parInstrumentOffset1 = arrayOfInputs(7) 'D7
    parInstrumentHeight = arrayOfInputs(8)  'D8
    parInstrumentWidth = arrayOfInputs(9)   'D9
    parInstrumentWidth1 = arrayOfInputs(10) 'D10
    parInsulationThickness = arrayOfInputs(11)
    
    iOutput = 0
    
'   Center position is assumed at the intersection of the flow axis and
'   verticle axis through the point at distance D2 from port1.
 
' Insert your code for output 3(FlowmtrBase)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace2toCenter, 0, 0
    Set ObjFlowmtrBase = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFlowDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowmtrBase
    Set ObjFlowmtrBase = Nothing
    
' Insert your code for output 4(Connector)
'   Assumption: Connector diameter is 0.1 times that of flow diameter
    Dim ConnectorDiameter As Double
    Dim MeetRadius As Double
    
    ConnectorDiameter = 0.1 * parFlowDiameter
    MeetRadius = Sqr((0.5 * parFlowDiameter) ^ 2 - (0.5 * ConnectorDiameter) ^ 2)
    stPoint.Set parInstrumentOffset, MeetRadius, 0
    enPoint.Set parInstrumentOffset, parInstrumentOffset1, 0
    Set ObjConnector = PlaceCylinder(m_OutputColl, stPoint, enPoint, ConnectorDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector
    Set ObjConnector = Nothing
    
' Insert your code for output 5(Transmitter)
    stPoint.Set parInstrumentOffset - parInstrumentLength / 2, _
                parInstrumentOffset1, _
                -parInstrumentWidth
    enPoint.Set parInstrumentOffset + parInstrumentLength / 2, _
                parInstrumentOffset1 + parInstrumentHeight, _
                parInstrumentWidth1
    Set ObjTransmitter = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTransmitter
    Set ObjTransmitter = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
