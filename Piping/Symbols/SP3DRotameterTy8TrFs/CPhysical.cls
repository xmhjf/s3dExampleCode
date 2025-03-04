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
'   Creation Date:  Friday, Sep 20 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR8,
'   PDMS Category Name=IFNN). Physical aspect consist of two nozzles,one vertical cylinder,
'   one horizontal cylinder on right side, one box on left side of the vertical cylinder
'   respectively. Two nozzles are placed one facing negative Y and the other facing
'   positive Y direction.
'   Insulation aspect consist of two cylinders one along X-axis and the other along Z-axis and
'   a box on left side of the vertical cylinder.
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
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjRotaMtrVerBody As Object
    Dim ObjRotaMtrSwitch As Object
    Dim ObjRotaMtrTrBox As Object
    
    Dim parFacetoFace As Double
    Dim parFlowDiameter As Double
    Dim parFace1toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInsulationThickness As Double
    Dim parInstrumentHeight2 As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)           'D2
    parFlowDiameter = arrayOfInputs(3)         'D3
    parFace1toCenter = arrayOfInputs(4)        'D4
    parInstrumentHeight = arrayOfInputs(5)     'D5
    parInstrumentHeight1 = arrayOfInputs(6)    'D6
    parInstrumentWidth = arrayOfInputs(7)      'D7
    parInstrumentWidth1 = arrayOfInputs(8)     'D8
    parInstrumentLength = arrayOfInputs(9)     'D9
    parInstrumentLength1 = arrayOfInputs(10)   'D10
    parInstrumentDiameter = arrayOfInputs(11)  'D11
    parInstrumentHeight2 = arrayOfInputs(12)   'D12
    parInsulationThickness = arrayOfInputs(13)
    
    iOutput = 0
    
'   Origin of the symbol is assumed at the intersection of vertical cylinder axis
'   horizontal axis at D4 offset from  face1

' Insert your code for output 4(Rotameter Vertical Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick1, flangeDiam, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam, sptOffset2, depth2

    stPoint.Set 0, -parFace1toCenter + flangeThick1, 0
    enPoint.Set 0, parFacetoFace - parFace1toCenter - flangeThick2, 0
    Set ObjRotaMtrVerBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFlowDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotaMtrVerBody
    Set ObjRotaMtrVerBody = Nothing
    
' Insert your code for output 5(RotaMeter Switch)
    stPoint.Set parFlowDiameter / 2, parInstrumentHeight2 - parFace1toCenter, 0
    enPoint.Set parInstrumentLength1, parInstrumentHeight2 - parFace1toCenter, 0
    Set ObjRotaMtrSwitch = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotaMtrSwitch
    Set ObjRotaMtrSwitch = Nothing
    
' Insert your code for output 6(RotaMeter Transmitter Box)
    stPoint.Set -parInstrumentLength, -parInstrumentHeight, parInstrumentWidth1
    enPoint.Set -parFlowDiameter / 2, parInstrumentHeight1 - parInstrumentHeight, _
                 -parInstrumentWidth
    Set ObjRotaMtrTrBox = PlaceBox(m_OutputColl, stPoint, enPoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotaMtrTrBox
    Set ObjRotaMtrTrBox = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set 0, -parFace1toCenter - sptOffset1 + depth1, 0
    oDir.Set 0, -1, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set 0, parFacetoFace - parFace1toCenter + sptOffset2 - depth2, 0
    oDir.Set 0, 1, 0
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
