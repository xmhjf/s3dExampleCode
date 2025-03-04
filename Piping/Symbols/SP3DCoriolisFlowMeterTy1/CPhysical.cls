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
'   Author:          MS
'   Creation Date:  Monday, Sep 16 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ICM1)
'The symbol consists of seven outputs, five physical and two insulation outputs.
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
    Dim ObjFlowMeterBody As Object
    Dim ObjVerCyl As Object
    Dim ObjFlowMeterTopBody As Object
    Dim ObjFlowMeterBodyIns As Object
    Dim ObjFlowMeterTopBodyIns As Object
    
    Dim parFacetoFace As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)   'D2
    parFlowDiameter = arrayOfInputs(3)          'D3
    parInstrumentHeight = arrayOfInputs(4)       'D4
    parInstrumentDiameter = arrayOfInputs(5)     'D5
    parInstrumentWidth = arrayOfInputs(6)       'D6
    parInstrumentWidth1 = arrayOfInputs(7)      'D7
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0
'Origin is considered to be mid of parFacetoFace and along the flow axis
 ' Insert your code for output 1(Body of Flowmeter)
    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    Set ObjFlowMeterBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFlowDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowMeterBody
    Set ObjFlowMeterBody = Nothing
    
 ' Insert your code for output 2(Vertical Cylinder)
    Dim meetHeight As Double
    Dim verCylDia As Double
'Diameter of Vertical cylinder is assumed to be 0.2 times the parInstrumentDiameter
    verCylDia = 0.2 * parInstrumentDiameter
    meetHeight = Sqr((parInstrumentDiameter / 2 * parInstrumentDiameter / 2) - (verCylDia / 2 * verCylDia / 2))
    stPoint.Set 0, 0, 0
' The Dimension D4 is taken from the flowaxis to the Instrument Top
    enPoint.Set 0, parInstrumentHeight - parInstrumentDiameter / 2 - meetHeight, 0
    Set ObjVerCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, verCylDia, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCyl
    Set ObjVerCyl = Nothing
    
 ' Insert your code for output 3(Top Body of Flowmeter)
    stPoint.Set 0, _
                        parInstrumentHeight - parInstrumentDiameter / 2, _
                        parInstrumentWidth
                        
    enPoint.Set 0, _
                        parInstrumentHeight - parInstrumentDiameter / 2, _
                        -parInstrumentWidth1
                        
    Set ObjFlowMeterTopBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowMeterTopBody
    Set ObjFlowMeterTopBody = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
