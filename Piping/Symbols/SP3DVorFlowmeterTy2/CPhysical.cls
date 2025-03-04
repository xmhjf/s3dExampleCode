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
'   Creation Date:  Monday, Sep 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IVM2,
'   PDMS Category Name=IFND). Physical aspect consist of one horizontal cylinder, two vertical cylinders
'   and two nozzles one facing -nagative an the other in positive direction of X-axis.
'   Insulation aspect consist of two cylinders one along X-axis, and the other along Y-axis.
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
    Dim ObjInsHorCylinder As Object
    Dim ObjInsVerCylinder As Object
    Dim ObjHorCylinder As Object
    Dim ObjConnector As Object
    Dim ObjVerCylinder As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim Diameter As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)      'D2
    parFace2toCenter = arrayOfInputs(3)      'D3
    parFlowDiameter = arrayOfInputs(4)       'D4
    parInstrumentHeight = arrayOfInputs(5)   'D5
    parInstrumentHeight1 = arrayOfInputs(6)  'D6
    parInstrumentDiameter = arrayOfInputs(7) 'D7
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0

'   Origin of the symbol is assumed at the intersection of vertical cylinder axis
'   horizontal axis.

' Insert your code for output 3(Horizontal Cylinder)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set parFace2toCenter, 0, 0
    Set ObjHorCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFlowDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHorCylinder
    Set ObjHorCylinder = Nothing
    
' Insert your code for output 4(Connector)
'   Assumption: Connector diamter is 0.2 times D7
    Dim radPlaceDist As Double
    
    radPlaceDist = Sqr((parFlowDiameter / 2) ^ 2 - (parInstrumentDiameter * 0.2 / 2) ^ 2)
    stPoint.Set 0, radPlaceDist, 0
    enPoint.Set 0, parInstrumentHeight - parInstrumentHeight1, 0
    Set ObjConnector = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                    parInstrumentDiameter * 0.2, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConnector
    Set ObjConnector = Nothing
    
' Insert your code for output 5(Vertical Cylinder)
    stPoint.Set 0, parInstrumentHeight - parInstrumentHeight1, 0
    enPoint.Set 0, parInstrumentHeight, 0
    Set ObjVerCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                                parInstrumentDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylinder
    Set ObjVerCylinder = Nothing
    
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
