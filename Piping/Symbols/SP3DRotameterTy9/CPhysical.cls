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
'   Author:         svsmylav
'   Creation Date:  Friday, Sep 20 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR9,
'   PDMS Category Name=IFNO). Physical aspect consist of one vertical cylinder, one horizontal cylinder
'   on right side of the vertical cylinder and two nozzles both facing -negative direction direction of X-axis.
'   Insulation aspect consist of two cylinders one along X-axis and the other along Y-axis.
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
    Dim ObjRotaMtrVerCylinder As Object
    Dim ObjRotaMtrHorCylinder As Object
    
    Dim parInstrumentHeight As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight1 As Double
    Dim parFacetoFace As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentDiameter As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)    'D2
    parFlowDiameter = arrayOfInputs(3)        'D3
    parInstrumentHeight1 = arrayOfInputs(4)   'D4
    parFacetoFace = arrayOfInputs(5)          'D5
    parInstrumentLength = arrayOfInputs(6)    'D6
    parInstrumentDiameter = arrayOfInputs(7)  'D7
    parInsulationThickness = arrayOfInputs(8)

    iOutput = 0
    
'   Origin of the symbol is assumed at the intersection of vertical cylinder axis
'   horizontal axis at D4 offset from  face1

' Insert your code for output 3(Rotameter Vertical Cylinder)
    stPoint.Set 0, -parInstrumentHeight1, 0
    enPoint.Set 0, parInstrumentHeight - parInstrumentHeight1, 0
    Set ObjRotaMtrVerCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parFlowDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotaMtrVerCylinder
    Set ObjRotaMtrVerCylinder = Nothing
    
' Insert your code for output 4(RotaMeter Horizontal Cylinder)
    Dim radialPlaceDist As Double
    radialPlaceDist = Sqr(Abs((parFlowDiameter / 2) ^ 2 - (parInstrumentDiameter / 2) ^ 2))
    stPoint.Set radialPlaceDist, 0, 0
    enPoint.Set parInstrumentLength, 0, 0
    Set ObjRotaMtrHorCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotaMtrHorCylinder
    Set ObjRotaMtrHorCylinder = Nothing

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFlowDiameter / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set -parFlowDiameter / 2 - sptOffset + depth, parFacetoFace, 0
    oDir.Set -1, 0, 0
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
