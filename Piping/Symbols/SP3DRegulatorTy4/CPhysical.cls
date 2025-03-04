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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IREG4, PDMS Category Name=IVRI)
'   The symbol consists of  eight outputs, seven physical and one insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
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
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick1     As Double
    Dim flangeThick2     As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjConeBody1 As Object
    Dim ObjConeBody2 As Object
    Dim ObjConeBody3 As Object
    Dim ObjConeBody4 As Object
    Dim ObjTopCylBody As Object
    Dim ObjRegulatorBodyIns As Object
    Dim ObjConeBody4Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight2 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim BaseRadius As Double
    Dim TopRadius As Double
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFacetoFace = arrayOfInputs(3)        'D3
    parInstrumentHeight = arrayOfInputs(4)   'D4
    parInstrumentHeight1 = arrayOfInputs(5)  'D5
    parInstrumentDiameter = arrayOfInputs(6)  'D6
    parInstrumentHeight2 = arrayOfInputs(7)  'D9
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0
'The origin is taken to be along the flow axis and mid of parFacetoFace
 ' Insert your code for output 1(Cone Body1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam, sptOffset2, depth2
    stPoint.Set -parFace1toCenter + flangeThick1, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjConeBody1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam1 / 2, 0, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConeBody1
    Set ObjConeBody1 = Nothing
 
 ' Insert your code for output 2(Cone Body2)
    stPoint.Set parFacetoFace - parFace1toCenter - flangeThick2, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjConeBody2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam2 / 2, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConeBody2
    Set ObjConeBody2 = Nothing

 ' Insert your code for output 3(Cone Body3)
    stPoint.Set 0, parInstrumentHeight, 0
    enPoint.Set 0, 0, 0
'Diameter of third cone base is assumed to be pipediameter of nozzle2
    Set ObjConeBody3 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam2 / 2, 0, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConeBody3
    Set ObjConeBody3 = Nothing
    
 ' Insert your code for output 4(Cone Body4)
    stPoint.Set 0, -parInstrumentHeight1, 0
    enPoint.Set 0, 0, 0
    Set ObjConeBody4 = PlaceCone(m_OutputColl, stPoint, enPoint, parInstrumentDiameter / 2, 0, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConeBody4
    Set ObjConeBody4 = Nothing
    
 ' Insert your code for output 5(Top Cylinder Body)
'The diameter of Top Cylinder Body is assumed to be 0.2 times PipeDiameter of Nozzle2
    stPoint.Set 0, parInstrumentHeight, 0
    enPoint.Set 0, parInstrumentHeight2, 0
    Set ObjTopCylBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, 0.2 * pipeDiam2, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopCylBody
    Set ObjTopCylBody = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFacetoFace - parFace1toCenter + sptOffset2 - depth2, 0, 0
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
