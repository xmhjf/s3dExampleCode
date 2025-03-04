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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IMRSV, PDMS Category Name=IFRD)
'   The symbol consists of eleven outputs, eight physical and three insulation outputs.
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
    Dim pipeDiam        As Double
    Dim pipeDiam1       As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim depth1           As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjCylinderBody1 As Object
    Dim ObjCylinderBody2 As Object
    Dim ObjReducingCone As Object
    Dim ObjDivergingCone As Object
    Dim ObjVenturiBodyIns As Object
    Dim ObjNozz3BodyIns As Object
    Dim ObjNozz4BodyIns As Object
    
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentLength2 As Double
    Dim parInstrumentLength3 As Double
    Dim parInstrumentDiameter As Double
    Dim parNozzleOffset As Double
    Dim parNozzleOffset1 As Double
    Dim parNozzleOffset2 As Double
    Dim parNozzleOffset3 As Double
    Dim parPortRotation1 As Double
    Dim parPortRotation2 As Double
    Dim parInstrumentLength4 As Double
    Dim parInstrumentLength5 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim BaseRadius As Double
    Dim TopRadius As Double
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentLength = arrayOfInputs(2)  'D2
    parInstrumentLength1 = arrayOfInputs(3) 'D3
    parInstrumentLength2 = arrayOfInputs(4) 'D4
    parInstrumentLength3 = arrayOfInputs(5) 'D5
    parInstrumentDiameter = arrayOfInputs(6)       'D6
    parNozzleOffset = arrayOfInputs(7)     'D7
    parNozzleOffset1 = arrayOfInputs(8)     'D8
    parNozzleOffset2 = arrayOfInputs(9)        'D9
    parNozzleOffset3 = arrayOfInputs(10)  'D10
    parPortRotation1 = arrayOfInputs(11)                'D11
    parPortRotation2 = arrayOfInputs(12)         'D12
    parInstrumentLength4 = arrayOfInputs(13)        'D13
    parInstrumentLength5 = arrayOfInputs(14)        'D14
    parInsulationThickness = arrayOfInputs(15)
    
    iOutput = 0
'The orgin is considered to be along the flow axis and mid of parInstrumentLength2
 ' Insert your code for output 1(Cylinder Body1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset2, depth2
    stPoint.Set -parInstrumentLength2 / 2 - parInstrumentLength1 - parInstrumentLength, 0, 0
    enPoint.Set -parInstrumentLength2 / 2 - parInstrumentLength1, 0, 0
    Set ObjCylinderBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam1, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinderBody1
    Set ObjCylinderBody1 = Nothing
    
 ' Insert your code for output 2(Cylinder Body2)
    stPoint.Set parInstrumentLength2 / 2, 0, 0
    enPoint.Set -parInstrumentLength2 / 2, 0, 0
    Set ObjCylinderBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinderBody2
    Set ObjCylinderBody2 = Nothing

 ' Insert your code for output 3(Reducing Cone)
    stPoint.Set -parInstrumentLength2 / 2 - parInstrumentLength1, 0, 0
    enPoint.Set -parInstrumentLength2 / 2, 0, 0
    Set ObjReducingCone = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam1 / 2, parInstrumentDiameter / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjReducingCone
    Set ObjReducingCone = Nothing

 ' Insert your code for output 4(Diverging Cone)
    stPoint.Set parInstrumentLength2 / 2 + parInstrumentLength3, 0, 0
    enPoint.Set parInstrumentLength2 / 2, 0, 0
    Set ObjDivergingCone = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam2 / 2, parInstrumentDiameter / 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergingCone
    Set ObjDivergingCone = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleLength As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parInstrumentLength2 / 2 - parInstrumentLength1 - parInstrumentLength - parInstrumentLength4 - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    NozzleLength = parInstrumentLength4
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parInstrumentLength2 / 2 + parInstrumentLength3 + parInstrumentLength5 + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    NozzleLength = parInstrumentLength5
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set -(parNozzleOffset1 - parNozzleOffset), _
                                (parNozzleOffset2 + sptOffset - depth) * Sin(parPortRotation1), _
                                (parNozzleOffset2 + sptOffset - depth) * Cos(parPortRotation1)
                                
    oDir.Set 0, Sin(parPortRotation1), Cos(parPortRotation1)
    NozzleLength = parNozzleOffset2
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 4
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set 0, (parNozzleOffset3 + sptOffset - depth) * Sin(parPortRotation2), _
                                    (parNozzleOffset3 + sptOffset - depth) * Cos(parPortRotation2)

    oDir.Set 0, Sin(parPortRotation2), Cos(parPortRotation2)
    NozzleLength = parNozzleOffset3
    Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

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
