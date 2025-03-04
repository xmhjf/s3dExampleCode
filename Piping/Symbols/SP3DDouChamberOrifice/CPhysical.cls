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
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ICM1)
'   The symbol consists of nine outputs, seven physical and two insulation outputs.
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
    Dim pipeDiam1        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim pipeDiam2        As Double
    Dim sptOffset1       As Double
    Dim depth1           As Double
    Dim pipeDiam        As Double
    
    
    Dim iOutput     As Double
    Dim ObjChamberBody As Object
    Dim ObjArmCyl1 As Object
    Dim ObjArmBody2 As Object
    Dim ObjNozzCylIns As Object
    Dim ObjChamberBodyIns As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parArmLength As Double
    Dim parNozzleOffset As Double
    Dim parNozzleOffset1 As Double
    Dim parArmHeight As Double
    Dim parArmHeight1 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFace2toCenter = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)  'D4
    parInstrumentHeight1 = arrayOfInputs(5) 'D5
    parInstrumentLength = arrayOfInputs(6)  'D6
    parInstrumentLength1 = arrayOfInputs(7) 'D7
    parInstrumentWidth = arrayOfInputs(8)   'D8
    parInstrumentWidth1 = arrayOfInputs(9)  'D9
    parArmLength = arrayOfInputs(10)    'D10
    parNozzleOffset = arrayOfInputs(11)   'D11
    parNozzleOffset1 = arrayOfInputs(12)  'D12
    parArmHeight = arrayOfInputs(13)    'D14
    parArmHeight1 = arrayOfInputs(14)   'D15
    parInsulationThickness = arrayOfInputs(15)
    
    iOutput = 0
'The orgin is considered to be along the flow axis and at parFace1toCenter from nozzle 1
 ' Insert your code for output 1(Chamber Body)
    stPoint.Set -parInstrumentLength, -parInstrumentHeight, parInstrumentWidth
    enPoint.Set parInstrumentLength1, parInstrumentHeight1, -parInstrumentWidth1
    Set ObjChamberBody = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjChamberBody
    Set ObjChamberBody = Nothing
    
 ' Insert your code for output 2(Arm Cylinder1)
'Assuming the Horizontal Arm Diameter to be 0.2 times the parInstrumentLength
    Dim arm1Diameter As Double
    arm1Diameter = 0.2 * parInstrumentLength
    stPoint.Set 0, parArmHeight, -parInstrumentWidth1
    enPoint.Set 0, parArmHeight, -parArmLength
    Set ObjArmCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, arm1Diameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArmCyl1
    Set ObjArmCyl1 = Nothing
    
 ' Insert your code for output 3(Arm Body2)
'Assuming the Arm to be rectangular cross-section, with width being 0.2 times the parInstrumentLength, and Thickness being 0.1times the parInstrumentLength
    Dim arm2Width As Double
    Dim arm2Thickness As Double
    arm2Width = 0.6 * parInstrumentLength
    arm2Thickness = 0.1 * parInstrumentLength
    
    stPoint.Set arm2Width / 2, parArmHeight + arm1Diameter / 2 - parArmHeight1, -parArmLength + arm2Thickness
    enPoint.Set -arm2Width / 2, parArmHeight + arm1Diameter / 2, -parArmLength
    Set ObjArmBody2 = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjArmBody2
    Set ObjArmBody2 = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleLength As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    NozzleLength = parFace1toCenter - parInstrumentLength
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    NozzleLength = parFace2toCenter - parInstrumentLength1
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set -parNozzleOffset, 0, parInstrumentWidth + sptOffset - depth
    oDir.Set 0, 0, 1
    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 4
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parNozzleOffset1, 0, parInstrumentWidth + sptOffset - depth
    oDir.Set 0, 0, 1
    Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, oPlacePoint)

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
