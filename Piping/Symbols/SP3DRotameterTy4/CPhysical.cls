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
'   Creation Date:  Tuesday, Sep 17 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR2, PDMS Category Name=IFNH)
'   The symbol consists of four outputs, three physical and one insulation outputs.
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
    Dim ObjRotameterBody As Object
    Dim ObjRotameterBodyIns As Object
    Dim ObjTopCylBody As Object
    Dim ObjBottCylBody As Object
    Dim ObjNozz1BodyIns As Object
    Dim ObjNozz2BodyIns As Object
        
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parFacetoFace As Double
    Dim parPortRotation1 As Double
    Dim parFace1toCenter As Double
    Dim parPortRotation2 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)   'D2
    parInstrumentHeight1 = arrayOfInputs(3)  'D3
    parInstrumentDiameter = arrayOfInputs(4)      'D4
    parInstrumentLength = arrayOfInputs(5)        'D5
    parInstrumentLength1 = arrayOfInputs(6)       'D6
    parInstrumentWidth = arrayOfInputs(7)       'D7
    parInstrumentWidth1 = arrayOfInputs(8)       'D8
    parFacetoFace = arrayOfInputs(9)  'D9
    parPortRotation1 = arrayOfInputs(10)      'D10
    parFace1toCenter = arrayOfInputs(11)     'D11
    parPortRotation2 = arrayOfInputs(12)      'D12
    parInsulationThickness = arrayOfInputs(13)

    iOutput = 0
'The origin is considered to be mid of parInstrumentHeight1 and on the vertical axis
 ' Insert your code for output 1(Rotameter Body)
    stPoint.Set -parInstrumentLength, -parInstrumentHeight1 / 2, parInstrumentWidth
    enPoint.Set parInstrumentLength1, parInstrumentHeight1 / 2, -parInstrumentWidth1
    Set ObjRotameterBody = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotameterBody
    Set ObjRotameterBody = Nothing
    
 ' Insert your code for output 3(Top Cylinder Body)
    stPoint.Set 0, parInstrumentHeight1 / 2, 0
    enPoint.Set 0, parInstrumentHeight / 2, 0
    Set ObjTopCylBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopCylBody
    Set ObjTopCylBody = Nothing
    
 ' Insert your code for output 4(Bottom Cylinder Body)
    stPoint.Set 0, -parInstrumentHeight1 / 2, 0
    enPoint.Set 0, -parInstrumentHeight / 2, 0
    Set ObjBottCylBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottCylBody
    Set ObjBottCylBody = Nothing

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set (parFace1toCenter + sptOffset - depth) * Cos(parPortRotation2), _
                                -parFacetoFace / 2, _
                                -(parFace1toCenter + sptOffset - depth) * Sin(parPortRotation2)

    oDir.Set Cos(parPortRotation2), 0, -Sin(parPortRotation2)
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace1toCenter)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
        
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set (parFace1toCenter + sptOffset - depth) * Cos(parPortRotation1), _
                                parFacetoFace / 2, _
                                -(parFace1toCenter + sptOffset - depth) * Sin(parPortRotation1)

    oDir.Set Cos(parPortRotation1), 0, -Sin(parPortRotation1)
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace1toCenter)

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
