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
'   Author:          BG
'   Creation Date:  Monday, Nov 11 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR1,
'   PDMS Category Name=IFNG)
'   The symbol consists of seven outputs, five physical and two insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim ObjCyl1 As Object
    Dim ObjCyl2 As Object
    Dim ObjCyl3 As Object
    Dim ObjInsulationBody1 As Object
    Dim ObjInsulationBody2 As Object
    
    Dim parInstrumentHeight As Double            'D2
    Dim parInstrumentDiameter As Double          'D3
    Dim parInstrumentHeight1 As Double           'D4
    Dim parFacetoFace As Double                  'D5
    Dim parInstrumentLength As Double            'D6
    Dim parInstrumentDiameter1 As Double         'D7
    Dim parInstrumentDiameter2 As Double         'D8
    Dim parInstrumentLength1 As Double           'D9
    Dim parInstrumentHeight2 As Double           'D10
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)
    parInstrumentDiameter = arrayOfInputs(3)
    parInstrumentHeight1 = arrayOfInputs(4)
    parFacetoFace = arrayOfInputs(5)
    parInstrumentLength = arrayOfInputs(6)
    parInstrumentDiameter1 = arrayOfInputs(7)
    parInstrumentDiameter2 = arrayOfInputs(8)
    parInstrumentLength1 = arrayOfInputs(9)
    parInstrumentHeight2 = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)
    
    iOutput = 0

 ' Insert your code for output 1(Cyl1)
    stPoint.Set -(parInstrumentDiameter / 2 + parInstrumentLength1), 0, 0
    enPoint.Set -parInstrumentDiameter / 2, 0, 0
    
    Set ObjCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl1
    Set ObjCyl1 = Nothing
    
 ' Insert your code for output 2(Cyl2)
    stPoint.Set 0, -(parInstrumentHeight1 - (parInstrumentHeight2 - parInstrumentHeight)), _
                0
    enPoint.Set 0, parInstrumentHeight2 - parInstrumentHeight1, 0
    
    Set ObjCyl2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl2
    Set ObjCyl2 = Nothing
    
 ' Insert your code for output 3(Cyl3)
    stPoint.Set parInstrumentDiameter / 2, 0, 0
    
' This checks in case D6 = D3/2 which causes a failure of placecylinder.
    If parInstrumentLength = parInstrumentDiameter / 2 Then
        enPoint.Set stPoint.x + 0.001, 0, 0
    Else
        enPoint.Set parInstrumentLength, 0, 0
    End If
    
    Set ObjCyl3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl3
    Set ObjCyl3 = Nothing
    
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -(parInstrumentDiameter / 2 + parInstrumentLength1) - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set -parInstrumentDiameter / 2 - sptOffset + depth, parFacetoFace, 0
    oDir.Set -1, 0, 0
'   In normal scenario nozzle length will be equal to parInstrumentDiameter / 2,
'   if parInstrumentDiameter / 2 is less than (or equal to) the flange thickness then nozzle
'   length will be equal to flange thickness
    If parInstrumentDiameter / 2 <= flangeThick Then
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    Else
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, parInstrumentDiameter / 2)
    End If
    
' Set the output
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
