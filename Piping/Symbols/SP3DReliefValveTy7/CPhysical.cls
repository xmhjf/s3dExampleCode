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
'   Creation Date:  Monday, Nov 11 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT7, PDMS Category Name=IVPT)
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
    Dim ObjBox As Object
    Dim ObjInsulationBody As Object
    
    Dim parFace1toCenter As Double           'D2
    Dim parFace2toCenter As Double           'D3
    Dim parInstrumentHeight As Double        'D4
    Dim parInstrumentLength As Double        'D6
    Dim parInstrumentLength1 As Double       'D7
    Dim parInstrumentWidth As Double         'D8
    Dim parInstrumentWidth1 As Double        'D9
    Dim parInstrumentHeight1 As Double       'D25   'As per Eden ,its used to place a box.
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parInstrumentHeight = arrayOfInputs(4)
    parInstrumentLength = arrayOfInputs(5)
    parInstrumentLength1 = arrayOfInputs(6)
    parInstrumentWidth = arrayOfInputs(7)
    parInstrumentWidth1 = arrayOfInputs(8)
    parInstrumentHeight1 = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0

 ' Insert your code for output 1(Box)
    stPoint.Set -parInstrumentLength1, -parInstrumentHeight, parInstrumentWidth1
    enPoint.Set parInstrumentLength, parInstrumentHeight1, -parInstrumentWidth
    
    Set ObjBox = PlaceBox(m_OutputColl, stPoint, enPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox
    Set ObjBox = Nothing
    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim nozzleLength As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    nozzleLength = parFace1toCenter - parInstrumentLength1
    
    Set objNozzle = CreateNozzleWithLength( _
                        1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, nozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset + depth, 0
    oDir.Set 0, -1, 0

    nozzleLength = parFace2toCenter - parInstrumentHeight

    Set objNozzle = CreateNozzleWithLength( _
                        2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, nozzleLength)
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

