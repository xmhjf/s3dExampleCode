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
'   Creation Date:  Tuesday, Oct 1 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT3,
'   PDMS Category Name=IVPP). Physical aspect consist of Relief valve body of Box shape, 1 Cylinder
'   Nozzle-1 facing (-)ive side of X-Axis and Nozzle-2 facing (-)ive side of Y-axis.
'   Insulation aspect consist of a single box place covering the whole valve.
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
    
    Dim ObjRefValBody As Object
    Dim ObjVerCylin As Object
        
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentDiameter As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)            ''D2
    parFace2toCenter = arrayOfInputs(3)            ''D3
    parInstrumentHeight = arrayOfInputs(4)          'D4
    parInstrumentHeight1 = arrayOfInputs(5)         'D5
    parInstrumentLength = arrayOfInputs(6)          'D6
    parInstrumentLength1 = arrayOfInputs(7)         'D7
    parInstrumentWidth = arrayOfInputs(8)           'D8
    parInstrumentWidth1 = arrayOfInputs(9)          'D9
    parInstrumentHeight2 = arrayOfInputs(10)        'D10
    parInstrumentDiameter = arrayOfInputs(11)       'D11
    parInsulationThickness = arrayOfInputs(12)

    iOutput = 0
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

'Insert your code for output 1(Relief Valve Body)
'Origin of the symbol is assumed at the intersection of Nozzle-1 and Nozzle-2 Axes
    ''Left top point.
    stPoint.Set parInstrumentLength, parInstrumentHeight1, -parInstrumentWidth
    enPoint.Set -parInstrumentLength1, -parInstrumentHeight, parInstrumentWidth1
    Set ObjRefValBody = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRefValBody
    Set ObjRefValBody = Nothing
    
' Insert your code for output 2 (Vertical Cylinder)
    
    stPoint.Set 0, parInstrumentHeight1, 0
    enPoint.Set 0, parInstrumentHeight1 + parInstrumentHeight2, 0
    
    Set ObjVerCylin = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylin
    Set ObjVerCylin = Nothing

' Place Nozzle 1 for Output 3 (Nozzle1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim NozzleLength As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    NozzleLength = parFace1toCenter - parInstrumentLength1         'D2-D7
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2 for output 4 (Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset + depth, 0                       'D3
    oDir.Set 0, -1, 0
    NozzleLength = parFace2toCenter - parInstrumentHeight         'D3-D4
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
    
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
