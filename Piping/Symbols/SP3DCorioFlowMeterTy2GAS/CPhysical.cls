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
'   Creation Date:  Monday, Oct 7 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=ICM3,
'   PDMS Category Name=IFMZ). Physical aspect consist of 1 main cylinder body, 2 Support
'   Cylinder and a Box. Insulation aspect consist of a Cylinder and a box
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
    Dim flangeThick1     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim pipeDiam2        As Double
    Dim flangeThick2     As Double
    Dim sptOffset2       As Double
    Dim depth2           As Double
    Dim iOutput     As Double
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjFlowMtrBody As Object
    Dim ObjCylinder1 As Object
    Dim ObjCylinder2 As Object
    Dim ObjBox As Object
    Dim ObjBodyIns As Object
    Dim ObjBoxIns As Object
    
    Dim parFacetoFace As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)       'D2
    parFlowDiameter = arrayOfInputs(3)       'D3
    parInstrumentHeight = arrayOfInputs(4)       'D4
    parInstrumentLength = arrayOfInputs(5)       'D5
    parInstrumentHeight1 = arrayOfInputs(6)       'D6
    parInstrumentWidth = arrayOfInputs(7)       'D7
    parInstrumentWidth1 = arrayOfInputs(8)       'D8
    parInsulationThickness = arrayOfInputs(9)
    
    iOutput = 0

  ' Insert your code for output 1(Flow Meter Body)
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam, sptOffset2, depth2
    
    startPoint.Set -parFacetoFace / 2 + flangeThick1, 0, 0
    endPoint.Set parFacetoFace / 2 - flangeThick2, 0, 0
    Set ObjFlowMtrBody = PlaceCylinder(m_OutputColl, startPoint, endPoint, parFlowDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowMtrBody
    Set ObjFlowMtrBody = Nothing
    
 ' Insert your code for output 2(Cylinder 1)
 'Center to center between the Cylinders is assumed to be 0.7 times of the variable parInstrumentLength
 'Cylinder diameter is assumed to be 0.1 times of the variable parInstrumentLength
    Dim cylinCenter As Double
    Dim cylinDiameter As Double
    
    cylinCenter = 0.7 * parInstrumentLength
    cylinDiameter = 0.1 * parInstrumentLength
    startPoint.Set -cylinCenter / 2, ((parFlowDiameter / 2) ^ 2 - (cylinDiameter / 2) ^ 2), 0
    endPoint.Set -cylinCenter / 2, (parInstrumentHeight - parInstrumentHeight1), 0
    Set ObjCylinder1 = PlaceCylinder(m_OutputColl, startPoint, endPoint, cylinDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1
    Set ObjCylinder1 = Nothing
    
 ' Insert your code for output 3(Cylinder 2)
    startPoint.Set cylinCenter / 2, ((parFlowDiameter / 2) ^ 2 - (cylinDiameter / 2) ^ 2), 0
    endPoint.Set cylinCenter / 2, (parInstrumentHeight - parInstrumentHeight1), 0
    Set ObjCylinder2 = PlaceCylinder(m_OutputColl, startPoint, endPoint, cylinDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2
    Set ObjCylinder2 = Nothing
    
 ' Insert your code for output 4(Box)
    startPoint.Set -parInstrumentLength / 2, parInstrumentHeight, -parInstrumentWidth1
    endPoint.Set parInstrumentLength / 2, (parInstrumentHeight - parInstrumentHeight1), _
                            parInstrumentWidth
    Set ObjBox = PlaceBox(m_OutputColl, startPoint, endPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox
    Set ObjBox = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
        ' Place Nozzle 1 for Output 4 (Nozzle1)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2 for output 5 (Nozzle2)
    oPlacePoint.Set parFacetoFace / 2 + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
