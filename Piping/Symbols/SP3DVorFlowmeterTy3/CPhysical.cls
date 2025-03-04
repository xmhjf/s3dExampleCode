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
'   Creation Date:  Tuesday, Sep 23 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IVM3,
'   PDMS Category Name=IFNE). Physical aspect consist of one horizontal cylinder, one top cylinder
'   along Y axis, a top box and two nozzles one facing -ive direction direction of X-axis and the other in +ive
'   direction of X-axis. Insulation aspect consist of a cylinders along X-axis and a box to cover both
'   top cylinder and top box
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
    
    
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjHoriCylin As Object
    Dim ObjTopCylin As Object
    Dim ObjTopBox As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFlowDiameter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFace2toCenter = arrayOfInputs(3) 'D3
    parFlowDiameter = arrayOfInputs(4)  'D4
    parInstrumentHeight = arrayOfInputs(5)  'D5
    parInstrumentHeight1 = arrayOfInputs(6) 'D6
    parInstrumentWidth = arrayOfInputs(7)   'D7
    parInstrumentWidth1 = arrayOfInputs(8)  'D8
    parInstrumentLength = arrayOfInputs(9)  'D9
    parInstrumentLength1 = arrayOfInputs(10) 'D10
    parInsulationThickness = arrayOfInputs(11)
    
    iOutput = 0
    
' Insert your code for output 1(Vortex Flowmeter HorizntalCylinder)
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
'   Origin of the symbol is assumed at the intersection of Top cylinder axis and nozzles  axes
    startPoint.Set -parFace1toCenter, 0, 0
    endPoint.Set parFace2toCenter, 0, 0
    Set ObjHoriCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, parFlowDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjHoriCylin
    Set ObjHoriCylin = Nothing

' Insert your code for output 2(Vortex Flowmeter Top Cylinder)
'   Top Cylinder Diameter is assumed as 0.1 times of Flow Diameter
    Dim TopCylinDia As Double
    
    TopCylinDia = 0.1 * parFlowDiameter
    startPoint.Set 0, Sqr((parFlowDiameter / 2) ^ 2 - (TopCylinDia / 2) ^ 2), 0
    endPoint.Set 0, (parInstrumentHeight - parInstrumentHeight1), 0
    Set ObjTopCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, TopCylinDia, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopCylin
    Set ObjTopCylin = Nothing
    
' Insert your code for output 3(Top Instrument Box)
    startPoint.Set -parInstrumentLength, (parInstrumentHeight - parInstrumentHeight1), _
                    -parInstrumentWidth1
    endPoint.Set parInstrumentLength1, parInstrumentHeight, parInstrumentWidth
    Set ObjTopBox = PlaceBox(m_OutputColl, startPoint, endPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopBox
    Set ObjTopBox = Nothing

' Place Nozzle 1 for Output 4 (Nozzle1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2 for output 5 (Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
