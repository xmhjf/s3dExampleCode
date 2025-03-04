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
'   Author:          SSP
'   Creation Date:  Wed, Jan 29 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This Symbol details are available in PDS Piping Data Reference Guide at page no 378 and SN=1219AZ
'
'   Consist of Physical and Insulation Aspects. Physical aspect is made up of Two Nozzles, One FlowNozzleCylBody of Cylinder,
'   and Filter body of truncated cone. Insulation aspect consist of a Cylinder
'   for FlowNozzleBody Insulation.Option Explicit
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

''''

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
    Dim ObjFlowNozzleCylBody As Object
    Dim ObjFlowNozzleConeBody As Object
    Dim ObjFlowNozzleBodyIns As Object
    
    Dim parFaceToFace As Double
    Dim parFlowNozzleDiameter1 As Double
    Dim parFlowNozzleLength As Double
    Dim parFlowNozzleEndDiameter As Double
    Dim parInsulationThickness As Double
    Dim parFlowNozzleDiameter2 As Double
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFaceToFace = arrayOfInputs(2)            'P2
    parFlowNozzleDiameter1 = arrayOfInputs(3)             'P3
    parFlowNozzleLength = arrayOfInputs(4)             'P4
    parFlowNozzleEndDiameter = arrayOfInputs(5)             'P5
    parFlowNozzleDiameter2 = arrayOfInputs(6)            'P6
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0

 ' Insert your code for output 1(FlowNozzleConeBody)
    Dim baseCenPt As AutoMath.DPosition
    Dim topCenPt As AutoMath.DPosition
    Set baseCenPt = New AutoMath.DPosition
    Set topCenPt = New AutoMath.DPosition
        
    baseCenPt.Set parFaceToFace / 2, 0, 0
    topCenPt.Set parFaceToFace / 2 + parFlowNozzleLength, 0, 0
    Set ObjFlowNozzleConeBody = PlaceCone(m_OutputColl, baseCenPt, topCenPt, parFlowNozzleDiameter2 / 2, parFlowNozzleEndDiameter / 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowNozzleConeBody
    Set ObjFlowNozzleConeBody = Nothing
    Set baseCenPt = Nothing
    Set topCenPt = Nothing
 
 ' Insert your code for output 2(FlowNozzleCylBody)
    Dim startPt As AutoMath.DPosition
    Dim endPt As AutoMath.DPosition
    Set startPt = New AutoMath.DPosition
    Set endPt = New AutoMath.DPosition
    
    startPt.Set -parFaceToFace / 2, 0, 0
    endPt.Set parFaceToFace / 2, 0, 0
    Set ObjFlowNozzleCylBody = PlaceCylinder(m_OutputColl, startPt, endPt, parFlowNozzleDiameter1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlowNozzleCylBody
    Set ObjFlowNozzleCylBody = Nothing
    Set startPt = Nothing
    Set endPt = Nothing
 
' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFaceToFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFaceToFace / 2 + sptOffset - depth, 0, 0
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
