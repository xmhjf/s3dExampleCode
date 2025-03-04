VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          BG
'   Creation Date:  Wednesday, May 29 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'   22.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

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
    Dim ObjConvergingCone As Object
    Dim ObjVenturi As Object
    Dim ObjDivergingCone As Object
    Dim ObjNozzle1Ins As Object
    Dim ObjNozzle2Ins As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parVenturiLength As Double
    Dim parVenturiDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parVenturiLength = arrayOfInputs(4)
    parVenturiDiameter = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
    iOutput = 0

 ' Insert your code for output 1(Converging Cone)
    Dim basecenter As AutoMath.DPosition
    Dim topcenter As AutoMath.DPosition
    
    Set basecenter = New AutoMath.DPosition
    Set topcenter = New AutoMath.DPosition
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    basecenter.Set -parFace1toCenter + flangeThick, 0, 0
    topcenter.Set -parVenturiLength / 2, 0, 0
    Set ObjConvergingCone = PlaceCone(m_OutputColl, basecenter, topcenter, pipeDiam / 2 + parInsulationThickness, _
                                                                parVenturiDiameter / 2 + parInsulationThickness, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergingCone
    Set ObjConvergingCone = Nothing
    
 ' Insert your code for output 2(Venturi)
    Dim startPoint As AutoMath.DPosition
    Dim endPoint  As AutoMath.DPosition
    
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set -parVenturiLength / 2, 0, 0
    endPoint.Set parVenturiLength / 2, 0, 0
    Set ObjVenturi = PlaceCylinder(m_OutputColl, startPoint, endPoint, parVenturiDiameter + parInsulationThickness * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVenturi
    Set ObjVenturi = Nothing
    
 ' Insert your code for output 3(Nozzle1 Insulation)
    Dim parInsulationDiameter As Double
    
    startPoint.Set -parFace1toCenter, 0, 0
    endPoint.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, 0
    
    parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    Set ObjNozzle1Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1Ins
    Set ObjNozzle1Ins = Nothing
    
 ' Insert your code for output 4(Diverging Cone)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001    
    basecenter.Set parFace2toCenter - flangeThick, 0, 0
    topcenter.Set parVenturiLength / 2, 0, 0
    Set ObjDivergingCone = PlaceCone(m_OutputColl, basecenter, topcenter, pipeDiam / 2 + parInsulationThickness, _
                                                                parVenturiDiameter / 2 + parInsulationThickness, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergingCone
    Set ObjDivergingCone = Nothing
    Set basecenter = Nothing
    Set topcenter = Nothing
    
 ' Insert your code for output 5(Nozzle2 Insulation)
    startPoint.Set parFace2toCenter, 0, 0
    endPoint.Set parFace2toCenter - flangeThick - parInsulationThickness, 0, 0
    
    parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    If pipeDiam > flangeDiam Then parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    Set ObjNozzle2Ins = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInsulationDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2Ins
    Set ObjNozzle2Ins = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
        
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
