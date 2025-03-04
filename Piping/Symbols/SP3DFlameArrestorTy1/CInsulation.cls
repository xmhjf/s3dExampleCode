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
'   Author:          svsmylav
'   Creation Date:  Tuesday, Jun 4 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Five Insulation aspect outputs are created. ObjInsFABody is created using PlaceBox
' function and the other insulation outputs are created using PlaceCylinder function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
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
    Dim ObjInsFABody As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsBody1 As Object
    Dim ObjInsPort2 As Object
    Dim ObjInsBody2 As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parFlArrestorBodyLength1 As Double
    Dim parFlArrestorBodyLength2 As Double
    Dim parFlArrestorBodyHeight1 As Double
    Dim parFlArrestorBodyHeight2 As Double
    Dim parFlArrestorBodyWidth1 As Double
    Dim parFlArrestorBodyWidth2 As Double
    Dim parInsulationThickness As Double
    
    Dim Point1   As New AutoMath.DPosition
    Dim Point2   As New AutoMath.DPosition
    Dim parInsulationDiameter   As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parOffset1 = arrayOfInputs(4)
    parOffset2 = arrayOfInputs(5)
    parFlArrestorBodyLength1 = arrayOfInputs(6)
    parFlArrestorBodyLength2 = arrayOfInputs(7)
    parFlArrestorBodyHeight1 = arrayOfInputs(8)
    parFlArrestorBodyHeight2 = arrayOfInputs(9)
    parFlArrestorBodyWidth1 = arrayOfInputs(10)
    parFlArrestorBodyWidth2 = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    
    iOutput = 0

' Insert your code for output 1(Insulation for Flame Arrestor Body)
    Point1.Set -parFlArrestorBodyLength1 - parInsulationThickness, _
                -parFlArrestorBodyWidth1 - parInsulationThickness, _
                -parFlArrestorBodyHeight1 - parInsulationThickness
    Point2.Set parFlArrestorBodyLength2 + parInsulationThickness, _
                parFlArrestorBodyWidth2 + parInsulationThickness, _
                parFlArrestorBodyHeight2 + parInsulationThickness
    Set ObjInsFABody = PlaceBox(m_OutputColl, Point1, Point2)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsFABody
    Set ObjInsFABody = Nothing
    
' Insert your code for output 2(Insulation Port1 Side)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001    
    Point1.Set -parFace1toCenter, 0, -parOffset1
    Point2.Set -parFace1toCenter + flangeThick + parInsulationThickness, 0, -parOffset1
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort1 = PlaceCylinder(m_OutputColl, Point1, Point2, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort1
    Set ObjInsPort1 = Nothing
    
' Insert your code for output 3(Insulation for body Port1 Side)
    Point1.Set -parFace1toCenter + flangeThick, 0, -parOffset1
    Point2.Set -parFlArrestorBodyLength1, 0, -parOffset1
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody1 = PlaceCylinder(m_OutputColl, Point1, Point2, parInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody1
    Set ObjInsBody1 = Nothing
    
' Insert your code for output 4(Insulation Port2 Side)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    If flangeThick = 0 Then flangeThick = 0.0001
    Point1.Set parFace2toCenter, 0, parOffset2
    Point2.Set parFace2toCenter - flangeThick - parInsulationThickness, 0, parOffset2
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    If flangeDiam > pipeDiam Then parInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Set ObjInsPort2 = PlaceCylinder(m_OutputColl, Point1, Point2, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsPort2
    Set ObjInsPort2 = Nothing
    
' Insert your code for output 5(Insulation for body Port2 Side)
    Point1.Set parFace2toCenter - flangeThick, 0, parOffset2
    Point2.Set parFlArrestorBodyLength2, 0, parOffset2
    parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    Set ObjInsBody2 = PlaceCylinder(m_OutputColl, Point1, Point2, parInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody2
    Set ObjInsBody2 = Nothing
    Set Point1 = Nothing
    Set Point2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
