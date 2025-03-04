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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MA
'   Creation Date:  Tuesday, May 2007
'   Description:
'   The following Part data basis cases are addressed for the parameters specified:
'   Case A (Part data Basis value -343): (Gauge Root Valve with Single Outlet)
'                                          FacetoFace,ValveWidth and Offset
'   Case B (Part data Basis value -346): (Gauge Root Valve with Multiple Outlet)
'                                 FacetoFace,ValveWidth,Offset,Port3Offset and Port4Offset
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   24.May.2007    MA       CR-113431: Implemented Part data basis for values 343 and 346.
'   25.Sep.2007    MA       CR-121926 Enhanced Gauge Root Valve to support four ports case.
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

    Dim parFacetoFace As Double
    Dim parValveWidth As Double
    Dim parOffset As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double
    Dim parPort3Offset As Double
    Dim parPort4Offset As Double
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parValveWidth = arrayOfInputs(3)
    parOffset = arrayOfInputs(4)
    parHandwheelAngle = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    parPort3Offset = arrayOfInputs(7)
    parPort4Offset = arrayOfInputs(8)
    
    iOutput = 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                  sptOffset, depth
    
    Dim dInsulationDiameter As Double
     If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        dInsulationDiameter = flangeDiam + 2 * parInsulationThickness
    Else
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If

' Insert your code for output 5(Cone Insulation)
    Dim objConeIns As Object
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    oStPoint.Set -0.3 * parFacetoFace, 0, 0
    oEnPoint.Set -0.25 * parFacetoFace, 0, 0

    Set objConeIns = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                dInsulationDiameter, parValveWidth / 2 + parInsulationThickness)
        
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objConeIns
    Set objConeIns = Nothing
        
 ' Insert your code for output 6(CBox Insulation)
    Dim objBoxIns As Object
    
    oStPoint.Set -0.25 * parFacetoFace - parInsulationThickness, _
                -0.5 * parValveWidth - parInsulationThickness, _
                  0.5 * parValveWidth + parInsulationThickness
    oEnPoint.Set 0.5 * parFacetoFace + parInsulationThickness, _
                    0.5 * parValveWidth + parInsulationThickness, _
                    -0.5 * parValveWidth - parInsulationThickness
    
    Set objBoxIns = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBoxIns
    Set objBoxIns = Nothing

 ' Insert your code for output 7(Nozzle 1 Insulation)
    oStPoint.Set -0.5 * parFacetoFace, 0, 0
    oEnPoint.Set -0.3 * parFacetoFace, 0, 0
    
    Dim objNoz1Ins As Object
    Set objNoz1Ins = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNoz1Ins
    Set objNoz1Ins = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
