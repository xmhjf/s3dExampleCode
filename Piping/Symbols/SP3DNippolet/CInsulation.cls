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
'   Copyright (c) 2005-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         kkk
'   Creation Date:  Wednesday, November 16 2005
'   Description:
'   The Details were taken from PDS Piping Component Data - Reference Guide at Page no D-78
'   SN=F176 and MC=NOL. Symbol consist of Physical and Insulation aspects. Physical aspect is
'   made up of 1.Diverging Tapering Cone, 2.Tapering Cone, 3.Nozzle-1 with zero length, 4.Nozzle-2
'   with length.
'   Insulation Aspect consist of Insulation for Diverging cone, Converging cone, insulation for the
'   straight portion of nozzle-2.
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  18.APR.2007     KKC  CR-117893 Modified Olet symbol to be Trimmable
'  2.Jul.2007     KKC  TR-123363: Added PI/2 value to parAngle(if Olet is not trimmable)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double


Private Sub Class_Initialize()
 PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjDivergConeIns As Object
    Dim ObjConvergConeIns As Object
    Dim ObjNoz2StraightIns As Object
    
    Dim parFacetoHeaderCenter As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
    Dim DivergConeLen As Double
    Dim ConvergConeLen As Double
    Dim NipoletIntruIntoPipe As Double
    Dim NipoletConeLen As Double
    Dim NozzleLength As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoHeaderCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    If UBound(arrayOfInputs) >= 4 Then 'Not to fail the prior to V7 instances by accessing index 31
        parAngle = arrayOfInputs(4)
    Else
        parAngle = PI / 2 'To suppport the prior to V7 instances
    End If
    
    iOutput = 0
    
'   Extracted from PDS EDEN CODE used in the same way as defined in PDS.
'   olet_body = F_to_F_Dim - Dimension_10
'   olet_body_1 = 0.2 * olet_body 'DivergConeLen
'   olet_body_2 = 0.4 * olet_body 'ConvergConeLen
'   olet_body_3 = 0.4 * olet_body 'NozzleLength
'   waist_dimension = 1.5 * Body_OD_2
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset2, depth2
    Dim dOlet_body As Double
    NipoletIntruIntoPipe = pipeDiam1 / 2 - _
                                Sqr( _
                                (pipeDiam1 / 2 * pipeDiam1 / 2) - (pipeDiam2 / 2 * pipeDiam2 / 2) _
                                )
    dOlet_body = parFacetoHeaderCenter - pipeDiam1 / 2 + NipoletConeLen
    NipoletConeLen = 3 / 5 * dOlet_body
    DivergConeLen = dOlet_body * 1 / 5
    ConvergConeLen = dOlet_body * 2 / 5
    NozzleLength = 2 / 5 * dOlet_body
    
  ' Insert your code for output 1(Diverging Cone)
    Dim baseCenterPt As AutoMath.DPosition
    Dim topCenterPt As AutoMath.DPosition
    Set baseCenterPt = New AutoMath.DPosition
    Set topCenterPt = New AutoMath.DPosition
    
    baseCenterPt.Set 0, parFacetoHeaderCenter - dOlet_body, 0
    topCenterPt.Set 0, baseCenterPt.y + DivergConeLen, 0
    'The enlarged portion Diameter of Endolet is 1.5 times of Pipe-2 Diameter, as given in PDS catalog
    Set ObjDivergConeIns = PlaceCone(m_OutputColl, baseCenterPt, topCenterPt, pipeDiam2 / 2 + parInsulationThickness, _
                                                pipeDiam2 * 1.5 / 2 + parInsulationThickness, True)
    Dim oDirVector As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotAboutZaxis As Double
    
    Set oDirVector = New AutoMath.DVector
    oTransformationMat.LoadIdentity
    dRotAboutZaxis = PI / 2 - parAngle
    oDirVector.Set 0, 0, -1
    oTransformationMat.Rotate dRotAboutZaxis, oDirVector
    ObjDivergConeIns.Transform oTransformationMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergConeIns
    Set ObjDivergConeIns = Nothing
    
 ' Insert your code for output 2(Converging Cone)
    baseCenterPt.Set 0, parFacetoHeaderCenter - dOlet_body + DivergConeLen, 0
    topCenterPt.Set 0, baseCenterPt.y + ConvergConeLen, 0
    'The enlarged portion Diameter of Nippolet is 1.5 times of Pipe-2 Diameter, as given in PDS catalog
    Set ObjConvergConeIns = PlaceCone(m_OutputColl, baseCenterPt, topCenterPt, pipeDiam2 * 1.5 / 2 + parInsulationThickness, _
                                                pipeDiam2 / 2 + parInsulationThickness, True)
    ObjConvergConeIns.Transform oTransformationMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergConeIns
    Set ObjConvergConeIns = Nothing
    
' Insert your code for output 3(Insulation for the straight portion of Nozzle-2)
    Dim InsulationDiameter As Double
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
    startPoint.Set 0, baseCenterPt.y + ConvergConeLen, 0
    endPoint.Set 0, startPoint.y + NozzleLength, 0
    InsulationDiameter = pipeDiam2 + parInsulationThickness * 2
    Set ObjNoz2StraightIns = PlaceCylinder(m_OutputColl, startPoint, endPoint, InsulationDiameter, True)
    ObjNoz2StraightIns.Transform oTransformationMat
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNoz2StraightIns
    Set ObjNoz2StraightIns = Nothing
    Set baseCenterPt = Nothing
    Set topCenterPt = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Set oTransformationMat = Nothing
    Set oDirVector = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
