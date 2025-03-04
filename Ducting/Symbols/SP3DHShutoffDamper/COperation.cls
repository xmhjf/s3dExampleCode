VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   COperation.cls
'   Author:         RRK
'   Creation Date:  Tuesday, Aug 7 2007
'   Description:
'   This symbol is prepared for Shut off damper that caters to Round and Rectangular shapes as per CR-123916

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Operation:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

   Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parDamperLength As Double
    Dim parActuatorPosition As Double
    Dim parCentertoSlot As Double
    Dim parSlotToLeverEnd As Double
    Dim parActuatorWidth As Double
    Dim parRodExtension As Double
    Dim parAngle As Double
    Dim parHVACShape As Double
    Dim parInsulationThickness As Double
    
    Dim CornerRadius As Double
    CornerRadius = 0

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parDamperLength = arrayOfInputs(4)
    parActuatorPosition = arrayOfInputs(5)
    parCentertoSlot = arrayOfInputs(6)
    parSlotToLeverEnd = arrayOfInputs(7)
    parActuatorWidth = arrayOfInputs(8)
    parRodExtension = arrayOfInputs(9)
    parAngle = arrayOfInputs(10)
    parHVACShape = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    
    
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
'Creating Operation Envelope (Output 1)
    
    'Assuming lever width to be 25mm
    Dim dLeverWidth As Double
    dLeverWidth = 0.025
    Dim LineStrPoints(0 To 8) As Double
    
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim dActPosFromCenter As Double
    dActPosFromCenter = (parDamperLength / 2) - parActuatorPosition
    
    'Creating points for the outline of operation envelope
    LineStrPoints(0) = -(dLeverWidth / 2) + dActPosFromCenter
    LineStrPoints(1) = -(dLeverWidth / 2)
    LineStrPoints(2) = (parWidth / 2) + parActuatorWidth
    
    LineStrPoints(3) = LineStrPoints(0)
    LineStrPoints(4) = parCentertoSlot + parSlotToLeverEnd
    LineStrPoints(5) = LineStrPoints(2)
    
    Dim dLowestAngle As Double
    Dim dTipOfLever As Double
    
    Dim dTipOfLeverAng As Double
    dTipOfLeverAng = Atn((dLeverWidth / 2) / (parCentertoSlot + parSlotToLeverEnd))
    
    dTipOfLever = Sqr((parCentertoSlot + parSlotToLeverEnd) ^ 2 + _
        (dLeverWidth / 2) ^ 2)
    
    If parHVACShape = 4 And CmpDblGreaterthanOrEqualTo(parWidth, 0.6) Then
        dLowestAngle = 5 * PI / 180
    Else
        dLowestAngle = 0
    End If
    
    LineStrPoints(6) = (dTipOfLever * Cos(dTipOfLeverAng + dLowestAngle)) + dActPosFromCenter
    LineStrPoints(7) = -dTipOfLever * Sin(dTipOfLeverAng + dLowestAngle)
    LineStrPoints(8) = LineStrPoints(2)
    
    Dim objCurvesCollection  As Collection
    Set objCurvesCollection = New Collection
    
    'Defining lines (or curves) for the outline of operation envelope
    Dim oLine1 As IngrGeom3D.Line3d
    Set oLine1 = New IngrGeom3D.Line3d
    Set oLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            LineStrPoints(0), LineStrPoints(1), LineStrPoints(2), _
                            LineStrPoints(3), LineStrPoints(4), LineStrPoints(5))
    objCurvesCollection.Add oLine1
    Set oLine1 = Nothing
    
    Dim oArc As IngrGeom3D.Arc3d
    Set oArc = New IngrGeom3D.Arc3d
    Set oArc = geomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, dActPosFromCenter, 0, LineStrPoints(2), _
                                LineStrPoints(3), LineStrPoints(4), LineStrPoints(5), _
                                LineStrPoints(6), LineStrPoints(7), LineStrPoints(8))
    objCurvesCollection.Add oArc
    Set oArc = Nothing
                            
    Dim oLine2 As IngrGeom3D.Line3d
    Set oLine2 = New IngrGeom3D.Line3d
    Set oLine2 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            LineStrPoints(6), LineStrPoints(7), LineStrPoints(8), _
                            LineStrPoints(0), LineStrPoints(1), LineStrPoints(2))
    objCurvesCollection.Add oLine2
    Set oLine2 = Nothing
        
    Dim oStPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set LineStrPoints(0), LineStrPoints(1), LineStrPoints(2)
    
    Dim objOpEnvOutline  As IngrGeom3D.ComplexString3d
    Set objOpEnvOutline = PlaceTrCString(oStPoint, objCurvesCollection)
    Set objCurvesCollection = Nothing
    
    Dim objOpEnvelope As Object
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    
    oAxisVec.Set 0, 0, -1
    Set objOpEnvelope = PlaceProjection(m_OutputColl, objOpEnvOutline, oAxisVec, 0.006, True)
    
    Dim oTransformationMat  As AutoMath.DT4x4
    Set oTransformationMat = New AutoMath.DT4x4
    
    Dim oDirVector As AutoMath.DVector
    Set oDirVector = New AutoMath.DVector
    
    'Transforming lever in case of Rectangular shape
    If (parHVACShape = Rectangular) Then
        oTransformationMat.LoadIdentity
        oTransformationMat.IndexValue(0) = -1
        objOpEnvelope.Transform oTransformationMat
        
        oTransformationMat.LoadIdentity
        oDirVector.Set 2 * dActPosFromCenter, 0, 0
        oTransformationMat.Translate oDirVector
        objOpEnvelope.Transform oTransformationMat
    End If
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objOpEnvelope
    Set objOpEnvelope = Nothing
    Set oTransformationMat = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub


