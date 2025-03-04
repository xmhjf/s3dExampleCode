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
'   Author:         KKC
'   Creation Date:  Tuesday, Jan 16 2007
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   16.JAN.2007     KKC       Created :CR-112137 Provide additional symbols required for Victaulic General Catalog
'   7.Nov.2007      RRK          TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                                in which case small cylinders of negligible thickness will be created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI           As Double

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim pipeDiam1       As Double
    Dim pipeDiam2       As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim oCenterPos      As AutoMath.DPosition
    Set oCenterPos = New AutoMath.DPosition
    
    oCenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjInsulatedBody1 As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedBody2 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsBranchCyl1 As Object
    Dim ObjInsBranchTorSeg As Object
    Dim ObjInsBranchCyl2 As Object
    Dim ObjInsulatedPort3 As Object
    Dim ObjBranchTor As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3Offset1 As Double
    Dim parFace3Offset2 As Double
    Dim parInsulationThickness As Double
    Dim parBendCenterOffset As Double
    
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Dim dInsulationDiameter   As Double
    Dim dInsulationDiameter1 As Double
    Dim dTorPriRad               As Double
    Dim dTorSecRad               As Double
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3Offset1 = arrayOfInputs(4)
    parFace3Offset2 = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    
' Assumed Bend Center offset as 0.2 times Face 3 to offset 2
    parBendCenterOffset = 0.2 * parFace3Offset2
    
    iOutput = 0

' Insert your code for output 1(Insulated Body1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oStPoint.Set oCenterPos.x - (parFace1toCenter - flangeThick), oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    Set ObjInsulatedBody1 = PlaceCylinder(m_OutputColl, oStPoint, _
                                        oEnPoint, dInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody1
    Set ObjInsulatedBody1 = Nothing
    
' Insert your code for output 2(Insulated Port1)
    oStPoint.Set oCenterPos.x - parFace1toCenter, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x - (parFace1toCenter - flangeThick - parInsulationThickness), _
                oCenterPos.y, oCenterPos.z
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        oEnPoint.Set oCenterPos.x - 0.0001, _
                oCenterPos.y, oCenterPos.z
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, _
                                        oEnPoint, dInsulationDiameter, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
' Insert your code for output 3(Insulated Body2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset, depth

    dInsulationDiameter = pipeDiam1 + parInsulationThickness * 2
    dInsulationDiameter1 = pipeDiam2 + parInsulationThickness * 2
    
    Dim ObjInsulatedCylinder1 As Object
    Dim objInsulatedCone1 As Object
    
    Dim oStPosition As AutoMath.DPosition
    Set oStPosition = New AutoMath.DPosition
    Dim oEnPosition As AutoMath.DPosition
    Set oEnPosition = New AutoMath.DPosition
    
    oStPosition.Set 0, 0, 0
    oEnPosition.Set parFace2toCenter / 2, 0, 0
    
    Set ObjInsulatedCylinder1 = PlaceCylinder(m_OutputColl, oStPosition, oEnPosition, dInsulationDiameter, False)

'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedCylinder1
    Set ObjInsulatedCylinder1 = Nothing
    
'Place cylinder on port 2 side
    Dim oConeEnPosition As AutoMath.DPosition
    Set oConeEnPosition = New AutoMath.DPosition
    
    oConeEnPosition.Set 0.7 * parFace2toCenter, 0, 0
    
    Set objInsulatedCone1 = PlaceCone(m_OutputColl, oEnPosition, oConeEnPosition, dInsulationDiameter / 2, dInsulationDiameter1 / 2, False)

'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedCone1
    Set objInsulatedCone1 = Nothing
    
'Place Cylinder for the Port
    oStPoint.Set oCenterPos.x + parFace2toCenter - flangeThick, oCenterPos.y, oCenterPos.z
    oEnPoint.Set parFace2toCenter * 0.7, oCenterPos.y, oCenterPos.z
    
    Set ObjInsulatedBody2 = PlaceCylinder(m_OutputColl, oStPoint, _
                                     oEnPoint, dInsulationDiameter1, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody2
    Set ObjInsulatedBody2 = Nothing
    
' Insert your code for output 4(Insulated Port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset, depth
    
    oStPoint.Set oCenterPos.x + parFace2toCenter, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace2toCenter - flangeThick - parInsulationThickness, _
                oCenterPos.y, oCenterPos.z
                
    If CmpDblGreaterthan(flangeDiam, pipeDiam2) Then dInsulationDiameter1 = flangeDiam + parInsulationThickness * 2
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        oEnPoint.Set oCenterPos.x + parFace2toCenter - 0.0001, _
                oCenterPos.y, oCenterPos.z
        dInsulationDiameter1 = pipeDiam2 + 2 * parInsulationThickness
    End If
    
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter1, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    
' Insert your code for output 5(Insulated Branch Cylider1)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    dTorPriRad = (parFace3Offset2 - parFace3Offset1 - parBendCenterOffset) / (Sqr(2) - 1)
    dTorSecRad = dInsulationDiameter / 2
    
    oStPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace3Offset1 - dTorPriRad * (1 - Cos(PI / 4)), _
                oCenterPos.y + parFace3Offset2 - parBendCenterOffset - dTorPriRad * Sin(PI / 4), _
                oCenterPos.z
    
    Set ObjInsBranchCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranchCyl1
    Set ObjInsBranchCyl1 = Nothing
    
' Insert your code for output 6(Insulated Branch Torus Segment)
    Dim objCirc1 As IngrGeom3D.Circle3d
    
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    cirNormalX = 0
    cirNormalY = 1
    cirNormalZ = 0
    
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    cir1CenterX = oCenterPos.x + parFace3Offset1
    cir1CenterY = oCenterPos.y + parFace3Offset2 - parBendCenterOffset
    cir1CenterZ = 0
    
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                Nothing, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, dTorSecRad)
' Insert your code for output BranchTor
    Dim oCentPoint   As AutoMath.DPosition
    Dim oAxis        As AutoMath.DVector
    Set oCentPoint = New AutoMath.DPosition
    Set oAxis = New AutoMath.DVector

    oCentPoint.Set oCenterPos.x + parFace3Offset1 - dTorPriRad, _
                        oCenterPos.y + parFace3Offset2 - parBendCenterOffset, _
                        oCenterPos.z
    oAxis.Set 0, 0, -1
    
    Set ObjInsBranchTorSeg = PlaceRevolution(m_OutputColl, objCirc1, _
                                    oAxis, oCentPoint, PI / 4, _
                                    False)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranchTorSeg
    Set ObjInsBranchTorSeg = Nothing
    
' Insert your code for output 7(Insulated Branch Cylinder 2)
    oStPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + parFace3Offset2 - parBendCenterOffset, _
                oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + parFace3Offset2, _
                oCenterPos.z
    
    Set ObjInsBranchCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranchCyl2
    Set ObjInsBranchCyl2 = Nothing
    
' Insert your code for output 8(Insulated Port3)
    oStPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + parFace3Offset2 - flangeThick - parInsulationThickness, _
                oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + parFace3Offset2, oCenterPos.z
    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
    
    'Providing a check on start point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        oStPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + parFace3Offset2 - 0.0001, _
                oCenterPos.z
        dInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCentPoint = Nothing
    Set oCenterPos = Nothing
    Set oStPosition = Nothing
    Set oEnPosition = Nothing
    Set oConeEnPosition = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
