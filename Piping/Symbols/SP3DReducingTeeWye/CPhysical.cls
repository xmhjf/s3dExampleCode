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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         KKC
'   Creation Date:  Tuesday, Jan 16 2007
'   This symbol is graphical representation of the Reducing Tee Wye developed from
'   the victaulic general catalog pg 39 No. 32-R
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   16.JAN.2007     KKC       Created :CR-112137 Provide additional symbols required for Victaulic General Catalog
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
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
    Dim oCenterPos       As AutoMath.DPosition
    
    Set oCenterPos = New AutoMath.DPosition
    
    oCenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjBranchTor As Object
    Dim ObjCylinder     As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3Offset1 As Double
    Dim parFace3Offset2 As Double
    Dim parInsulationThickness As Double
    Dim parBendCenterOffset As Double

    Dim oStPoint                 As AutoMath.DPosition
    Dim oEnPoint                 As AutoMath.DPosition
    Dim oCentPoint               As AutoMath.DPosition
    Dim dTorPriRad               As Double
    Dim dTorSecRad               As Double
    Dim oTorAxis                As AutoMath.DVector
    Dim oAxis                   As AutoMath.DVector
    Dim objCirc1                As IngrGeom3D.Circle3d
    Dim dCir1CenterX             As Double
    Dim dCir1CenterY             As Double
    Dim dCir1CenterZ             As Double
    Dim dCirNormalX              As Double
    Dim dCirNormalY              As Double
    Dim dCirNormalZ              As Double
    
    Set oTorAxis = New AutoMath.DVector
    Set oAxis = New AutoMath.DVector
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
'  Place cylinder
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    dTorPriRad = (parFace3Offset2 - parFace3Offset1 - parBendCenterOffset) / (Sqr(2) - 1)
    dTorSecRad = pipeDiam / 2

    oStPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace3Offset1 - dTorPriRad * (1 - Cos(PI / 4)), _
                oCenterPos.y + parFace3Offset2 - parBendCenterOffset - dTorPriRad * Sin(PI / 4), _
                oCenterPos.z
                
    Set ObjCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder
    Set ObjCylinder = Nothing
    
'  Place Torus segment
'   Prepare points needed for circle, to be used to generate Torus segment
    dCir1CenterX = oCenterPos.x + parFace3Offset1
    dCir1CenterY = oCenterPos.y + parFace3Offset2 - parBendCenterOffset
    dCir1CenterZ = 0

    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    dCirNormalX = 0
    dCirNormalY = 1
    dCirNormalZ = 0
    
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                        Nothing, dCir1CenterX, _
                                dCir1CenterY, dCir1CenterZ, dCirNormalX, _
                                dCirNormalY, dCirNormalZ, pipeDiam / 2)

' Insert your code for output 2(Branch Torus)
    Set oCentPoint = New AutoMath.DPosition
    oCentPoint.Set oCenterPos.x + parFace3Offset1 - dTorPriRad, _
                        oCenterPos.y + parFace3Offset2 - parBendCenterOffset, _
                        oCenterPos.z
    oAxis.Set 0, 0, -1
    
    Set ObjBranchTor = PlaceRevolution(m_OutputColl, objCirc1, _
                                    oAxis, oCentPoint, PI / 4, _
                                    False)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchTor
    Set ObjBranchTor = Nothing
    
' Place Cylinder
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset, depth
    Dim dconcentriclength As Double
    Dim dCylindricalLength As Double
' Assuming concentric reducing length as 0.2 times Face 2 to Center
' Assuming cylindrical Length from origin is 0.5 times Face 2 to Center
    Dim ObjCylinder1 As Object
    Dim objCone1 As Object
    Dim oStPosition As AutoMath.DPosition
    Dim oEnPosition As AutoMath.DPosition
    Dim oConeEnPosition As AutoMath.DPosition
    
    Set oStPosition = New AutoMath.DPosition
    Set oEnPosition = New AutoMath.DPosition
    Set oConeEnPosition = New AutoMath.DPosition
    
    oConeEnPosition.Set 0.7 * parFace2toCenter, 0, 0
    oStPosition.Set 0, 0, 0
    oEnPosition.Set parFace2toCenter / 2, 0, 0
    
    Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPosition, oEnPosition, pipeDiam1, False)
    
'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1
    Set ObjCylinder1 = Nothing
' Assumed Cone end position is at 0.7 times Face 2 to Center.
    oConeEnPosition.Set 0.7 * parFace2toCenter, 0, 0
    
    Set objCone1 = PlaceCone(m_OutputColl, oEnPosition, oConeEnPosition, pipeDiam1 / 2, pipeDiam2 / 2, False)

'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCone1
    Set objCone1 = Nothing
 
 ' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set oCenterPos.x - parFace1toCenter - sptOffset + depth, _
                    oCenterPos.y, oCenterPos.z
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzleWithLength( _
                        1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, parFace1toCenter)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
 'Place Nozzle 2
    oPlacePoint.Set oCenterPos.x + parFace2toCenter + sptOffset - depth, _
                    oCenterPos.y, oCenterPos.z
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzleWithLength( _
                        2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, 0.3 * parFace2toCenter)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                            flangeThick, flangeDiam, sptOffset, depth

    oPlacePoint.Set oCenterPos.x + parFace3Offset1, _
                    oCenterPos.y + parFace3Offset2 + sptOffset - depth, _
                    oCenterPos.z
    oDir.Set 0, 1, 0

    Set objNozzle = CreateNozzleWithLength( _
                        3, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, parBendCenterOffset)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set objCirc1 = Nothing
    Set oPlacePoint = Nothing
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

