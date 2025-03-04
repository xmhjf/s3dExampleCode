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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Wednesday, Jan 2 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The Details were taken from PDS Piping Component Data - Reference Guide, PDS symbol SN=F154.
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   22.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  08.FEB.2007     KKC                CR- 112137 Updated the symbol with
'                                          the partdatabasis property to have BendCenterOffset as optional
'   7.Nov.2007     RRK           TR-128456: Provided a check on end points of non flanged insulation port cylinders
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
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim oCenterPos       As New AutoMath.DPosition

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
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3Offset1 As Double
    Dim parFace3Offset2 As Double
    Dim parInsulationThickness As Double
    Dim parBendCenterOffset As Double

'   Variables defined by svsmylav
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    Dim parInsulationDiameter   As Double
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
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    iOutput = 0
    
    If (lPartDataBasis = 80) Then
    ' Assuming Bend Center Offset as 0.2 Times Face 3 to Center.
        parBendCenterOffset = 0.2 * parFace3Offset2
    Else
        parBendCenterOffset = arrayOfInputs(7)
    End If
    
    iOutput = 0

' Insert your code for output 1(Insulated Body1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
'    If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
    oStPoint.Set oCenterPos.x - (parFace1toCenter - flangeThick), oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    
    Set ObjInsulatedBody1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody1
    Set ObjInsulatedBody1 = Nothing
    
' Insert your code for output 2(Insulated Port1)
    oStPoint.Set oCenterPos.x - parFace1toCenter, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x - (parFace1toCenter - flangeThick - parInsulationThickness), _
                oCenterPos.y, oCenterPos.z
                
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        oEnPoint.Set oCenterPos.x - (parFace1toCenter - 0.0001), _
                oCenterPos.y, oCenterPos.z
        'To ensure insulation diameter doesn't exceed insulated pipe diameter
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
' Insert your code for output 3(Insulated Body2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oStPoint.Set oCenterPos.x + parFace2toCenter - flangeThick, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    Set ObjInsulatedBody2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody2
    Set ObjInsulatedBody2 = Nothing
    
' Insert your code for output 4(Insulated Port2)
    oStPoint.Set oCenterPos.x + parFace2toCenter, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace2toCenter - flangeThick - parInsulationThickness, _
                oCenterPos.y, oCenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    
    'Providing a check on end point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        oEnPoint.Set oCenterPos.x + parFace2toCenter - 0.0001, _
                oCenterPos.y, oCenterPos.z
        'To ensure insulation diameter doesn't exceed insulated pipe diameter
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    
' Insert your code for output 5(Insulated Branch Cylider1)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    dTorPriRad = (parFace3Offset2 - parFace3Offset1 - parBendCenterOffset) / (Sqr(2) - 1)
    dTorSecRad = parInsulationDiameter / 2
    
    oStPoint.Set oCenterPos.x, oCenterPos.y, oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace3Offset1 - dTorPriRad * (1 - Cos(PI / 4)), _
                oCenterPos.y + parFace3Offset2 - parBendCenterOffset - dTorPriRad * Sin(PI / 4), _
                oCenterPos.z
    Set ObjInsBranchCyl1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
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
    
    Dim oCentPoint   As AutoMath.DPosition
    Dim oAxis        As AutoMath.DVector
    Set oCentPoint = New AutoMath.DPosition
    oCentPoint.Set oCenterPos.x + parFace3Offset1 - dTorPriRad, _
                        oCenterPos.y + parFace3Offset2 - parBendCenterOffset, _
                       oCenterPos.z
                       
    Dim dRadius As Double
    dRadius = dTorSecRad
    
    If CmpDblGreaterthan(oCentPoint.x, cir1CenterX - dRadius) Then dRadius = -oCentPoint.x + cir1CenterX - 0.001
    
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                m_OutputColl.ResourceManager, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, dRadius)
' Insert your code for output BranchTor
    
    Set oAxis = New AutoMath.DVector
    oAxis.Set 0, 0, -1
    
    Set ObjInsBranchTorSeg = PlaceRevolution(m_OutputColl, objCirc1, _
                                    oAxis, oCentPoint, PI / 4, _
                                    False)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranchTorSeg
    Set ObjInsBranchTorSeg = Nothing
    
' Insert your code for output 7(Insulated Branch Cylinder2)
    oStPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + parFace3Offset2 - parBendCenterOffset, _
                oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + parFace3Offset2, _
                oCenterPos.z
    
    If CmpDblGreaterthan(oCentPoint.x, cir1CenterX - dTorSecRad) Then
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        oStPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + pipeDiam / 2 + parInsulationThickness, _
                oCenterPos.z
    End If
    
    Set ObjInsBranchCyl2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBranchCyl2
    Set ObjInsBranchCyl2 = Nothing
    
' Insert your code for output 8(Insulated Port3)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oStPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + parFace3Offset2 - flangeThick - parInsulationThickness, _
                oCenterPos.z
    oEnPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + parFace3Offset2, oCenterPos.z
    parInsulationDiameter = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
    
    'Providing a check on start point for non flanged ends in which case a
    'small cylinder of negligible thickness is created
    If CmpDblEqual(flangeThick, 0) Then
        oStPoint.Set oCenterPos.x + parFace3Offset1, _
                oCenterPos.y + parFace3Offset2 - 0.0001, _
                oCenterPos.z
        'To ensure insulation diameter doesn't exceed insulated pipe diameter
        parInsulationDiameter = pipeDiam + 2 * parInsulationThickness
    End If
    
    Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
    Set ObjInsulatedPort3 = Nothing

    Set objCirc1 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCentPoint = Nothing
    Set oAxis = Nothing
    Set oCenterPos = Nothing
    Set geomFactory = Nothing
    Set oPipeComponent = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
