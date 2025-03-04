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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          svsmylav
'   Creation Date:  Wednesday, Jan 2 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy      who                  change description
'   -----------      -----                ------------------
'   09.Jul.2003     SymbolTeam(India)     Copyright Information, Header  is added.
'   22.Aug.2005     svsmylav              TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                         cylinder to avoid PlaceCylinder function failure.
'   08.SEP.2006      KKC                  DI-95670  Replace names with initials in all revision history sheets and symbols
'   12.SEP.2007      VRG                  CR-126671 Modified symbol code in order to support Face to Center and Diameter dimension basis
'   7.Nov.2007       VRG                  TR-128456: Provided a check on end points of non flanged insulation port cylinders
'                                         in which case small cylinders of negligible thickness will be created
'   12.AUG.2008      PK                   CR-145847 Modified the symbol code to support the partdatabasis values 1055,1056,1057,1058,1059,1060,1061,1062 and 1063
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
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
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Integer
    Dim ObjSP3DReturn As Object
    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    
    Dim parFacetoCenter As Double
    Dim parInsulationThickness As Double
    Dim parDiameter As Double
    Dim dElbowRadius As Double
    
    'Variables defined by svsmylav
    Dim stPoint               As New AutoMath.DPosition
    Dim enPoint               As New AutoMath.DPosition
    Dim dInsDiameterPort      As Double
    Dim dInsulationRadius     As Double
    
    Dim parCentertoCenter       As Double
    Dim parBacktoFace           As Double
    Dim parBacktoTangentLine    As Double
    Dim parTangentLength        As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    'parDiameter = arrayOfInputs(4)
    'parCentertoCenter=arrayOfInputs(5)
    'parBacktoFace = arrayOfInputs(6)
    'parBacktoTangentLine = arrayOfInputs(7)
    'parTangentLength = arrayOfInputs(8)
        
    iOutput = 0
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
      
    Select Case lPartDataBasis
    
        Case 414    '180 degree return, open, specified by center-to-center and back-centerline-to-face
            parFacetoCenter = arrayOfInputs(2)
            parDiameter = arrayOfInputs(4)
            dElbowRadius = parDiameter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                dElbowRadius = pipeDiam / 2 + parInsulationThickness + NEGLIGIBLE_THICKNESS
            End If
            
        Case Is <= 1, 413   'Default or 180 degree return, open, specified by back-centerline-to-face
            parFacetoCenter = arrayOfInputs(2)
            dElbowRadius = parFacetoCenter - flangeThick
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                dElbowRadius = pipeDiam / 2 + parInsulationThickness + NEGLIGIBLE_THICKNESS
            End If
        
        Case 1055   '180 degree return, closed, specified by center-to-center
            parCentertoCenter = arrayOfInputs(5)
            dElbowRadius = parCentertoCenter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                pipeDiam = dElbowRadius
            End If
            
        Case 1056   '180 degree return, open, specified by center-to-center
            parCentertoCenter = arrayOfInputs(5)
            dElbowRadius = parCentertoCenter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                dElbowRadius = pipeDiam / 2 + parInsulationThickness + NEGLIGIBLE_THICKNESS
            End If
    
        Case 1057   '180 degree return, open, specified by back-to-face
            parBacktoFace = arrayOfInputs(6)
            dElbowRadius = parBacktoFace - pipeDiam / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                dElbowRadius = pipeDiam / 2 + parInsulationThickness + NEGLIGIBLE_THICKNESS
            End If
                    
        Case 1058   '180 degree return, open, single tangent, specified by center-to-center and tangent, port 1
            parCentertoCenter = arrayOfInputs(5)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parCentertoCenter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                dElbowRadius = pipeDiam / 2 + parInsulationThickness + NEGLIGIBLE_THICKNESS
            End If
            
        Case 1059   '180 degree return, open, single tangent, specified by back-to-tangent-line and tangent, port 1
            parBacktoTangentLine = arrayOfInputs(7)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parBacktoTangentLine - pipeDiam / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                dElbowRadius = pipeDiam / 2 + parInsulationThickness + NEGLIGIBLE_THICKNESS
            End If
            
        Case 1060   '180 degree return, open, double tangent, specified by center-to-center and tangent
            parCentertoCenter = arrayOfInputs(5)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parCentertoCenter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                dElbowRadius = pipeDiam / 2 + parInsulationThickness + NEGLIGIBLE_THICKNESS
            End If
            
        Case 1061   '180 degree return, open, double tangent, specified by back-to-tangent-line and tangent
            parBacktoTangentLine = arrayOfInputs(7)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parBacktoTangentLine - pipeDiam / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                dElbowRadius = pipeDiam / 2 + parInsulationThickness + NEGLIGIBLE_THICKNESS
            End If
            
        Case 1062   '180 degree return, open, single tangent, specified by center-to-center and tangent, port 2
            parCentertoCenter = arrayOfInputs(5)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parCentertoCenter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                dElbowRadius = pipeDiam / 2 + parInsulationThickness + NEGLIGIBLE_THICKNESS
            End If
            
        Case 1063   '180 degree return, open, single tangent, specified by back-to-face and tangent, port 2
            parBacktoFace = arrayOfInputs(6)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parBacktoFace - pipeDiam / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2 + parInsulationThickness) Then
                dElbowRadius = pipeDiam / 2 + parInsulationThickness + NEGLIGIBLE_THICKNESS
            End If
    End Select
    
    Dim objCirc1        As IngrGeom3D.Circle3d
    Dim geomFactory     As New IngrGeom3D.GeometryFactory
    
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    cirNormalX = -1
    cirNormalY = 0
    cirNormalZ = 0
    
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    If lPartDataBasis <= 1 Or lPartDataBasis = 413 Or lPartDataBasis = 414 Then
        cir1CenterX = CenterPos.x - parFacetoCenter + flangeThick
        cir1CenterY = CenterPos.y
        cir1CenterZ = CenterPos.z
    ElseIf lPartDataBasis = 1055 Or lPartDataBasis = 1056 Or lPartDataBasis = 1057 _
        Or lPartDataBasis = 1058 Or lPartDataBasis = 1059 Or lPartDataBasis = 1060 _
        Or lPartDataBasis = 1061 Or lPartDataBasis = 1062 Or lPartDataBasis = 1063 Then
            cir1CenterX = CenterPos.x
            cir1CenterY = CenterPos.y + dElbowRadius
            cir1CenterZ = CenterPos.z
    End If
    
    dInsulationRadius = pipeDiam / 2 + parInsulationThickness
    'Check to ensure that the insulation radius is always less than the calculated elbow radius
    If lPartDataBasis = 1055 Then
        If CmpDblGreaterthan(dInsulationRadius, parCentertoCenter / 2) Then
            dInsulationRadius = parCentertoCenter / 2 - NEGLIGIBLE_THICKNESS
        End If
    End If
        
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                m_OutputColl.ResourceManager, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, dInsulationRadius)

    
    'Insert your code for output 1(Fitting - Return )
    Dim centPoint   As New AutoMath.DPosition
    Dim axis        As New AutoMath.DVector
 
    If lPartDataBasis <= 1 Or lPartDataBasis = 413 Or lPartDataBasis = 414 Then
        centPoint.Set CenterPos.x - parFacetoCenter + flangeThick, _
                      CenterPos.y + dElbowRadius, _
                      CenterPos.z
        axis.Set 0, 0, 1
    ElseIf lPartDataBasis = 1055 Or lPartDataBasis = 1056 Or lPartDataBasis = 1057 Or _
           lPartDataBasis = 1058 Or lPartDataBasis = 1059 Or lPartDataBasis = 1060 Or _
           lPartDataBasis = 1061 Or lPartDataBasis = 1062 Or lPartDataBasis = 1063 Then
           
        centPoint.Set CenterPos.x, CenterPos.y, CenterPos.z
        axis.Set 0, 0, -1
    End If
    
    Set ObjInsulatedBody = PlaceRevolution(m_OutputColl, objCirc1, _
                                    axis, centPoint, PI, _
                                        True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
    Set ObjInsulatedBody = Nothing
    
    If lPartDataBasis = 1055 Then
        Dim oEndCurve       As Object
        Dim oInsProjection  As IngrGeom3D.Projection3d
        Dim oCurves         As Collection
        Dim oLine           As Object
        Dim oArc            As IngrGeom3D.Arc3d
        Dim oStPoint        As AutoMath.DPosition
        Dim oVec            As AutoMath.DVector
        Dim dInsElbowRadius As Double
                
        Set oStPoint = New AutoMath.DPosition
        Set oCurves = New Collection
        Set oVec = New AutoMath.DVector
        dInsElbowRadius = dElbowRadius + parInsulationThickness
                
        Set oArc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                            CenterPos.x, CenterPos.y + dInsElbowRadius, CenterPos.z - flangeDiam / 2 - parInsulationThickness, _
                            CenterPos.x, CenterPos.y + dInsElbowRadius + flangeDiam / 2 + parInsulationThickness, CenterPos.z, _
                            CenterPos.x, CenterPos.y + dInsElbowRadius, CenterPos.z + flangeDiam / 2 + parInsulationThickness)
        oCurves.Add oArc
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            CenterPos.x, CenterPos.y + dInsElbowRadius, CenterPos.z + flangeDiam / 2 + parInsulationThickness, _
                            CenterPos.x, CenterPos.y - dInsElbowRadius, CenterPos.z + flangeDiam / 2 + parInsulationThickness)
        oCurves.Add oLine
        Set oArc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                            CenterPos.x, CenterPos.y - dInsElbowRadius, CenterPos.z + flangeDiam / 2 + parInsulationThickness, _
                            CenterPos.x, CenterPos.y - dInsElbowRadius - flangeDiam / 2 - parInsulationThickness, CenterPos.z, _
                            CenterPos.x, CenterPos.y - dInsElbowRadius, CenterPos.z - flangeDiam / 2 - parInsulationThickness)
        oCurves.Add oArc
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            CenterPos.x, CenterPos.y - dInsElbowRadius, CenterPos.z - flangeDiam / 2 - parInsulationThickness, _
                            CenterPos.x, CenterPos.y + dInsElbowRadius, CenterPos.z - flangeDiam / 2 - parInsulationThickness)
        oCurves.Add oLine

        oStPoint.Set CenterPos.x, CenterPos.y + dInsElbowRadius, CenterPos.z - flangeDiam / 2 - parInsulationThickness
        oVec.Set 1, 0, 0
        Set oEndCurve = PlaceTrCString(oStPoint, oCurves)
        
        Set oInsProjection = PlaceProjection(m_OutputColl, oEndCurve, oVec, flangeThick + parInsulationThickness, True)
        
        m_OutputColl.AddOutput "Projection", oInsProjection
        
        Set oEndCurve = Nothing
        Set oInsProjection = Nothing
        Set oCurves = Nothing
        Set oLine = Nothing
        Set oArc = Nothing
        Set oStPoint = Nothing
        Set oVec = Nothing
        
    End If
    
    'Insert your code for output 2 (Insulated Port1)
    Select Case lPartDataBasis
    
        Case Is <= 1, 413, 414
            
            stPoint.Set CenterPos.x - parFacetoCenter, cir1CenterY, CenterPos.z
            enPoint.Set CenterPos.x - (parFacetoCenter - flangeThick - parInsulationThickness), _
                        cir1CenterY, CenterPos.z
                    
            If CmpDblEqual(flangeThick, 0) Then
                enPoint.Set CenterPos.x - parFacetoCenter + NEGLIGIBLE_THICKNESS, cir1CenterY, CenterPos.z
            End If
            
        Case 1055, 1056, 1057, 1062, 1063
            stPoint.Set CenterPos.x, CenterPos.y + dElbowRadius, CenterPos.z
            enPoint.Set CenterPos.x + flangeThick + parInsulationThickness, _
                                CenterPos.y + dElbowRadius, _
                                CenterPos.z
            
            If CmpDblEqual(flangeThick, 0) Then
                enPoint.Set CenterPos.x + NEGLIGIBLE_THICKNESS, _
                                CenterPos.y + dElbowRadius, _
                                CenterPos.z
            End If

        Case 1058, 1059, 1060, 1061
            stPoint.Set CenterPos.x - parTangentLength, _
                                CenterPos.y + dElbowRadius, _
                                CenterPos.z
            enPoint.Set stPoint.x + flangeThick + parInsulationThickness, _
                                stPoint.y, _
                                stPoint.z
            If CmpDblEqual(flangeThick, 0) Then
                enPoint.Set stPoint.x + NEGLIGIBLE_THICKNESS, _
                                stPoint.y, _
                                stPoint.z
            End If
                          
    End Select
    
    dInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsDiameterPort = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsDiameterPort, True)
            
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
    Set ObjInsulatedPort1 = Nothing
    
    'Create an Insulation Cylinder for the nozzle length
    Dim oInsTangent As Object
    If lPartDataBasis = 1058 Or lPartDataBasis = 1059 Or lPartDataBasis = 1060 Or lPartDataBasis = 1061 Then
        stPoint.Set CenterPos.x - parTangentLength, CenterPos.y + dElbowRadius, CenterPos.z
        enPoint.Set CenterPos.x, CenterPos.y + dElbowRadius, CenterPos.z
    
        Set oInsTangent = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                    pipeDiam + 2 * parInsulationThickness, True)
                    
        m_OutputColl.AddOutput "InsTangent1", oInsTangent
    End If
            
    'Insert your code for output 3(Insulated Port2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 413, 414
        
            stPoint.Set CenterPos.x - parFacetoCenter, _
                        CenterPos.y + 2 * dElbowRadius, CenterPos.z
            enPoint.Set CenterPos.x - (parFacetoCenter - flangeThick - parInsulationThickness), _
                        CenterPos.y + 2 * dElbowRadius, CenterPos.z
            If CmpDblEqual(flangeThick, 0) Then
                enPoint.Set CenterPos.x - parFacetoCenter + NEGLIGIBLE_THICKNESS, _
                        CenterPos.y + 2 * dElbowRadius, CenterPos.z
            End If
        
        Case 1055, 1056, 1057, 1058, 1059
            
            stPoint.Set CenterPos.x, CenterPos.y - dElbowRadius, CenterPos.z
            enPoint.Set CenterPos.x + flangeThick + parInsulationThickness, _
                        CenterPos.y - dElbowRadius, CenterPos.z
            
            If CmpDblEqual(flangeThick, 0) Then
                enPoint.Set CenterPos.x + NEGLIGIBLE_THICKNESS, _
                        CenterPos.y - dElbowRadius, _
                        CenterPos.z
            End If
            
        Case 1060, 1061, 1062, 1063
            
            stPoint.Set CenterPos.x - parTangentLength, _
                            CenterPos.y - dElbowRadius, _
                            CenterPos.z
            enPoint.Set stPoint.x + flangeThick + parInsulationThickness, _
                            stPoint.y, _
                            stPoint.z
            If CmpDblEqual(flangeThick, 0) Then
                enPoint.Set stPoint.x + NEGLIGIBLE_THICKNESS, _
                        stPoint.y, _
                        stPoint.z
            End If
         
    End Select
    
    dInsDiameterPort = pipeDiam + parInsulationThickness * 2
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsDiameterPort = flangeDiam + parInsulationThickness * 2
    Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, dInsDiameterPort, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
    Set ObjInsulatedPort2 = Nothing
    
    'Create an Insulation Cylinder for the nozzle length
    If lPartDataBasis = 1060 Or lPartDataBasis = 1061 Or lPartDataBasis = 1062 Or lPartDataBasis = 1063 Then
        stPoint.Set CenterPos.x - parTangentLength, CenterPos.y - dElbowRadius, CenterPos.z
        enPoint.Set CenterPos.x, CenterPos.y - dElbowRadius, CenterPos.z
        
        Set oInsTangent = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                    pipeDiam + 2 * parInsulationThickness, True)
                    
        m_OutputColl.AddOutput "InsTangent2", oInsTangent
    End If
    
    Set oInsTangent = Nothing
    
    'remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCirc1
    Set objCirc1 = Nothing
    sp3dElem1.Remove
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
