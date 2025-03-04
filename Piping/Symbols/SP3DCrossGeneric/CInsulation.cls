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
'   Author:          BG
'   Creation Date:  Monday, Dec 31 2001
'   Description:
'        Cross Specified by face-to-center 1, face-to ceter 2, face-to-center 3, face-to-center 4 (default)
'
'        "CR-127569  Create sanitary cross symbol" and "CR-127570  Create reducing sanitary cross symbol"
'        Source for below PDB's: A74.17725-1.pdf, A861.26224-1.pdf, ABI Foundry Catalog.pdf, C110-A21.10-03.pdf, No Hub Fitting sheet.pdf, Sect4(Fittings).pdf and Tyler Soil Pipe Catalog.pdf
'        Cross, specified by  face-to-center 1 and face-to-center 3 (589)
'        Cross, specified by seat-to-center 1, seat-to-center 2, and seat-to-center 3 (590)
'        Sanitary cross, specified by header face-to-face dimension, face-to-center 1, and face-to center 3 (591)
'        Sanitary cross, specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3 (592)
'        Sanitary cross, specified by header seat-to-seat dimension, seat-to-center 1, and face-to center 3 (593)
'        Sanitary cross with side opening, specified by header seat-to-seat dimension, seat-to-center 1, seat-to-center 3, offset for side opening (594)
'        Sanitary cross with two 90o side openings above center, specified by header face-to-face dimension, face-to-center 1, horizontal offset for opening above center, and vertical offset for opening above center (595)
'        Sanitary cross with two 45o openings on same side, specified by header face-to-face dimension, face-to-center 1, face-to-center for side opening, horizontal offset for 45o side opening (596)
'        T branches, double, specified by header face-to-face dimension, face-to-center 1, and face-to-center 3 (597)
'        T branches, double, specified by face-to-center 1, face-to-center 2, and face-to-center 3 (598)
'        T branches, double, specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3 (599)
'        Sanitary T branches, double, specified by header seat-to-seat dimension, seat-to-center 1, seat-to-center 3, branch radius, and tangent length 3 (600)
'        Sanitary T branches, double, specified by header face-to-face dimension, face-to-center 1, face-to-center 3, branch radius, and tangent length 3 (601)
'        Sanitary T branches, double, specified by face-to-center 1, face-to-center 2, face-to-center 3, branch radius, and tangent length 3 (602)
'        Sanitary T branches, double, specified by header face-to-face dimension, tangent length 1, branch radius, face-to-center 3, and tangent length 3 (603)
'        Sanitary T branches, double, specified by tangent length 1, branch radius, face-to-center 2, face-to-center 3, and tangent length 3 (604)
'        Sanitary T branches, double, specified by header seat-to-seat dimension, seat-to-center 1, face-to-center 3, and branch radius (605)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   19.Aug.2005     svsmylav        TR-83739: Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   11.Jan.2008     RUK                     CR-127570  Create reducing sanitary cross symbol
'   21.Jan.2008     RUK                     CR-127569  Create sanitary cross symbol
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

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim sptOffset As Double
    Dim flangeDiam As Double
    Dim depth As Double

    Dim iOutput As Double
    Dim ObjPort1Insulation As Object
    Dim ObjPort2Insulation As Object
    Dim ObjPort3Insulation As Object
    Dim ObjPort4Insulation As Object
    Dim ObjMainBodyInsul As Object
    Dim ObjTaperConeInsul As Object
    Dim ObjBranchInsulation As Object
    Dim ObjBranchInsulation1 As Object

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parInsulationThickness As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parFace4toCenter = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)

    iOutput = 0
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    Select Case lPartDataBasis
        Case Is <= 1
        ' Insert your code for output 6(Insulation for Port 1)
            Dim stPoint As New AutoMath.DPosition
            Dim enPoint As New AutoMath.DPosition
            Dim parInsulationDiameter As Double
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
            If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
            stPoint.Set -parFace1toCenter, 0, 0
            enPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
            Set ObjPort1Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort1Insulation
            Set ObjPort1Insulation = Nothing
    
            ' Insert your code for output 10(Insulation for the Main Body)
            stPoint.Set -(parFace1toCenter - flangeThick), 0, 0
            enPoint.Set 0, 0, 0
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
            Set ObjMainBodyInsul = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainBodyInsul
            Set ObjMainBodyInsul = Nothing
    
            ' Insert your code for output 7(Insulation for Port 2)
            Dim pipeDiam2 As Double
            Dim flangeThick2 As Double
            Dim sptOffset2 As Double
            Dim flangeDiam2 As Double
            Dim depth2 As Double
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
            If CmpDblEqual(flangeThick2, 0) Then flangeThick2 = 0.0001
            stPoint.Set parFace2toCenter, 0, 0
            enPoint.Set (parFace2toCenter - flangeThick2 - parInsulationThickness), 0, 0
            parInsulationDiameter = pipeDiam2 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam2, pipeDiam2) Then parInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
            Set ObjPort2Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort2Insulation
            Set ObjPort2Insulation = Nothing
    
            ' Insert your code for output 11(Insulation for Taper Cone)
            Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
            Dim objCircle1 As IngrGeom3D.Circle3d, objCircle2 As IngrGeom3D.Circle3d
            Dim cir2CenterX As Double, cir2CenterY As Double, cir2CenterZ As Double
            Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
            Dim geomFactory As New IngrGeom3D.GeometryFactory
            pipeDiam = pipeDiam + parInsulationThickness * 2
            pipeDiam2 = pipeDiam2 + parInsulationThickness * 2
            cirNormalX = 1
            cirNormalY = 0
            cirNormalZ = 0
            cir1CenterX = 0
            cir1CenterY = 0
            cir1CenterZ = 0
            Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                             Nothing, cir1CenterX, _
                             cir1CenterY, cir1CenterZ, cirNormalX, _
                             cirNormalY, cirNormalZ, pipeDiam * 0.5)
            'Create Circle2
            cir2CenterX = parFace2toCenter - flangeThick2     'Should be parFace2toFace - flangeThick2-sptOffset
            cir2CenterY = 0
            cir2CenterZ = 0
            Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                             Nothing, cir2CenterX, _
                             cir2CenterY, cir2CenterZ, cirNormalX, _
                             cirNormalY, cirNormalZ, pipeDiam2 * 0.5)
            'Create Ruled Surface for body
            Set ObjTaperConeInsul = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                                               objCircle1, objCircle2, True)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperConeInsul
            Set ObjTaperConeInsul = Nothing
    
            ' Insert your code for output 8(Insulation for Port 3)
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
            If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
            stPoint.Set 0, parFace3toCenter, 0
            enPoint.Set 0, (parFace3toCenter - flangeThick - parInsulationThickness), 0
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
            Set ObjPort3Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort3Insulation
            Set ObjPort3Insulation = Nothing
    
            ' Insert your code for output 12(Insulation for Branch)
            stPoint.Set 0, (parFace3toCenter - flangeThick), 0
            enPoint.Set 0, 0, 0
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
            Set ObjBranchInsulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchInsulation
            Set ObjBranchInsulation = Nothing
    
            ' Insert your code for output 9(Insulation for Port 4)
            RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
            If CmpDblEqual(flangeThick, 0) Then flangeThick = 0.0001
            stPoint.Set 0, -parFace4toCenter, 0
            enPoint.Set 0, -(parFace4toCenter - flangeThick - parInsulationThickness), 0
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then parInsulationDiameter = flangeDiam + parInsulationThickness * 2
            Set ObjPort4Insulation = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPort4Insulation
            Set ObjPort4Insulation = Nothing
    
            ' Insert your code for output 13(Insulation for Branch 1)
            stPoint.Set 0, -(parFace4toCenter - flangeThick), 0
            enPoint.Set 0, 0, 0
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
            Set ObjBranchInsulation1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchInsulation1
            Set ObjBranchInsulation1 = Nothing

        Case 589 To 605
            Dim parFacetoFace As Double
            Dim parAngle As Double
            Dim parFace5toCenter As Double
            Dim parSeattoSeat As Double
            Dim parSeat1toOpening As Double
            Dim parSeat1toCenter As Double
            Dim parSeat2toCenter As Double
            Dim parSeat3toCenter As Double
            Dim parPorttoPortOffset As Double
            Dim parBranchRadius As Double
            Dim parTangentLength1 As Double
            Dim parTangentLength3 As Double
    
            parFace1toCenter = arrayOfInputs(2)
            parFace3toCenter = arrayOfInputs(4)
            parFacetoFace = arrayOfInputs(7)
            parAngle = arrayOfInputs(8)
            If UBound(arrayOfInputs) > 8 Then
                parFace5toCenter = arrayOfInputs(9)
                parSeattoSeat = arrayOfInputs(10)
                parSeat1toOpening = arrayOfInputs(11)
                parSeat1toCenter = arrayOfInputs(12)
                parSeat2toCenter = arrayOfInputs(13)
                parSeat3toCenter = arrayOfInputs(14)
                parPorttoPortOffset = arrayOfInputs(15)
                parBranchRadius = arrayOfInputs(16)
                parTangentLength1 = arrayOfInputs(17)
                parTangentLength3 = arrayOfInputs(18)
            End If
    
            Dim pipeDiam1 As Double
            Dim flangeThick1 As Double
            Dim flangeDiam1 As Double
            Dim depth1 As Double
            Dim CptOffset1 As Double
            
            Dim CptOffset2 As Double
    
            Dim pipeDiam3 As Double
            Dim flangeThick3 As Double
            Dim CptOffset3 As Double
            Dim flangeDiam3 As Double
            Dim depth3 As Double
    
            Dim pipeDiam4 As Double
            Dim flangeThick4 As Double
            Dim CptOffset4 As Double
            Dim flangeDiam4 As Double
            Dim depth4 As Double
            
            'Retrieve the Nozzles properties
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, _
                                                                                    CptOffset1, depth1
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                                    CptOffset2, depth2
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, _
                                                                                    CptOffset3, depth3
            RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, _
                                                                                    CptOffset4, depth4
            
            Dim oGeomFactory As IngrGeom3D.GeometryFactory
            Dim oStPoint As AutoMath.DPosition
            Dim oEnPoint As AutoMath.DPosition
            Dim oCenPoint As AutoMath.DPosition
            Dim oAxisVect As AutoMath.DVector
            Dim oCircle As Object
            
            Dim objInsPort As Object
            Dim dInsThick As Double
            
            Set oGeomFactory = New GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oCenPoint = New AutoMath.DPosition
            Set oAxisVect = New AutoMath.DVector
    
            Dim dFace1toCenter As Double
            Dim dFace2toCenter As Double
            Dim dFace3toCenter As Double
            Dim dFacetoFace As Double
            Dim dFace1toOpening As Double
            Dim dElbowRadius As Double
            Dim dTangentLength1 As Double
            Dim dTangentLength3 As Double
            Dim dRevAngle As Double
            
            dRevAngle = PI / 2
            
            'Determine Face1toCenter, Face3toCenter, FacetoFace, Elbow Radius, Tangent Length 1 and
            'Tangent Length 3 for each PDB
            Select Case lPartDataBasis
                Case 589    'Cross, specified by  face-to-center 1 and face-to-center 3
                    dFace1toCenter = parFace1toCenter
                    dFace3toCenter = parFace3toCenter
                    dFacetoFace = 2 * parFace1toCenter
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                     dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                     dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 590    'Cross, specified by seat-to-center 1, seat-to-center 2, and seat-to-center 3
                    dFace1toCenter = parSeat1toCenter - CptOffset1 + depth1
                    dFace3toCenter = parSeat3toCenter - CptOffset3 + depth3
                    If CmpDblEqual(parSeat2toCenter, LINEAR_TOLERANCE) Then
                        dFacetoFace = 2 * dFace1toCenter
                    Else
                        dFacetoFace = dFace1toCenter + parSeat2toCenter - CptOffset2 + depth2
                    End If
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 591    'Sanitary cross, specified by header face-to-face dimension, face-to-center 1, and face-to center 3
                    dFace1toCenter = parFace1toCenter
                    dFace3toCenter = parFace3toCenter
                    dFacetoFace = parFacetoFace
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 592    'Sanitary cross, specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3
                    dFace1toCenter = parSeat1toCenter - CptOffset1 + depth1
                    dFace3toCenter = parSeat3toCenter - CptOffset3 + depth3
                    dFacetoFace = parSeattoSeat - (CptOffset1 + CptOffset2) + (depth1 + depth2)
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 593    'Sanitary cross, specified by header seat-to-seat dimension, seat-to-center 1, and face-to center 3
                    dFace1toCenter = parSeat1toCenter - CptOffset1 + depth1
                    dFace3toCenter = parFace3toCenter
                    dFacetoFace = parSeattoSeat - (CptOffset1 + CptOffset2) + (depth1 + depth2)
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 594    'Sanitary cross with side opening, specified by header seat-to-seat dimension, seat-to-center 1, seat-to-center 3, offset for side opening
                    dFace1toCenter = parSeat1toCenter - CptOffset1 + depth1
                    dFace3toCenter = parSeat3toCenter - CptOffset3 + depth3
                    dFacetoFace = parSeattoSeat - (CptOffset1 + CptOffset2) + (depth1 + depth2)
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 595    'Sanitary cross with two 90o side openings above center, specified by header face-to-face dimension, face-to-center 1, horizontal offset for opening above center, and vertical offset for opening above center
                    dFace1toCenter = parFace1toCenter
                    dFace3toCenter = parFace1toCenter
                    dFacetoFace = parFacetoFace
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 596    'Sanitary cross with two 45o openings on same side, specified by header face-to-face dimension, face-to-center 1, face-to-center for side opening, horizontal offset for 45o side opening
                    dFace1toCenter = parFace1toCenter
                    dFace3toCenter = parFace1toCenter
                    dFacetoFace = parFacetoFace
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 597    'T branches, double, specified by header face-to-face dimension, face-to-center 1, and face-to-center 3
                    dFace1toCenter = parFace1toCenter
                    dFace3toCenter = parFace1toCenter
                    dFacetoFace = parFacetoFace
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 598    'T branches, double, specified by face-to-center 1, face-to-center 2, and face-to-center 3
                    dFace1toCenter = parFace1toCenter
                    dFace3toCenter = parFace1toCenter
                    dFacetoFace = parFace1toCenter + parFace2toCenter
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 599    'T branches, double, specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3
                    dFace1toCenter = parSeat1toCenter - CptOffset1 + depth1
                    dFace3toCenter = parSeat3toCenter - CptOffset3 + depth3
                    dFacetoFace = parSeattoSeat - (CptOffset1 + CptOffset2) + (depth1 + depth2)
                    'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                    If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                        dElbowRadius = 0.8 * dFace3toCenter
                    Else
                        dElbowRadius = 0.8 * dFace1toCenter
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 600    'Sanitary T branches, double, specified by header seat-to-seat dimension, seat-to-center 1, seat-to-center 3, branch radius, and tangent length 3
                    dFace1toCenter = parSeat1toCenter - CptOffset1 + depth1
                    dFace3toCenter = parSeat3toCenter - CptOffset3 + depth3
                    dFacetoFace = parSeattoSeat - (CptOffset1 + CptOffset2) + (depth1 + depth2)
                    If CmpDblEqual(parSeat3toCenter, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parTangentLength3, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) Then
                        dElbowRadius = parBranchRadius
                        dFace3toCenter = parTangentLength3 + parBranchRadius
                    ElseIf CmpDblEqual(parTangentLength3, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parSeat3toCenter, parBranchRadius) Then
                        dElbowRadius = parBranchRadius
                    Else
                        'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                        If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                            dElbowRadius = 0.8 * dFace3toCenter
                        Else
                            dElbowRadius = 0.8 * dFace1toCenter
                        End If
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 601    'Sanitary T branches, double, specified by header face-to-face dimension, face-to-center 1, face-to-center 3, branch radius, and tangent length 3
                    dFace1toCenter = parFace1toCenter
                    dFace3toCenter = parFace3toCenter
                    dFacetoFace = parFacetoFace
                    If CmpDblEqual(parFace3toCenter, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parTangentLength3, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) Then
                        dElbowRadius = parBranchRadius
                        dFace3toCenter = parTangentLength3 + parBranchRadius
                    ElseIf CmpDblEqual(parTangentLength3, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parFace3toCenter, parBranchRadius) Then
                        dElbowRadius = parBranchRadius
                    Else
                        'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                        If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                            dElbowRadius = 0.8 * dFace3toCenter
                        Else
                            dElbowRadius = 0.8 * dFace1toCenter
                        End If
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 602    'Sanitary T branches, double, specified by face-to-center 1, face-to-center 2, face-to-center 3, branch radius, and tangent length 3
                    dFace1toCenter = parFace1toCenter
                    dFace3toCenter = parFace3toCenter
                    dFacetoFace = parFace1toCenter + parFace2toCenter
                    If CmpDblEqual(parFace3toCenter, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parTangentLength3, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) Then
                        dElbowRadius = parBranchRadius
                        dFace3toCenter = parTangentLength3 + parBranchRadius
                    ElseIf CmpDblEqual(parTangentLength3, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parFace3toCenter, parBranchRadius) Then
                        dElbowRadius = parBranchRadius
                    Else
                        'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                        If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                            dElbowRadius = 0.8 * dFace3toCenter
                        Else
                            dElbowRadius = 0.8 * dFace1toCenter
                        End If
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 603    'Sanitary T branches, double, specified by header face-to-face dimension, tangent length 1, branch radius, face-to-center 3, and tangent length 3
                    dFace1toCenter = parTangentLength1 + parBranchRadius
                    dFace3toCenter = parFace3toCenter
                    dFacetoFace = parFacetoFace
                    If CmpDblEqual(parFace3toCenter, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parTangentLength3, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) Then
                        dElbowRadius = parBranchRadius
                        dFace3toCenter = parTangentLength3 + parBranchRadius
                    ElseIf CmpDblEqual(parTangentLength3, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parFace3toCenter, parBranchRadius) Then
                        dElbowRadius = parBranchRadius
                    Else
                        'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                        If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                            dElbowRadius = 0.8 * dFace3toCenter
                        Else
                            dElbowRadius = 0.8 * dFace1toCenter
                        End If
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 604    'Sanitary T branches, double, specified by tangent length 1, branch radius, face-to-center 2, face-to-center 3, and tangent length 3
                    dFace1toCenter = parTangentLength1 + parBranchRadius
                    dFace3toCenter = parFace3toCenter
                    dFacetoFace = dFace1toCenter + parFace2toCenter
                    If CmpDblEqual(parFace3toCenter, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parTangentLength3, LINEAR_TOLERANCE) And _
                               CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) Then
                        dElbowRadius = parBranchRadius
                        dFace3toCenter = parTangentLength3 + parBranchRadius
                    ElseIf CmpDblEqual(parTangentLength3, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parFace3toCenter, parBranchRadius) Then
                        dElbowRadius = parBranchRadius
                    Else
                        'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                        If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                            dElbowRadius = 0.8 * dFace3toCenter
                        Else
                            dElbowRadius = 0.8 * dFace1toCenter
                        End If
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
                Case 605    'Sanitary T branches, double, specified by header seat-to-seat dimension, seat-to-center 1, face-to-center 3, and branch radius
                    dFace1toCenter = parSeat1toCenter - CptOffset1 + depth1
                    dFace3toCenter = parFace3toCenter
                    dFacetoFace = parSeattoSeat - (CptOffset1 + CptOffset2) + (depth1 + depth2)
                    If CmpDblGreaterthan(parBranchRadius, LINEAR_TOLERANCE) And _
                                CmpDblGreaterthan(parFace3toCenter, parBranchRadius) Then
                        dElbowRadius = parBranchRadius
                    Else
                        'Elbow radius = 80% of maximum of parFace1toCenter and parFace3toCenter
                        If CmpDblGreaterthanOrEqualTo(dFace1toCenter, dFace3toCenter) Then
                            dElbowRadius = 0.8 * dFace3toCenter
                        Else
                            dElbowRadius = 0.8 * dFace1toCenter
                        End If
                    End If
                    dTangentLength1 = dFace1toCenter - dElbowRadius * Tan(dRevAngle / 2)
                    dTangentLength3 = dFace3toCenter - dElbowRadius * Tan(dRevAngle / 2)
            End Select
            
            If CmpDblEqual(dFace1toCenter, LINEAR_TOLERANCE) Or _
                CmpDblEqual(dFace3toCenter, LINEAR_TOLERANCE) Or _
                CmpDblEqual(dFacetoFace, LINEAR_TOLERANCE) Then
                GoTo ErrorLabel
            End If
            
            Dim dInsDia As Double
            'Create the Insulation of Port 1
            Dim dPortLen As Double
            If CmpDblGreaterthanOrEqualTo(pipeDiam1, flangeDiam1) Then
                dInsDia = pipeDiam1 + 2 * parInsulationThickness
            Else
                dInsDia = flangeDiam1 + 2 * parInsulationThickness
            End If
    
            If CmpDblEqual(flangeThick1, LINEAR_TOLERANCE) Then
                flangeThick1 = 0.0001
                dInsDia = pipeDiam1 + 2 * parInsulationThickness
                dInsThick = 0
            Else
                dInsThick = parInsulationThickness
            End If
    
            If CmpDblGreaterthan(dFacetoFace, flangeThick1 + dInsThick) Then
                dPortLen = flangeThick1 + dInsThick
            Else
                dPortLen = dFacetoFace
            End If
            
            oCenPoint.Set -dFace1toCenter * Sin(parAngle), dFace1toCenter * Cos(parAngle), 0
            oAxisVect.Set -Sin(parAngle), Cos(parAngle), 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                        oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                        oAxisVect.x, oAxisVect.y, oAxisVect.z, dInsDia / 2)
            
            oAxisVect.Set Sin(parAngle), -Cos(parAngle), 0
            Set objInsPort = PlaceProjection(m_OutputColl, oCircle, oAxisVect, dPortLen, True)
            
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
            Set objInsPort = Nothing
            
            'Create the Insulation for Body between Nozzle 1 and Nozzle 2
            Dim ObjMainRunInsul As Object
            dInsDia = pipeDiam1 + 2 * parInsulationThickness
            
            oCenPoint.Set -dFace1toCenter * Sin(parAngle), dFace1toCenter * Cos(parAngle), 0
            oAxisVect.Set -Sin(parAngle), Cos(parAngle), 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                        oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                        oAxisVect.x, oAxisVect.y, oAxisVect.z, dInsDia / 2)
            
            oAxisVect.Set Sin(parAngle), -Cos(parAngle), 0
            Set ObjMainRunInsul = PlaceProjection(m_OutputColl, oCircle, oAxisVect, dFacetoFace, True)
            
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainRunInsul
            Set ObjMainRunInsul = Nothing
            
            'Create the Insulation for Port 2
            If CmpDblGreaterthanOrEqualTo(pipeDiam2, flangeDiam2) Then
                dInsDia = pipeDiam2 + 2 * parInsulationThickness
            Else
                dInsDia = flangeDiam2 + 2 * parInsulationThickness
            End If
    
            If CmpDblEqual(flangeThick2, LINEAR_TOLERANCE) Then
                flangeThick2 = 0.0001
                dInsDia = pipeDiam2 + 2 * parInsulationThickness
                dInsThick = 0
            Else
                dInsThick = parInsulationThickness
            End If
    
            If CmpDblGreaterthan(dFacetoFace, flangeThick2 + dInsThick) Then
                dPortLen = flangeThick2 + dInsThick
            Else
                dPortLen = dFacetoFace
            End If
            oCenPoint.Set (dFacetoFace - dFace1toCenter) * Sin(parAngle), _
                               -(dFacetoFace - dFace1toCenter) * Cos(parAngle), 0
            oAxisVect.Set Sin(parAngle), -Cos(parAngle), 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                        oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                        oAxisVect.x, oAxisVect.y, oAxisVect.z, dInsDia / 2)
            
            oAxisVect.Set -Sin(parAngle), Cos(parAngle), 0
            Set objInsPort = PlaceProjection(m_OutputColl, oCircle, oAxisVect, dPortLen, True)
    
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
            Set objInsPort = Nothing
            
            'Create insulation for elbow 1
            Dim dArcStPointToCenter As Double
            Dim dInsulationRadius As Double
            
            Dim dArcCenX As Double
            Dim dArcCenY As Double
            
            dArcCenX = dFace1toCenter - dTangentLength1
            dArcCenY = dFace3toCenter - dTangentLength3 - dElbowRadius
            
            dInsulationRadius = pipeDiam3 / 2 + parInsulationThickness
    
            If CmpDblGreaterthan(dInsulationRadius, dElbowRadius) Then
                dInsulationRadius = dElbowRadius - 0.0001
            End If
    
            Dim objInsulatedBody1 As Object
            oCenPoint.Set -dArcCenX, dArcCenY, 0
            oAxisVect.Set 1, 0, 0
            
            Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                 oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                 oAxisVect.x, oAxisVect.y, oAxisVect.z, dInsulationRadius)
            
            dArcCenX = dFace1toCenter - dTangentLength1
            dArcCenY = dFace3toCenter - dTangentLength3
            
            oCenPoint.Set -dArcCenX, dArcCenY, 0
            oAxisVect.Set 0, 0, 1
            
            Set objInsulatedBody1 = PlaceRevolution(m_OutputColl, oCircle, _
                                                    oAxisVect, oCenPoint, dRevAngle, True)
            iOutput = 0
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody1
            Set objInsulatedBody1 = Nothing
    
            'Create insulation for port 3 body
            dInsulationRadius = pipeDiam3 / 2 + parInsulationThickness
            
            oStPoint.Set 0, dFace3toCenter, 0
            oEnPoint.Set 0, dFace3toCenter - dTangentLength3, 0
            
            Dim objNozzI1 As Object
            Set objNozzI1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                    2 * dInsulationRadius, True)
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzI1
            Set objNozzI1 = Nothing
    
            'Create the insulation for elbow 2
            Dim objInsulatedBody2 As Object
    
            dArcCenX = dFace1toCenter - dTangentLength1
            dArcCenY = dFace3toCenter - dTangentLength3 - dElbowRadius
            
            oCenPoint.Set -dArcCenX, -dArcCenY, 0
            oAxisVect.Set 1, 0, 0
            
            dInsulationRadius = pipeDiam4 / 2 + parInsulationThickness
    
            If CmpDblGreaterthan(dInsulationRadius, dElbowRadius) Then
                dInsulationRadius = dElbowRadius - 0.0001
            End If
    
            Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                 oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                 oAxisVect.x, oAxisVect.y, oAxisVect.z, dInsulationRadius)
            
            dArcCenX = dFace1toCenter - dTangentLength1
            dArcCenY = dFace3toCenter - dTangentLength3
            
            oCenPoint.Set -dArcCenX, -dArcCenY, 0
            oAxisVect.Set 0, 0, -1
    
            Set objInsulatedBody2 = PlaceRevolution(m_OutputColl, oCircle, _
                                                    oAxisVect, oCenPoint, dRevAngle, True)
    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsulatedBody2
            Set objInsulatedBody2 = Nothing
    
            'Create insulation for Port 4 body
            dInsulationRadius = pipeDiam4 / 2 + parInsulationThickness
            
            oStPoint.Set 0, -dFace3toCenter, 0
            oEnPoint.Set 0, -(dFace3toCenter - dTangentLength3), 0
    
            Dim objNozzI2 As Object
            Set objNozzI2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                2 * dInsulationRadius, True)
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzI2
            Set objNozzI2 = Nothing
            
            'Create the Insulation of Port 3
            If CmpDblGreaterthanOrEqualTo(pipeDiam3, flangeDiam3) Then
                dInsDia = pipeDiam3 + 2 * parInsulationThickness
            Else
                dInsDia = flangeDiam3 + 2 * parInsulationThickness
            End If
            
            If CmpDblEqual(flangeThick3, LINEAR_TOLERANCE) Then
                flangeThick3 = 0.0001
                dInsDia = pipeDiam3 + 2 * parInsulationThickness
                dInsThick = 0
            Else
                dInsThick = parInsulationThickness
            End If
            
            oStPoint.Set 0, dFace3toCenter, 0
            If CmpDblGreaterthan(2 * dFace3toCenter, Abs((dFace3toCenter - flangeThick3 - dInsThick))) Then
                oEnPoint.Set 0, (dFace3toCenter - flangeThick3 - dInsThick), 0
            Else
                oEnPoint.Set 0, -dFace3toCenter, 0
            End If
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
            Set objInsPort = Nothing
            
            'Create the Insulation for Port 4
            If CmpDblGreaterthanOrEqualTo(pipeDiam4, flangeDiam4) Then
                dInsDia = pipeDiam4 + 2 * parInsulationThickness
            Else
                dInsDia = flangeDiam4 + 2 * parInsulationThickness
            End If
            
            If CmpDblEqual(flangeThick4, LINEAR_TOLERANCE) Then
                flangeThick4 = 0.0001
                dInsDia = pipeDiam4 + 2 * parInsulationThickness
                dInsThick = 0
            Else
                dInsThick = parInsulationThickness
            End If
            
            oStPoint.Set 0, -dFace3toCenter, 0
            If CmpDblGreaterthan(2 * dFace3toCenter, Abs((dFace3toCenter - flangeThick3 - dInsThick))) Then
                oEnPoint.Set 0, -(dFace3toCenter - flangeThick3 - dInsThick), 0
            Else
                oEnPoint.Set 0, dFace3toCenter, 0
            End If
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            
            ' Set the Output
            m_OutputColl.AddOutput "InsPort4", objInsPort
            Set objInsPort = Nothing
            
            'Create the side opening(s) for the PDB 594, 595 and 596
            Dim objInsOpening As Object
            Dim dPlacePtX As Double     'X - Co-ordinate of the Nozzle 5
            Dim dPlacePtZ As Double     'Z - Co-ordinate of the Nozzle 5
            Dim dPortAngle As Double    'Port orientation w.r.t Z axis
            Dim iCount As Integer
            Dim pipeDiam5 As Double
            Dim flangeThick5 As Double
            Dim flangeDiam5 As Double
            Dim depth5 As Double
            Dim CptOffset5 As Double
            If lPartDataBasis = 594 Or lPartDataBasis = 595 Or lPartDataBasis = 596 Then
                            'Sanitary cross with side opening, specified by header seat-to-seat dimension, seat-to-center 1, seat-to-center 3, offset for side opening (594)
                            'Sanitary cross with two 90o side openings above center, specified by header face-to-face dimension, face-to-center 1, horizontal offset for opening above center, and vertical offset for opening above center (595)
                            'Sanitary cross with two 45o openings on same side, specified by header face-to-face dimension, face-to-center 1, face-to-center for side opening, horizontal offset for 45o side opening (596)
                
                'Determine Face1toOpening
                If Not CmpDblEqual(parSeat1toOpening, LINEAR_TOLERANCE) Then
                    dFace1toOpening = parSeat1toOpening - CptOffset1 + depth1
                ElseIf Not CmpDblEqual(parPorttoPortOffset, LINEAR_TOLERANCE) Then
                    dFace1toOpening = parPorttoPortOffset
                Else
                    GoTo ErrorLabel
                End If
                
                dPlacePtX = dFace1toOpening - dFace1toCenter
                dPortAngle = PI / 4
                If lPartDataBasis = 594 Then
                    dPlacePtZ = pipeDiam1 / 2
                Else
                    dPlacePtZ = parFace5toCenter
                End If
                
                Dim dTempPipeDia As Double
                Dim dTempFlangeDia As Double
                
                For iCount = 0 To 1
                    RetrieveParameters 5 + iCount, oPartFclt, m_OutputColl, dTempPipeDia, flangeThick, _
                                        dTempFlangeDia, sptOffset, depth
                    'Calculate the Insulation Diameter
                    If CmpDblGreaterthanOrEqualTo(dTempPipeDia, dTempFlangeDia) Then
                        dInsDia = dTempPipeDia + 2 * parInsulationThickness
                    Else
                        dInsDia = dTempFlangeDia + 2 * parInsulationThickness
                    End If
                    If lPartDataBasis = 594 Or lPartDataBasis = 595 Then
                        oStPoint.Set dPlacePtX, 0, ((-1) ^ iCount) * dPlacePtZ
                        oEnPoint.Set dPlacePtX, 0, 0
                    Else
                        oStPoint.Set dPlacePtX, ((-1) ^ iCount) * dPlacePtZ * Sin(dPortAngle), _
                                            dPlacePtZ * Cos(dPortAngle)
                        oEnPoint.Set dPlacePtX, 0, 0
                    End If
                    Set objInsOpening = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
                    
                    'Set the Output
                    m_OutputColl.AddOutput "InsSideOpening" & iCount + 1, objInsOpening
                    Set objInsOpening = Nothing
                    
                    If lPartDataBasis = 594 Then    'For PDB 594 there is only one opening
                        Exit For
                    End If
                Next iCount
            End If
            Set oGeomFactory = Nothing
            Set oStPoint = Nothing
            Set oEnPoint = Nothing
            Set oCenPoint = Nothing
            Set oAxisVect = Nothing
            Set oCircle = Nothing
        
        Case Else
            GoTo ErrorLabel
    End Select
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
