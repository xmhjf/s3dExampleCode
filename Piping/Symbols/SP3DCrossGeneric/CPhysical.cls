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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
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
'   01.Aug.2006     Kishore(kkkottur)       TR: 80254 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   11.Jan.2008     RUK                     CR-127570  Create reducing sanitary cross symbol
'   21.Jan.2008     RUK                     CR-127569  Create sanitary cross symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
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
    Dim ObjMainNozwithLen As Object
    Dim ObjTaperCone As Object
    Dim ObjBranchNozLen As Object
    Dim ObjBranchNozLen1 As Object
    Dim objNozzle As Object
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    Dim parInsulationThickness As Double

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parFace2toCenter = arrayOfInputs(3)
'    parFace3toCenter = arrayOfInputs(4)
'    parFace4toCenter = arrayOfInputs(5)
'    parInsulationThickness = arrayOfInputs(6)

    iOutput = 0

    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing

    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Select Case lPartDataBasis
        Case Is <= 1
            parFace1toCenter = arrayOfInputs(2)
            parFace2toCenter = arrayOfInputs(3)
            parFace3toCenter = arrayOfInputs(4)
            parFace4toCenter = arrayOfInputs(5)
            parInsulationThickness = arrayOfInputs(6)
            ' Insert your code for output 1 (Main run Nozzle with Length)  Place Nozzle 1
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            Dim Port1Length As Double
            Port1Length = parFace1toCenter
            oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
            oDir.Set -1, 0, 0
            Set ObjMainNozwithLen = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMainNozwithLen
            Set ObjMainNozwithLen = Nothing
    
            ' Place Nozzle 2  Insert your code for output 2(Nozzle in Main Line)
            Dim pipeDiam2 As Double
            Dim flangeThick2 As Double
            Dim sptOffset2 As Double
            Dim flangeDiam2 As Double
            Dim depth2 As Double
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
            oPlacePoint.Set parFace2toCenter + sptOffset2 - depth2, 0, 0
            oDir.Set 1, 0, 0
            Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
    
            ' Insert your code for output 2(Main run Reducer)
            Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
            Dim objCircle1 As IngrGeom3D.Circle3d, objCircle2 As IngrGeom3D.Circle3d
            Dim cir2CenterX As Double, cir2CenterY As Double, cir2CenterZ As Double
            Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
            Dim geomFactory As New IngrGeom3D.GeometryFactory
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
            cir2CenterX = parFace2toCenter - flangeThick2
            cir2CenterY = 0
            cir2CenterZ = 0
            Set objCircle2 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                             Nothing, cir2CenterX, _
                             cir2CenterY, cir2CenterZ, cirNormalX, _
                             cirNormalY, cirNormalZ, pipeDiam2 * 0.5)
            'Create Ruled Surface for body
            Set ObjTaperCone = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                                          objCircle1, objCircle2, True)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTaperCone
            Set ObjTaperCone = Nothing
    
            ' Insert your code for output 3(Branch run Nozzle with Length)
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
            oPlacePoint.Set 0, parFace3toCenter + sptOffset - depth, 0
            oDir.Set 0, 1, 0
            Dim Port3Length As Double
            Port3Length = parFace3toCenter
            Set ObjBranchNozLen = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port3Length)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchNozLen
            Set ObjBranchNozLen = Nothing
    
            ' Insert your code for output 4(Branch run Nozzle with Length)
            RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
            oPlacePoint.Set 0, -parFace4toCenter - sptOffset + depth, 0
            Dim Port4Length As Double
            Port4Length = parFace4toCenter
            oDir.Set 0, -1, 0
            Set ObjBranchNozLen1 = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port4Length)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBranchNozLen1
            Set ObjBranchNozLen1 = Nothing
        
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
            parFace2toCenter = arrayOfInputs(3)
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
            Dim flangeDiam3 As Double
            Dim depth3 As Double
            Dim CptOffset3 As Double
            
            Dim pipeDiam4 As Double
            Dim flangeThick4 As Double
            Dim flangeDiam4 As Double
            Dim depth4 As Double
            Dim CptOffset4 As Double
            
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
            
            'Retrieve the Nozzles properties
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, CptOffset1, depth1
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, CptOffset2, depth2
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick, flangeDiam, CptOffset3, depth3
            RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick, flangeDiam, CptOffset4, depth4
                        
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
            
            Set oPlacePoint = New AutoMath.DPosition
            Set oDir = New AutoMath.DVector
            
            'Create Nozzle 1
            oPlacePoint.Set -(dFace1toCenter + CptOffset1 - depth1) * Sin(parAngle), _
                                    (dFace1toCenter - CptOffset1 + depth1) * Cos(parAngle), 0
            Port1Length = dFacetoFace - flangeThick2
            oDir.Set -Sin(parAngle), Cos(parAngle), 0
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, _
                                                   Port1Length)
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
    
            'Create Nozzle 2 (without Length)
            oPlacePoint.Set (dFacetoFace - dFace1toCenter + CptOffset2 - depth2) * Sin(parAngle), _
                            -(dFacetoFace - dFace1toCenter + sptOffset2 - depth2) * Cos(parAngle), 0
            oDir.Set Sin(parAngle), -Cos(parAngle), 0
            Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
    
            'Create Elbow 1
            Dim dArcStPointToCenter As Double
            Dim dArcCenX As Double
            Dim dArcCenY As Double
            
            dArcCenX = dFace1toCenter - dTangentLength1
            dArcCenY = dFace3toCenter - dTangentLength3 - dElbowRadius
            
            Dim oCircle As Object
            Dim oCenPoint As AutoMath.DPosition
            Dim oAxisVect As AutoMath.DVector
            Dim objElbow1 As Object
            
            Set oCenPoint = New AutoMath.DPosition
            Set oAxisVect = New AutoMath.DVector
        
            oCenPoint.Set -dArcCenX, dArcCenY, 0
            oAxisVect.Set 1, 0, 0
            
            Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                 oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                 oAxisVect.x, oAxisVect.y, oAxisVect.z, pipeDiam3 / 2)
            
            
            dArcCenX = dFace1toCenter - dTangentLength1
            dArcCenY = dFace3toCenter - dTangentLength3
            
            oCenPoint.Set -dArcCenX, dArcCenY, 0
            oAxisVect.Set 0, 0, 1
            
            Set objElbow1 = PlaceRevolution(m_OutputColl, oCircle, oAxisVect, oCenPoint, _
                                                                dRevAngle, True)
    
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objElbow1
            Set objElbow1 = Nothing
    
            'Create Nozzle 3
            oPlacePoint.Set 0, dFace3toCenter + CptOffset3 - depth3, 0
            oDir.Set 0, 1, 0
            Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, _
                                                dTangentLength3)
            
            'Set the Output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
    
            'Create Elbow 2
            Dim objElbow2 As Object
            
            dArcCenX = dFace1toCenter - dTangentLength1
            dArcCenY = dFace3toCenter - dTangentLength3 - dElbowRadius
            
            oCenPoint.Set -dArcCenX, -dArcCenY, 0
            oAxisVect.Set 1, 0, 0
            
            Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                 oCenPoint.x, oCenPoint.y, oCenPoint.z, _
                                 oAxisVect.x, oAxisVect.y, oAxisVect.z, pipeDiam3 / 2)
            
            dArcCenX = dFace1toCenter - dTangentLength1
            dArcCenY = dFace3toCenter - dTangentLength3
            
            oCenPoint.Set -dArcCenX, -dArcCenY, 0
            oAxisVect.Set 0, 0, -1
    
            Set objElbow2 = PlaceRevolution(m_OutputColl, oCircle, oAxisVect, oCenPoint, _
                                            dRevAngle, True)
    
            'Set the outputs
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objElbow2
            Set objElbow2 = Nothing
    
            'Create Nozzle 4
            oPlacePoint.Set 0, -(dFace3toCenter + CptOffset4 - depth4), 0
            oDir.Set 0, -1, 0
            Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, oPlacePoint, _
                                                   dTangentLength3)
            'Set the Output
            m_OutputColl.AddOutput "Nozzle4", objNozzle
            Set objNozzle = Nothing
            
            'Create the side opening(s) for the PDB 594, 595 and 596
            Dim dPlacePtX As Double     'X - Co-ordinate of the Nozzle 5 and 6
            Dim dPlacePtZ As Double     'Z - Co-ordinate of the Nozzle 5 and 6
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
                
                Dim dTempCpt As Double
                Dim dTempDepth As Double
                
                For iCount = 0 To 1
                    RetrieveParameters 5 + iCount, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                        flangeDiam, dTempCpt, dTempDepth
                    If lPartDataBasis = 594 Or lPartDataBasis = 595 Then
                        oPlacePoint.Set dPlacePtX, 0, ((-1) ^ iCount) * _
                                                            (dPlacePtZ + dTempCpt - dTempDepth)
                        oDir.Set 0, 0, (-1) ^ iCount
                    Else
                        oPlacePoint.Set dPlacePtX, _
                                ((-1) ^ iCount) * (dPlacePtZ + dTempCpt - dTempDepth) * Sin(dPortAngle), _
                                    (dPlacePtZ + dTempCpt - dTempDepth) * Cos(dPortAngle)
                        oDir.Set 0, ((-1) ^ iCount) * Sin(dPortAngle), Cos(dPortAngle)
                    End If
                    Set objNozzle = CreateNozzleWithLength(5 + iCount, oPartFclt, m_OutputColl, _
                                                    oDir, oPlacePoint, dPlacePtZ)
                    
                    'Set the Output
                    m_OutputColl.AddOutput "SideOpening" & iCount + 1, objNozzle
                    Set objNozzle = Nothing
                    
                    If lPartDataBasis = 594 Then 'For PDB 594 there is only one opening
                        Exit For
                    End If
                Next iCount
            End If
    
            Set oCircle = Nothing
            Set oAxisVect = Nothing
            Set oCenPoint = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing
        
        Case Else
            GoTo ErrorLabel
    End Select

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' This is a sample of placing cylinder and 2 nozzles
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' Use nozzle 1 as pipe port for flange and pipe parameters
'''    Dim pipeDiam        As Double
'''    Dim flangeThick     As Double
'''    Dim sptOffset       As Double
'''    Dim flangeDiam      As Double
'''    Dim depth           As Double
'''    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
'''    Dim parOperatorWidth    As Double
'''    parOperatorWidth = arrayOfInputs(2)
'''    Dim stPoint   As New AutoMath.DPosition
'''    Dim enPoint   As New AutoMath.DPosition
'''    stPoint.Set -parOperatorWidth / 2 + flangeThick, 0, 0
'''    enPoint.Set parOperatorWidth / 2 - flangeThick, 0, 0
'''    Dim objCylinder As Object
'''
'''    Set objCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinder
'''    Set objCylinder = Nothing
'''
'''    Dim oPlacePoint As AutoMath.DPosition
'''    Dim oDir        As AutoMath.DVector
'''    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
'''
'''    Set oPlacePoint = New AutoMath.DPosition
'''    Set oDir = New AutoMath.DVector
'''    oPlacePoint.Set -parOperatorWidth / 2 - sptOffset + depth, 0, 0
'''    oDir.Set -1, 0, 0
'''    Set oPartFclt = arrayOfInputs(1)
'''    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''
'''    oPlacePoint.Set parOperatorWidth / 2 + sptOffset - depth, 0, 0
'''    oDir.Set 1, 0, 0
'''
'''    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
'''
'''' Set the output
'''    iOutput = iOutput + 1
'''    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
'''    Set objNozzle = Nothing
'''


