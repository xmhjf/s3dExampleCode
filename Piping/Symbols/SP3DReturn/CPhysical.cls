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
'   Author:          svsmylav
'   Creation Date:  Wednesday, Jan 2 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                  change description
'   -----------     -----                 ------------------
'   09.Jul.2003     SymbolTeam(India)     Copyright Information, Header  is added.
'   01.Aug.2006     Kishore(kkkottur)     TR: 80254 Replaced 'm_OutputColl.ResourceManager' with "Nothing'.

'   08.SEP.2006     KKC                   DI-95670  Replace names with initials in all revision history sheets and symbols
'   12.SEP.2007     VRG                   CR-126671 Modified symbol code in order to support Face to Center and Diameter dimension basis
'   12.AUG.2008     PK                    CR-145847 Modified the symbol code to support the partdatabasis values 1055,1056,1057,1058,1059,1060,1061,1062 and 1063
'   09.Aug.2008      MP                   TR-149518 Unable to route out of other end of the 180 degree return component (Modified the position of Port2)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
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
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput             As Integer
    Dim ObjSP3DReturn       As Object
    Dim ObjInsulatedBody    As Object
    
    Dim parFacetoCenter         As Double
    Dim parInsulationThickness  As Double
    Dim parDiameter             As Double
    Dim dElbowRadius            As Double
    
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
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 dElbowRadius = pipeDiam / 2 + NEGLIGIBLE_THICKNESS
            End If
            
        Case Is <= 1, 413   'Default or 180 degree return, open, specified by back-centerline-to-face
            parFacetoCenter = arrayOfInputs(2)
            dElbowRadius = parFacetoCenter - flangeThick
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 dElbowRadius = pipeDiam / 2 + NEGLIGIBLE_THICKNESS
            End If
            
        
        Case 1055   '180 degree return, closed, specified by center-to-center
            parCentertoCenter = arrayOfInputs(5)
            dElbowRadius = parCentertoCenter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 pipeDiam = dElbowRadius
            End If
    
        Case 1056   '180 degree return, open, specified by center-to-center
            parCentertoCenter = arrayOfInputs(5)
            dElbowRadius = parCentertoCenter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 dElbowRadius = pipeDiam / 2 + NEGLIGIBLE_THICKNESS
            End If
    
        Case 1057   '180 degree return, open, specified by back-to-face
            parBacktoFace = arrayOfInputs(6)
            dElbowRadius = parBacktoFace - pipeDiam / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 dElbowRadius = pipeDiam / 2 + NEGLIGIBLE_THICKNESS
            End If
                    
        Case 1058   '180 degree return, open, single tangent, specified by center-to-center and tangent, port 1
            parCentertoCenter = arrayOfInputs(5)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parCentertoCenter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 dElbowRadius = pipeDiam / 2 + NEGLIGIBLE_THICKNESS
            End If
            
        Case 1059   '180 degree return, open, single tangent, specified by back-to-tangent-line and tangent, port 1
            parBacktoTangentLine = arrayOfInputs(7)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parBacktoTangentLine - pipeDiam / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 dElbowRadius = pipeDiam / 2 + NEGLIGIBLE_THICKNESS
            End If
            
        Case 1060   '180 degree return, open, double tangent, specified by center-to-center and tangent
            parCentertoCenter = arrayOfInputs(5)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parCentertoCenter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 dElbowRadius = pipeDiam / 2 + NEGLIGIBLE_THICKNESS
            End If
            
        Case 1061   '180 degree return, open, double tangent, specified by back-to-tangent-line and tangent
            parBacktoTangentLine = arrayOfInputs(7)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parBacktoTangentLine - pipeDiam / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 dElbowRadius = pipeDiam / 2 + NEGLIGIBLE_THICKNESS
            End If
            
        Case 1062   '180 degree return, open, single tangent, specified by center-to-center and tangent, port 2
            parCentertoCenter = arrayOfInputs(5)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parCentertoCenter / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 dElbowRadius = pipeDiam / 2 + NEGLIGIBLE_THICKNESS
            End If
            
        Case 1063   '180 degree return, open, single tangent, specified by back-to-face and tangent, port 2
            parBacktoFace = arrayOfInputs(6)
            parTangentLength = arrayOfInputs(8)
            dElbowRadius = parBacktoFace - pipeDiam / 2
            If CmpDblLessThanOrEqualTo(dElbowRadius, pipeDiam / 2) Then
                 dElbowRadius = pipeDiam / 2 + NEGLIGIBLE_THICKNESS
            End If
    End Select
    
    Dim objCirc1 As IngrGeom3D.Circle3d
    Dim geomFactory As New IngrGeom3D.GeometryFactory
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
  
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                      Nothing, cir1CenterX, _
                                      cir1CenterY, cir1CenterZ, cirNormalX, _
                                      cirNormalY, cirNormalZ, pipeDiam / 2)

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
    
    Set ObjSP3DReturn = PlaceRevolution(m_OutputColl, objCirc1, _
                                    axis, centPoint, PI, _
                                    True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSP3DReturn
    Set ObjSP3DReturn = Nothing
    
    If lPartDataBasis = 1055 Then
        Dim oEndCurve   As Object
        Dim oProjection As IngrGeom3D.Projection3d
        Dim oCurves     As Collection
        Dim oLine       As Object
        Dim oArc        As IngrGeom3D.Arc3d
        Dim oStPoint    As AutoMath.DPosition
        Dim oVec        As AutoMath.DVector


        Set oStPoint = New AutoMath.DPosition
        Set oCurves = New Collection
        Set oVec = New AutoMath.DVector
        
        Set oArc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                            CenterPos.x, CenterPos.y + dElbowRadius, CenterPos.z - flangeDiam / 2, _
                            CenterPos.x, CenterPos.y + dElbowRadius + flangeDiam / 2, CenterPos.z, _
                            CenterPos.x, CenterPos.y + dElbowRadius, CenterPos.z + flangeDiam / 2)
        oCurves.Add oArc
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            CenterPos.x, CenterPos.y + dElbowRadius, CenterPos.z + flangeDiam / 2, _
                            CenterPos.x, CenterPos.y - dElbowRadius, CenterPos.z + flangeDiam / 2)
        oCurves.Add oLine
        Set oArc = geomFactory.Arcs3d.CreateBy3Points(Nothing, _
                            CenterPos.x, CenterPos.y - dElbowRadius, CenterPos.z + flangeDiam / 2, _
                            CenterPos.x, CenterPos.y - dElbowRadius - flangeDiam / 2, CenterPos.z, _
                            CenterPos.x, CenterPos.y - dElbowRadius, CenterPos.z - flangeDiam / 2)
        oCurves.Add oArc
        Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                            CenterPos.x, CenterPos.y - dElbowRadius, CenterPos.z - flangeDiam / 2, _
                            CenterPos.x, CenterPos.y + dElbowRadius, CenterPos.z - flangeDiam / 2)
        oCurves.Add oLine

        oStPoint.Set CenterPos.x, CenterPos.y + dElbowRadius, CenterPos.z - flangeDiam / 2
        oVec.Set 1, 0, 0
        Set oEndCurve = PlaceTrCString(oStPoint, oCurves)
        
        Set oProjection = PlaceProjection(m_OutputColl, oEndCurve, oVec, flangeThick, True)
        
        m_OutputColl.AddOutput "Projection", oProjection
        
        Set oEndCurve = Nothing
        Set oProjection = Nothing
        Set oCurves = Nothing
        Set oLine = Nothing
        Set oArc = Nothing
        Set oStPoint = Nothing
        Set oVec = Nothing
        
    End If
    
    'Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    'Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
    'faceToFace = arrayOfInputs(2)
    
    Select Case lPartDataBasis
        Case Is <= 1, 413, 414
            oPlacePoint.Set CenterPos.x - parFacetoCenter - sptOffset + depth, _
                         CenterPos.y, _
                         CenterPos.z
            Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
        
        Case 1055, 1056, 1057, 1062, 1063
            oPlacePoint.Set CenterPos.x - sptOffset + depth, _
                         CenterPos.y + dElbowRadius, _
                         CenterPos.z
            Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
            
        Case 1058, 1059, 1060, 1061
            oPlacePoint.Set CenterPos.x - parTangentLength - sptOffset + depth, _
                         CenterPos.y + dElbowRadius, _
                         CenterPos.z
            Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                         oPlacePoint, parTangentLength)
    End Select

    'Set oPartFclt = arrayOfInputs(1)
 
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oDir.Set -1, 0, 0
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 413, 414
            oPlacePoint.Set CenterPos.x - parFacetoCenter - sptOffset + depth, _
                    CenterPos.y + 2 * dElbowRadius, _
                    CenterPos.z
            Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
            
        Case 1055, 1056, 1057, 1058, 1059
            oPlacePoint.Set CenterPos.x - sptOffset + depth, _
                         CenterPos.y - dElbowRadius, _
                         CenterPos.z
            Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
               
        Case 1060, 1061, 1062, 1063
            oPlacePoint.Set CenterPos.x - parTangentLength, _
                        CenterPos.y - dElbowRadius, _
                        CenterPos.z
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, parTangentLength)
            
    End Select
       
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCirc1
    Set objCirc1 = Nothing
    sp3dElem1.Remove
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

