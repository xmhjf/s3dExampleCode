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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Monday, Feb 04, 2008
'   Description:
'       This is Open/Blind spcer symbol. This is prepared based Appendex E-94 in Piping Design Document L57.
'       CR-134984  Provide symbol for open spacer and blind spacer set
'       Implemented the following PDBs
'               Blind spacer and open spacer set, specified by spacer length(758)
'               Blind spacer and open spacer set, specified by handle length from flow centerline(759)
'               Blind spacer and open spacer set, specified by handle length from spacer outside diameter(760)
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   04.Feb.2008     RUK     CR-134984  Provide symbol for open spacer and blind spacer set
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim parInstallationState As Double
    Dim parFacetoFace As Double
    Dim parSpacerLength As Double
    Dim parSpacerOutsideDia As Double
    Dim parSpacerInsideDia As Double
    Dim parHandleLength As Double
    Dim parHandleWidth As Double
    Dim parHandleThickness As Double
    Dim parHandleRadius As Double
    Dim parHoleDiameter As Double
    Dim parInsulationThickness As Double
    
    Set oPartFclt = arrayOfInputs(1)
    parInstallationState = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parSpacerLength = arrayOfInputs(4)
    parSpacerOutsideDia = arrayOfInputs(5)
    parSpacerInsideDia = arrayOfInputs(6)
    parHandleLength = arrayOfInputs(7)
    parHandleWidth = arrayOfInputs(8)
    parHandleThickness = arrayOfInputs(9)
    parHandleRadius = arrayOfInputs(10)
    parHoleDiameter = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)
    
    'Origin is at the center of inlet and outlet nozzles
    Dim iOutput As Long
    iOutput = 0
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oCenter As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    
    Set oGeomFactory = New GeometryFactory
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oCenter = New DPosition
    Set oVector = New DVector
    
    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    'Retrieve the nozzle parameters
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        cptOffset, depth
    'Compute the values for Optional parameters
    'Compute the value for Spacer Inside Diameter (If value is zero).
    If CmpDblEqual(parSpacerInsideDia, LINEAR_TOLERANCE) Then
        parSpacerInsideDia = pipeDiam
    End If
    
    'Compute the value for Handle Thickness (If value is zero).
    If CmpDblEqual(parHandleThickness, LINEAR_TOLERANCE) Then
        parHandleThickness = 0.5 * parFacetoFace
    End If
    
    'The Handle Thickness shall be 6.4mm minimum, except when Face to Face
    'is less than 6.4mm. If Face to Face is less than 6.4mm then
    'Handle Thickness equals to Face to Face.
    If CmpDblGreaterthan(parFacetoFace, 6.4 / 1000) And _
                            CmpDblLessThan(parHandleThickness, 6.4 / 1000) Then
       parHandleThickness = 6.4 / 1000
    ElseIf CmpDblLessThan(parFacetoFace, 6.4 / 1000) Then
        parHandleThickness = parFacetoFace
    End If
    
    'Compute the value for Handle Radius (If value is zero).
    If CmpDblEqual(parHandleRadius, LINEAR_TOLERANCE) Then
        parHandleRadius = parHandleThickness / 2
    End If
    
    'Compute the value for Hole Diameter (If value is zero).
    If CmpDblEqual(parHoleDiameter, LINEAR_TOLERANCE) Or _
            CmpDblGreaterthanOrEqualTo(parHoleDiameter, parHandleWidth) Then
        parHoleDiameter = 0.75 * parHandleWidth
        'Hole Radius should be less tham Handle Radius.
        If CmpDblGreaterthanOrEqualTo(parHoleDiameter / 2, parHandleRadius) Then
            parHoleDiameter = 2 * Abs((parHandleRadius - 0.0001))
        End If
    End If
    
    'Create the Handle
    Dim objHandle As Object
    Dim Surfset As IngrGeom3D.IJElements
    Dim oCompStr As IngrGeom3D.ComplexString3d
    Dim oCollection As Collection
    Dim oLine As IngrGeom3D.Line3d
    Dim oStArc As IngrGeom3D.Arc3d
    Dim oEnArc As IngrGeom3D.EllipticalArc3d
    Dim oPlane As IngrGeom3D.Plane3d
    
    Dim dHandleLength As Double
    Dim stnorm() As Double
    Dim ednorm() As Double
    Dim iCount As Integer
    
    Set oPlane = New Plane3d
    Set oCompStr = New ComplexString3d
    Set oCollection = New Collection
    Set oLine = New Line3d
    Set oStArc = New Arc3d
    Set oEnArc = New EllipticalArc3d
    
    'Determine the Handle Length
    Select Case lPartDataBasis
        Case Is <= 1 'default
            dHandleLength = parSpacerLength - parSpacerOutsideDia
        Case 758  'Blind spacer and open spacer set, specified by spacer length (758)
            dHandleLength = parSpacerLength - parSpacerOutsideDia
        Case 759    'Blind spacer and open spacer set, specified by handle length from flow centerline  (759)
            dHandleLength = parHandleLength - parSpacerOutsideDia / 2
        Case 760   'Blind spacer and open spacer set, specified by handle length from spacer outside diameter (760)
            dHandleLength = parHandleLength
        Case Else
            GoTo ErrorLabel
    End Select
    
    Dim dTempAng As Double 'Angle made by the handle top point at the center of the spacer
    dTempAng = Atn((parHandleWidth / 2) / _
                (Sqr((parSpacerOutsideDia / 2) ^ 2 - (parHandleWidth / 2) ^ 2)))
    
    'Line 1
    oStPoint.Set parHandleThickness / 2, (parSpacerOutsideDia / 2) * Sin(dTempAng), _
                                (parSpacerOutsideDia / 2) * Cos(dTempAng)
    oEnPoint.Set oStPoint.x, oStPoint.y, _
                                parSpacerOutsideDia / 2 + dHandleLength - parHandleRadius
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine
    
    'Elliptical Arc
    Dim dMMRatio As Double
    oCenter.Set parHandleThickness / 2, 0, parSpacerOutsideDia / 2 + dHandleLength - parHandleRadius
    oVector.Set 1, 0, 0
    oStPoint.Set 0, parHandleRadius, 0
    dMMRatio = parHandleRadius / (parHandleWidth / 2)
    Set oEnArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                            oCenter.x, oCenter.y, oCenter.z, oVector.x, oVector.y, oVector.z, _
                            oStPoint.x, oStPoint.y, oStPoint.z, dMMRatio, 0, PI)
    oCollection.Add oEnArc
    Set oEnArc = Nothing
    
    'Line 2
    oStPoint.Set parHandleThickness / 2, -(parSpacerOutsideDia / 2) * Sin(dTempAng), _
                                parSpacerOutsideDia / 2 + dHandleLength - parHandleRadius
    oEnPoint.Set oStPoint.x, oStPoint.y, (parSpacerOutsideDia / 2) * Cos(dTempAng)
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    oCollection.Add oLine

    'Circular Arc
    oCenter.Set parHandleThickness / 2, 0, 0
    oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
    oEnPoint.Set parHandleThickness / 2, (parSpacerOutsideDia / 2) * Sin(dTempAng), _
                                (parSpacerOutsideDia / 2) * Cos(dTempAng)
    Set oStArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                oCenter.x, oCenter.y, oCenter.z, oStPoint.x, oStPoint.y, oStPoint.z, _
                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
    oCollection.Add oStArc
    Set oStArc = Nothing
    
    'Create the Complext string
    oStPoint.Set parHandleThickness / 2, (parSpacerOutsideDia / 2) * Sin(dTempAng), (parSpacerOutsideDia / 2) * Cos(dTempAng)
    Set oCompStr = PlaceTrCString(oStPoint, oCollection)
        
    For iCount = 1 To oCollection.Count
        oCollection.Remove 1
    Next iCount
    
    'Create the Plane by using above complex string as outer boundary
    Set oPlane = oGeomFactory.Planes3d.CreateByOuterBdry(Nothing, oCompStr)
    Set oCompStr = Nothing
    
    If parInstallationState <= 1 Or parInstallationState = OPEN_SPACER_INSTALLED Then
        Dim oHoleArc1 As IngrGeom3D.Arc3d
        Dim oHoleArc2 As IngrGeom3D.Arc3d
        Dim oHoleCompStr As IngrGeom3D.ComplexString3d
        
        Set oHoleArc1 = New Arc3d
        Set oHoleArc2 = New Arc3d
        Set oHoleCompStr = New ComplexString3d
        
        oCenter.Set parHandleThickness / 2, 0, _
                            parSpacerOutsideDia / 2 + dHandleLength - parHandleRadius
        oVector.Set 1, 0, 0
        
        oStPoint.Set oCenter.x, parHoleDiameter / 2, oCenter.z
        oEnPoint.Set oStPoint.x, -parHoleDiameter / 2, oStPoint.z
        Set oHoleArc1 = PlaceTrArcByCenterNorm(oStPoint, oEnPoint, oCenter, oVector)
        oCollection.Add oHoleArc1
        
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set oStPoint.x, parHoleDiameter / 2, oStPoint.z
        Set oHoleArc2 = PlaceTrArcByCenterNorm(oStPoint, oEnPoint, oCenter, oVector)
        oCollection.Add oHoleArc2
        
        oStPoint.Set oCenter.x, parHoleDiameter / 2, oCenter.z
        Set oHoleCompStr = PlaceTrCString(oStPoint, oCollection)
        
        oPlane.AddHole oHoleCompStr
        
        Set oHoleArc1 = Nothing
        Set oHoleArc2 = Nothing
        Set oHoleCompStr = Nothing
        For iCount = 1 To oCollection.Count
            oCollection.Remove 1
        Next iCount
    End If
    
    oStPoint.Set parHandleThickness / 2, (parSpacerOutsideDia / 2) * Sin(dTempAng), _
                                (parSpacerOutsideDia / 2) * Cos(dTempAng)
    oEnPoint.Set -parHandleThickness / 2, oStPoint.y, oStPoint.z
    Set oLine = PlaceTrLine(oStPoint, oEnPoint)
    
    Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                                        oLine, oPlane, CircularCorner, 0, stnorm, ednorm, True)
    For Each objHandle In Surfset
        m_OutputColl.AddOutput "Handle_", objHandle
    Next objHandle
    Set objHandle = Nothing
    For iCount = 1 To Surfset.Count
        Surfset.Remove 1
    Next iCount
    Set Surfset = Nothing
    Set oPlane = Nothing
    Set oLine = Nothing
    Set oCollection = Nothing
    
    'Create the Body
    Dim objBody As Object
    oStPoint.Set parFacetoFace / 2, 0, 0
    oEnPoint.Set -parFacetoFace / 2, 0, 0
    Set objBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parSpacerOutsideDia, True)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBody
    Set objBody = Nothing
        
    'Create Nozzles
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oPlacementPoint As AutoMath.DPosition
    Dim oDirVec As AutoMath.DVector
    
    Set oPlacementPoint = New DPosition
    Set oDirVec = New DVector
    
    'Create the Inlet Nozzle
    oPlacementPoint.Set -parFacetoFace / 2, 0, 0
    oDirVec.Set -1, 0, 0
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDirVec, _
                                                        oPlacementPoint, parFacetoFace)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    'Create Outlet Nozzle
    oPlacementPoint.Set parFacetoFace / 2, 0, 0
    oDirVec.Set 1, 0, 0
    
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDirVec, oPlacementPoint)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacementPoint = Nothing
    Set oDirVec = Nothing
    
    'Remove the References
    Set oGeomFactory = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oCenter = Nothing
    Set oVector = Nothing
    
    Exit Sub
    
ErrorLabel:
   ReportUnanticipatedError MODULE, METHOD
End Sub
