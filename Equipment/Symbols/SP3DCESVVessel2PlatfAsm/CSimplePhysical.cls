VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:01:08 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          MS
'   Creation Date:  Wednesday, Feb 5 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003  SymbolTeam(India)   Copyright Information, Header  is added.
'   22.Aug.2003  SymbolTeam(India)   TR 46728 Change Port Index of a nozzle to 15 used to corrupt Equipment.
'                                    Modified port index logic.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   28.Sep.2006    RRK  TR-84724  If variables of type double are compared, appropriate CmpDblXXX functions are used.
'   04.Oct.2006     dkl  TR-100087 Changed the direction of normal vector for creating default surface to -Z (from earlier Z)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private m_bAngle As Boolean
Private Const MODULE = "Physical:" 'Used for error messages
Private Const TOLERANCEVAL = 0.000001


Private Sub Class_Initialize()
        PI = 4 * Atn(1)
End Sub
Private Sub Class_Terminate()
''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    Dim oPartFclt       As PartFacelets.IJDPart

    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint

    Dim iOutput     As Double
    Dim ObjVessel As Object
    Dim ObjSupport As Object
    
    Dim parVesselHeight As Double
    Dim parVesselDiameter As Double
    Dim parSkirtTopDiameter As Double   'Skirt Top Diameter
    Dim parSkirtBottomDiameter As Double   'Skirt Bottom Diameter
    Dim parSkirtHeight As Double
    Dim parVesselStartHeight As Double
    Dim parInsulationThickness As Double

    'Declaration of Platform Input Parameters
    Dim parPlatformWidth As Double
    Dim parPlatformHeight As Double
    Dim parPlatformRadius As Double
    Dim parPlatformAngle As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parSkirtTopDiameter = arrayOfInputs(4)
    parSkirtBottomDiameter = arrayOfInputs(5)
    parSkirtHeight = arrayOfInputs(6)
    parVesselStartHeight = arrayOfInputs(7)
    parPlatformWidth = arrayOfInputs(8)
    parPlatformHeight = arrayOfInputs(9)
    parPlatformRadius = arrayOfInputs(10)
    parPlatformAngle = arrayOfInputs(11)
    
    iOutput = 0


    
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
 ' Insert your code for output 1(Vessel (Shell))
    Dim cenX As Double, cenY As Double, cenZ As Double
    Dim MajorX As Double, MajorY As Double
    Dim MajorZ As Double, mMRatio As Double, StartAngle As Double, SweepAngle As Double
    Dim norX As Double, norY As Double, norZ As Double
    
    Dim oEqpLine As IngrGeom3D.Line3d
    Dim oEqpArcTop As IngrGeom3D.EllipticalArc3d
    Dim oEqpArcBottom As IngrGeom3D.EllipticalArc3d
    Dim oEqpComplexStr As IngrGeom3D.ComplexString3d

' Vessel is formed by creating a Complex string and rotating about Z-Axis
'Top Convex Dome Curve
    cenX = 0
    cenY = 0
    cenZ = parVesselHeight + parVesselStartHeight
    MajorX = parVesselDiameter / 2
    MajorY = 0
    MajorZ = 0
    mMRatio = 0.5
    StartAngle = PI / 2
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEqpArcTop = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
                                                        
    'Vertical Line
    Dim Linepts(0 To 5) As Double
    Linepts(0) = parVesselDiameter / 2
    Linepts(1) = 0
    Linepts(2) = parVesselHeight + parVesselStartHeight
    Linepts(3) = parVesselDiameter / 2
    Linepts(4) = 0
    Linepts(5) = parVesselStartHeight

    Set oEqpLine = geomFactory.Lines3d.CreateBy2Points(Nothing, Linepts(0), Linepts(1), Linepts(2), _
                                                                        Linepts(3), Linepts(4), Linepts(5))
    'Bottom Dome Curve
    cenX = 0
    cenY = 0
    cenZ = parVesselStartHeight
    MajorX = parVesselDiameter / 2
    MajorY = 0
    MajorZ = 0
    mMRatio = 0.5
    StartAngle = PI
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0

    Set oEqpArcBottom = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
                                                            
    Dim EleCollection           As Collection
    Dim startPoint As New AutoMath.DPosition
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    
    Set EleCollection = New Collection
    EleCollection.Add oEqpArcTop
    EleCollection.Add oEqpLine
    EleCollection.Add oEqpArcBottom
    
    startPoint.Set 0, 0, (parVesselHeight + parVesselStartHeight + parVesselDiameter / 4)
    Set oEqpComplexStr = PlaceTrCString(startPoint, EleCollection)
    axisVect.Set 0, 0, 1
    revCenPt.Set 0, 0, (parVesselHeight / 2 + parVesselStartHeight)
    'Revolve it about Z-Axiz

    Set ObjVessel = PlaceRevolution(m_OutputColl, oEqpComplexStr, axisVect, revCenPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing
    Set axisVect = Nothing
    Set revCenPt = Nothing
    
    Dim Objcurves As IJDObject
    Set Objcurves = oEqpComplexStr
    Objcurves.Remove
    
    Set Objcurves = Nothing
    Set oEqpArcBottom = Nothing
    Set oEqpLine = Nothing
    Set oEqpArcTop = Nothing
    
    Dim count As Integer
    For count = 1 To EleCollection.count
        EleCollection.Remove 1
    Next count
    Set EleCollection = Nothing
    Set oEqpComplexStr = Nothing
    
' Insert your code for output 2(Truncated Cone Skirt)
    Dim baseCenter As New AutoMath.DPosition
    Dim topCenter As New AutoMath.DPosition
    Dim topRadius As Double, baseRadius As Double
    
    topCenter.Set 0, 0, parSkirtHeight
    baseCenter.Set 0, 0, 0
    baseRadius = parSkirtBottomDiameter / 2
    topRadius = parSkirtTopDiameter / 2
    Set ObjSupport = PlaceCone(m_OutputColl, baseCenter, topCenter, baseRadius, topRadius, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSupport
    Set ObjSupport = Nothing
    Set topCenter = Nothing
    
'' Insert your code for output 3(Default Surface )
    Dim ObjDefSurface As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oComplexStr As New IngrGeom3D.ComplexString3d
    Dim normalVect As New AutoMath.DVector
    
    normalVect.Set 0, 0, -1
    Set oCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    baseCenter.x, baseCenter.y, baseCenter.z, _
                                                    normalVect.x, normalVect.y, normalVect.z, _
                                                    parSkirtBottomDiameter / 2)
    Call oComplexStr.AddCurve(oCircle, False)
    Set ObjDefSurface = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                                        baseCenter.x, baseCenter.y, baseCenter.z, _
                                                        normalVect.x, normalVect.y, normalVect.z)
    Call ObjDefSurface.AddBoundary(oComplexStr)
'Set  the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefSurface
    Set ObjDefSurface = Nothing
    
    Dim ObjTmpcurves As IJDObject
    Set ObjTmpcurves = oComplexStr
    ObjTmpcurves.Remove
    Set oCircle = Nothing
    Set oComplexStr = Nothing
    Set normalVect = Nothing
    Set baseCenter = Nothing
     
'Creation of Bottom Platform
Dim oComplexString     As IngrGeom3D.ComplexString3d
Dim oPlane             As IngrGeom3D.Plane3d
Dim arrPoints(0 To 11) As Double
Dim objRevolution      As Object
Dim oLine              As IngrGeom3D.Line3d
Dim oAxisVector        As IJDVector
Dim oCenterPoint       As IJDPosition

    'Case of 360 degrees.
    If (Abs(parPlatformAngle - (2 * PI)) < TOLERANCEVAL) Then
        m_bAngle = True
    Else
        While (CmpDblGreaterthan(parPlatformAngle, (2 * PI)))
            parPlatformAngle = parPlatformAngle - (2 * PI)
            If (parPlatformAngle < TOLERANCEVAL) Then
                 m_bAngle = True
                 parPlatformAngle = (2 * PI)
            End If
        Wend
    End If

     If (m_bAngle = True) Then
            Dim oPlCircle As IngrGeom3D.Circle3d
            Dim oLines As Collection
            Dim iCount As Integer

            Set oPlCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, parSkirtHeight + parPlatformHeight, 0, 0, 1, parPlatformRadius + parPlatformWidth)
            Set oComplexString = New ComplexString3d
            oComplexString.AddCurve oPlCircle, True

            Set oPlane = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, 0, 0, parSkirtHeight + parPlatformHeight, 0, 0, 1)
            oPlane.AddBoundary oComplexString

            Set oComplexString = Nothing
            Set oComplexString = New ComplexString3d
            Set oPlCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, parSkirtHeight + parPlatformHeight, 0, 0, 1, parPlatformRadius)
            oComplexString.AddCurve oPlCircle, True
            oPlane.AddBoundary oComplexString
            m_OutputColl.AddOutput "Plat1Plane1", oPlane
            Set oPlane = Nothing
            Set oComplexString = Nothing

            Set oPlCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, parSkirtHeight, 0, 0, -1, parPlatformRadius + parPlatformWidth)
            Set oComplexString = New ComplexString3d
            oComplexString.AddCurve oPlCircle, True
            Set oPlane = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, 0, 0, parSkirtHeight, 0, 0, -1)
            oPlane.AddBoundary oComplexString
            Set oComplexString = Nothing
            Set oComplexString = New ComplexString3d
            Set oPlCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, parSkirtHeight, 0, 0, -1, parPlatformRadius)
            oComplexString.AddCurve oPlCircle, True
            oPlane.AddBoundary oComplexString
            m_OutputColl.AddOutput "Plat1Plane2", oPlane
            Set oPlane = Nothing
            Set oComplexString = Nothing

            Set oLines = New Collection
            Set oPlCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, 0, 0, parSkirtHeight + parPlatformHeight, 0, 0, 1, parPlatformRadius)
            oLines.Add oPlCircle
            m_OutputColl.AddOutput "Plat1Line2", oPlCircle

            Set oPlCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, 0, 0, parSkirtHeight + parPlatformHeight, 0, 0, 1, parPlatformRadius + parPlatformWidth)
            oLines.Add oPlCircle
            m_OutputColl.AddOutput "Plat1Line4", oPlCircle
            Set oPlCircle = Nothing
            For iCount = 1 To oLines.count
                oLines.Remove 1
            Next iCount
            Set oLines = Nothing
    End If

    If (m_bAngle = False) Then
        arrPoints(0) = -parPlatformRadius * Sin(parPlatformAngle / 2)
        arrPoints(1) = parPlatformRadius * Cos(parPlatformAngle / 2)
        arrPoints(2) = parSkirtHeight + parPlatformHeight
        arrPoints(3) = -(parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2)
        arrPoints(4) = (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2)
        arrPoints(5) = parSkirtHeight + parPlatformHeight
        arrPoints(6) = -(parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2)
        arrPoints(7) = (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2)
        arrPoints(8) = parSkirtHeight
        arrPoints(9) = -parPlatformRadius * Sin(parPlatformAngle / 2)
        arrPoints(10) = parPlatformRadius * Cos(parPlatformAngle / 2)
        arrPoints(11) = parSkirtHeight

'Creates a plane based on the above points
        Set oPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, arrPoints)
        m_OutputColl.AddOutput "Plat1Plane1", oPlane
        Set oPlane = Nothing

        arrPoints(0) = parPlatformRadius * Sin(parPlatformAngle / 2)
        arrPoints(1) = parPlatformRadius * Cos(parPlatformAngle / 2)
        arrPoints(2) = parSkirtHeight + parPlatformHeight
        arrPoints(3) = parPlatformRadius * Sin(parPlatformAngle / 2)
        arrPoints(4) = parPlatformRadius * Cos(parPlatformAngle / 2)
        arrPoints(5) = parSkirtHeight
        arrPoints(6) = (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2)
        arrPoints(7) = (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2)
        arrPoints(8) = parSkirtHeight
        arrPoints(9) = (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2)
        arrPoints(10) = (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2)
        arrPoints(11) = parSkirtHeight + parPlatformHeight

'Creates a plane based on the above points
        Set oPlane = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, arrPoints)
        m_OutputColl.AddOutput "Plat1Plane2", oPlane
        Set oPlane = Nothing
    End If

'Create a line and revolve it around the AxisVector with centre point and angle
    Set oAxisVector = New AutoMath.DVector
    Set oCenterPoint = New AutoMath.DPosition
    oCenterPoint.Set 0, 0, parSkirtHeight + parPlatformHeight
    oAxisVector.Set 0, 0, 1
    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, parPlatformRadius * Sin(parPlatformAngle / 2), _
                                        parPlatformRadius * Cos(parPlatformAngle / 2), _
                                        parSkirtHeight, _
                                        parPlatformRadius * Sin(parPlatformAngle / 2), _
                                        parPlatformRadius * Cos(parPlatformAngle / 2), _
                                        parSkirtHeight + parPlatformHeight)
    Set objRevolution = PlaceRevolution(m_OutputColl, oLine, oAxisVector, oCenterPoint, parPlatformAngle, False)
    Set oLine = Nothing
    m_OutputColl.AddOutput "Plat1Revolution1", objRevolution
    Set objRevolution = Nothing

'Create a line and revolve it around the AxisVector with centre point and angle
    Set oLine = geomFactory.Lines3d.CreateBy2Points(Nothing, (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), _
                                (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), _
                                parSkirtHeight, _
                                (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), _
                                (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), _
                                parSkirtHeight + parPlatformHeight)
    Set objRevolution = PlaceRevolution(m_OutputColl, oLine, oAxisVector, oCenterPoint, parPlatformAngle, False)
    Set oLine = Nothing
    m_OutputColl.AddOutput "Plat1Revolution2", objRevolution
    Set objRevolution = Nothing
    Set oAxisVector = Nothing

    If (m_bAngle = False) Then
'        Set oComplexString = CreateTopCurves(Nothing, oCenterPoint, parPlatformRadius, parPlatformAngle, parPlatformWidth, parPlatformHeight)
        Dim oLine1           As IngrGeom3D.Line3d
        Dim oArc1            As Object
        Dim arrPt(4)       As New AutoMath.DPosition
        Dim oNormalVector1   As IJDVector
        Dim oLines1          As Collection
        Set oNormalVector1 = New AutoMath.DVector
        Set oLines1 = New Collection

        arrPt(1).Set -parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), parSkirtHeight + parPlatformHeight
        arrPt(2).Set -(parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), parSkirtHeight + parPlatformHeight
        arrPt(3).Set (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), parSkirtHeight + parPlatformHeight
        arrPt(4).Set parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), parSkirtHeight + parPlatformHeight
        Set oLine1 = PlaceLine(m_OutputColl.ResourceManager, arrPt(1), arrPt(2))
        m_OutputColl.AddOutput "Plat1Line1", oLine1
        oLines1.Add oLine1
        oNormalVector1.Set 0, 0, -1
        Set oArc1 = PlaceArcByCen(m_OutputColl.ResourceManager, oCenterPoint, arrPt(2), arrPt(3), oNormalVector1)
        m_OutputColl.AddOutput "Plat1Line2", oArc1
        oLines1.Add oArc1
        Set oLine1 = PlaceLine(m_OutputColl.ResourceManager, arrPt(3), arrPt(4))
        m_OutputColl.AddOutput "Plat1Line3", oLine1
        oLines1.Add oLine1
        oNormalVector1.Set 0, 0, 1
        Set oArc1 = PlaceArcByCen(m_OutputColl.ResourceManager, oCenterPoint, arrPt(4), arrPt(1), oNormalVector1)
        m_OutputColl.AddOutput "Plat1Line4", oArc1
        oLines1.Add oArc1
        Set oComplexString = PlaceTrCString(arrPt(1), oLines1)
        Set oLines1 = Nothing

        Set oPlane = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, -parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), parSkirtHeight + parPlatformHeight, 0, 0, 1)
        oPlane.AddBoundary oComplexString
        m_OutputColl.AddOutput "Plat1TopSurface", oPlane
        Set oPlane = Nothing

        oCenterPoint.Set 0, 0, parSkirtHeight

'        Set oComplexString = CreateTopCurves(Nothing, oCenterPoint, parPlatformRadius, parPlatformAngle, parPlatformWidth, parPlatformHeight)

        Dim oLines2          As Collection
        Set oLines2 = New Collection
        arrPt(1).Set -parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), parSkirtHeight
        arrPt(2).Set -(parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), parSkirtHeight
        arrPt(3).Set (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), parSkirtHeight
        arrPt(4).Set parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), parSkirtHeight
        Set oLine1 = PlaceLine(Nothing, arrPt(1), arrPt(2))
        oLines2.Add oLine1
        oNormalVector1.Set 0, 0, -1
        Set oArc1 = PlaceArcByCen(Nothing, oCenterPoint, arrPt(2), arrPt(3), oNormalVector1)
        oLines2.Add oArc1
        Set oLine1 = PlaceLine(Nothing, arrPt(3), arrPt(4))
        oLines2.Add oLine1
        oNormalVector1.Set 0, 0, 1
        Set oArc1 = PlaceArcByCen(Nothing, oCenterPoint, arrPt(4), arrPt(1), oNormalVector1)
        oLines2.Add oArc1
        Set oComplexString = PlaceTrCString(arrPt(1), oLines2)
        Set oLines2 = Nothing

        Set oPlane = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, -parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), parSkirtHeight, 0, 0, -1)
        oPlane.AddBoundary oComplexString
        m_OutputColl.AddOutput "Plat1BottomSurface", oPlane
        
        
    End If
    Set oCenterPoint = Nothing
    Set oComplexString = Nothing
    Set oLine1 = Nothing
    Set oPlane = Nothing
    Set oArc1 = Nothing
    Set oNormalVector1 = Nothing
    
'Creation of Top Platform
'The top platform is taken to be at the top third of the vessel
Dim TopPlatHeightFromBase As Double
TopPlatHeightFromBase = parSkirtHeight + (2 * parVesselHeight / 3)

Dim oComplexStringT     As IngrGeom3D.ComplexString3d
Dim oPlaneT             As IngrGeom3D.Plane3d
Dim arrPointsT(0 To 11) As Double
Dim objRevolutionT      As Object
Dim oLineT              As IngrGeom3D.Line3d
Dim oAxisVectorT        As IJDVector
Dim oCenterPointT       As IJDPosition
     
    'Case of 360 degrees.
    'This validation code is kept, so that it can be easily modified if Platform Angle is different for Top and bottom platforms
    If (Abs(parPlatformAngle - (2 * PI)) < TOLERANCEVAL) Then
        m_bAngle = True
    Else
        While CmpDblGreaterthan(parPlatformAngle, (2 * PI))
            parPlatformAngle = parPlatformAngle - (2 * PI)
            If (parPlatformAngle < TOLERANCEVAL) Then
                 m_bAngle = True
                 parPlatformAngle = (2 * PI)
            End If
        Wend
    End If
                
     If (m_bAngle = True) Then
            Dim oPlCircleT As IngrGeom3D.Circle3d
            Dim oLinesT As Collection
            Dim iCountT As Integer
            
            Set oPlCircleT = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, TopPlatHeightFromBase, 0, 0, 1, parPlatformRadius + parPlatformWidth)
            Set oComplexStringT = New ComplexString3d
            oComplexStringT.AddCurve oPlCircleT, True
            
            Set oPlaneT = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, 0, 0, TopPlatHeightFromBase, 0, 0, 1)
            oPlaneT.AddBoundary oComplexStringT
            
            Set oComplexStringT = Nothing
            Set oComplexStringT = New ComplexString3d
            Set oPlCircleT = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, TopPlatHeightFromBase, 0, 0, 1, parPlatformRadius)
            oComplexStringT.AddCurve oPlCircleT, True
            oPlaneT.AddBoundary oComplexStringT
            m_OutputColl.AddOutput "Plat2Plane1", oPlaneT
            Set oPlaneT = Nothing
            Set oComplexStringT = Nothing
            
            Set oPlCircleT = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, TopPlatHeightFromBase - parPlatformHeight, 0, 0, -1, parPlatformRadius + parPlatformWidth)
            Set oComplexStringT = New ComplexString3d
            oComplexStringT.AddCurve oPlCircleT, True
            Set oPlaneT = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, 0, 0, TopPlatHeightFromBase - parPlatformHeight, 0, 0, -1)
            oPlaneT.AddBoundary oComplexStringT
            Set oComplexStringT = Nothing
            Set oComplexStringT = New ComplexString3d
            Set oPlCircleT = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, TopPlatHeightFromBase - parPlatformHeight, 0, 0, -1, parPlatformRadius)
            oComplexStringT.AddCurve oPlCircleT, True
            oPlaneT.AddBoundary oComplexStringT
            m_OutputColl.AddOutput "Plat2Plane2", oPlaneT
            Set oPlaneT = Nothing
            Set oComplexStringT = Nothing
            
            Set oLinesT = New Collection
            Set oPlCircleT = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, 0, 0, TopPlatHeightFromBase, 0, 0, 1, parPlatformRadius)
            oLinesT.Add oPlCircleT
            m_OutputColl.AddOutput "Plat2Line2", oPlCircleT
            
            Set oPlCircleT = geomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, 0, 0, TopPlatHeightFromBase, 0, 0, 1, parPlatformRadius + parPlatformWidth)
            oLinesT.Add oPlCircleT
            m_OutputColl.AddOutput "Plat2Line4", oPlCircleT
                    
            For iCountT = 1 To oLinesT.count
                oLinesT.Remove 1
            Next iCountT
            Set oLinesT = Nothing
    End If
    
    If (m_bAngle = False) Then
        arrPointsT(0) = -parPlatformRadius * Sin(parPlatformAngle / 2)
        arrPointsT(1) = parPlatformRadius * Cos(parPlatformAngle / 2)
        arrPointsT(2) = TopPlatHeightFromBase
        arrPointsT(3) = -(parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2)
        arrPointsT(4) = (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2)
        arrPointsT(5) = TopPlatHeightFromBase
        arrPointsT(6) = -(parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2)
        arrPointsT(7) = (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2)
        arrPointsT(8) = TopPlatHeightFromBase - parPlatformHeight
        arrPointsT(9) = -parPlatformRadius * Sin(parPlatformAngle / 2)
        arrPointsT(10) = parPlatformRadius * Cos(parPlatformAngle / 2)
        arrPointsT(11) = TopPlatHeightFromBase - parPlatformHeight
        
'Creates a plane based on the above points
        Set oPlaneT = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, arrPointsT)
        m_OutputColl.AddOutput "Plat2Plane1", oPlaneT
        Set oPlaneT = Nothing
       
        arrPointsT(0) = parPlatformRadius * Sin(parPlatformAngle / 2)
        arrPointsT(1) = parPlatformRadius * Cos(parPlatformAngle / 2)
        arrPointsT(2) = TopPlatHeightFromBase
        arrPointsT(3) = parPlatformRadius * Sin(parPlatformAngle / 2)
        arrPointsT(4) = parPlatformRadius * Cos(parPlatformAngle / 2)
        arrPointsT(5) = TopPlatHeightFromBase - parPlatformHeight
        arrPointsT(6) = (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2)
        arrPointsT(7) = (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2)
        arrPointsT(8) = TopPlatHeightFromBase - parPlatformHeight
        arrPointsT(9) = (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2)
        arrPointsT(10) = (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2)
        arrPointsT(11) = TopPlatHeightFromBase
        
'Creates a plane based on the above points
        Set oPlaneT = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, arrPointsT)
        m_OutputColl.AddOutput "Plat2Plane2", oPlaneT
        Set oPlaneT = Nothing
    End If

'Create a line and revolve it around the AxisVector with centre point and angle
    Set oAxisVectorT = New AutoMath.DVector
    Set oCenterPointT = New AutoMath.DPosition
    oCenterPointT.Set 0, 0, TopPlatHeightFromBase
    oAxisVectorT.Set 0, 0, 1
    Set oLineT = geomFactory.Lines3d.CreateBy2Points(Nothing, parPlatformRadius * Sin(parPlatformAngle / 2), _
                                        parPlatformRadius * Cos(parPlatformAngle / 2), _
                                        TopPlatHeightFromBase - parPlatformHeight, _
                                        parPlatformRadius * Sin(parPlatformAngle / 2), _
                                        parPlatformRadius * Cos(parPlatformAngle / 2), _
                                        TopPlatHeightFromBase)
    Set objRevolutionT = PlaceRevolution(m_OutputColl, oLineT, oAxisVectorT, oCenterPointT, parPlatformAngle, False)
    Set oLineT = Nothing
    m_OutputColl.AddOutput "Plat2Revolution1", objRevolutionT
    Set objRevolutionT = Nothing
    
'Create a line and revolve it around the AxisVector with centre point and angle
    Set oLineT = geomFactory.Lines3d.CreateBy2Points(Nothing, (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), _
                                (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), _
                                TopPlatHeightFromBase - parPlatformHeight, _
                                (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), _
                                (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), _
                                TopPlatHeightFromBase)
    Set objRevolutionT = PlaceRevolution(m_OutputColl, oLineT, oAxisVectorT, oCenterPointT, parPlatformAngle, False)
    Set oLineT = Nothing
    m_OutputColl.AddOutput "Plat2Revolution2", objRevolutionT
    Set objRevolutionT = Nothing
    Set oAxisVectorT = Nothing

    If (m_bAngle = False) Then
'        Set oComplexStringT = CreateTopCurves(Nothing, oCenterPointT, parPlatformRadius, parPlatformAngle, parPlatformWidth, parPlatformHeight)

        Dim oLine1T           As IngrGeom3D.Line3d
        Dim oArc1T            As Object
        Dim arrPtT(4)       As New AutoMath.DPosition
        Dim oNormalVector1T   As IJDVector
        Dim oLines1T          As Collection
        Set oNormalVector1T = New AutoMath.DVector
        Set oLines1T = New Collection

        arrPtT(1).Set -parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), TopPlatHeightFromBase
        arrPtT(2).Set -(parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), TopPlatHeightFromBase
        arrPtT(3).Set (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), TopPlatHeightFromBase
        arrPtT(4).Set parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), TopPlatHeightFromBase
        Set oLine1T = PlaceLine(m_OutputColl.ResourceManager, arrPtT(1), arrPtT(2))
        m_OutputColl.AddOutput "Plat2Line1", oLine1T
        oLines1T.Add oLine1T
        oNormalVector1T.Set 0, 0, -1
        Set oArc1T = PlaceArcByCen(m_OutputColl.ResourceManager, oCenterPointT, arrPtT(2), arrPtT(3), oNormalVector1T)
        m_OutputColl.AddOutput "Plat2Line2", oArc1T
        oLines1T.Add oArc1T
        Set oLine1T = PlaceLine(m_OutputColl.ResourceManager, arrPtT(3), arrPtT(4))
         m_OutputColl.AddOutput "Plat2Line3", oLine1T
        oLines1T.Add oLine1T
        oNormalVector1T.Set 0, 0, 1
        Set oArc1T = PlaceArcByCen(m_OutputColl.ResourceManager, oCenterPointT, arrPtT(4), arrPtT(1), oNormalVector1T)
        m_OutputColl.AddOutput "Plat2Line4", oArc1T
        oLines1T.Add oArc1T
        Set oComplexStringT = PlaceTrCString(arrPtT(1), oLines1T)
        Set oLines1T = Nothing

        Set oPlaneT = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, -parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), TopPlatHeightFromBase, 0, 0, 1)
        oPlaneT.AddBoundary oComplexStringT
        m_OutputColl.AddOutput "Plat2TopSurface", oPlaneT
        Set oPlaneT = Nothing
        
        oCenterPointT.Set 0, 0, TopPlatHeightFromBase - parPlatformHeight
        
'        Set oComplexStringT = CreateTopCurves(Nothing, oCenterPointT, parPlatformRadius, parPlatformAngle, parPlatformWidth, parPlatformHeight)

        Dim oLines2T          As Collection
        Set oLines2T = New Collection
        arrPtT(1).Set -parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), TopPlatHeightFromBase - parPlatformHeight
        arrPtT(2).Set -(parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), TopPlatHeightFromBase - parPlatformHeight
        arrPtT(3).Set (parPlatformWidth + parPlatformRadius) * Sin(parPlatformAngle / 2), (parPlatformWidth + parPlatformRadius) * Cos(parPlatformAngle / 2), TopPlatHeightFromBase - parPlatformHeight
        arrPtT(4).Set parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), TopPlatHeightFromBase - parPlatformHeight
        Set oLine1T = PlaceLine(Nothing, arrPtT(1), arrPtT(2))
        oLines2T.Add oLine1T
        oNormalVector1T.Set 0, 0, -1
        Set oArc1T = PlaceArcByCen(Nothing, oCenterPointT, arrPtT(2), arrPtT(3), oNormalVector1T)
        oLines2T.Add oArc1T
        Set oLine1T = PlaceLine(Nothing, arrPtT(3), arrPtT(4))
        oLines2T.Add oLine1T
        oNormalVector1T.Set 0, 0, 1
        Set oArc1T = PlaceArcByCen(Nothing, oCenterPointT, arrPtT(4), arrPtT(1), oNormalVector1T)
        oLines2T.Add oArc1T
        Set oComplexStringT = PlaceTrCString(arrPtT(1), oLines2T)
        Set oLines2T = Nothing

        Set oPlaneT = geomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, -parPlatformRadius * Sin(parPlatformAngle / 2), parPlatformRadius * Cos(parPlatformAngle / 2), TopPlatHeightFromBase - parPlatformHeight, 0, 0, -1)
        oPlaneT.AddBoundary oComplexStringT
        m_OutputColl.AddOutput "Plat2BottomSurface", oPlaneT
        Set oPlaneT = Nothing

    End If
    Set oCenterPointT = Nothing
    Set oComplexStringT = Nothing
    Set oLine1T = Nothing
    Set oArc1T = Nothing
    Set oNormalVector1 = Nothing
    
'==========================================
'Construction of  Control Point at tangent line, start of vessel
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, parVesselStartHeight, 0.05, , , True)
    
    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "CESimVerVesselControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
''' This function creates persistent/Transient Line based on
''' start and end points of the line
'''<{(Line begin)}>
Public Function PlaceLine(pResourceMgr As IUnknown, ByRef startPoint As IJDPosition, _
                            ByRef endPoint As IJDPosition) _
                            As IngrGeom3D.Line3d

    Const METHOD = "PlaceLine"
    On Error GoTo ErrorHandler
        
    Dim oLine As IngrGeom3D.Line3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'MsgBox "about to create the line"
    ' Create Line object
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(pResourceMgr, _
                                startPoint.x, startPoint.y, startPoint.z, _
                                endPoint.x, endPoint.y, endPoint.z)
    
    
    Set PlaceLine = oLine
    Set oLine = Nothing
    Set oGeomFactory = Nothing
    Exit Function
    
ErrorHandler:
   ReportUnanticipatedError MODULE, METHOD
    Resume Next
End Function
'''<{(Line end)}>

''' This function creates persistent/Transient Arc based on centre point
''' NormalVector, start and end points of the curve
'''<{(PlaceArcByCen begin)}>
Public Function PlaceArcByCen(pResourceMgr As IUnknown, _
                            ByRef centerPoint As IJDPosition, _
                            ByRef startPoint As IJDPosition, _
                            ByRef endPoint As IJDPosition, _
                            ByRef normVector As IJDVector) _
                            As IngrGeom3D.Arc3d
    Const METHOD = "PlaceArcByCen"
    On Error GoTo ErrorHandler
        
    Dim oArc As IngrGeom3D.Arc3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(pResourceMgr, _
                              centerPoint.x, centerPoint.y, centerPoint.z, _
                             normVector.x, normVector.y, normVector.z, _
                             startPoint.x, startPoint.y, startPoint.z, _
                            endPoint.x, endPoint.y, endPoint.z)
   
    Set PlaceArcByCen = oArc
    Set oArc = Nothing
    Set oGeomFactory = Nothing

    Exit Function
    
ErrorHandler:
   ReportUnanticipatedError MODULE, METHOD
   
End Function
'''<{(PlaceArcByCen end)}>



