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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:21 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RUK
'   Creation Date:  Wednesday, Mar 14 2007
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Six Outputs:
'      In which Default Surface is created at the Origin.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CPhysical:" 'Used for error messages
Private m_oSymGeomHelper As IJSymbolGeometryHelper

Private Const CYLINDRICAL_SUPPORTS = 1
Private Const CUBOID_SUPPORTS = 2
Private PI       As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oSymGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt           As PartFacelets.IJDPart
    Dim iOutput             As Double
    Dim ObjVessel           As Object

    Dim parVesselDiameter           As Double
    Dim parSupportAngularLocation   As Double
    Dim parNumberOfSupports         As Long
    Dim parVesselCenterHeight       As Double
    Dim parSupportLength            As Double
    Dim parSupportThickness         As Double
    Dim parSupportRadialLocation    As Double
    Dim parSupportHeight            As Double
    Dim parInsulationThickness      As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)            'P2
    parVesselCenterHeight = arrayOfInputs(3)        'P5
    parSupportLength = arrayOfInputs(4)             'P6
    parSupportHeight = arrayOfInputs(5)             'P9
    parSupportThickness = arrayOfInputs(6)          'P7
    parNumberOfSupports = arrayOfInputs(7)          'P4
    parSupportAngularLocation = arrayOfInputs(8)    'P3
    parSupportRadialLocation = arrayOfInputs(9)     'P8
    parInsulationThickness = arrayOfInputs(10)

    iOutput = 0
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    'Origin is taken at Place point1(DP/PP 1)
    'Insert your code for Vessel
    Dim centPoint As New AutoMath.DPosition
    centPoint.Set 0, 0, parVesselCenterHeight
    
    Set ObjVessel = PlaceSphere(m_OutputColl, centPoint, parVesselDiameter / 2)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVessel
    Set ObjVessel = Nothing
    Set centPoint = Nothing
    
    'Insert code for Circular Foundation Port
    Dim ObjFoundationPort As IJEqpFoundationPort
    Dim NozzlePHFactory As NozzlePHFactory
    Set NozzlePHFactory = New NozzlePHFactory
    Dim dOrigin(0 To 2) As Double
    Dim dXaxis(0 To 2) As Double
    Dim dZaxis(0 To 2) As Double
    
    'The origin of the port is taken to be at the centre point of the support base.
    dOrigin(0) = 0
    dOrigin(1) = 0
    dOrigin(2) = 0
    
    dXaxis(0) = 0
    dXaxis(1) = 1
    dXaxis(2) = 0#
    
    dZaxis(0) = 0#
    dZaxis(1) = 0#
    dZaxis(2) = -1#

    Set ObjFoundationPort = NozzlePHFactory.CreateNozzlePHGivenPartAndID(oPartFclt, "STFndPort1", _
                                                     False, m_OutputColl.ResourceManager)
    Call ObjFoundationPort.PutCS(dOrigin(0), dOrigin(1), dOrigin(2), _
                            dXaxis(0), dXaxis(1), dXaxis(2), _
                            dZaxis(0), dZaxis(1), dZaxis(2))
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFoundationPort
    Set ObjFoundationPort = Nothing
    Set NozzlePHFactory = Nothing
    
    'Insert your code for Default Surface
    'Create non-persistent circle to use for creating default surface ---
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim objPlane As IngrGeom3D.Plane3d
    Dim objCircle As IngrGeom3D.Circle3d
    
    Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                             0, 0, 0, _
                            0, 0, -1, _
                            0.5 * parVesselDiameter / 2)

    'Create persistent default surface plane - the plane can mate ---
    Set objPlane = oGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objPlane
    Set objPlane = Nothing
    Set objCircle = Nothing
   
   'Insert code for supports, edges, surfaces of Cylindrical and Cuboidal Supports
    If parNumberOfSupports > 0 And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportHeight, LINEAR_TOLERANCE) Then
        If CmpDblLessThan(2 * PI * parSupportRadialLocation / parSupportLength, parNumberOfSupports) Then
            GoTo ErrorLabel
        End If
        'Insert your code for output  (Supports/DefaultSurfaces)
        Dim iSupportType As Integer
        Dim oNormalVect As New AutoMath.DVector
        Dim oCircle As Object
        Dim dAngle As Double
        Dim iCount As Integer
        
        'Set the number and types of Supports/DefaultSurfaces and their type
        Dim ObjSupport As Object
        Dim ObjSurface As Object
        Dim ObjEdge As Object
        Dim ObjPoint As Object
        Dim stPoint   As New AutoMath.DPosition
        Dim enPoint   As New AutoMath.DPosition
        
        If CmpDblEqual(parSupportThickness, LINEAR_TOLERANCE) Then
            iSupportType = CYLINDRICAL_SUPPORTS
        Else
            iSupportType = CUBOID_SUPPORTS
        End If
        oNormalVect.Set 0, 0, -1

        'Place supports
        Select Case iSupportType
           Case CYLINDRICAL_SUPPORTS
                For iCount = 1 To parNumberOfSupports
                    'Center point position in horizontal plane is needed
                    dAngle = parSupportAngularLocation + _
                                    (iCount - 1) * (2 * PI / parNumberOfSupports)
                    stPoint.Set parSupportRadialLocation * Sin(dAngle), _
                                parSupportRadialLocation * Cos(dAngle), 0
                    enPoint.Set stPoint.x, stPoint.y, stPoint.z + parSupportHeight
                    Set ObjSupport = PlaceCylinder(m_OutputColl, stPoint, enPoint, parSupportLength, False)
                    'Set the output
                    iOutput = iOutput + 1
                    m_OutputColl.AddOutput "Support_", ObjSupport
                    Set ObjSupport = Nothing
                        
                    'Create the Edge(Circle) for Bottom Surface of Support
                    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(m_OutputColl.ResourceManager, _
                                                    stPoint.x, stPoint.y, stPoint.z, _
                                                    oNormalVect.x, oNormalVect.y, oNormalVect.z, _
                                                    parSupportLength / 2)
                    'Set the output
                    iOutput = iOutput + 1
                    m_OutputColl.AddOutput "Edge_", oCircle
                    Set oCircle = Nothing
                     
                    'Place Point at both Ends
                    Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, enPoint.x, enPoint.y, enPoint.z)
                    iOutput = iOutput + 1
                    m_OutputColl.AddOutput "Point1_", ObjPoint
                    Set ObjPoint = Nothing
                    
                    Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, stPoint.x, stPoint.y, stPoint.z)
                    iOutput = iOutput + 1
                    m_OutputColl.AddOutput "Point2_", ObjPoint
                    Set ObjPoint = Nothing
                     
                    'Place Surface at bottom of support
                    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    stPoint.x, stPoint.y, stPoint.z, _
                                                    oNormalVect.x, oNormalVect.y, oNormalVect.z, _
                                                    parSupportLength / 2)
                    
                    Dim oComplexStr As New IngrGeom3D.ComplexString3d
                    Call oComplexStr.AddCurve(oCircle, False)
                    Set ObjSurface = oGeomFactory.Planes3d.CreateByPointNormal(m_OutputColl.ResourceManager, _
                                                                        stPoint.x, stPoint.y, stPoint.z, _
                                                                        oNormalVect.x, oNormalVect.y, oNormalVect.z)
                    Call ObjSurface.AddBoundary(oComplexStr)
                    'Set the output
                    iOutput = iOutput + 1
                    m_OutputColl.AddOutput "Surface_", ObjSurface
                    Set ObjSurface = Nothing
                    
                    'Remove curve
                    Dim ObjTmpcurves As IJDObject
                    Set ObjTmpcurves = oComplexStr
                    ObjTmpcurves.Remove
                    Set oComplexStr = Nothing
                    Set ObjTmpcurves = oCircle
                    ObjTmpcurves.Remove
                    Set oCircle = Nothing
                    Set ObjTmpcurves = Nothing
                Next iCount
            
            Case CUBOID_SUPPORTS
                Dim oTopPts(0 To 3) As IJDPosition
                Dim oBotPts(0 To 3) As IJDPosition
                
                For iCount = 0 To 3
                    Set oTopPts(iCount) = New DPosition
                    Set oBotPts(iCount) = New DPosition
                Next iCount

                Dim ObjColl As Collection
                For iCount = 1 To parNumberOfSupports
                    dAngle = parSupportAngularLocation + _
                                    (iCount - 1) * (2 * PI / parNumberOfSupports)
                    'Point 1
                    oBotPts(0).x = Sin(dAngle) * (parSupportRadialLocation) + Cos(dAngle) * parSupportLength / 2
                    oBotPts(0).y = Cos(dAngle) * (parSupportRadialLocation) - Sin(dAngle) * parSupportLength / 2
                    oBotPts(0).z = 0

                    'Point 2
                    oBotPts(1).x = oBotPts(0).x + Sin(dAngle) * parSupportThickness
                    oBotPts(1).y = oBotPts(0).y + parSupportThickness * Cos(dAngle)
                    oBotPts(1).z = 0

                    'Point 4
                    oBotPts(3).x = Sin(dAngle) * (parSupportRadialLocation) - Cos(dAngle) * parSupportLength / 2
                    oBotPts(3).y = Cos(dAngle) * (parSupportRadialLocation) + Sin(dAngle) * parSupportLength / 2
                    oBotPts(3).z = 0

                    'Point3
                    oBotPts(2).x = oBotPts(3).x + parSupportThickness * Sin(dAngle)
                    oBotPts(2).y = oBotPts(3).y + parSupportThickness * Cos(dAngle)
                    oBotPts(2).z = 0

                    oTopPts(0).Set oBotPts(0).x, oBotPts(0).y, parSupportHeight
                    oTopPts(1).Set oBotPts(1).x, oBotPts(1).y, parSupportHeight
                    oTopPts(2).Set oBotPts(2).x, oBotPts(2).y, parSupportHeight
                    oTopPts(3).Set oBotPts(3).x, oBotPts(3).y, parSupportHeight

                    Set ObjColl = PlaceTrapezoidWithPlanes(m_OutputColl, oTopPts, oBotPts)
                    'Set the output
                    iOutput = iOutput + 1
                    Dim jCount As Integer
                        For jCount = 1 To ObjColl.Count
                            m_OutputColl.AddOutput "Support_", ObjColl(jCount)
                        Next jCount
                    Set ObjColl = Nothing
                                 
                    'Create the Edges(Lines) by two points at Bottom Surface
                        Dim kCount As Integer
                        For kCount = 0 To 3
                            If kCount = 3 Then
                                Set ObjEdge = oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, oBotPts(kCount).x, oBotPts(kCount).y, oBotPts(kCount).z, oBotPts(0).x, oBotPts(0).y, oBotPts(0).z)
                            Else
                                Set ObjEdge = oGeomFactory.Lines3d.CreateBy2Points(m_OutputColl.ResourceManager, oBotPts(kCount).x, oBotPts(kCount).y, oBotPts(kCount).z, oBotPts(kCount + 1).x, oBotPts(kCount + 1).y, oBotPts(kCount + 1).z)
                            End If
                        'Set the Output for Edges
                        iOutput = iOutput + 1
                        m_OutputColl.AddOutput "Edge_", ObjEdge
                        Set ObjEdge = Nothing
                        Next kCount
                        
                        'Create Points on End faces of Support
                        Dim x, y, z As Double
                        'Bottom FacePoint
                        x = (oBotPts(0).x + oBotPts(2).x) / 2
                        y = (oBotPts(0).y + oBotPts(2).y) / 2
                        z = (oBotPts(0).z + oBotPts(2).z) / 2
                        Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
                        iOutput = iOutput + 1
                        m_OutputColl.AddOutput "Point1_", ObjPoint
                        Set ObjPoint = Nothing
                        
                        'Top FacePoint
                        x = (oTopPts(0).x + oTopPts(2).x) / 2
                        y = (oTopPts(0).y + oTopPts(2).y) / 2
                        z = (oTopPts(0).z + oTopPts(2).z) / 2
                        Set ObjPoint = oGeomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, x, y, z)
                        iOutput = iOutput + 1
                        m_OutputColl.AddOutput "Point2_", ObjPoint
                        Set ObjPoint = Nothing
                Next iCount
        End Select
        For jCount = 0 To 3
            Set oTopPts(jCount) = Nothing
            Set oBotPts(jCount) = Nothing
        Next jCount
        
        Set oNormalVect = Nothing
        Set stPoint = Nothing
        Set enPoint = Nothing
    End If
    Set oGeomFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

