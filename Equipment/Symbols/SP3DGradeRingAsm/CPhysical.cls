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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:02:00 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author       :  PK
'   Creation Date:  Wednesday, June 11 2008
'
'   Description:
'       CR-105119  Deliver Common Equipment Symbols Required for Designing Drain and Waste Piping
'            Grade Ring Assembly.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Four Outputs
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   11.06.2008      PK     Created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages
Private Const NEGLIGIBLE_THICKNESS = 0.001
Private PI As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parRingHeight             As Double
    Dim parRingOuterDia           As Double
    Dim parRingInnerDia           As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parRingHeight = arrayOfInputs(2)
    parRingOuterDia = arrayOfInputs(3)
    parRingInnerDia = arrayOfInputs(4)
    
    iOutput = 0
    
    If CmpDblEqual(parRingInnerDia, 0) Then parRingInnerDia = NEGLIGIBLE_THICKNESS

    Dim oCenPoint    As AutoMath.DPosition
    Set oCenPoint = New AutoMath.DPosition
        
    'Create Grade Ring
    Dim oGradeRing              As Object
    Dim oGeomFactory            As IngrGeom3D.GeometryFactory
    Dim dLineStrPoints(0 To 14) As Double
    Dim oLineString             As IngrGeom3D.LineString3d
    Dim oAxis                   As AutoMath.DVector
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oAxis = New AutoMath.DVector
    
    oAxis.Set 0, 0, 1
    oCenPoint.Set 0, 0, parRingHeight / 2
    
    dLineStrPoints(0) = parRingInnerDia / 2
    dLineStrPoints(1) = 0
    dLineStrPoints(2) = 0
    
    dLineStrPoints(3) = parRingOuterDia / 2
    dLineStrPoints(4) = 0
    dLineStrPoints(5) = 0
    
    dLineStrPoints(6) = parRingOuterDia / 2
    dLineStrPoints(7) = 0
    dLineStrPoints(8) = parRingHeight
    
    dLineStrPoints(9) = parRingInnerDia / 2
    dLineStrPoints(10) = 0
    dLineStrPoints(11) = parRingHeight
    
    dLineStrPoints(12) = parRingInnerDia / 2
    dLineStrPoints(13) = 0
    dLineStrPoints(14) = 0
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dLineStrPoints)

    Set oGradeRing = PlaceRevolution(m_OutputColl, oLineString, oAxis, oCenPoint, 2 * PI, True)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oGradeRing
    Set oGradeRing = Nothing
    Set oCenPoint = Nothing
    Set oLineString = Nothing
    Set oAxis = Nothing

    'Create Connect Point 1
    Dim oConnectPoint1  As Object
    Set oConnectPoint1 = oGeomFactory.Points3d.CreateByPoint( _
                                                m_OutputColl.ResourceManager, _
                                                0, 0, 0)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oConnectPoint1
    Set oConnectPoint1 = Nothing

    'Create Connect Point 2
    Dim oConnectPoint2 As Object
    Set oConnectPoint2 = oGeomFactory.Points3d.CreateByPoint( _
                                                m_OutputColl.ResourceManager, _
                                                0, 0, parRingHeight)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oConnectPoint2
    Set oConnectPoint2 = Nothing

    'Create Default Surface
    'Create non-persistent circle to use for creating default surface
    Dim oDefaultSurface As IngrGeom3D.Plane3d
    Dim oCircle As IngrGeom3D.Circle3d
    
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                      0, 0, 0, _
                                      0, 0, -1, _
                                     0.1 * parRingInnerDia / 2)

        
    'Create persistent default surface plane - the plane can mate ---
    Set oDefaultSurface = oGeomFactory.Planes3d.CreateByOuterBdry _
                                               (m_OutputColl.ResourceManager, _
                                               oCircle)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oDefaultSurface
    Set oDefaultSurface = Nothing
    Set oCircle = Nothing
    Set oGeomFactory = Nothing
    
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
