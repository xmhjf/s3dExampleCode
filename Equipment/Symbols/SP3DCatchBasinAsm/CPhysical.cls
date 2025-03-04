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
'   Creation Date:  Tuesday, June 10 2008
'
'   Description:
'       CR-105119  Deliver Common Equipment Symbols Required for Designing Drain and Waste Piping
'               Catch Basin Assembly.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Four Outputs
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   10.06.2008      PK     Created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize:"
    On Error GoTo Errx

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
    
    Dim parBasinHeight              As Double
    Dim parBasinDiameter            As Double
    Dim parBasinBottoNozzleHeight   As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBasinHeight = arrayOfInputs(2)
    parBasinDiameter = arrayOfInputs(3)
    parBasinBottoNozzleHeight = arrayOfInputs(4)
    
    iOutput = 0

    Dim oStPoint    As AutoMath.DPosition
    Dim oEnPoint    As AutoMath.DPosition

    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Create Catch Basin
    Dim oCatchBasin    As Object
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, 0, parBasinHeight
    Set oCatchBasin = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBasinDiameter, True)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oCatchBasin
    Set oCatchBasin = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    'Create Connect Point 1
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oConnectPoint1  As Object
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
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
                                                0, 0, parBasinHeight)

    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oConnectPoint2
    Set oConnectPoint2 = Nothing

    'Create Default Surface
    'Create non-persistent circle to use for creating default surface
    Dim oDefaultSurface As IngrGeom3D.Plane3d
    Dim oCircle         As IngrGeom3D.Circle3d
            
    Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                      0, 0, 0, _
                                      0, 0, -1, _
                                     parBasinDiameter / 2)

        
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
