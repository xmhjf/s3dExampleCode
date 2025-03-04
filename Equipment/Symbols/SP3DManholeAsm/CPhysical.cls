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
'               Manhole Assembly.
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This class module has Five Outputs
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   10.06.2008      PK     Created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private Const NEGLIGIBLE_THICKNESS = 0.0001

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
    
    Dim parBaseHeight               As Double
    Dim parConeHeight               As Double
    Dim parBaseDiameter             As Double
    Dim parConeTopDiameter          As Double
    Dim parBaseBottoNozzleHeight    As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBaseHeight = arrayOfInputs(2)
    parConeHeight = arrayOfInputs(3)
    parBaseDiameter = arrayOfInputs(4)
    parConeTopDiameter = arrayOfInputs(5)
    parBaseBottoNozzleHeight = arrayOfInputs(6)
    
    iOutput = 0
    
    'When the Cone Height is zero,set it to Negligible Thickness
    If CmpDblEqual(parConeHeight, 0) Then parConeHeight = NEGLIGIBLE_THICKNESS
        
    Dim oStPoint    As AutoMath.DPosition
    Dim oEnPoint    As AutoMath.DPosition
    Dim iCount      As Integer
    
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    
    'Create the Manhole Base
    Dim oManholeBase    As Object
    oStPoint.Set 0, 0, 0
    oEnPoint.Set 0, 0, parBaseHeight
    Set oManholeBase = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parBaseDiameter, True)
    
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oManholeBase
    Set oManholeBase = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    'Create Manhole Eccentric Cone
    Dim oEccentricCone  As Object
    Dim oBaseCircle     As IngrGeom3D.Circle3d
    Dim oTopCircle      As IngrGeom3D.Circle3d
    Dim oBase           As IJComplexString
    Dim oTop            As IJComplexString
    Dim oElements       As IJElements
    
    Set oElements = New JObjectCollection
    
    Set oBaseCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    0, 0, parBaseHeight, _
                                                    0, 0, 1, _
                                                    parBaseDiameter / 2)
                                                    
    Set oTopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                    -(parBaseDiameter - parConeTopDiameter) / 2 _
                                                    , 0, parBaseHeight + parConeHeight, _
                                                    0, 0, 1, _
                                                    parConeTopDiameter / 2)
    oElements.Add oBaseCircle
    Set oBase = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
    
    oElements.Clear
    
    oElements.Add oTopCircle
    Set oTop = oGeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oElements)
    
    Set oEccentricCone = oGeomFactory.RuledSurfaces3d.CreateByCurves( _
                                                    m_OutputColl.ResourceManager, _
                                                    oBase, oTop, True)
        
    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oEccentricCone
    
    Set oEccentricCone = Nothing
    Set oTop = Nothing
    Set oBase = Nothing
    Set oElements = Nothing
    Set oBaseCircle = Nothing
    Set oTopCircle = Nothing
    
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
                                                -(parBaseDiameter - parConeTopDiameter) / 2, _
                                                0, _
                                                parBaseHeight + parConeHeight)

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
                                     parBaseDiameter / 2)

        
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
