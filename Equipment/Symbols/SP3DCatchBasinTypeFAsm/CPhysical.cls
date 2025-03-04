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
'               Catch Basin Type F Assembly.
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
    Dim parBasinWidth               As Double
    Dim parBasinDepth               As Double
    Dim parBasinBottoNozzleHeight   As Double

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBasinHeight = arrayOfInputs(2)
    parBasinWidth = arrayOfInputs(3)
    parBasinDepth = arrayOfInputs(4)
    parBasinBottoNozzleHeight = arrayOfInputs(5)
    
    iOutput = 0

    Dim oStPoint    As AutoMath.DPosition
    Dim oEnPoint    As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    'Create Recatangular Solid Catch Basin
    Dim oCatchBasin    As Object
    oStPoint.Set -parBasinWidth / 2, -parBasinDepth / 2, 0
    oEnPoint.Set parBasinWidth / 2, parBasinDepth / 2, parBasinHeight
    Set oCatchBasin = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    
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
    Dim oDefaultSurface     As Object
    Dim oPoints(0 To 11)    As Double
      
    oPoints(0) = -parBasinWidth / 2
    oPoints(1) = parBasinDepth / 2
    oPoints(2) = 0
    
    oPoints(3) = parBasinWidth / 2
    oPoints(4) = parBasinDepth / 2
    oPoints(5) = 0
    
    oPoints(6) = parBasinWidth / 2
    oPoints(7) = -parBasinDepth / 2
    oPoints(8) = 0
    
    oPoints(9) = -parBasinWidth / 2
    oPoints(10) = -parBasinDepth / 2
    oPoints(11) = 0
        
    Set oDefaultSurface = oGeomFactory.Planes3d.CreateByPoints( _
                                            m_OutputColl.ResourceManager, _
                                            4, _
                                            oPoints)

    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oDefaultSurface
    Set oDefaultSurface = Nothing
    Set oGeomFactory = Nothing
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub
