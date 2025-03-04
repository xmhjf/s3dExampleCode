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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:11:29 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          SSP
'   Creation Date:  Monday, May 12, 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This is a Shell and Tube Exchanger shell Body.
'   In this Symbol there are total 11 outputs.9 outputs are physical ,out of which 6 are pipe Nozzles.
'   Nozzles are fully parametric.
'   There are 2 Reference Geometry objects which are Control-Point at origin and a reference plane at Origin.
'   This is a Dow symbol Equipment
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------    -----            ------------------
'   09.Jul.2003  SymbolTeam(India)   Copyright Information, Header  is added.
'   22.Aug.2003  SymbolTeam(India)   TR 46728 Change Port Index of a nozzle to 15 used to corrupt Equipment.
'                                    Modified port index logic.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private PI As Double
Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    PI = Atn(1) * 4
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
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim parVesselDiameter As Double
    Dim parCPtoFace As Double
    Dim parCPtoShellTangent As Double
    Dim parFaceDiameter As Double
    Dim parFaceThickness As Double
    
    Dim iOutput     As Double
    
    Dim ObjRearHeadBonnet As Object
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parCPtoFace = arrayOfInputs(3)
    parCPtoShellTangent = arrayOfInputs(4)
    parFaceDiameter = arrayOfInputs(5)
    parFaceThickness = arrayOfInputs(6)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    Dim dActualVesselDiamter As Double
    Dim dDomeHeight As Double

'    'Assumptions
    'The input parameter for Vessel Diameter is Internal diameter .We need to take care for Thickness of Body.
    dActualVesselDiamter = parVesselDiameter + 0.0508   ' 2 Inch
    dDomeHeight = dActualVesselDiamter / 4  '2:1 ELiptical Head (TYP)
    
' Insert your code for output 1(ObjShellHeadFlange)As a Cylinder
    oStPoint.Set parCPtoFace, 0, 0
    oEnPoint.Set parCPtoFace + parFaceThickness, 0, 0
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parFaceDiameter
    
' Insert your code for output 2(ObjShellBodyCyl)As a Cylinder
    oStPoint.Set parCPtoFace + parFaceThickness, 0, 0
    oEnPoint.Set parCPtoShellTangent, 0, 0
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dActualVesselDiamter

' Insert your code for outputs 3(ObjRearHeadBonnet)
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Dim oAxisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
   Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim norX As Double
    Dim norY As Double
    Dim norZ As Double
    Dim Linepts(0 To 5) As Double
    
    Dim oExchangerArc1 As IngrGeom3D.EllipticalArc3d
    
    cenX = parCPtoShellTangent
    cenY = 0
    cenZ = 0

    MajorX = 0
    MajorY = 0
    MajorZ = dActualVesselDiamter / 2

    mMRatio = dDomeHeight / (dActualVesselDiamter / 2)
    StartAngle = PI / 2
    SweepAngle = PI / 2
    
    norX = 0
    norY = 1
    norZ = 0
    
    Set oExchangerArc1 = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                            cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                            StartAngle, SweepAngle)
                                            
    oAxisVect.Set 1, 0, 0
    revCenPt.Set 0, 0, 0
                        
    'Revolve it about X-Axiz
    Set ObjRearHeadBonnet = PlaceRevolution(m_OutputColl, oExchangerArc1, oAxisVect, revCenPt, PI * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRearHeadBonnet
    Set ObjRearHeadBonnet = Nothing
    
    Dim Objcurves As IJDObject
    Set Objcurves = oExchangerArc1
    Objcurves.Remove
    
    Set Objcurves = Nothing
    Set revCenPt = Nothing
    
   
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oAxisVect = Nothing
    Set oGeomFactory = Nothing
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
