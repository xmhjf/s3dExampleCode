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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:17:29 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:        MS
'   Creation Date: Wednesday, Feb 18 2004
'   Description:
'    The symbol is prepared from PDS Equipment guide E210. The symbol consists of 3 Physical outputs. The skirt is variable output.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl       TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
 Private PI As Double
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
     
    Dim iOutput     As Double
    Dim ObjVesselTopDome As Object
    Dim ObjVesselBottomDome  As Object
    Dim objVesselBody As Object
    
    
    Dim parVesselHeight As Double
    Dim parVesselDiameter As Double
    Dim parSkirtTopDiameter As Double
    Dim parSkirtBottomDiameter As Double
    Dim parSkirtHeight As Double
    Dim parVesselStartHeight As Double
    Dim parInsulationThickness  As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight = arrayOfInputs(2)
    parVesselDiameter = arrayOfInputs(3)
    parSkirtTopDiameter = arrayOfInputs(4)
    parSkirtBottomDiameter = arrayOfInputs(5)
    parSkirtHeight = arrayOfInputs(6)
    parVesselStartHeight = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
   
    iOutput = 0
        
'' Origin is at Mid of Bottom plane of Skirt.
        
 ' Insert your code for output 1(Vessel Top Dome)
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
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oEqpArcTop As IngrGeom3D.EllipticalArc3d
    Dim oEqpArcBottom As IngrGeom3D.EllipticalArc3d
    
' Vessel is formed by creating a Complex string and rotating about Z-Axis
' Top Convex Dome Curve
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
                                                        
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    axisVect.Set 0, 0, 1
    revCenPt.Set 0, 0, (parVesselHeight / 2 + parVesselStartHeight)
    'Revolve it about Z-Axiz

    Set ObjVesselTopDome = PlaceRevolution(m_OutputColl, oEqpArcTop, axisVect, revCenPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselTopDome
    Set ObjVesselTopDome = Nothing
    Set oEqpArcTop = Nothing
  
' Insert your code for output 2(Vessel Bottom Dome)
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

    Set ObjVesselBottomDome = PlaceRevolution(m_OutputColl, oEqpArcBottom, axisVect, revCenPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselBottomDome
    Set ObjVesselBottomDome = Nothing
    Set oEqpArcBottom = Nothing
    Set axisVect = Nothing
    Set revCenPt = Nothing
 
 ' Insert your code for output 3(Vessel Body)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    stPoint.Set 0, 0, parVesselStartHeight
    enPoint.Set 0, 0, parVesselStartHeight + parVesselHeight
    
    Set objVesselBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parVesselDiameter, True)
    
   ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objVesselBody
    Set objVesselBody = Nothing
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
' Insert your code for output 4(Truncated Cone Skirt)
If (CmpDblGreaterthan(parVesselStartHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSkirtTopDiameter, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSkirtBottomDiameter, LINEAR_TOLERANCE)) Then
    Dim objSkirt As Object
    Dim baseCenter As New AutoMath.DPosition
    Dim topCenter As New AutoMath.DPosition
    
    topCenter.Set 0, 0, parSkirtHeight
    baseCenter.Set 0, 0, 0

    Set objSkirt = PlaceCone(m_OutputColl, baseCenter, topCenter, parSkirtBottomDiameter / 2, parSkirtTopDiameter / 2, False)
    
   ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput "SupportSkirt", objSkirt
    Set objSkirt = Nothing
    Set baseCenter = Nothing
    Set topCenter = Nothing
End If
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub




