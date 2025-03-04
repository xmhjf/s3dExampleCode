VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:09 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Tuesday, Feb 17 2004
'   Description:
'       This is a Simple Vertical Cylider Equipment component with Legs. This is developed based on PDS E-215.
'       This is a variable out put symbol. Number of Support legs is a variable.
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double
Private Sub Class_Initialize()
        PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim ObjVesselTopEllipBodyIns As Object
    Dim ObjVesselCylBodyIns As Object
    Dim ObjVesselBottomEllipBodyIns As Object
        
    Dim parVesselHeight As Double
    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports  As Long
    Dim parSupportHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parVesselStartPoint  As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight = arrayOfInputs(2)  'P2
    parVesselDiameter = arrayOfInputs(3)    'P3
    parSupportAngularLocation = arrayOfInputs(4)  'P4
    parNumberOfSupports = arrayOfInputs(5)  'P5
    parSupportHeight = arrayOfInputs(6)     'P6
    parSupportLength = arrayOfInputs(7)     'P7
    parSupportThickness = arrayOfInputs(8)  'P8
    parVesselStartPoint = arrayOfInputs(9)      'P9
    parInsulationThickness = arrayOfInputs(10)
    
     Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

 ' Insert your code for output 1(Vessel Top Elliptical Body)
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
    Dim oEllipticalArc As IngrGeom3D.EllipticalArc3d
    
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition


' Vessel is formed by creating a Complex string and rotating about Z-Axis
'Top Convex Dome Curve
    cenX = 0
    cenY = 0
    cenZ = parVesselHeight + parVesselStartPoint
    MajorX = parVesselDiameter / 2 + parInsulationThickness
    MajorY = 0
    MajorZ = 0
    mMRatio = (parVesselDiameter / 4 + parInsulationThickness) / (parVesselDiameter / 2 + parInsulationThickness)
    StartAngle = PI / 2
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    axisVect.Set 0, 0, 1
    revCenPt.Set 0, 0, parVesselHeight + parVesselStartPoint

    Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)

    Set ObjVesselTopEllipBodyIns = PlaceRevolution(m_OutputColl, oEllipticalArc, axisVect, revCenPt, PI * 2, False)


' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselTopEllipBodyIns
    Set ObjVesselTopEllipBodyIns = Nothing
    Set oEllipticalArc = Nothing
    
 
 ' Insert your code for output 2(Vessel Cylindrical Body Insulation)
   Dim oStartPoint As New AutoMath.DPosition
   Dim oEndPoint As New AutoMath.DPosition
   
   oStartPoint.Set 0, 0, parVesselStartPoint
   oEndPoint.Set 0, 0, parVesselHeight + parVesselStartPoint
   
   Dim dInsulationDiameter As Double
   dInsulationDiameter = parVesselDiameter + 2 * parInsulationThickness
   
    Set ObjVesselCylBodyIns = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, _
                                                                 dInsulationDiameter, False)
  
   ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselCylBodyIns
    Set ObjVesselCylBodyIns = Nothing
    
  ' Insert your code for output 3(Vessel Bottom Elliptical Body Insulation)
     
    'Bottom Dome Curve
    cenX = 0
    cenY = 0
    cenZ = parVesselStartPoint
    MajorX = parVesselDiameter / 2 + parInsulationThickness
    MajorY = 0
    MajorZ = 0
    mMRatio = (parVesselDiameter / 4 + parInsulationThickness) / (parVesselDiameter / 2 + parInsulationThickness)
    StartAngle = PI
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)

    revCenPt.Set 0, 0, parVesselStartPoint
  
    Set ObjVesselBottomEllipBodyIns = PlaceRevolution(m_OutputColl, oEllipticalArc, axisVect, revCenPt, PI * 2, False)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselBottomEllipBodyIns
    Set ObjVesselBottomEllipBodyIns = Nothing
    Set oEllipticalArc = Nothing
    
    Set axisVect = Nothing
    Set revCenPt = Nothing
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
        
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
