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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:08 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         ACM
'   Creation Date:  Tuesday, Feb 17 2004
'   Description:
'       This is a Simple Vertical Cylider Equipment component with Legs. This is developed based on PDS E-215.
'       This is a variable out put symbol. Number of Support legs is a variable.
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl       TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
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
    Dim iOutput     As Double
        
    Dim ObjVesselTopEllipBody As Object
    Dim ObjVesselCylBody As Object
    Dim ObjVesselBottomEllipBody As Object
    
    Dim parVesselHeight As Double
    Dim parVesselDiameter As Double
    Dim parSupportAngularLocation As Double
    Dim parNumberOfSupports  As Long
    Dim parSupportHeight As Double
    Dim parSupportLength As Double
    Dim parSupportThickness As Double
    Dim parVesselStartPoint  As Double
    Dim parInsulationThickness As Double
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselHeight = arrayOfInputs(2)                      'P2
    parVesselDiameter = arrayOfInputs(3)                    'P3
    parSupportAngularLocation = arrayOfInputs(4)       'P4
    parNumberOfSupports = arrayOfInputs(5)      'P5
    parSupportHeight = arrayOfInputs(6)             'P6
    parSupportLength = arrayOfInputs(7)            'P7
    parSupportThickness = arrayOfInputs(8)      'P8
    parVesselStartPoint = arrayOfInputs(9)         'P9
    parInsulationThickness = arrayOfInputs(10)
    
    iOutput = 0
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
''Origin is at Place Point 1.
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

' Vessel Top Dome is formed by creating a Elliptical arc and rotating about Z-Axis
'Top Convex Dome Curve
    cenX = 0
    cenY = 0
    cenZ = parVesselHeight + parVesselStartPoint
    MajorX = parVesselDiameter / 2
    MajorY = 0
    MajorZ = 0
    mMRatio = 0.5
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
                                                        
    Set ObjVesselTopEllipBody = PlaceRevolution(m_OutputColl, oEllipticalArc, axisVect, revCenPt, PI * 2, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselTopEllipBody
    Set ObjVesselTopEllipBody = Nothing
    Set oEllipticalArc = Nothing
 
 ' Insert your code for output 2(Vessel Cylindrical Body)
   Dim oStartPoint As New AutoMath.DPosition
   Dim oEndPoint As New AutoMath.DPosition
   
   oStartPoint.Set 0, 0, parVesselStartPoint
   oEndPoint.Set 0, 0, parVesselStartPoint + parVesselHeight
   
    Set ObjVesselCylBody = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parVesselDiameter, False)
   
   ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselCylBody
    Set ObjVesselCylBody = Nothing
    
  ' Insert your code for output 3(Vessel Bottom Elliptical Body)
     
    'Bottom Dome Curve
    cenX = 0
    cenY = 0
    cenZ = parVesselStartPoint
    MajorX = parVesselDiameter / 2
    MajorY = 0
    MajorZ = 0
    mMRatio = 0.5
    StartAngle = PI
    SweepAngle = PI / 2
    norX = 0
    norY = -1
    norZ = 0
    Set oEllipticalArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    
    revCenPt.Set 0, 0, parVesselStartPoint
  
    Set ObjVesselBottomEllipBody = PlaceRevolution(m_OutputColl, oEllipticalArc, axisVect, revCenPt, PI * 2, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVesselBottomEllipBody
    Set ObjVesselBottomEllipBody = Nothing
    Set oEllipticalArc = Nothing
    
    Set axisVect = Nothing
    Set revCenPt = Nothing
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing

' Insert your code for output (Leg Support )

    Dim IncludedAng As Double
    Dim dYaxisToSupportAngle As Double
    Dim iCount As Integer
    If (parNumberOfSupports > 0) And CmpDblGreaterthan(parSupportHeight, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) _
                                                    And CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE) Then
            IncludedAng = (PI * 2) / parNumberOfSupports
            
            dYaxisToSupportAngle = parSupportAngularLocation
            
            Dim topSurPts(0 To 3) As IJDPosition
            Dim botSurPts(0 To 3) As IJDPosition
            
            For iCount = 0 To 3
                Set topSurPts(iCount) = New DPosition
                Set botSurPts(iCount) = New DPosition
            Next iCount
            
            For iCount = 0 To parNumberOfSupports - 1
                    'form Line String and Project it
                    'Point 1
                    botSurPts(0).x = Sin(dYaxisToSupportAngle) * (parVesselDiameter / 2) + Cos(dYaxisToSupportAngle) * parSupportLength / 2
                    botSurPts(0).y = Cos(dYaxisToSupportAngle) * (parVesselDiameter / 2) - Sin(dYaxisToSupportAngle) * parSupportLength / 2
                    botSurPts(0).z = 0
                    
                    'Point 2
                    botSurPts(1).x = botSurPts(0).x + Sin(dYaxisToSupportAngle) * parSupportThickness
                    botSurPts(1).y = botSurPts(0).y + parSupportThickness * Cos(dYaxisToSupportAngle)
                    botSurPts(1).z = botSurPts(0).z
                    
                    'Point 4
                    botSurPts(3).x = Sin(dYaxisToSupportAngle) * (parVesselDiameter / 2) - Cos(dYaxisToSupportAngle) * parSupportLength / 2
                    botSurPts(3).y = Cos(dYaxisToSupportAngle) * (parVesselDiameter / 2) + Sin(dYaxisToSupportAngle) * parSupportLength / 2
                    botSurPts(3).z = botSurPts(0).z
                    
                     'Point3
                    botSurPts(2).x = botSurPts(3).x + parSupportThickness * Sin(dYaxisToSupportAngle)
                    botSurPts(2).y = botSurPts(3).y + parSupportThickness * Cos(dYaxisToSupportAngle)
                    botSurPts(2).z = botSurPts(0).z
                    
                    topSurPts(0).Set botSurPts(0).x, botSurPts(0).y, botSurPts(0).z + parSupportHeight
                    topSurPts(1).Set botSurPts(1).x, botSurPts(1).y, topSurPts(0).z
                    topSurPts(2).Set botSurPts(2).x, botSurPts(2).y, topSurPts(0).z
                    topSurPts(3).Set botSurPts(3).x, botSurPts(3).y, topSurPts(0).z
                    
                    Dim ObjCollection As Collection
                    
                    Set ObjCollection = PlaceTrapezoidWithPlanes(m_OutputColl, topSurPts, botSurPts)
                    
                    m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane1"), ObjCollection(1)
                    m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane2"), ObjCollection(2)
                    m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane3"), ObjCollection(3)
                    m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane4"), ObjCollection(4)
                    m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane5"), ObjCollection(5)
                    m_OutputColl.AddOutput ("Support" & CStr(iCount + 1) & "Plane6"), ObjCollection(6)
            
                    Set ObjCollection = Nothing
            
                    'Increment the Angle
                    dYaxisToSupportAngle = dYaxisToSupportAngle + IncludedAng
            Next iCount
            
            For iCount = 0 To 3
                Set topSurPts(iCount) = Nothing
                Set botSurPts(iCount) = Nothing
            Next iCount

End If

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext

End Sub
