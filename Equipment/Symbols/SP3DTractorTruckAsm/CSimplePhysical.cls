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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:18:09 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          KV
'   Creation Date:  Monday, Dec 9 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim geomFactory           As New IngrGeom3D.GeometryFactory

    Dim StartPoints(0 To 16)     As New AutoMath.DPosition
    Dim EndPoints(0 To 16)       As New AutoMath.DPosition
    Dim StartPointsBox(0 To 29)  As New AutoMath.DPosition
    Dim EndPointsBox(0 To 29)    As New AutoMath.DPosition
    
    Dim iOutput               As Double
    Dim Diameter(0 To 16)     As Double
    Dim domeDiameter          As Double
    
    Dim Counter               As Integer
    Dim Pass                  As Integer
    Dim iCount                As Integer
    
    Dim objCyli               As Object
    Dim objBox                As Object
    Dim objDome               As Object
    Dim objPyra               As Object
    Dim objTorus              As Object

    Dim oriPartOrientation    As Orientation
    Set oriPartOrientation = New Orientation

    iOutput = 0

' Insert your code for outputs 1-17(Cyli)
'   For iCount = 1 To 17
    StartPoints(0).Set 1.2668, -1.1938, -0.8176    'rear tire
    StartPoints(1).Set 1.2668, 0.1524, -0.8176     'rear tire
    StartPoints(2).Set -0.5905, -1.1938, -0.8176   'rear tire
    StartPoints(3).Set -0.5905, 0.1524, -0.8176    'rear tire
    StartPoints(4).Set 0.6731, -1.1938, -0.808     'axle
    StartPoints(5).Set 0.6731, 0.1349, -0.808      'axle
    StartPoints(6).Set 0.9112, 3.51, -0.8176       'front tire
    StartPoints(7).Set -0.608, 3.51, -0.816        'front tire
    StartPoints(8).Set 0.789, 0.808, -0.6064       'gas tank
    StartPoints(9).Set -0.735, 0.808, -0.6064      'gas tank
    StartPoints(10).Set -0.0238, 0.1445, 0.0571    'hitch
    StartPoints(11).Set 0, -0.0183, 0.1008         'hitch
    StartPoints(12).Set 0.6445, 1.5399, 1.4724     'Exhaust
    StartPoints(13).Set -1.0123971, 2.694236, 0.51276    'mirror
    StartPoints(14).Set -1.01759, 2.68446, 0.89059      'mirror
    StartPoints(15).Set 0.74136, 2.69399, 0.89059      'mirror
    StartPoints(16).Set 0.751253, 2.7832966, 0.51276   'mirror
     
    EndPoints(0).Set 0.6318, -1.1938, -0.8176      'rear tire
    EndPoints(1).Set 0.6318, 0.1524, -0.8176       'rear tire
    EndPoints(2).Set -1.2256, -1.1938, -0.8176     'rear tire
    EndPoints(3).Set -1.2256, 0.1524, -0.8176      'rear tire
    EndPoints(4).Set -0.6731, -1.1938, -0.808      'axle
    EndPoints(5).Set -0.6731, 0.1349, -0.808       'axle
    EndPoints(6).Set 0.6318, 3.51, -0.8176         'front tire
    EndPoints(7).Set -0.8874, 3.51, -0.816         'front tire
    EndPoints(8).Set 0.789, 1.4176, -0.6064        'gas tank
    EndPoints(9).Set -0.735, 1.4176, -0.6064       'gas tank
    EndPoints(10).Set -0.0238, 0.1445, -0.2476     'hitch
    EndPoints(11).Set 0, -0.0007, 0.0008           'hitch
    EndPoints(12).Set 0.6445, 1.5399, -0.7771      'Exhaust
    EndPoints(13).Set -0.7179829, 2.773124, 0.51276   'mirror
    EndPoints(14).Set -0.71279, 2.68446, 0.89059     'mirror
    EndPoints(15).Set 1.04616, 2.69399, 0.89059      'mirror
    EndPoints(16).Set 1.039447, 2.6840634, 0.51276   'mirror
    

    Diameter(0) = 1.02872     'rear tire
    Diameter(1) = 1.02872     'rear tire
    Diameter(2) = 1.02872     'rear tire
    Diameter(3) = 1.02872     'rear tire
    Diameter(4) = 0.0762      'axle
    Diameter(5) = 0.0762      'axle
    Diameter(6) = 0.9         'front tire
    Diameter(7) = 0.9         'front tire
    Diameter(8) = 0.6096      'gas tank
    Diameter(9) = 0.6096      'gas tank
    Diameter(10) = 0.508      'hitch
    Diameter(11) = 1.016      'hitch
    Diameter(12) = 0.1016     'Exhaust
    Diameter(13) = 0.01905    'mirror
    Diameter(14) = 0.01905    'mirror
    Diameter(15) = 0.01905    'mirror
    Diameter(16) = 0.01905    'mirror
    
   For Counter = 0 To 16
      Set objCyli = PlaceCylinder(m_outputColl, StartPoints(Counter), EndPoints(Counter), Diameter(Counter), True)
   
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objCyli
      Set objCyli = Nothing
   Next Counter
    
'Insert your code for outputs 18-47(boxes)
'  For iCount = 18 To 47
'The first three are the only rotatable boxes
   StartPointsBox(0).Set 0.7715, 2.6185, 1.1092       'windshield
   StartPointsBox(1).Set 1.0528, 2.6848, 0.9017       'passenger mirror
   StartPointsBox(2).Set -0.8745, 2.7156, 0.9017      'driver mirror
   StartPointsBox(3).Set 0.9827, 1.4407, -0.7271
   StartPointsBox(4).Set 1.1874, -1.7923, -0.1905
   StartPointsBox(5).Set -0.5652, -1.7923, -0.1905
   StartPointsBox(6).Set 0.7382, 1.7447, 0.327
   StartPointsBox(7).Set -0.7636, 1.7447, 0.327
   StartPointsBox(8).Set 0.3675, 3.0155, 0.3667
   StartPointsBox(9).Set 1.1874, -1.778, -0.2532
   StartPointsBox(10).Set -0.5652, -1.778, -0.2532
   StartPointsBox(11).Set 0.7858, 1.6962, 1.047
   StartPointsBox(12).Set 1.0858, 4.5418, -0.8763
   StartPointsBox(13).Set -0.7699, 1.4414, -0.7525
   StartPointsBox(14).Set -0.7668, 4.4053, -0.0397
   StartPointsBox(15).Set 1.0303, 4.4053, -0.0397
   StartPointsBox(16).Set 0.7858, 2.6233, 1.1224
   StartPointsBox(17).Set -0.7128, 1.8018, 1.0462
   StartPointsBox(18).Set -0.7668, 3.6274, -0.0349
   StartPointsBox(19).Set 1.0319, 3.6259, -0.0349
   StartPointsBox(20).Set 1.1874, 1.3875, -0.1984
   StartPointsBox(21).Set 0.6795, 3.1686, -0.7803
   StartPointsBox(22).Set -0.7572, 3.1686, -0.7803
   StartPointsBox(23).Set -0.7668, 3.6211, -0.0429
   StartPointsBox(24).Set 1.0319, 3.6211, -0.0429
   StartPointsBox(25).Set 0.789, 3.0083, 0.3675
   StartPointsBox(26).Set -0.7096, 3.0083, 0.3675
   StartPointsBox(27).Set -0.7128, 2.8115, 0.327
   StartPointsBox(28).Set 0.7787, 4.4053, 0.0143
   StartPointsBox(29).Set 0.789, 1.7256, 0.3945
   
'The first three are the only rotatable boxes
   EndPointsBox(0).Set -0.7525, 2.8075, 0.3671        'windshield
   EndPointsBox(1).Set 0.8744, 2.6969, 0.4699         'passenger mirror
   EndPointsBox(2).Set -1.0463, 2.666, 0.4699         'driver mirror
   EndPointsBox(3).Set 0.7922, 2.798, -0.7779
   EndPointsBox(4).Set 0.5778, 1.3827, -0.2159
   EndPointsBox(5).Set -1.1747, 1.3827, -0.2159
   EndPointsBox(6).Set 0.789, 2.8115, -0.7144
   EndPointsBox(7).Set -0.7128, 2.8115, -0.7144
   EndPointsBox(8).Set -0.3516, 4.3918, 0.2651
   EndPointsBox(9).Set 0.5778, -1.7907, -0.947
   EndPointsBox(10).Set -1.1747, -1.7907, -0.947
   EndPointsBox(11).Set -0.7636, 1.5931, -0.7517
   EndPointsBox(12).Set -1.0985, 4.4656, -0.5715
   EndPointsBox(13).Set -0.9604, 2.8131, -0.8033
   EndPointsBox(14).Set -1.0081, 4.1005, -0.7795
   EndPointsBox(15).Set 0.789, 4.1005, -0.7795
   EndPointsBox(16).Set -0.7382, 1.5931, 1.0462
   EndPointsBox(17).Set -0.7636, 1.6113, -0.7541
   EndPointsBox(18).Set -1.0081, 4.3672, -0.2889
   EndPointsBox(19).Set 0.7906, 4.3656, -0.2889
   EndPointsBox(20).Set -1.1747, 1.3367, 0.4112
   EndPointsBox(21).Set 0.7811, 4.3879, -0.0421
   EndPointsBox(22).Set -0.6556, 4.3879, -0.0421
   EndPointsBox(23).Set -1.0081, 3.5195, -0.2461
   EndPointsBox(24).Set 0.7906, 3.5195, -0.2461
   EndPointsBox(25).Set 0.7382, 2.8178, -0.7803
   EndPointsBox(26).Set -0.7604, 2.8178, -0.7803
   EndPointsBox(27).Set -0.7636, 1.8209, -0.7144
   EndPointsBox(28).Set -0.7533, 4.3799, -0.7795
   EndPointsBox(29).Set 0.7382, 1.5351, -0.7533
   
'for the windshield
   For Counter = 0 To 0
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = 90
    oriPartOrientation.RotationAboutX = 20
    oriPartOrientation.ApplyRotations
    Set objBox = CreateBox(m_outputColl, "e 0.0095200 n 2.7130400 u 0.7381900", oriPartOrientation, 0.076195, 1.524, 0.7619948, True)

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox
      Set objBox = Nothing
   Next Counter
   
'passenger mirror
   For Counter = 1 To 1
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -10
    oriPartOrientation.RotationAboutX = 5
    oriPartOrientation.ApplyRotations
    Set objBox = CreateBox(m_outputColl, "e 0.9604400 n 2.6908100 u 0.6858000", oriPartOrientation, 0.1778, 0.01905, 0.4318, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox
      Set objBox = Nothing
   Next Counter
   
'driver mirror
   For Counter = 2 To 2
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = 10
    oriPartOrientation.RotationAboutX = 5
    oriPartOrientation.ApplyRotations
    Set objBox = CreateBox(m_outputColl, "w 0.9604400 n 2.6908100 u 0.6858000", oriPartOrientation, 0.1778, 0.01905, 0.4318, True)

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox
      Set objBox = Nothing
   Next Counter
   
   For Counter = 3 To 29
      Set objBox = PlaceBox(m_outputColl, StartPointsBox(Counter), EndPointsBox(Counter))

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objBox
      Set objBox = Nothing
   Next Counter
 
' Insert your code for output 48 to 49 (ObjDomes)

'       For iCount = 48 To 49
        Dim EllipNormX     As Double, EllipNormY As Double, EllipNormZ As Double
        Dim majorX         As Double, majorY As Double, majorZ As Double
        Dim Axis           As New AutoMath.DVector
        Dim PI             As Double
        Dim parRatio       As Double
        Dim oDomeEllip     As IngrGeom3D.EllipticalArc3d
        Dim DishStart(0 To 1)      As New AutoMath.DPosition
        
        domeDiameter = 0.1778

        DishStart(0).Set 0.9033, 4.399, -0.2572
        DishStart(1).Set -0.8922, 4.399, -0.2572

PI = 4 * Atn(1)
      For Counter = 0 To 1
        Axis.Set 0, 1, 0
        majorX = 0
        majorY = 0
        majorZ = domeDiameter / 2
        'The Ratio of minor axis over major axis
        parRatio = (0.0508 / (0.1778 / 2)) 'domeHeight / (domeDiameter /2)
        'Normal vector to plane of ellipce)
        EllipNormX = -1
        EllipNormY = 0
        EllipNormZ = 0

       Set oDomeEllip = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle( _
                                                Nothing, DishStart(Counter).x, _
                                                DishStart(Counter).y, DishStart(Counter).z, EllipNormX, _
                                                EllipNormY, EllipNormZ, majorX, majorY, majorZ, _
                                                parRatio, 0, PI / 2)

        Set objDome = PlaceRevolution(m_outputColl, oDomeEllip, Axis, DishStart(Counter), PI * 2, False)

'  Set the output
        iOutput = iOutput + 1
        m_outputColl.AddOutput arrayOfOutputs(iOutput), objDome
        Set objDome = Nothing
      Next Counter
'End Dome Construction

' Begin construction of RTor Driver side fender
    
    iCount = 50
    
   Set objTorus = CreateRectangularTorus(m_outputColl, "W 0.88741 N 3.51949 D 0.77946", 0.5334, 0.7366, 180, 0, 0.2413, "N 0 U 90", "N 0 U 0")
   iOutput = iOutput + 1
   m_outputColl.AddOutput arrayOfOutputs(iOutput), objTorus

   Set objTorus = Nothing
   
' Begin construction of RTor Passenger side fender
   iCount = 51
    
   Set objTorus = CreateRectangularTorus(m_outputColl, "E 0.91123 N 3.51949 D 0.77946", 0.5334, 0.7366, 180, 0, 0.2413, "N 0 U 90", "N 0 U 0")
   iOutput = iOutput + 1
   m_outputColl.AddOutput arrayOfOutputs(iOutput), objTorus

   Set objTorus = Nothing
     
' Begin construction of RTor (Driver side hood)
   iCount = 52
    
   Set objTorus = CreateRectangularTorus(m_outputColl, "W 0.35084 N 3.70364 D 0.04128", 0.03048, 0.4064, 90, 0, 1.3716, "W 270 U 0", "N 0 D -90")
   iOutput = iOutput + 1
   m_outputColl.AddOutput arrayOfOutputs(iOutput), objTorus

   Set objTorus = Nothing

' Begin construction of RTor (Passenger side Hood)
   iCount = 53
       
   Set objTorus = CreateRectangularTorus(m_outputColl, "E 0.37465 N 3.70205 D 0.04128", 0.3048, 0.4064, 90, 0, 1.3716, "E 90 U 0", "N 0 D -90")
   iOutput = iOutput + 1
   m_outputColl.AddOutput arrayOfOutputs(iOutput), objTorus
   
   Set objTorus = Nothing

 '  Insert your code for output (Back of hood Pyramid)
   iCount = 54
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutY = -180
    oriPartOrientation.RotationAboutX = -180
    oriPartOrientation.ApplyRotations
    Set objPyra = CreatePyramid(m_outputColl, "E 0.0079400 N 2.9114800 U 0.4048100", oriPartOrientation, 1.524, 0.1905, 1.524, 0.0047625, 0, 0.0508, 0.0762)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objPyra
    Set objPyra = Nothing
   
' Output (Control Point)
    iCount = 55
    
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim centerPoint As Point3d
    
    
    Set centerPoint = geomFactory.Points3d.CreateByPoint(m_outputColl.ResourceManager, 0, 0, 0)

    '  Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), centerPoint
    Set centerPoint = Nothing
    Set geomFactory = Nothing
   
   Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

