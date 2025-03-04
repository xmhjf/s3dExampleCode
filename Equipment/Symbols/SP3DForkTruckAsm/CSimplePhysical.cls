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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:05:22 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          nka8226
'   Creation Date:  Friday, Nov 22 2002
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
    
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'   NOTE: All distances, sizes and coordinates are in meters

'   --- some fork truck constants used in multiple output sections ---
    Const WHEEL_THICKNESS  As Double = 0.1524
    Const WHEEL_DIAMETER   As Double = 0.381
        
'   --- general purpose item declarations ---
    Dim iOutput              As Double
    Dim projDistance         As Double
    Dim i                    As Integer
    Dim nVertices            As Integer
    
    Dim objGeomFactory       As IngrGeom3D.GeometryFactory
    Dim startPoint           As IJDPosition
    Dim endPoint             As IJDPosition
    Dim posCylinderCenter    As IJDPosition
    Dim polygonVertices()    As IJDPosition
    Dim vecProjDir           As IJDVector

'   --- fork truck item declarations ---
    Dim objForkTruckPart         As Object
    Dim posForkTruckControlPoint As IJDPosition
    Dim oriPartOrientation       As Orientation
            
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

    Set objGeomFactory = New IngrGeom3D.GeometryFactory
    Set vecProjDir = New DVector
    
    Set startPoint = New DPosition
    Set endPoint = New DPosition
    
    Set posForkTruckControlPoint = New DPosition
    
    Set posCylinderCenter = New DPosition
    
    Set oriPartOrientation = New Orientation
    
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'   --- set fork truck control point origin ---
    posForkTruckControlPoint.Set 0, 0, 0
    
'   --- do the outputs ---
    iOutput = 0

'   ---------
'   WHEELS
'   ---------
' Insert your code for output 1(Left Front Wheel)
'   --- create a cylinder to model left front tire wheel ---
    Set objForkTruckPart = CreateCylinder(m_outputColl, "w 0.4445 n 0.77311 u 0.1905", WHEEL_DIAMETER, WHEEL_THICKNESS, "E 90 U 0")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 2(Right Front Wheel)
'   --- create a cylinder to model right front tire wheel ---
    Set objForkTruckPart = CreateCylinder(m_outputColl, "e 0.4445 n 0.77311 u 0.1905", WHEEL_DIAMETER, WHEEL_THICKNESS, "E 90 U 0")

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    
    

' Insert your code for output 3(Left Rear Wheel)
'   --- create a cylinder to model left rear tire wheel ---
    Set objForkTruckPart = CreateCylinder(m_outputColl, "w 0.4445 s 0.54769 u 0.1905", WHEEL_DIAMETER, WHEEL_THICKNESS, "E 90 U 0")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 4(Right Rear Wheel)
'   --- create a cylinder to model right rear tire wheel ---
    Set objForkTruckPart = CreateCylinder(m_outputColl, "e 0.4445 s 0.54769 u 0.1905", WHEEL_DIAMETER, WHEEL_THICKNESS, "E 90 U 0")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    
    

'   ---------------
'   FORK TRUCK BODY
'   ---------------
' Insert your code for output 5(Main Body Part)
'   --- create projected polygon to model main body part ---
    nVertices = 16
    ReDim polygonVertices(0 To (nVertices - 1)) As IJDPosition
   
'   --- instantiate each array item ---
    For i = 0 To (nVertices - 1)
       Set polygonVertices(i) = New DPosition
    Next i

'   --- load the polygon vertices array with x,y,z data ---
    polygonVertices(0).Set -0.5207, -0.76359, 1.03981
    polygonVertices(1).Set -0.5207, -0.48419, 1.03981
    polygonVertices(2).Set -0.5207, -0.48419, 0.99219
    polygonVertices(3).Set -0.5207, 0.17621, 0.99219
    polygonVertices(4).Set -0.5207, 0.17621, 0.635
    polygonVertices(5).Set -0.5207, 0.61222, 0.635
    polygonVertices(6).Set -0.5207, 0.55289, 0.9715
    polygonVertices(7).Set -0.5207, 0.70298, 0.99796
    polygonVertices(8).Set -0.5207, 0.76698, 0.635
    polygonVertices(9).Set -0.5207, 0.92076, 0.635
    polygonVertices(10).Set -0.5207, 0.92076, 0.4318
    polygonVertices(11).Set -0.5207, 0.592653, 0.4318
    polygonVertices(12).Set -0.5207, 0.4826, 0.04763
    polygonVertices(13).Set -0.5207, -0.2794, 0.04763
    polygonVertices(14).Set -0.5207, -0.389453, 0.4318
    polygonVertices(15).Set -0.5207, -0.76199, 0.4318
    
'   --- establish the direction of projection ---
    vecProjDir.Set 1, 0, 0
   
'   --- establish the projection distance ---
    projDistance = 1.0414

'   --- project the polygon ---
    Set objForkTruckPart = placeProjectedPolygonFromVertices(m_outputColl, polygonVertices(), vecProjDir, projDistance, True)

'   --- release each array item ---
    For i = 0 To (nVertices - 1)
       Set polygonVertices(i) = Nothing
    Next i
   
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    
    

' Insert your code for output 6(Rear Body Part)
'   --- create a polygon projection to model rear body part ---
    nVertices = 4
    ReDim polygonVertices(0 To (nVertices - 1)) As IJDPosition
    
'   --- instantiate each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = New DPosition
    Next i
    
'   --- load the polygon vertices array with x,y,z data ---
    polygonVertices(0).Set -0.5207, -0.76359, 1.03981
    polygonVertices(1).Set 0.5207, -0.76359, 1.03981
    polygonVertices(2).Set 0.1524, -0.94139, 1.03981
    polygonVertices(3).Set -0.1524, -0.94139, 1.03981

'   --- establish the direction of projection ---
    vecProjDir.Set 0, 0, -1
    
'   --- establish the projection distance ---
    projDistance = 0.9398

'   --- project the polygon ---
    Set objForkTruckPart = placeProjectedPolygonFromVertices(m_outputColl, polygonVertices(), vecProjDir, projDistance, True)

'   --- release each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = Nothing
    Next i
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    
    

'   -----------------
'   STEERING ASSEMBLY
'   -----------------
' Insert your code for output 7(Steering Column Housing)
'   --- create a rotated box to model steering column housing ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.RotationAboutX = 55
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "e 0 n 0.49848 u 1.10014", oriPartOrientation, 0.1524, 0.2286, 0.4572, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing



' Insert your code for output 8(Steering Column Shaft)
'   --- create a cylinder to model steering column shaft, tilted down 35 deg ---
    Set objForkTruckPart = CreateCylinder(m_outputColl, "e 0 n 0.2794 u 1.25254", 0.0635, 0.0762, "S 180 U 35")

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    
    

' Insert your code for output 9(Steering Wheel)
'   --- create a cylinder to model steering wheel, tilted down 35 deg ---
    Set objForkTruckPart = CreateCylinder(m_outputColl, "e 0 n 0.23971 u 1.28111", 0.3302, 0.01905, "S 180 U 35")

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    
    

'   -------------
'   SEAT ASSEMBLY
'   -------------
' Insert your code for output 10(Seat Bottom)
'   --- create a box to model seat bottom ---
    oriPartOrientation.ResetDefaultAxis
    Set objForkTruckPart = CreateBox(m_outputColl, "e 0 s 0.1524 u 1.03346", oriPartOrientation, 0.4572, 0.4826, 0.1016, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 11(Seat Back)
'   --- create a polygon projection to model seat back ---
    nVertices = 4
    ReDim polygonVertices(0 To (nVertices - 1)) As IJDPosition
    
'   --- instantiate each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = New DPosition
    Next i
    
'   --- load the polygon vertices array with x,y,z data ---
    polygonVertices(0).Set -0.1778, -0.47308, 1.397
    polygonVertices(1).Set 0.1778, -0.47308, 1.397
    polygonVertices(2).Set 0.2413, -0.47308, 0.9906
    polygonVertices(3).Set -0.2413, -0.47308, 0.9906

'   --- establish the direction of projection ---
    vecProjDir.Set 0, 1, 0
    
'   --- establish the projection distance ---
    projDistance = 0.1016

'   --- project the polygon ---
    Set objForkTruckPart = placeProjectedPolygonFromVertices(m_outputColl, polygonVertices(), vecProjDir, projDistance, True)

'   --- release each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = Nothing
    Next i
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


'   -------------
'   FORK ASSEMBLY
'   -------------
' Insert your code for output 12(Left Fork)
'   --- create a triangular projection to model left fork ---
    nVertices = 3
    ReDim polygonVertices(0 To (nVertices - 1)) As IJDPosition
    
'   --- instantiate each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = New DPosition
    Next i
    
'   --- load the polygon vertices array with x,y,z data ---
    polygonVertices(0).Set -0.38894, 2.3114, 0.06509
    polygonVertices(1).Set -0.38894, 1.2192, 0.0889
    polygonVertices(2).Set -0.38894, 1.2192, 0.0381

'   --- establish the direction of projection ---
    vecProjDir.Set 1, 0, 0
    
'   --- establish the projection distance ---
    projDistance = 0.1016

'   --- project the polygon ---
    Set objForkTruckPart = placeProjectedPolygonFromVertices(m_outputColl, polygonVertices(), vecProjDir, projDistance, True)

'   --- release each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = Nothing
    Next i
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 13(Right Fork)
'   --- create a triangular projection to model right fork ---
    nVertices = 3
    ReDim polygonVertices(0 To (nVertices - 1)) As IJDPosition
    
'   --- instantiate each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = New DPosition
    Next i
    
'   --- load the polygon vertices array with x,y,z data ---
    polygonVertices(0).Set 0.35719, 2.3114, 0.06509
    polygonVertices(1).Set 0.35719, 1.2192, 0.0889
    polygonVertices(2).Set 0.35719, 1.2192, 0.0381

'   --- establish the direction of projection ---
    vecProjDir.Set 1, 0, 0
    
'   --- establish the projection distance ---
    projDistance = 0.1016

'   --- project the polygon ---
    Set objForkTruckPart = placeProjectedPolygonFromVertices(m_outputColl, polygonVertices(), vecProjDir, projDistance, True)

'   --- release each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = Nothing
    Next i
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 14(Fork Upper Cross Over)
'   --- create a box to model fork upper cross over ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "e 0 n 0.92075 u 2.02089", oriPartOrientation, 0.2032, 0.7366, 0.1524, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 15(Fork Left Upright)
'   --- create a box to model fork left upright ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "w 0.29369 n 1.09538 u 1.08426", oriPartOrientation, 0.1524, 0.1524, 2.032, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 16(Fork Right Upright)
'   --- create a box to model fork right upright ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "e 0.29051 n 1.09538 u 1.08426", oriPartOrientation, 0.1524, 0.1524, 2.032, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 17(Fork Left Bumper)
'   --- create a box to model fork left bumper ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "w 0.33814 n 1.25095 u 0.32861", oriPartOrientation, 0.0508, 0.1016, 0.5334, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 18(Fork Right Bumper)
'   --- create a box to model fork right bumper ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "e 0.41116 n 1.25095 u 0.33496", oriPartOrientation, 0.0508, 0.1016, 0.5334, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 19(Fork Grill)
'   --- create a box to model fork grill ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "e 0.02699 n 1.20015 u 0.5842", oriPartOrientation, 0.0508, 1.1016, 1.0668, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
     
     

'   ----------------
'   DRIVER ROLL CAGE
'   ----------------
' Insert your code for output 20(left rear post)
'   --- create a box to model left rear post ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "w 0.46038 s 0.48419 u 1.61449", oriPartOrientation, 0.0635, 0.0381, 1.143, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing



' Insert your code for output 21(right rear post)
'   --- create a box to model right rear post ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "e 0.45244 s 0.48419 u 1.61449", oriPartOrientation, 0.0635, 0.0381, 1.143, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing



' Insert your code for output 22(left front post)
'   --- create a rotated box to model left front post ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.RotationAboutX = 10
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "w 0.46038 n 0.68421 u 1.41764", oriPartOrientation, 0.0635, 0.0381, 1.55575, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing



' Insert your code for output 23(right front post)
'   --- create a rotated box to model right front post ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.RotationAboutX = 10
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "e 0.46514 n 0.68263 u 1.41605", oriPartOrientation, 0.0635, 0.0381, 1.558925, True)
    
   ' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing



' Insert your code for output 24(left top frame piece)
'   --- create a box to model left top frame piece ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "w 0.46355 N 0.06191 u 2.15741", oriPartOrientation, 1.0302875, 0.0381, 0.0635, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing



' Insert your code for output 25(right top frame piece)
'   --- create a box to model right top frame piece ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "e 0.45085 N 0.06191 u 2.15741", oriPartOrientation, 1.0302875, 0.0381, 0.0635, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing



' Insert your code for output 26(front top frame piece)
'   --- create a box to model front top frame piece ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "w 0.00476 n 0.55404 u 2.15741", oriPartOrientation, 0.0381, 0.8747125, 0.0635, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing



' Insert your code for output 27(rear top frame piece)
'   --- create a box to model rear top frame piece ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutZ = -90
    oriPartOrientation.ApplyRotations
    Set objForkTruckPart = CreateBox(m_outputColl, "w 0.00476 s 0.48419 u 2.15424", oriPartOrientation, 0.0381, 0.8747125, 0.0635, True)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing



'   -------------
'   MUFFLER PARTS
'   -------------
' Insert your code for output 28(large muffler part)
    Set objForkTruckPart = CreateCylinder(m_outputColl, "w 0.12541 s 0.72866 u 1.2065", 0.3048, 0.508, "E 90 U 0")

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


' Insert your code for output 29(small muffler part)
'   --- create a cylinder to model small muffler part ---
    Set objForkTruckPart = CreateCylinder(m_outputColl, "e 0.21908 s 0.73025 u 1.2065", 0.2032, 0.1778, "E 90 U 0")

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


'Insert your code for output 30(Fork Truck Position Point)
'   --- create a persistent point at the fork truck control point position ---
    Set objForkTruckPart = objGeomFactory.Points3d.CreateByPoint _
              (m_outputColl.ResourceManager, _
               posForkTruckControlPoint.x, posForkTruckControlPoint.y, posForkTruckControlPoint.z)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    


'Insert your code for output 31(Default Surface)
'   --- create non-persistent circle to use for creating default surface ---
    Dim objTmpCircle As Object
    Set objTmpCircle = objGeomFactory.Circles3d.CreateByCenterNormalRadius _
         (Nothing, posForkTruckControlPoint.x, posForkTruckControlPoint.y, posForkTruckControlPoint.z, _
          posForkTruckControlPoint.x, posForkTruckControlPoint.y, posForkTruckControlPoint.z - 1, _
          0.05)
    
'   --- create persistent default surface plane - the plane can mate ---
    Set objForkTruckPart = objGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_outputColl.ResourceManager, objTmpCircle)
                                       
    Set objTmpCircle = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objForkTruckPart
    Set objForkTruckPart = Nothing
    
    

'   --- release outstanding objects ---
    Set objGeomFactory = Nothing
    Set vecProjDir = Nothing
    Set posCylinderCenter = Nothing
    Set posForkTruckControlPoint = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Set oriPartOrientation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

