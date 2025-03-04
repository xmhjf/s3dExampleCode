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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:16:39 AM"
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

'   --- some six foot worker constants used in multiple output sections ---
        
'   --- general purpose item declarations ---
    Dim iOutput              As Double
    Dim dblProjDistance      As Double
    Dim i                    As Integer
    Dim nVertices            As Integer
    Dim dblDiameter1         As Double
    Dim dblDiameter2         As Double
    Dim dblHeight            As Double
    
    Dim objGeomFactory       As IngrGeom3D.GeometryFactory
    Dim posStartPoint        As IJDPosition
    Dim posEndPoint          As IJDPosition
    Dim posCylinderCenter    As IJDPosition
    Dim vecProjDir           As IJDVector
    Dim polygonVertices()    As IJDPosition

'   --- six foot worker declarations ---
    Dim objWorkerPart        As Object
    
    Dim posWorkerControlPoint  As IJDPosition
    Dim posBottomCenter        As IJDPosition
    Dim posTopCenter           As IJDPosition
    Dim posDishPosition        As IJDPosition
    
    Dim vecDirOutTop       As IJDVector
    Dim vecDirOutBottom    As IJDVector
    
    Dim oriPartOrientation As Orientation
    
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

    Set objGeomFactory = New IngrGeom3D.GeometryFactory
    Set vecProjDir = New DVector
    
    Set posStartPoint = New DPosition
    Set posEndPoint = New DPosition
    
    Set posWorkerControlPoint = New DPosition
    Set posCylinderCenter = New DPosition
    Set posBottomCenter = New DPosition
    Set posTopCenter = New DPosition
    Set posDishPosition = New DPosition
    
    Set vecDirOutTop = New DVector
    Set vecDirOutBottom = New DVector

    Set oriPartOrientation = New Orientation
    
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'   --- set worker control point origin ---
    posWorkerControlPoint.Set 0, 0, 0

'   --- do the outputs ---
    iOutput = 0



'   ------
'   HEAD
'   ------
' Insert your code for output 1(Head)
'   --- create a snout to model head ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutX = -90
    oriPartOrientation.RotationAboutY = 90
    oriPartOrientation.ApplyRotations
    Set objWorkerPart = CreateSnout(m_outputColl, "n 0.20003 e 0.12382 u 1.6383", oriPartOrientation, 0.041275, 0.0889, 0.225425, 0.123825, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing
    
    
' Insert your code for output 2(Left Eye)
'   --- create a sphere to model left eye ---
    Set objWorkerPart = CreateSphere(m_outputColl, "e 0.06985 n 0.26671 u 1.6637", 0.034925)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 3(Right Eye)
'   --- create a sphere to model right eye ---
    Set objWorkerPart = CreateSphere(m_outputColl, "e 0.1778 n 0.26671 u 1.6637", 0.034925)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 4(Hat)
'   --- create a dish to model hat ---
    Set objWorkerPart = CreateDish(m_outputColl, "n 0.18416 e 0.12382 u 1.68592", "N 0 U 90", 0.2286, 0.130175, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 5(Hat Brim)
'   --- create a cylinder to model hat brim ---
    Set objWorkerPart = CreateCylinder(m_outputColl, "N 0.18416 E 0.12382 U 1.69227", 0.250825, 0.0127, "N 0 U 90")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 6(Neck)
'   --- create a cylinder to model neck ---
    Set objWorkerPart = CreateCylinder(m_outputColl, "N 0.18416 E 0.12382 U 1.57638", 0.111125, 0.200025, "N 0 U 90")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



'   ----------------
'   SHOULDERS / ARMS
'   ----------------
' Insert your code for output 7(shoulders)
'   --- create a cylinder to model shoulders ---
    Set objWorkerPart = CreateCylinder(m_outputColl, "N 0.16828 E 0.12382 U 1.45732", 0.098425, 0.4064, "E 90 U 0")
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 8(Left Shoulder Joint)
'   --- create a sphere to model left shoulder joint ---
    Set objWorkerPart = CreateSphere(m_outputColl, "w 0.07938 n 0.16828 u 1.45732", 0.1143)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 9(Right Shoulder Joint)
'   --- create a sphere to model right shoulder joint ---
    Set objWorkerPart = CreateSphere(m_outputColl, "e 0.32702 n 0.16828 u 1.45732", 0.1143)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 10(Left Upper Arm)
'   --- create a cylinder to model left upper arm ---
    Set objWorkerPart = CreateCylinder(m_outputColl, "N 0.21908 W 0.09208 U 1.3208", 0.0889, 0.269875, "S 180 U 70")

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 11(Right Upper Arm)
'   --- create a cylinder to model right upper arm ---
    Set objWorkerPart = CreateCylinder(m_outputColl, "N 0.31751 E 0.33972 U 1.48907", 0.0889, 0.269875, "N 0 U 15")
   
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 12(Left Elbow Joint)
'   --- create a sphere to model left elbow joint ---
    Set objWorkerPart = CreateSphere(m_outputColl, "w 0.09208 n 0.27941 u 1.1684", 0.079375)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 13(Right Elbow Joint)
'   --- create a sphere to model right elbow joint ---
    Set objWorkerPart = CreateSphere(m_outputColl, "e 0.33972 n 0.46038 u 1.524", 0.079375)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 14(Left Forearm)
'   --- create a snout to model left forearm ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutY = 15
    oriPartOrientation.RotationAboutZ = -90


    oriPartOrientation.ApplyRotations
    Set objWorkerPart = CreateSnout(m_outputColl, "n 0.43816 w 0.09208 u 1.20332", oriPartOrientation, 0, 0.0889, 0.0762, 0.263525, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 15(Right Forearm)
'   --- create a snout to model right forearm ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutY = 90
    oriPartOrientation.RotationAboutX = -33.50001109
    oriPartOrientation.RotationAboutZ = 0.05254222
    oriPartOrientation.ApplyRotations
    Set objWorkerPart = CreateSnout(m_outputColl, "n 0.54769 e 0.34131 u 1.651", oriPartOrientation, 0, 0.0889, 0.0762, 0.263525, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 16(Left Hand)
'   --- create a sphere to model left hand ---
    Set objWorkerPart = CreateSphere(m_outputColl, "w 0.09208 n 0.59056 u 1.24459", 0.0762)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 17(Right Hand)
'   --- create a sphere to model right hand ---
    Set objWorkerPart = CreateSphere(m_outputColl, "e 0.34131 n 0.63183 u 1.78117", 0.0762)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



'   ---------
'   TORSO
'   ---------
' Insert your code for output 18(Torso Bottom)
'   --- create a dish to model torso bottom ---
    Set objWorkerPart = CreateDish(m_outputColl, "n 0.18416 e 0.12065 u 0.96043", "N 0 D -90", 0.339725, 0.149225, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 19(Lower Middle)
' --- create a snout to model lower middle of torso ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutY = -90
    oriPartOrientation.ApplyRotations
    Set objWorkerPart = CreateSnout(m_outputColl, "n 0.18416 e 0.12382 u 1.03663", oriPartOrientation, 0, 0.31115, 0.339725, 0.149225, True)
'
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 20(Upper Middle)
'   --- create a snout to model upper middle of torso ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutY = -90
    oriPartOrientation.ApplyRotations
    Set objWorkerPart = CreateSnout(m_outputColl, "n 0.18416 e 0.1238200 u 1.26365", oriPartOrientation, 0, 0.339725, 0.31115, 0.3048, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing


' Insert your code for output 21(Torso Top)
'   --- create a dish to model torso top ---
    Set objWorkerPart = CreateDish(m_outputColl, "n 0.18416 e 0.1285800 u 1.41446", "N 0 U 90", 0.339725, 0.098425, 0, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



'   -------------
'   LEGS AND FEET
'   -------------
' Insert your code for output 22(Left Thigh)
'   --- create a snout to model left thigh ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutY = -90
    oriPartOrientation.RotationAboutX = -10
    oriPartOrientation.ApplyRotations
    Set objWorkerPart = CreateSnout(m_outputColl, "n 0.18733 e 0.04127 u 0.72707", oriPartOrientation, 0, 0.17145, 0.11125, 0.428625, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



' Insert your code for output 23(Right Thigh)
'   --- create a snout to model right thigh ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutY = -90
    oriPartOrientation.RotationAboutX = 10
    oriPartOrientation.ApplyRotations
    Set objWorkerPart = CreateSnout(m_outputColl, "n 0.18733 e 0.21272 u 0.72866", oriPartOrientation, 0, 0.17145, 0.11125, 0.428625, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



' Insert your code for output 24(Left Knee)
'   --- create a sphere to model left knee ---
    Set objWorkerPart = CreateSphere(m_outputColl, "e 0.03969 n 0.13971 u 0.4699", 0.12065)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



' Insert your code for output 25(Right Knee)
'   --- create a sphere to model right knee ---
    Set objWorkerPart = CreateSphere(m_outputColl, "e 0.21272 n 0.23813 u 0.4699", 0.12065)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



' Insert your code for output 26(Left Lower Leg)
'   --- create a snout to model left lower leg ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutY = -90
    oriPartOrientation.RotationAboutX = -10
    oriPartOrientation.ApplyRotations
    Set objWorkerPart = CreateSnout(m_outputColl, "n 0.08891 e 0.0381 u 0.24923", oriPartOrientation, 0, 0.111125, 0.0635, 0.333375, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



' Insert your code for output 27(Right Lower Leg)
'   --- create a snout to model right lower leg ---
    oriPartOrientation.ResetDefaultAxis
    oriPartOrientation.RotationAboutY = -90
    oriPartOrientation.RotationAboutX = -10
    oriPartOrientation.ApplyRotations
    Set objWorkerPart = CreateSnout(m_outputColl, "n 0.18733 e 0.2159 u 0.25241", oriPartOrientation, 0, 0.111125, 0.0635, 0.333375, True)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



' Insert your code for output 28(Left Ankle)
'   --- create a sphere to model left ankle ---
    Set objWorkerPart = CreateSphere(m_outputColl, "e 0.03651 n 0.0381 u 0.05873", 0.0635)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



' Insert your code for output 29(Right Ankle)
'   --- create a sphere to model right ankle ---
    Set objWorkerPart = CreateSphere(m_outputColl, "e 0.2159 n 0.13971 u 0.06826", 0.0635)

' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing



' Insert your code for output 30(Left Foot)
'   --- create a polygon projection to model left foot ---
    nVertices = 4
    ReDim polygonVertices(0 To (nVertices - 1)) As IJDPosition
    
'   --- instantiate each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = New DPosition
    Next i
    
'   --- load the polygon vertices array with x,y,z data ---
    polygonVertices(0).Set 0#, 0#, 0.02857
    polygonVertices(1).Set 0.00953, 0.24131, 0.01428
    polygonVertices(2).Set 0.00953, 0.24131, 0.00159
    polygonVertices(3).Set 0#, 0#, 0#

'   --- establish the direction of projection ---
    vecProjDir.Set 1, 0, 0
    
'   --- establish the projection distance ---
    dblProjDistance = 0.0762

'   --- project the polygon ---
    Set objWorkerPart = placeProjectedPolygonFromVertices(m_outputColl, polygonVertices(), vecProjDir, dblProjDistance, True)

'   --- release each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = Nothing
    Next i
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing
    


' Insert your code for output 31(Right Foot)
'   --- create a triangular projection to model right foot ---
    nVertices = 4
    ReDim polygonVertices(0 To (nVertices - 1)) As IJDPosition
    
'   --- instantiate each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = New DPosition
    Next i
    
'   --- load the polygon vertices array with x,y,z data ---
    polygonVertices(0).Set 0.17621, 0.09843, 0.03809
    polygonVertices(1).Set 0.17938, 0.33973, 0.0238
    polygonVertices(2).Set 0.17938, 0.33973, 0.01111
    polygonVertices(3).Set 0.17621, 0.09843, 0.00952

'   --- establish the direction of projection ---
    vecProjDir.Set 1, 0, 0
    
'   --- establish the projection distance ---
    dblProjDistance = 0.0762

'   --- project the polygon ---
    Set objWorkerPart = placeProjectedPolygonFromVertices(m_outputColl, polygonVertices(), vecProjDir, dblProjDistance, True)

'   --- release each array item ---
    For i = 0 To (nVertices - 1)
      Set polygonVertices(i) = Nothing
    Next i
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing
    


'Insert your code for output 32(Control Point)
'   --- create a persistent point at the control point position ---
    Set objWorkerPart = objGeomFactory.Points3d.CreateByPoint _
              (m_outputColl.ResourceManager, _
               posWorkerControlPoint.x, posWorkerControlPoint.y, posWorkerControlPoint.z)
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing
    


'Insert your code for output 33(Default Surface)
'   --- create non-persistent circle to use for creating default surface ---
    Dim objTmpCircle As Object
    Set objTmpCircle = objGeomFactory.Circles3d.CreateByCenterNormalRadius _
         (Nothing, posWorkerControlPoint.x, posWorkerControlPoint.y, posWorkerControlPoint.z, _
          posWorkerControlPoint.x, posWorkerControlPoint.y, posWorkerControlPoint.z - 1, _
          0.05)
    
'   --- create persistent default surface plane - the plane can mate ---
    Set objWorkerPart = objGeomFactory.Planes3d.CreateByOuterBdry _
                                       (m_outputColl.ResourceManager, objTmpCircle)
                                       
    Set objTmpCircle = Nothing
    
' Set the output
    iOutput = iOutput + 1
    m_outputColl.AddOutput arrayOfOutputs(iOutput), objWorkerPart
    Set objWorkerPart = Nothing
    
    

'   --- release outstanding objects ---
    Set objGeomFactory = Nothing
    Set vecProjDir = Nothing
    Set posCylinderCenter = Nothing
    Set posWorkerControlPoint = Nothing
    Set posStartPoint = Nothing
    Set posEndPoint = Nothing
    Set posBottomCenter = Nothing
    Set posTopCenter = Nothing
    Set vecDirOutTop = Nothing
    Set vecDirOutBottom = Nothing
    Set oriPartOrientation = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
