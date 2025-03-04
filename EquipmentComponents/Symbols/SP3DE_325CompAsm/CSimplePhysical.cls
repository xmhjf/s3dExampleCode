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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:49 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:        MS
'   Creation Date: Tuesday, Feb 17 2004
'   Description:
'    The symbol is prepared from PDS Equipment guide E325. The symbol consists of 11 physical outputs.
'
'   Change History:
'   dd.mmm.yyyy     who             change description
'   -----------     ---             ------------------
'
'   29.Nov.2004     V6UpgradeSO     Made compatible with Smart Occurrence based Equipments
'   01.Jun.2006     svsmylav        CR- 90896: Center of Plate 2 is modified to be same Z-coordinate as DP1 (origin)
'                                   and side plate top is aligned to the top of Plate 2.
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

    Private m_oSymGeomHelper As IJSymbolGeometryHelper
    Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
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
    
    Dim iOutput     As Double
    Dim ObjPlate1 As Object
    Dim ObjExchangerBody As Object
    Dim ObjPlate2 As Object
    Dim ObjSidePlate As Object
    
    Dim parBotPlate1toPP As Double 'P1
    Dim parPlate1Height As Double 'P2
    Dim parPlate1Thickness As Double 'P3
    Dim parPlatesWidth As Double 'P4
    Dim parExchangerLength As Double 'P5
    Dim parExchangerHeight As Double 'P6
    Dim parExchangerWidth As Double 'P7
    Dim parPlate2Height As Double 'P8
    Dim parPlate2Thickness As Double 'P9
    Dim parRodsLength As Double 'P10
    Dim parRodsDiameter As Double 'P11
    Dim parBotSidePlatetoRod2Cen As Double 'P12
    Dim parRodsCentoCen As Double 'P13
    Dim parSidePlateThickness As Double 'P14
    Dim parSidePlateWidth As Double 'P15

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBotPlate1toPP = arrayOfInputs(2)
    parPlate1Height = arrayOfInputs(3)
    parPlate1Thickness = arrayOfInputs(4)
    parPlatesWidth = arrayOfInputs(5)
    parExchangerLength = arrayOfInputs(6)
    parExchangerHeight = arrayOfInputs(7)
    parExchangerWidth = arrayOfInputs(8)
    parPlate2Height = arrayOfInputs(9)
    parPlate2Thickness = arrayOfInputs(10)
    parRodsLength = arrayOfInputs(11)
    parRodsDiameter = arrayOfInputs(12)
    parBotSidePlatetoRod2Cen = arrayOfInputs(13)
    parRodsCentoCen = arrayOfInputs(14)
    parSidePlateThickness = arrayOfInputs(15)
    parSidePlateWidth = arrayOfInputs(16)
    
    m_oSymGeomHelper.OutputCollection = m_OutputColl
     
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
    iOutput = 0
    
    ''Origin is at DP/PP1
    
 ' Insert your code for output 1(Plate1)
 
 Dim topSurPts(3) As IJDPosition
    Dim botSurPts(3) As IJDPosition
    Dim iCount As Integer
    
    For iCount = 0 To 3
        Set topSurPts(iCount) = New DPosition
        Set botSurPts(iCount) = New DPosition
    Next iCount
  
    'Point 1
    botSurPts(0).x = 0
    botSurPts(0).y = parPlatesWidth / 2
    botSurPts(0).z = -parBotPlate1toPP
    'Point 2
    botSurPts(1).x = 0
    botSurPts(1).y = -parPlatesWidth / 2
    botSurPts(1).z = -parBotPlate1toPP
    
    'Point3
    botSurPts(2).x = parPlate1Thickness
    botSurPts(2).y = -parPlatesWidth / 2
    botSurPts(2).z = -parBotPlate1toPP
    
    'Point 4
    botSurPts(3).x = parPlate1Thickness
    botSurPts(3).y = parPlatesWidth / 2
    botSurPts(3).z = -parBotPlate1toPP
    
    topSurPts(0).Set botSurPts(0).x, botSurPts(0).y, botSurPts(0).z + parPlate1Height
    topSurPts(1).Set botSurPts(1).x, botSurPts(1).y, botSurPts(1).z + parPlate1Height
    topSurPts(2).Set botSurPts(2).x, botSurPts(2).y, botSurPts(2).z + parPlate1Height
    topSurPts(3).Set botSurPts(3).x, botSurPts(3).y, botSurPts(3).z + parPlate1Height
    
    Dim ObjCollection As Collection
    
    Set ObjCollection = PlaceTrapezoidWithPlanes(m_OutputColl, topSurPts, botSurPts)
    For iCount = 1 To 6
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCollection(iCount)
    Next iCount

    Set ObjCollection = Nothing
     For iCount = 0 To 3
        Set topSurPts(iCount) = Nothing
        Set botSurPts(iCount) = Nothing
    Next iCount
   
 ' Insert your code for output 2(ExchangerBody)
    stPoint.Set parPlate1Thickness, parExchangerWidth / 2, parExchangerHeight / 2
    enPoint.Set parPlate1Thickness + parExchangerLength, -parExchangerWidth / 2, -parExchangerHeight / 2
    
    Set ObjExchangerBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
    
 ' Insert your code for output 3(Plate2)
    Dim ActualBody As Double 'Sum of ExchangerBody, Plate1 and Plate2 lengths
    ActualBody = parPlate1Thickness + parExchangerLength + parPlate2Thickness 'P3+P5+P9
     stPoint.Set parPlate1Thickness + parExchangerLength, _
                    parPlatesWidth / 2, _
                    parPlate2Height / 2
    enPoint.Set ActualBody, -parPlatesWidth / 2, stPoint.z - parPlate2Height
    
    Set ObjPlate2 = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlate2
    Set ObjPlate2 = Nothing
    
 ' Insert your code for output 4(Rod1)
    stPoint.Set ActualBody, 0, parRodsCentoCen / 2
    enPoint.Set ActualBody + parRodsLength, 0, parRodsCentoCen / 2
    
    ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parRodsDiameter
    
 ' Insert your code for output 5(Rod2)
    stPoint.Set ActualBody, 0, -parRodsCentoCen / 2
    enPoint.Set ActualBody + parRodsLength, 0, -parRodsCentoCen / 2
    
    ' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, parRodsDiameter
    
 ' Insert your code for output 6(SidePlate)
     stPoint.Set ActualBody + parRodsLength, _
                    parSidePlateWidth / 2, _
                    parPlate2Height / 2
    enPoint.Set ActualBody + parRodsLength + parSidePlateThickness, _
                -parSidePlateWidth / 2, -parRodsCentoCen / 2 - parBotSidePlatetoRod2Cen
    
    Set ObjSidePlate = PlaceBox(m_OutputColl, stPoint, enPoint)

    Set stPoint = Nothing
    Set enPoint = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSidePlate
    Set ObjSidePlate = Nothing
      
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub


Private Function PlaceTrapezoidWithPlanes(ByVal objOutputColl As Object, _
                        ByRef topSurfacePoints() As IJDPosition, _
                        ByRef bottomSurfacePoints() As IJDPosition _
                        ) As Collection
''This function returns six Planes of Trapezoid as "Collection". The order of planes in the
''collection is Top,Bottom and four sides. The top and bottom surface points are taken in
''anti-clockwise direction when viewing from top.
    Const METHOD = "PlaceTrapezoidWithPlanes:"
    On Error GoTo ErrorHandler
    Dim iCount As Integer
    Dim arrayTopPoints(0 To 11) As Double
    Dim arrayBottomPoints(0 To 11) As Double
    Dim objGeomFactory As IngrGeom3D.GeometryFactory
    Dim oTmpCollection As New Collection
    
'   These are to obtain the coordinates of various vertices.
    For iCount = 0 To 3
        arrayTopPoints(3 * iCount) = topSurfacePoints(iCount).x
        arrayTopPoints(3 * iCount + 1) = topSurfacePoints(iCount).y
        arrayTopPoints(3 * iCount + 2) = topSurfacePoints(iCount).z
        
        arrayBottomPoints(3 * iCount) = bottomSurfacePoints(iCount).x
        arrayBottomPoints(3 * iCount + 1) = bottomSurfacePoints(iCount).y
        arrayBottomPoints(3 * iCount + 2) = bottomSurfacePoints(iCount).z
    Next iCount
    Set objGeomFactory = New IngrGeom3D.GeometryFactory

    Dim arrayPlanePoints(0 To 11) As Double
    Dim objPlane As IngrGeom3D.Plane3d
    
'   Top plane
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                        4, arrayTopPoints)
    oTmpCollection.Add objPlane
    
'   Bottom plane
'To orient normal of the bottom plane outside the body, the plane points are taken in reverse order.
    Dim botPtCount As Integer
    For iCount = 1 To 4
        botPtCount = 5 - iCount
        arrayPlanePoints(3 * iCount - 3) = arrayBottomPoints(3 * botPtCount - 3)
        arrayPlanePoints(3 * iCount - 2) = arrayBottomPoints(3 * botPtCount - 2)
        arrayPlanePoints(3 * iCount - 1) = arrayBottomPoints(3 * botPtCount - 1)
    Next iCount
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    oTmpCollection.Add objPlane
    
'   Front plane
    arrayPlanePoints(0) = arrayBottomPoints(0)
    arrayPlanePoints(1) = arrayBottomPoints(1)
    arrayPlanePoints(2) = arrayBottomPoints(2)
    
    arrayPlanePoints(3) = arrayBottomPoints(3)
    arrayPlanePoints(4) = arrayBottomPoints(4)
    arrayPlanePoints(5) = arrayBottomPoints(5)
    
    arrayPlanePoints(6) = arrayTopPoints(3)
    arrayPlanePoints(7) = arrayTopPoints(4)
    arrayPlanePoints(8) = arrayTopPoints(5)
    
    arrayPlanePoints(9) = arrayTopPoints(0)
    arrayPlanePoints(10) = arrayTopPoints(1)
    arrayPlanePoints(11) = arrayTopPoints(2)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    
    oTmpCollection.Add objPlane
    
'   Right hand side plane
    arrayPlanePoints(0) = arrayBottomPoints(3)
    arrayPlanePoints(1) = arrayBottomPoints(4)
    arrayPlanePoints(2) = arrayBottomPoints(5)
    
    arrayPlanePoints(3) = arrayBottomPoints(6)
    arrayPlanePoints(4) = arrayBottomPoints(7)
    arrayPlanePoints(5) = arrayBottomPoints(8)
    
    arrayPlanePoints(6) = arrayTopPoints(6)
    arrayPlanePoints(7) = arrayTopPoints(7)
    arrayPlanePoints(8) = arrayTopPoints(8)
    
    arrayPlanePoints(9) = arrayTopPoints(3)
    arrayPlanePoints(10) = arrayTopPoints(4)
    arrayPlanePoints(11) = arrayTopPoints(5)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                4, arrayPlanePoints)
    
    oTmpCollection.Add objPlane
    
'   Rear plane
    arrayPlanePoints(0) = arrayBottomPoints(6)
    arrayPlanePoints(1) = arrayBottomPoints(7)
    arrayPlanePoints(2) = arrayBottomPoints(8)
    
    arrayPlanePoints(3) = arrayBottomPoints(9)
    arrayPlanePoints(4) = arrayBottomPoints(10)
    arrayPlanePoints(5) = arrayBottomPoints(11)
    
    arrayPlanePoints(6) = arrayTopPoints(9)
    arrayPlanePoints(7) = arrayTopPoints(10)
    arrayPlanePoints(8) = arrayTopPoints(11)
    
    arrayPlanePoints(9) = arrayTopPoints(6)
    arrayPlanePoints(10) = arrayTopPoints(7)
    arrayPlanePoints(11) = arrayTopPoints(8)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    
    oTmpCollection.Add objPlane
    
'   Left hand side plane
    arrayPlanePoints(0) = arrayBottomPoints(9)
    arrayPlanePoints(1) = arrayBottomPoints(10)
    arrayPlanePoints(2) = arrayBottomPoints(11)
    
    arrayPlanePoints(3) = arrayBottomPoints(0)
    arrayPlanePoints(4) = arrayBottomPoints(1)
    arrayPlanePoints(5) = arrayBottomPoints(2)
    
    arrayPlanePoints(6) = arrayTopPoints(0)
    arrayPlanePoints(7) = arrayTopPoints(1)
    arrayPlanePoints(8) = arrayTopPoints(2)
    
    arrayPlanePoints(9) = arrayTopPoints(9)
    arrayPlanePoints(10) = arrayTopPoints(10)
    arrayPlanePoints(11) = arrayTopPoints(11)
    
    Set objPlane = objGeomFactory.Planes3d.CreateByPoints(objOutputColl.ResourceManager, _
                                                                    4, arrayPlanePoints)
    oTmpCollection.Add objPlane
    
    Set PlaceTrapezoidWithPlanes = oTmpCollection

    Set oTmpCollection = Nothing
    Set objGeomFactory = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Function
