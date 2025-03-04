Attribute VB_Name = "basVerPumpVS1Asm"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   basVerPumpVS1Asm.cls
'   Author: RUK
'   Creation Date:  Friday, August 31 2007
'
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
'
'******************************************************************************
Option Explicit

'Used to report truly unexpected errors - a last resort response
'As errors actually occur and are reported the calling code should then
'be modified to in anticipate and handle them and not call this sub

Public Sub ReportUnanticipatedError(InModule As String, InMethod As String, Optional errnumber As Long, Optional Context As String, Optional ErrDescription As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

Public Function CreateBoxEdges(Outputcoll As Object, oBoxTopCorner As IJDPosition, _
                                        oBoxBotCorner As IJDPosition) As Collection
    Const METHOD = "CreateBoxEdges:"
    On Error GoTo ErrorHandler
    
    Dim objEdgeColl As Collection
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim stPoint As AutoMath.DPosition
    
    Set objEdgeColl = New Collection
    Set oGeomFactory = New GeometryFactory
    Set stPoint = New DPosition
    
    Dim dBlockWidth As Double
    Dim dBlockHeight As Double
    Dim dBlockLength As Double
    Dim iCount As Integer
    Dim dPoints(0 To 23) As Double
    
    dBlockLength = oBoxTopCorner.x - oBoxBotCorner.x
    dBlockWidth = oBoxTopCorner.y - oBoxBotCorner.y
    dBlockHeight = oBoxTopCorner.z - oBoxBotCorner.z
    
    'Point on Left face
    stPoint.Set (oBoxTopCorner.x + oBoxBotCorner.x) / 2, _
                    (oBoxTopCorner.x + oBoxBotCorner.y) / 2, (oBoxTopCorner.z + oBoxBotCorner.z) / 2
    stPoint.x = stPoint.x - dBlockLength / 2
    'Points at corners of block
    
    '  0,1,2-----------------12,13,14
    '       '               '
    '       '  6,7,8-----------------18,19,20
    '       '       '        '      '
    '       '       '        '      '
    '       '       '        '      '
    '  3,4,5--------'---------15,16,17
    '               '               '
    '       9,10,11 -----------------21,22,23
    
    For iCount = 0 To 12 Step 12
        dPoints(0 + iCount) = stPoint.x
        dPoints(1 + iCount) = stPoint.y + dBlockWidth / 2
        dPoints(2 + iCount) = stPoint.z + dBlockHeight / 2

        dPoints(3 + iCount) = stPoint.x
        dPoints(4 + iCount) = stPoint.y + dBlockWidth / 2
        dPoints(5 + iCount) = stPoint.z - dBlockHeight / 2
        
        dPoints(6 + iCount) = stPoint.x
        dPoints(7 + iCount) = stPoint.y - dBlockWidth / 2
        dPoints(8 + iCount) = stPoint.z + dBlockHeight / 2
        
        dPoints(9 + iCount) = stPoint.x
        dPoints(10 + iCount) = stPoint.y - dBlockWidth / 2
        dPoints(11 + iCount) = stPoint.z - dBlockHeight / 2
        
        stPoint.x = stPoint.x + dBlockLength
    Next iCount
    
    'Edges
    'Edge1 Point1 to Point2
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(0), dPoints(1), dPoints(2), _
                                dPoints(3), dPoints(4), dPoints(5))
    'Edge2 Point1 to Point3
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(0), dPoints(1), dPoints(2), _
                                dPoints(6), dPoints(7), dPoints(8))
    'Edge3 Point1 to Point5
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(0), dPoints(1), dPoints(2), _
                                dPoints(12), dPoints(13), dPoints(14))
    'Edge4 Point4 to Point2
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(9), dPoints(10), dPoints(11), _
                                dPoints(3), dPoints(4), dPoints(5))
    'Edge5 Point4 to Point3
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(9), dPoints(10), dPoints(11), _
                                dPoints(6), dPoints(7), dPoints(8))
    'Edge6 Point4 to Point8
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(9), dPoints(10), dPoints(11), _
                                dPoints(21), dPoints(22), dPoints(23))
    'Edge7 Point6 to Point2
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(15), dPoints(16), dPoints(17), _
                                dPoints(3), dPoints(4), dPoints(5))
    'Edge8 Point6 to Point5
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(15), dPoints(16), dPoints(17), _
                                dPoints(12), dPoints(13), dPoints(14))
    'Edge9 Point6 to Point8
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(15), dPoints(16), dPoints(17), _
                                dPoints(21), dPoints(22), dPoints(23))
    'Edge10 Point7 to Point3
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(18), dPoints(19), dPoints(20), _
                                dPoints(6), dPoints(7), dPoints(8))
    'Edge11 Point7 to Point5
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(18), dPoints(19), dPoints(20), _
                                dPoints(12), dPoints(13), dPoints(14))
    'Edge12 Point7 to Point8
    objEdgeColl.Add oGeomFactory.Lines3d.CreateBy2Points(Outputcoll.ResourceManager, _
                                dPoints(18), dPoints(19), dPoints(20), _
                                dPoints(21), dPoints(22), dPoints(23))
    
    Set CreateBoxEdges = objEdgeColl
    Set objEdgeColl = Nothing
    Set stPoint = Nothing
    Set oGeomFactory = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
    Err.HelpFile, Err.HelpContext

End Function

Public Function CreatePointsOnBoxFaces(Outputcoll As Object, oBoxTopCorner As IJDPosition, _
                                oBoxBotCorner As IJDPosition) As Collection
    Const METHOD = "CreatePointsOnBoxFaces:"
    On Error GoTo ErrorHandler
    
    'Creating the Points on each surface of the Block
    Dim objPointColl As Collection
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim stPoint As AutoMath.DPosition
    
    Set oGeomFactory = New GeometryFactory
    Set stPoint = New DPosition
    Set objPointColl = New Collection
    
    Dim dBlockWidth As Double
    Dim dBlockHeight As Double
    Dim dBlockLength As Double
    
    dBlockLength = oBoxTopCorner.x - oBoxBotCorner.x
    dBlockWidth = oBoxTopCorner.y - oBoxBotCorner.y
    dBlockHeight = oBoxTopCorner.z - oBoxBotCorner.z
    
    'Point at Center of the Box
    stPoint.Set (oBoxTopCorner.x + oBoxBotCorner.x) / 2, _
                    (oBoxTopCorner.x + oBoxBotCorner.y) / 2, (oBoxTopCorner.z + oBoxBotCorner.z) / 2
    
    'Points on Right and Left surfaces
    objPointColl.Add oGeomFactory.Points3d.CreateByPoint(Outputcoll.ResourceManager, _
                            stPoint.x + dBlockLength / 2, stPoint.y, stPoint.z)
    objPointColl.Add oGeomFactory.Points3d.CreateByPoint(Outputcoll.ResourceManager, _
                            stPoint.x - dBlockLength / 2, stPoint.y, stPoint.z)
    'Points on Front and Back surfaces
    objPointColl.Add oGeomFactory.Points3d.CreateByPoint(Outputcoll.ResourceManager, _
                            stPoint.x, stPoint.y - dBlockWidth / 2, stPoint.z)
    objPointColl.Add oGeomFactory.Points3d.CreateByPoint(Outputcoll.ResourceManager, _
                            stPoint.x, stPoint.y + dBlockWidth / 2, stPoint.z)
    'Points on Top and Bottom surfaces
    objPointColl.Add oGeomFactory.Points3d.CreateByPoint(Outputcoll.ResourceManager, _
                            stPoint.x, stPoint.y, stPoint.z + dBlockHeight / 2)
    objPointColl.Add oGeomFactory.Points3d.CreateByPoint(Outputcoll.ResourceManager, _
                            stPoint.x, stPoint.y, stPoint.z - dBlockHeight / 2)
    Set CreatePointsOnBoxFaces = objPointColl
    
    Set objPointColl = Nothing
    Set stPoint = Nothing
    Set oGeomFactory = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
                            Err.HelpFile, Err.HelpContext
End Function

