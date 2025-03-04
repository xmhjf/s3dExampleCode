Attribute VB_Name = "MbrCustomRules"
'*******************************************************************
'
'Copyright (C) 2013-2014 Intergraph Corporation. All rights reserved.
'
'File : MbrCustomRules.bas
'
'Author : Alligators
'
'Description :
'   MbrMeasurementUtilities included all the routines necessary to determine the
'   relative position of the bounded member with respect to the bounding member.
'
'History :
'   17/Apr/2013   - Addedd file header.
'   17/Apr/2013   - Alligators
'           TR-230220 'GetSketchPlaneForTube' is modified to handle tube member
'                     bounded to Stiffener cases.
'    03/Nov/2014 - MDT/GH
'         CR-CP-250198  Lapped AC for traffic items
'*****************************************************************************
Option Explicit
Private Const MODULE = "StructDetail\Data\Include\MbrCustomRules"



'**********************************************************************************************
' Method      : GetSketchPlaneForTube
' Description : Sets the Sketching plane
'               This method should be used only when the Tube is Bounded and Non-Tubular
'               members(Standard) is Bounding and when Bounding Port is Lateral Port
'**********************************************************************************************
Public Function GetSketchPlaneForTube(oBoundingPort As IJPort, oBoundedPort As IJPort) As IJPlane
    
    Const MT = "GetSketchPlaneForTube"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    ' -------------------------------------------
    ' Verify bounded object is a member part port
    ' -------------------------------------------
    ' Note: prior comments said profile or member part, but the code will fail below if the input port is not a
    ' split axis port, which only applies to members.

    Dim oBoundedPartObject As Object
    Set oBoundedPartObject = oBoundedPort.Connectable
        
    If (Not TypeOf oBoundedPartObject Is IJStructProfilePart) Or (Not TypeOf oBoundedPort Is ISPSSplitAxisPort) Then
        Exit Function
    End If
    
    ' ---------------------------------
    ' Verify the bounding port is valid
    ' ---------------------------------
    If oBoundingPort Is Nothing Then
        Exit Function
    End If
    
    ' -------------------------------------------
    ' Get the start and end of the bounded member
    ' -------------------------------------------
    Dim oBoundedStart As IJPoint
    Dim oBoundedEnd As IJPoint
    Dim oBoundedPos As New DPosition
    Dim x As Double
    Dim y As Double
    Dim z As Double

    Dim oBoundedMemberPort As ISPSSplitAxisPort
    Dim oBoundedMemberPart As ISPSMemberPartCommon
    Dim oBoundingMemberPart As ISPSMemberPartCommon
    
    sMsg = "Checking Whether the member is bounded at Start and End"
        
    If TypeOf oBoundedPartObject Is ISPSMemberPartCommon Then
        Set oBoundedMemberPart = oBoundedPartObject
        Set oBoundedStart = oBoundedMemberPart.PointAtEnd(SPSMemberAxisStart)
        Set oBoundedEnd = oBoundedMemberPart.PointAtEnd(SPSMemberAxisEnd)
    End If
    
    Set oBoundedMemberPort = oBoundedPort
    If oBoundedMemberPort.PortIndex = SPSMemberAxisStart Then
        oBoundedStart.GetPoint x, y, z
        oBoundedPos.Set x, y, z
    ElseIf oBoundedMemberPort.PortIndex = SPSMemberAxisEnd Then
        oBoundedEnd.GetPoint x, y, z
        oBoundedPos.Set x, y, z
        ElseIf oBoundedMemberPort.PortIndex = SPSMemberAxisAlong Then
            'For Both Axis Along Case - For Split None case
            Dim oCurve1 As IJCurve
            Dim oCurve2 As IJCurve
            Dim dMinDist As Double

            Dim dx1 As Double, dy1 As Double, dz1 As Double

            'Get Bounded Location
            If TypeOf oBoundingPort.Connectable Is ISPSMemberPartCommon Then
                oBoundingMemberPart = oBoundingPort.Connectable
                oCurve1 = oBoundedMemberPart.Axis
                oCurve2 = oBoundingMemberPart.Axis
                oCurve1.DistanceBetween oCurve2, dMinDist, x, y, z, dx1, dy1, dz1
                oBoundedPos.Set x, y, z
            End If
    End If
    
    ' ---------------------------------------------------------
    ' The sketching plane U-direction is along the bounded axis
    ' ---------------------------------------------------------
    sMsg = "Getting the Transformation Matrix at Bounded Member"
    
    Dim oBoundedXSecMatrix As IJDT4x4
    Dim oBoundingXSecMatrix As IJDT4x4
    Dim oDummyMatrix As IJDT4x4
        
    oBoundedMemberPart.Rotation.GetTransformAtPosition oBoundedPos.x, oBoundedPos.y, oBoundedPos.z, oBoundedXSecMatrix, oDummyMatrix
    
    Dim oSketchPlane_U As New AutoMath.dVector
    
    oSketchPlane_U.Set oBoundedXSecMatrix.IndexValue(0), oBoundedXSecMatrix.IndexValue(1), oBoundedXSecMatrix.IndexValue(2)

    ' -----------------------------------------------------------------------------------------------------
    ' The U-direction should point away from the bounded member face, so reverse the vector if at the start
    ' -----------------------------------------------------------------------------------------------------
    ' Why is it the end instead?????
    If oBoundedMemberPort.PortIndex = SPSMemberAxisEnd Then
        oSketchPlane_U.Length = -1#
    End If

    ' ----------------------
    ' If bounded by a member
    ' ----------------------
    Dim oBoundingObject As IJConnectable
    Set oBoundingObject = oBoundingPort.Connectable
    
    Dim oBoundingMemberPort As ISPSSplitAxisPort
        
    Dim oSketchPlaneRootPt As New DPosition
    
    Dim oBounded_UDir As New AutoMath.dVector
    Dim oBounded_VDir As New AutoMath.dVector
        
    If TypeOf oBoundingObject Is ISPSMemberPartCommon Or TypeOf oBoundingObject Is IJStiffener Then
    
        If TypeOf oBoundingObject Is ISPSMemberPartCommon Then
            'Bounding is a Member Part
            ' The connection must be along the length of the bounding member
            If TypeOf oBoundingPort Is ISPSSplitAxisPort Then
                Set oBoundingMemberPort = oBoundingPort
                If Not oBoundingMemberPort.PortIndex = SPSMemberAxisAlong Then
                    Exit Function
                End If
            Else
                If Not oBoundingPort.Type = PortFace Then
                    Exit Function
                End If
            End If
            
            '----------------------------------------------
            ' Get the bounding axis at the bounded location
            '----------------------------------------------
            sMsg = "Getting the Transformation Matrix at Bounding member"
            
            Set oBoundingMemberPart = oBoundingObject
            oBoundingMemberPart.Rotation.GetTransformAtPosition oBoundedPos.x, oBoundedPos.y, oBoundedPos.z, oBoundingXSecMatrix, oDummyMatrix
            
            Dim oBounding_Axis As New AutoMath.dVector
            oBounding_Axis.Set oBoundingXSecMatrix.IndexValue(0), oBoundingXSecMatrix.IndexValue(1), oBoundingXSecMatrix.IndexValue(2)
        Else
            'Bounding is a Stiffener Part
            ' The connection must be along the length of the bounding member
            If Not oBoundingPort.Type = PortFace Then
                Exit Function
            End If
            
            '----------------------------------------------
            ' Get the bounding axis at the bounded location
            '----------------------------------------------
            sMsg = "Getting the Transformation Matrix at Bounding member"
            
            Dim oProfileUtils As IJProfileAttributes
            Dim oProfUVec As IJDVector
            Dim oProfVVec As IJDVector
            Dim oOriginPos As IJDPosition
            Set oProfileUtils = New ProfileUtils
            oProfileUtils.GetProfileOrientationAndLocation oBoundingObject, oBoundedPos, oProfUVec, oProfVVec, oOriginPos

            Set oBounding_Axis = oProfUVec.Cross(oProfVVec)
        End If
        
        ' -----------------------------------------------------------
        ' The sketching plane root is at cardinal point 15 of bounded
        ' ----------------------------------------------------------
        ' Get the catalog position of the points and transform to 3D

        oBounded_UDir.Set oBoundedXSecMatrix.IndexValue(4), oBoundedXSecMatrix.IndexValue(5), oBoundedXSecMatrix.IndexValue(6)


        oBounded_VDir.Set oBoundedXSecMatrix.IndexValue(8), oBoundedXSecMatrix.IndexValue(9), oBoundedXSecMatrix.IndexValue(10)
        
        Dim u15 As Double
        Dim v15 As Double
        oBoundedMemberPart.crossSection.GetCardinalPointDelta Nothing, oBoundedMemberPart.crossSection.CardinalPoint, 15, u15, v15
        
        oBounded_UDir.Length = -u15
        oBounded_VDir.Length = v15
        

        oSketchPlaneRootPt.Set oBoundedXSecMatrix.IndexValue(12) + oBounded_UDir.x + oBounded_VDir.x, _
                               oBoundedXSecMatrix.IndexValue(13) + oBounded_UDir.y + oBounded_VDir.y, _
                               oBoundedXSecMatrix.IndexValue(14) + oBounded_UDir.z + oBounded_VDir.z
     ElseIf TypeOf oBoundingObject Is IJPlate Then
     
        ' The connection must be to a lateral port



        Dim oStructPort As IJStructPort
        Set oStructPort = oBoundingPort
        
        If Not oStructPort.ContextID = CTX_LATERAL Then
            Exit Function
        End If
     
        ' ------------------------------------------------------------------------------------------------
        ' The bounding axis (as if bounding was a profile) is the base normal nearest the bounded location
        ' ------------------------------------------------------------------------------------------------
        ' Get the point nearest the bounded location
        Dim oPlate As New StructDetailObjects.PlatePart
        Set oPlate.object = oBoundingObject
        

        Dim oBasePort As IJPort




        Set oBasePort = oPlate.baseport(BPT_Base)
        If oBasePort Is Nothing Then
            GoTo ErrorHandler
        End If
        
        Dim oTopo As New TopologyLocate
        Dim oPointOnPort As IJDPosition
        Dim oDummyVector As IJDVector
        Dim oPortSurface As IJSurfaceBody
        Dim oOffsetVector As IJDVector
                    
        oTopo.GetProjectedPointOnModelBody oBasePort.Geometry, oBoundedPos, oPointOnPort, oDummyVector
        
        ' Get the surface normal
        Set oPortSurface = oBasePort.Geometry
        oPortSurface.GetNormalFromPosition oBoundedPos, oBounding_Axis
        
        ' ------------------------------------------------------------------------------------------
        ' The sketch plane origin is at the mid-thickness of the plate, nearest the bounded location
        ' ------------------------------------------------------------------------------------------
        Set oOffsetVector = oBounding_Axis
        oOffsetVector.Length = -oPlate.PlateThickness
        
        Set oSketchPlaneRootPt = oPointOnPort.Offset(oOffsetVector)
        
     Else
        Exit Function

     End If
     
    ' --------------------------------------------------------------
    ' Sketch V is the cross product of the bounding axis and sketchU
    ' --------------------------------------------------------------
    Dim oSketchPlane_V As New AutoMath.dVector

    Set oSketchPlane_V = oBounding_Axis.Cross(oSketchPlane_U)

    ' ----------------------------------------------------
    ' Sketch N is the cross product of sketchU and sketchV
    ' ----------------------------------------------------
    Dim oSketchPlane_N As New AutoMath.dVector

    Set oSketchPlane_N = oSketchPlane_U.Cross(oSketchPlane_V)
    
    oSketchPlane_U.Length = 1#
    oSketchPlane_N.Length = 1#
    
    ' -----------------------------------------------------------------------------
    ' Create a transient plane based on the root point and vectors calculated above
    ' -----------------------------------------------------------------------------
    Dim oGeometryFactory As IngrGeom3D.GeometryFactory
    Dim oPlanesFactory As IPlanes3d
    
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory
    Set oPlanesFactory = oGeometryFactory
    
    sMsg = "Setting the Sketching Plane"
    Dim oTubeCutSketchingPlane As IJPlane
    Set oTubeCutSketchingPlane = oPlanesFactory.CreateByPointNormal(Nothing, _
                                                                    oSketchPlaneRootPt.x, oSketchPlaneRootPt.y, oSketchPlaneRootPt.z, _
                                                                    oSketchPlane_N.x, oSketchPlane_N.y, oSketchPlane_N.z)
                                                                      
    sMsg = "Setting the U direction of Sketching Plane"
    

    oTubeCutSketchingPlane.SetUDirection oSketchPlane_U.x, oSketchPlane_U.y, oSketchPlane_U.z
                       
    Set GetSketchPlaneForTube = oTubeCutSketchingPlane
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT, sMsg).Number

End Function
