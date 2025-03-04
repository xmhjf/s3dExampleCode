Attribute VB_Name = "ComputeConverterHelper"
'--------------------------------------------------------------------------------------------'
'    Copyright (C) 1998, 1999 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'    This helper is to get the proper plane with U and V Vecotor of Cable way
'    CrossSection and Horientation
'
'Notes
'
'
'History
'
'    sypark@ship.samsung.co.kr    02/21/02                Creation.
'    Suresh  29-Oct-03  Added IJHoletrace varible as one of the Argument to
'                        ComputeMatrixForGeometryConverter function Tr 50621
'--------------------------------------------------------------------------------------------'
Option Explicit
Private Const MODULE = "HMCWSymbolServices.ComputeConverterHelper"
Private Const PI = 3.14159265

'********************************************************************
' ConvertWorkingPlane
'
'
'       In:   oHoleTrace
'             oHorientation
'       Out:
'             IJPlane
'
'********************************************************************
Public Function ConvertWorkingPlane(oHoleTrace As IJHoleTraceAE) As IJPlane
    Const MT = "ConvertWorkingPlane"
    On Error GoTo ErrorHandler

    Dim oNewPlane As IJPlane
    Set oNewPlane = New Plane3d
    
    Dim oWorkingPlane As IJPlane
    Set oWorkingPlane = oHoleTrace.GetWorkingPlane
    Dim oIntersection As IJDPosition
    
    'Get the normal vector and set again becasue setUDirection will change the normal vector.
    Dim dNX  As Double, dNY As Double, dNZ As Double
    oWorkingPlane.GetNormal dNX, dNY, dNZ
    
    'Get the Outfitting, Only Cable Way should be taken.
    Dim oOutfitting As IJDObjectCollection
    Set oOutfitting = oHoleTrace.GetParentOutfitting

    Dim oObj As Object
    Dim oSelectedType As CrossSectionShapeTypes

    'For Loop input object ( 1 .. icount .. n ),
    'Only one CW should be taken.
    For Each oObj In oOutfitting
        If TypeOf oObj Is IJRteCablewayPathFeat Then
            
            Dim oRteObj As Object
            Set oRteObj = oObj

            'Find the intersection between Cable Way and workingplane
            Set oIntersection = IntersectionPoint(oRteObj, oWorkingPlane)
            Set oRteObj = Nothing
            If Not oIntersection Is Nothing Then
                oNewPlane.SetRootPoint oIntersection.x, oIntersection.y, oIntersection.z
'                oWorkingPlane.SetRootPoint oIntersection.x, oIntersection.y, oIntersection.z
            End If
            
            'Set the UVector of Plane to be same to UVector of CW CrossSection.
            Dim oRtePathFeat As IJRtePathFeat
            Set oRtePathFeat = oObj
            Dim oRteFeatUtility As IJRtePathCrossSectUtility
            
            Set oRteFeatUtility = oRtePathFeat
            
            Dim dUVectorX As Double, dUVectorY As Double, dUVectorZ As Double
            Dim dVVectorX As Double, dVVectorY As Double, dVVectorZ As Double
            
            'Get the U and V vector of Cableway Crosssection
            
            If Not oRteFeatUtility Is Nothing Then
                oRteFeatUtility.GetWidthAndDepthAxis Nothing, dUVectorX, dUVectorY, dUVectorZ, _
                                        dVVectorX, dVVectorY, dVVectorZ
            End If
            
            
'            oWorkingPlane.SetUDirection dUVectorX, dUVectorY, dUVectorZ
            oNewPlane.SetUDirection dUVectorX, dUVectorY, dUVectorZ

            
            'Set Normal vector again because SetUDirection will change the normal vector.
'            oWorkingPlane.SetNormal dNX, dNY, dNZ
            oNewPlane.SetNormal dNX, dNY, dNZ
            
            
            Set oRtePathFeat = Nothing
            Set oRteFeatUtility = Nothing
            Set ConvertWorkingPlane = oNewPlane
            
            Exit For
        End If
    Next oObj
    
    Set oIntersection = Nothing
    Set oWorkingPlane = Nothing
    Set oNewPlane = Nothing
    Set oObj = Nothing
    Set oOutfitting = Nothing
    
    Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function

Public Function ComputeMatrixForGeometryConverter(pPlane As IJPlane, dHorientation As Double) As IJDT4x4
   'OBM: the idea is to set coordinates of plane in new space.
Const MT = "ComputeMatrixForGeometryConverter"
On Error GoTo ErrorHandler

    Dim Nx As Double, Ny As Double, Nz As Double
    Dim Rx As Double, Ry As Double, Rz As Double
    Dim uX As Double, uY As Double, uZ As Double
    Dim Vx As Double, Vy As Double, Vz As Double
    
    Dim TransfoMatrix As AutoMath.DT4x4
    Set TransfoMatrix = New AutoMath.DT4x4
   
    pPlane.GetNormal Nx, Ny, Nz
    pPlane.GetRootPoint Rx, Ry, Rz
    pPlane.GetUDirection uX, uY, uZ
    pPlane.GetVDirection Vx, Vy, Vz
    
    Fill4x4Matrix TransfoMatrix, uX, uY, uZ, 0, Vx, Vy, Vz, 0, Nx, Ny, Nz, 0, Rx, Ry, Rz, 1
    
    Set ComputeMatrixForGeometryConverter = TransfoMatrix
    
    Exit Function
ErrorHandler:
    HandleError MODULE, MT
End Function

Private Sub Fill4x4Matrix(p4x4Matrix As AutoMath.DT4x4, M11 As Double, M21 As Double, M31 As Double, _
                          M41 As Double, M12 As Double, M22 As Double, M32 As Double, _
                          M42 As Double, M13 As Double, M23 As Double, M33 As Double, _
                          M34 As Double, TransX As Double, TransY As Double, TransZ As Double, _
                          ScalingCoeff As Double)
Const MT = "Fill4x4Matrix"
On Error GoTo ErrorHandler

    Dim ValueTransf(16) As Double
    
    p4x4Matrix.Get ValueTransf(0)
    
    ValueTransf(0) = M11
    ValueTransf(1) = M21
    ValueTransf(2) = M31
    ValueTransf(3) = M41
    ValueTransf(4) = M12
    ValueTransf(5) = M22
    ValueTransf(6) = M32
    ValueTransf(7) = M42
    ValueTransf(8) = M13
    ValueTransf(9) = M23
    ValueTransf(10) = M33
    ValueTransf(11) = M34
    ValueTransf(12) = TransX
    ValueTransf(13) = TransY
    ValueTransf(14) = TransZ
    ValueTransf(15) = ScalingCoeff

    p4x4Matrix.Set ValueTransf(0)

Exit Sub
ErrorHandler:
    HandleError MODULE, MT
End Sub

'********************************************************************
' GetCWHoleSymbolParameterInputs
'    The CableWay Hole Symbol requires additional Parameter Inputs that are NOT
'    defined in the Rectangle .sym file
'    add the additional Parameter Inputs
'
'       In:   IDisp i/f of Symbol Definition
'       Out:
'             dHorientation As Double
'
'********************************************************************
Public Sub GetCWHoleSymbolParameterInputs(ByVal pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition, _
                                       dHorientation As Double)
Const MT = "GetCWHoleSymbolParameterInputs"
On Error GoTo ErrorHandler

    'Get the grahic inputs to the Collar symbol
    Dim DefPlayerEx As IMSSymbolEntities.IJDDefinitionPlayerEx
    Set DefPlayerEx = pSymbolDefinition
    
    Dim oSymbolOcc As IJDSymbol
    Set oSymbolOcc = DefPlayerEx.PlayingSymbol
    
    Dim oInputs As IMSSymbolEntities.IJDInputs
    Set oInputs = pSymbolDefinition
    
    Dim found As Long
    Dim pEnumArg As IJDArgument
    Dim pEnumJDArgument As IEnumJDArgument
    Dim oParameterContent As IJDParameterContent

    Dim oGraphicInput As IMSSymbolEntities.IJDInput

    'Get the enum of arguments set by reference on the symbol if any
    Set pEnumJDArgument = oSymbolOcc.IJDInputsArg.GetInputs(igINPUT_ARGUMENTS_MERGE)
    If Not pEnumJDArgument Is Nothing Then
        pEnumJDArgument.Reset
        Do
            pEnumJDArgument.Next 1, pEnumArg, found
            If found = 0 Then Exit Do
            
            Set oGraphicInput = oInputs.Item(pEnumArg.Index)

            If oGraphicInput.Properties = igINPUT_IS_A_PARAMETER Then
            
                If oGraphicInput.Name = "Horientation" Then
                    Set oParameterContent = pEnumArg.Entity
                    dHorientation = oParameterContent.UomValue
                End If
                            
            End If
            
            Set oGraphicInput = Nothing
        Loop
    End If
    
   Exit Sub
    
ErrorHandler:
    HandleError MODULE, MT
End Sub


'*************************************************************************
' Routine: IntersectionPoint
'
' Abstract: Return the position of intersection between plane and Cable way
'
' Description: This code was patterned from StructDetail.
'*************************************************************************
Private Function IntersectionPoint(pObj1 As Object, pObj2 As Object) As IJDPosition
    
    Const MT = "IntersectionPoint"
    On Error GoTo ErrorHandler
    
    Dim oIntersector As IMSModelGeomOps.DGeomOpsIntersect
    Dim oIntersectedUnknown As IUnknown
    Dim pAgtorUnk As IUnknown
    Dim NullObject As Object
    
    Dim sx As Double, sy As Double, sz As Double
    Dim ex As Double, ey As Double, ez As Double
    Dim oPathFeat As IJRtePathFeat
    Dim oGeomFactory As GeometryFactory
    Dim oLine3d As IJLine
    
    Set oIntersector = New IMSModelGeomOps.DGeomOpsIntersect
    'Intersect
    On Error Resume Next 'To avoid error in process when no intersection occurs
    
    'Get the pathfeat and create line
    Set oPathFeat = pObj1
    oPathFeat.GetStartLocation sx, sy, sz
    oPathFeat.GetEndLocation ex, ey, ez
        
    Set oGeomFactory = New GeometryFactory
    Set oLine3d = oGeomFactory.Lines3d.CreateBy2Points(Nothing, sx, sy, sz, ex, ey, ez)
    Set pObj1 = oLine3d
       
    oIntersector.PlaceIntersectionObject NullObject, pObj1, pObj2, pAgtorUnk, oIntersectedUnknown
    
    On Error GoTo ErrorHandler
   
    Dim oPointsGraphBody As IJPointsGraphBody
    Dim oSGOPointsGraphUtilities As SGOPointsGraphUtilities
    Dim oPointsCollection As Collection
    
    If Not oIntersectedUnknown Is Nothing Then
        Set oPointsGraphBody = oIntersectedUnknown
          If Not oPointsGraphBody Is Nothing Then
            ' return the first intersection point
            Set oSGOPointsGraphUtilities = New SGOPointsGraphUtilities
            Set oPointsCollection = oSGOPointsGraphUtilities.GetPositionsFromPointsGraph(oPointsGraphBody)
            If Not oPointsCollection Is Nothing Then
                If oPointsCollection.Count > 0 Then
                    Set IntersectionPoint = oPointsCollection.Item(1)
                End If
            End If
        End If
    End If
    
    Set oPointsGraphBody = Nothing
    Set oSGOPointsGraphUtilities = Nothing
    Set oPointsCollection = Nothing

Cleanup:
    Set oIntersector = Nothing
    Set oIntersectedUnknown = Nothing
    Set pAgtorUnk = Nothing
    Set NullObject = Nothing
    
    Set oPathFeat = Nothing
    Set oGeomFactory = Nothing
    Set oLine3d = Nothing
    Exit Function
    
ErrorHandler:
    HandleError MODULE, MT
    GoTo Cleanup
End Function
 