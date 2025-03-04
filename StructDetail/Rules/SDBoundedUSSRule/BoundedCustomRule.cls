VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoundedCustomRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'-----------------------------------------------------------------------------------------
'  Copyright (C) 2001 - 2011, Intergraph Corporation.  All rights reserved.
'
'
'  File Info:
'      Folder:   S:\StructDetail\Middle\Symbols
'      Project:  BoundedSymbol
'      Class:    BoundedCustomRule
'
'  Abstract:
'       This class module contains the code that:
'       Default Bounded Custom Rule that allows modification to the
'       Bounded Symbol (End Cut) process based on outside customer requirements
'
'  Notes:
'   These methods are called for every Bounded End Cut (Web Cut, Flange Cut)
'   It is recommended that only required processing is done so the performance
'   is NOT adversly affected.
'
'
'  History:
'-----------------------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------'

Private Const MODULE = "\StructDetail\Middle\Symbols\BoundedSymbol\BoundedCustomRule.cls"

' The SDBoundedUSSRule must support the following Interface
Implements IJBoundedUSSRule
'
'********************************************************************
'********************************************************************
'********************************************************************
Private Sub Class_Initialize()

End Sub

'********************************************************************
Private Sub Class_Terminate()

End Sub

'********************************************************************
'********************************************************************
'********************************************************************
' _PreComputeConverter
'
'   In:
'       oEndCutObject       : IJStructFeature
'       oDefinitionPlayerEx : IMSSymbolEntities.IJDDefinitionPlayerEx
'       oBoundedObject      : IJPort
'       oBoundingObject     : IJPort, IJplane
'       oConverterData      : IJSTGeomConverter
'       lEndCutType         : IMSProfileEntity.eEndCutTypes
'   Out:
'
' Usage case 1 (possible):
'   For End To End type cuts:
'   The user sometimes wants to apply the End Cut such that is based on a
'   Global Axis (Plane) based on the Landing Curve vector x,y,z values.
'   When this case is detected, the user may replaced the Bounding object
'   with a Struct Plane object that represents the Global Axis to be used
'
' Notes:
'
'********************************************************************



Public Sub IJBoundedUSSRule_PreComputeConverter(ByVal oEndCutObject As Object, _
                                              ByVal oDefinitionPlayerEx As Object, _
                                             ByVal oBoundedObject As Object, _
                                             ByVal oBoundingObject As Object, _
                                             ByVal oConverterData As Object, _
                                            ByVal lEndCutType As Long)
                                            
Const MT = "_PreComputeConverter"
On Error GoTo ErrorHandler

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub

'********************************************************************
' _PostComputeConverter
'
'   In:
'       oEndCutObject       : IJStructFeature
'       oDefinitionPlayerEx : IMSSymbolEntities.IJDDefinitionPlayerEx
'       oBoundedObject      : IJPort
'       oBoundingObject     : IJPort, IJplane
'       oConverterData      : IJSTGeomConverter
'       lEndCutType         : IMSProfileEntity.eEndCutTypes
'   Out:
'
' Usage case 1:
'   For the Bounded Member Tube by a Gusset Plate Lateral Edge
'   The user wants to control the Sketching Plane such that the U vector
'   is pointing iin the same direction as the Gusset Plate Base/Offset Surfaces
'   so that the Web Cut can be used to Snipe the Member Tube relative to the
'   Bounding Gusset Plate for applying end caps to the Bounded Member Tube
'
' Notes:
'   The Geometry Converter Data contains the following:
'   .ViewTransform
'       4x4 transformation Matrix used for converting 3D to 2D and 2D to 3D
'   .SketchingPlane
'       Normal vector projects elements contained on Symbol Default Layer
'   .CrossSectionMatrix
'       Normal vector projects elements contained on Symbol CrossSection Layer
'   .BasePlateDirection
'       Normal vector projects elements contained on Symbol BasePlate Layer
'   .BasePlatePortGeometry
'       geometry used as the Idealized Boundary
'
'********************************************************************
Public Sub IJBoundedUSSRule_PostComputeConverter(ByVal oEndCutObject As Object, _
                                               ByVal oDefinitionPlayerEx As Object, _
                                               ByVal oBoundedObject As Object, _
                                              ByVal oBoundingObject As Object, _
                                              ByVal oConverterData As Object, _
                                              ByVal lEndCutType As Long)
                                              
Const MT = "_PostComputeConverter"
On Error GoTo ErrorHandler
    
    Dim oConnectable As Object
    Dim lReturnCode As Long
    
    lReturnCode = 0
    If oBoundedObject Is Nothing Then
        Exit Sub
    ElseIf oBoundingObject Is Nothing Then
        Exit Sub
    ElseIf Not IsMemberPartTubeType(oBoundedObject) Then
        Exit Sub
    ElseIf Not TypeOf oBoundingObject Is IJPort Then
        Exit Sub
    
    ElseIf IsBoundingPlateLateralPort(oBoundingObject, oConnectable) Then
            
        AdjustTubeMatrixForGussetPlate oEndCutObject, oDefinitionPlayerEx, _
                                       oBoundedObject, oBoundingObject, _
                                       oConnectable, _
                                       oConverterData, lEndCutType, _
                                       lReturnCode
                                       
    ElseIf IsMemberPartTubeType(oBoundedObject) = True And IsMemberPartTubeType(oBoundingObject) = False Then
            
        ' Verify the bounding Port
        Dim sPortType As String
        sPortType = Get_PortFaceType(oBoundingObject)
            
        Select Case sPortType
            Case C_BaseSide:
            
            Case C_OffsetSide:
            
            Case Else:
                ' bounding Port is Lateral Face Port
            
                AdjustSketchingPlaneForBoundedTube oEndCutObject, oDefinitionPlayerEx, _
                                                oBoundedObject, oBoundingObject, _
                                                oConnectable, _
                                                oConverterData, lEndCutType, _
                                                lReturnCode
        End Select
    
    ElseIf IsMemberPartTubeType(oBoundedObject) Then
           
        If IsFreeEndCut(oEndCutObject) Then
           
            RotateSketchingPlaneForFreeEndCutOnTube oEndCutObject, oDefinitionPlayerEx, _
                                                    oBoundedObject, oBoundingObject, _
                                                    oConnectable, _
                                                    oConverterData, lEndCutType, _
                                                    lReturnCode
        End If
        
    End If


    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub

'********************************************************************
' _PreEvaluate
'
'   In:
'       oEndCutObject       : IJStructFeature
'       oDefinitionPlayerEx : IMSSymbolEntities.IJDDefinitionPlayerEx
'       oConverterData      : IJSTGeomConverter
'       oGameOutputRep      : IMSSymbolEntities.IJDRepresentationDuringGame
'       lEndCutType         : IMSProfileEntity.eEndCutTypes
'   Out:
'
' Usage case 1 (possible):
'   For the cases where the user need to control the projection direction
'   based on data outside the BoundedSymbol USS control
'   The user may modify the normals in contained in the ConverterData
'
' Notes:
'   The Geometry Converter Data contains the following:
'   .SketchingPlane
'       Normal vector projects elements contained on Symbol Default Layer
'   .CrossSectionMatrix
'       Normal vector projects elements contained on Symbol CrossSection Layer
'   .BasePlateDirection
'       Normal vector projects elements contained on Symbol BasePlate Layer
'
'********************************************************************

Private Sub IJBoundedUSSRule_PreEvaluate(ByVal oEndCutObject As Object, ByVal oDefinitionPlayerEx As Object, ByVal oConverterData As Object, ByVal oGameOutputRep As Object, ByVal lEndCutType As Long)
Const MT = "_PreEvaluate"
On Error GoTo ErrorHandler


    Dim sPrefix As String
    Dim oConverter As IJSTGeomConverter
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    
        
    Set oConverter = oConverterData
    On Error Resume Next
    
    Dim oFlangeCutPoints As Object
    Dim oSmartOcc As IJSmartOccurrence
    
    If lEndCutType = 0 Then
        sPrefix = "WebCut:"
        Exit Sub
        ' no Need
    Else
        sPrefix = "FlangeCut:"
    End If
    
    Get_BoundedSymbolGrahicInputs oDefinitionPlayerEx, _
                                  oBoundedObject, _
                                  oBoundingObject, _
                                  oFlangeCutPoints

    If oBoundedObject Is Nothing Then
    
        Exit Sub
    ElseIf oBoundingObject Is Nothing Then
    
        Exit Sub
    'Check1: Bounded member should not be PIPE
    ElseIf Not IsMemberPartTubeType(oBoundedObject) Then
        'Check2: Bounding should be PIPE
        If IsMemberPartTubeType(oBoundingObject) Then
    
        Dim bPenetratesWeb As Boolean
        bPenetratesWeb = IsWebPenetrated(oBoundingObject, oBoundedObject)
    
            If bPenetratesWeb Then
                If TypeOf oEndCutObject Is IJSmartOccurrence Then
                    Set oSmartOcc = oEndCutObject
                End If
                If sPrefix = "FlangeCut:" Then
                    'this method should trigger only when sprefix is FlangeCut
                    AdjustFlgCuttingDirection oEndCutObject, oConverterData, lEndCutType, oDefinitionPlayerEx
                End If
            End If
        End If

    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub


'********************************************************************
' _PostEvaluate
'
'   In:
'       oEndCutObject       : IJStructFeature
'       oDefinitionPlayerEx : IMSSymbolEntities.IJDDefinitionPlayerEx
'       oConverterData      : IJSTGeomConverter
'       oGameOutputRep      : IMSSymbolEntities.IJDRepresentationDuringGame
'       lEndCutType         : IMSProfileEntity.eEndCutTypes
'   Out:
'
' Usage case 1:
'   When The Bounding Object (Member Part Tube) is in a Nodal type connection
'   where there are Multiple Bounding Member Part Tubes
'   The user wants to control the Extrusion distances used such that the
'   Bounded member is NOT completely cut by each of the Bounding Member Tubes
'   When this case is detected, the user can reduce the PositiveExtrusion or the
'   NegativeExtrusion output parameter value
'
' Notes:
'
'********************************************************************
Public Sub IJBoundedUSSRule_PostEvaluate(ByVal oEndCutObject As Object, _
                                        ByVal oDefinitionPlayerEx As Object, _
                                        ByVal oConverterData As Object, _
                                       ByVal oGameOutputRep As Object, _
                                       ByVal lEndCutType As Long)
Const MT = "_PostEvaluate"
On Error GoTo ErrorHandler

    Dim sName As String
    Dim sPrefix As String
    Dim sPortType As String
    
    Dim eAxisPortIndex As SPSMemberAxisPortIndex
    
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    Dim oFlangeCutPoints As Object
    
    Dim oSplitAxisPort As ISPSSplitAxisPort
    Dim oRepresentation As IJDRepresentation
    Dim oGameRepresentation As IJDRepresentationDuringGame
    Dim lReturnCode As Long
    lReturnCode = 0
    
    If lEndCutType = 0 Then
        sPrefix = "WebCut:"
    Else
        ' no Need to adjust Flange Cut Extrusion Distances
        sPrefix = "FlangeCut:"
        Exit Sub
    End If
    
    ' Verify Representation
    Set oRepresentation = oGameOutputRep
    Set oGameRepresentation = oGameOutputRep
    If oRepresentation.Name <> "WebCut" Then
        Exit Sub
    End If
    
    Dim bBoxCut As Boolean
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim sItemName As String
    
    ' Check if Member Part Tube bounded by Member Part Tube
    Get_BoundedSymbolGrahicInputs oDefinitionPlayerEx, _
                                  oBoundedObject, _
                                  oBoundingObject, _
                                  oFlangeCutPoints
    
    If oBoundedObject Is Nothing Then
        Exit Sub
    ElseIf oBoundingObject Is Nothing Then
        Exit Sub
    ElseIf Not IsMemberPartTubeType(oBoundedObject) Then
        Exit Sub
    ElseIf IsMemberPartTubeType(oBoundingObject) Then
        'change made to exclude box cuts from calling method AdjustExtrusion()
        If TypeOf oEndCutObject Is IJSmartOccurrence Then
            Set oSmartOcc = oEndCutObject
            Set oSmartItem = oSmartOcc.SmartItemObject
            sItemName = oSmartItem.Name
        End If
        
        'check to see if the current generic member AC is part of the box cut
        If InStr(1, sItemName, "LongBox", vbTextCompare) > 0 Then
            bBoxCut = True
        ElseIf InStr(1, sItemName, "ShortBox", vbTextCompare) > 0 Then
            bBoxCut = True
        End If
        
        If bBoxCut = False Then
            ' Verify that bounding Port is Lateral Face Port
            sPortType = Get_PortFaceType(oBoundingObject)
            If sPortType = C_BaseSide Then
            ElseIf sPortType = C_OffsetSide Then
            Else
                AdjustExtrusionForBoundingTube oEndCutObject, oDefinitionPlayerEx, _
                                               oBoundedObject, oBoundingObject, _
                                               oConverterData, oGameOutputRep, _
                                               lEndCutType, lReturnCode
            End If
            Exit Sub
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub


'********************************************************************
' _ResolveBoundedEdge
'
'   In:
'       oEndCutObject       : IJStructFeature
'       oDefinitionPlayerEx : IMSSymbolEntities.IJDDefinitionPlayerEx
'       oConverterData      : IJSTGeomConverter
'       lEndCutType         : IMSProfileEntity.eEndCutTypes
'       oBoundedObject      : Bounded Object: Port
'       sEdgeId             : Name of requested Edge Id
'       sListEdgeIds        : List of Edge Ids in oListEdgeGeometry
'       oListEdgeGeometry   : List of Edge Geometry
'       lEdgeIndex          : Index in oListEdgeGeometry of Edge Id to use
'   Out:
'
' Usage case 1:
'   When The Bounded Object is a Member Part Tube
'   There are no valid Edge Ids
'   For Case, return the Member Part Axis curve for the requested Edge Id
'
' Notes:
'
'********************************************************************
Private Sub IJBoundedUSSRule_ResolveBoundedEdge(ByVal oEndCutObject As Object, ByVal oDefinitionPlayerEx As Object, ByVal oConverterData As Object, ByVal lEndCutType As Long, ByVal oBoundedObject As Object, ByVal sEdgeId As String, sListEdgeIds() As String, ByVal oListEdgeGeometry As IMSCoreCollections.IJElements, lEdgeIndex As Long)
Const MT = "_ResolveBoundedEdge"
On Error GoTo ErrorHandler

    Dim sMsg As String
    Dim nEdges As Long
    Dim oport As IJPort

    Dim oTempAxis As IJCurve
    Dim oConverter As IJSTGeomConverter
    Dim oBoundedAxis As IJCurve
    Dim oBoundedPart As ISPSMemberPartCommon
    Dim lReturnCode As Long
    lReturnCode = 0
    sMsg = "BoundedSymbol.BoundedCustomRule::ResymbBoundedPart"

    ' Verify that current Bounded object is a Member Part Tube
    If lEdgeIndex > 0 Then
        Exit Sub
    ElseIf oBoundedObject Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oBoundedObject Is IJPort Then
        Exit Sub
    ElseIf Not IsMemberPartTubeType(oBoundedObject) Then
        Exit Sub
    End If

    ' get Bounded Member Axis Curve and Tube Radius
    If TypeOf oBoundedObject Is IJPort Then
        Set oport = oBoundedObject
        Set oBoundedPart = oport.Connectable
        Set oBoundedAxis = oBoundedPart.Axis

    ElseIf TypeOf oBoundedObject Is ISPSMemberPartCommon Then
        Set oBoundedPart = oBoundedObject
        Set oBoundedAxis = oBoundedPart.Axis

    Else
        Exit Sub
    End If

    ' Check if want to simulate Top/Bottom for Web Cuts
    ' Check if want to simulate WebLeft/WebRight for Flange Cuts
    On Error Resume Next
    Set oTempAxis = oBoundedAxis
    If lEndCutType = WebCut Then
        If sEdgeId = "514" Then
            Set oConverter = oConverterData
            TranslateAxisCurve oBoundedObject, oBoundedPart, oBoundedAxis, _
                               oConverter.ViewTransform, True, oTempAxis, oEndCutObject
        ElseIf sEdgeId = "513" Then
            Set oConverter = oConverterData
            TranslateAxisCurve oBoundedObject, oBoundedPart, oBoundedAxis, _
                               oConverter.ViewTransform, False, oTempAxis, oEndCutObject
        End If
    Else
        If sEdgeId = "257" Then
            Set oConverter = oConverterData
            TranslateAxisCurve oBoundedObject, oBoundedPart, oBoundedAxis, _
                               oConverter.ViewTransform, True, oTempAxis, oEndCutObject
        ElseIf sEdgeId = "258" Then
            Set oConverter = oConverterData
            TranslateAxisCurve oBoundedObject, oBoundedPart, oBoundedAxis, _
                               oConverter.ViewTransform, False, oTempAxis, oEndCutObject
        End If
    End If

    On Error GoTo ErrorHandler
    If oTempAxis Is Nothing Then
        Set oTempAxis = oBoundedAxis
    End If

    ' add the Member Axis to the list of Edge Ids and Edge Geometry
    oListEdgeGeometry.Add oTempAxis
    nEdges = oListEdgeGeometry.Count
    If nEdges < 1 Then
        nEdges = 1
    End If

    ReDim Preserve sListEdgeIds(0 To nEdges)
    sListEdgeIds(nEdges - 1) = sEdgeId
    sMsg = sMsg & vbCrLf & _
           sEdgeId & " ... as Bounding Axis Curve"

    ' Return the Edge Index for the Member Axis
    lEdgeIndex = oListEdgeGeometry.Count

    lReturnCode = 1

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub


''********************************************************************
'' _ResolveBoundingEdge
''
''   In:
''       oEndCutObject       : IJStructFeature
''       oDefinitionPlayerEx : IMSSymbolEntities.IJDDefinitionPlayerEx
''       oConverterData      : CommonSymbolUtils.STGeomConverter
''       lEndCutType         : IMSProfileEntity.eEndCutTypes
''       oBoundingObject     : Bounding Object: Port, Plane, etc
''       sEdgeId             : Name of requested Edge Id
''       sListEdgeIds        : List of Edge Ids in oListEdgeGeometry
''       oListEdgeGeometry   : List of Edge Geometry
''       lEdgeIndex          : Index in oListEdgeGeometry of Edge Id to use
''   Out:
''
'' Usage case 1:
''   When The Bounding Object is a Member Part Tube
''   There are no valid Edge Ids
''   For Case, return the Member Part Axis curve for the requested Edge Id
''
'' Notes:
''
''********************************************************************
Private Sub IJBoundedUSSRule_ResolveBoundingEdge(ByVal oEndCutObject As Object, ByVal oDefinitionPlayerEx As Object, ByVal oConverterData As Object, ByVal lEndCutType As Long, ByVal oBoundingObject As Object, ByVal sEdgeId As String, sListEdgeIds() As String, ByVal oListEdgeGeometry As IJElements, lEdgeIndex As Long)
Const MT = "_ResolveBoundingEdge"
On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim nEdges As Long
    Dim oport As IJPort

    Dim oBoundingAxis As IJCurve
    Dim oBoundingPart As ISPSMemberPartCommon
       Dim lReturnCode As Long
    lReturnCode = 0
    sMsg = "BoundedSymbol.BoundedCustomRule::ResymbBoundingPart"

    ' Verify that current Bounding object is a Member Part Tube
    If lEdgeIndex > 0 Then
        Exit Sub
    ElseIf oBoundingObject Is Nothing Then
        Exit Sub
    ElseIf Not TypeOf oBoundingObject Is IJPort Then
        Exit Sub
    ElseIf Not IsMemberPartTubeType(oBoundingObject) Then
        Exit Sub
    End If

    ' get Bounded Member Axis Curve and Tube Radius
    If TypeOf oBoundingObject Is IJPort Then
        Set oport = oBoundingObject
        Set oBoundingPart = oport.Connectable
        Set oBoundingAxis = oBoundingPart.Axis

    ElseIf TypeOf oBoundingObject Is ISPSMemberPartCommon Then
        Set oBoundingPart = oBoundingObject
        Set oBoundingAxis = oBoundingPart.Axis

    Else
        Exit Sub
    End If

    ' add the Member Axis to the list of Edge Ids and Edge Geometry
    oListEdgeGeometry.Add oBoundingAxis
    nEdges = oListEdgeGeometry.Count
    If nEdges < 1 Then
        nEdges = 1
    End If

    ReDim Preserve sListEdgeIds(0 To nEdges)
    sListEdgeIds(nEdges - 1) = sEdgeId
    sMsg = sMsg & vbCrLf & _
           sEdgeId & " ... as Bounding Axis Curve"

    ' Return the Edge Index for the Member Axis
    lEdgeIndex = oListEdgeGeometry.Count

    lReturnCode = 1

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, MT).Number
End Sub

