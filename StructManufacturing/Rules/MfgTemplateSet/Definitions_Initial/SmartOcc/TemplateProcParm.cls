VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateProcParm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************************************
'  Copyright (C) 2004, Intergraph Corporation.  All rights reserved.
'
'  Project      : StructMfgSelectors
'  File         : TemplateProcParm.cls
'
'  Description  : Template process Parameter rule
'
'  Author       :
'
'  History      : 2004.07.24 MJV Modified the selection based on discussion with USC and their current
'                                Hicadec implementation
'
'
'**********************************************************************************************************
Option Explicit

Private Const IJDMfgTemplateSet = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"
Private m_SupportPlatePart As IJMfgWrapPlatePart

Const MODULE = "TemplateProcParm"

Const m_ParameterRuleProgid As String = "StructMfgSelectors.TemplateProcParm"
Const m_ParameterRuleName As String = "StructMfgSelectors.TemplateProcParm"
Const m_FamilyProgid As String = ""

Implements IJDUserSymbolServices

Public Sub ParameterRuleInputs(pIH As IJDInputsHelper)
Const method = "ParameterRuleLogic"
On Error GoTo ErrorHandler
    
  pIH.SetInput INPUT_PLATE
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, method).Number
  'pIH.ReportError
End Sub

Public Sub ParameterRuleLogic(pPLH As IJDParameterLogic)
  Const method = "ParameterRuleLogic"
  On Error GoTo ErrorHandler
    
Dim oPlate As Object
Set oPlate = pPLH.InputObject(INPUT_PLATE)

Dim ePlateType As StructPlateType

Dim oSDPlateWrapper As StructDetailObjects.PlatePart
Set oSDPlateWrapper = CreateObject("StructDetailObjects.PlatePart")
Dim oSDPlate As Object
Set oSDPlate = pPLH.InputObject(INPUT_PLATE)
Set oSDPlateWrapper.object = oSDPlate
ePlateType = oSDPlateWrapper.plateType

Dim nIndex As Long
Dim nIndex1 As Long

'* Determine the plate part category and see if there has been a special assignment
'* Currently we are only able to retrieve the naming category.

Dim PlateCat As Integer
PlateCat = oSDPlateWrapper.NamingCategory

Dim PlateCategory As String

Select Case PlateCat
       Case Is = 1      'Shell
            ' Do nothing
       Case Is = 24     'Echo Sounder Recess
            ' Do nothing
       Case Is = 26     'Anchor Recess
            ' Do nothing
       Case Is = 29     'Bilge Keel
            PlateCategory = "KEEL"
       Case Is = 32     'Flat of Side
            PlateCategory = "FOS"
       Case Is = 36     'Flat of Bottom
            PlateCategory = "FOB"
End Select


' * Determine if there is an intersect with the naval arch reference curves
'   SPL = Side Parallel Line  ==> Flat_of_Side_[P/S]_[A/F]
'   BFL = Bottom Flat Line ==> Flat_of_Bottom_[P/S]_[A/F]
Dim SPL As Boolean
Dim BFL As Boolean

SPL = False
BFL = False

'Get the Naval Arch reference curves that intersect the part
Dim oReferenceCurvesCol  As Collection
oSDPlateWrapper.ReferenceCurves PlateBaseSide, oReferenceCurvesCol

'* Determine if the adjacent part is positioned around the stem/stern (HasIntersectionline = true)
Dim AdjacentPartCenterline As Boolean
Dim MyPlateEdges As Collection
Dim oPort As IJPort
Dim oPortGeometry As IUnknown
Dim oAdjacentObject As StructDetailObjects.PlatePart
Dim oCenterline As Boolean

AdjacentPartCenterline = False

'If the part itself has a centerline intersection the code below will also return True
'Since one of the edges will return an object that is also along the centerline
'However this should be avoided

If oSDPlateWrapper.HasCenterLineIntersection = False Then
    'Retrieve all the edges to which a part could be connected
    Dim oSDO_Helper As StructDetailObjects.Helper

    On Error Resume Next
    Dim nConnectionData As Long
    Dim aConnectionData As ConnectionData
    Dim oConObjsCol() As ConnectionData

    nConnectionData = 0
    Set oSDO_Helper = New StructDetailObjects.Helper
    oSDO_Helper.Object_AppConnections oSDPlate, AppConnectionType_Physical, _
                                      nConnectionData, oConObjsCol
'    oSDO_Helper.Object_AppConnections oSDPlate, AppConnectionType_Assembly, _
'                                      nConnectionData, oConObjsCol
    Dim bIsSubPort As Boolean
    Dim lPortType As Long

    For nIndex = 1 To nConnectionData

        aConnectionData = oConObjsCol(nIndex)

        Set oPort = aConnectionData.ToConnectedPort
        lPortType = oSDO_Helper.GetPortType(oPort, bIsSubPort)

        If lPortType = BPT_Lateral Then
            'We found a part on the lateral face
            If TypeOf aConnectionData.ToConnectable Is IJPlatePart Then
            'Check if it is a plate part
                Set oAdjacentObject = New StructDetailObjects.PlatePart
                Set oAdjacentObject.object = aConnectionData.ToConnectable
                oCenterline = oAdjacentObject.HasCenterLineIntersection
                If oCenterline = True And AdjacentPartCenterline = False Then
                    AdjacentPartCenterline = True
                    Exit For
                End If
                Set oAdjacentObject = Nothing
            End If
        End If
        Set oPort = Nothing
    Next nIndex
End If

' * determine if we have a general plate or not
' * We are not able to detect the following:
' *   = No marking after bending
'* We should be able to determine the setting of the Frameline marking and therefore base our
'  value of:
'*   = Frame lines are used for reference lines.
'*   = Shape is either triangle or rectangle
'*   = Butt lines are parallel to the frame lines (Transversal, Vertical or MostlyVertical)

Dim General As Boolean
Dim PlateShape As Boolean
Dim ButtParallel As Boolean
Dim FrameLines As Boolean
Dim Template As Boolean

General = False
PlateShape = False
ButtParallel = False
FrameLines = False
Template = False

'*************************
'* Determine Plate Shape *
'*************************

'* Determine the number of lateral edge faces
Dim nEdges              As Long
Dim oCollectionOfEdges  As Collection

'Retrieve the Part's Edge Ports
Set oCollectionOfEdges = oSDPlateWrapper.PlateEdges

nEdges = oCollectionOfEdges.Count + 1

If nEdges = 9 Or nEdges = 12 Then
    ' Plate is either a triangle or a rectangular
    PlateShape = True
Else
    PlateShape = False
End If

'******************************************************************************
'* Determine Butt parallel and intersections with Naval Arch reference curves *
'******************************************************************************

Dim nButts As Long
Dim SDSeams As Collection
Dim oBoundary As BoundaryData
Dim oSeamOrientation As eSeamOrientationType
Dim oSeam As StructDetailObjects.Seam

'Get all the seams on the object regardless of the type of seam (straking, design,etc.)
Set SDSeams = oSDPlateWrapper.PlateBoundaries
nButts = 0

Dim nIndex2 As Integer
Dim iPos As Integer
Dim oCommonGeom As IUnknown
Dim oRefCurveData As IJRefCurveData
Dim RefCurveType As JSRCOS_TYPES

Dim oPortList As IJElements
Dim oBoundlist As IJElements
Dim oPartInfo As GSCADStructGeomUtilities.PartInfo
Dim nPorts As Integer
Dim iIndex As Integer

Set oPortList = New JObjectCollection
Set oBoundlist = New JObjectCollection
Set oPartInfo = New GSCADStructGeomUtilities.PartInfo

Dim oPlateSystem As IJPlateSystem
Dim oStructDetailHelper As StructDetailHelper
Set oStructDetailHelper = New StructDetailHelper
oStructDetailHelper.IsPartDerivedFromSystem oSDPlate, oPlateSystem
Set oStructDetailHelper = Nothing

If Not oPlateSystem Is Nothing Then
    oPartInfo.GetPlatePartPortBoundaries oSDPlate, oPortList, oBoundlist
    nPorts = oPortList.Count
Else
    nPorts = 0
End If

Dim oFacePort As IJPort

For nIndex = 1 To oBoundlist.Count

    If TypeOf oBoundlist.Item(nIndex) Is IJSeam Then
        Set oSeam = New StructDetailObjects.Seam
        Set oSeam.object = oBoundlist.Item(nIndex)
        oSeamOrientation = oSeam.Orientation

        If oReferenceCurvesCol.Count > 0 Then

            For nIndex1 = 1 To oReferenceCurvesCol.Count
                Dim oWB As IJWireBody
                Dim oMarkLines As Collection

                Set oRefCurveData = oReferenceCurvesCol.Item(nIndex1)
                If Not oRefCurveData Is Nothing Then
                    Set oMarkLines = oRefCurveData.GetMarkingLineCollection
                    For nIndex2 = 1 To oMarkLines.Count
                        Set oWB = oMarkLines.Item(nIndex2)
                        If Not oWB Is Nothing Then
                            If oSeamOrientation = MostlyVertical Or oSeamOrientation = Vertical Or oSeamOrientation = Transveral Then
                                nButts = nButts + 1
                                'Check if there is a common point between the butt and the Naval Arch line
                                RefCurveType = oRefCurveData.Type
                                If RefCurveType = JSRCOS_TANGENT Then
                                     'Search for String "Flat_of_bottom" in Curve name from start in text compare mode
                                     iPos = InStr(1, oRefCurveData.Name, "Flat_of_Bottom", vbTextCompare)
                                     If iPos = 1 Then 'ie at the start of string
                                        BFL = True
                                     Else
                                         iPos = InStr(1, oRefCurveData.Name, "Flat_of_Side", vbTextCompare)
                                         If iPos = 1 Then
                                            SPL = True
                                         End If
                                     End If
                                End If
                            End If
                        End If
                        Set oWB = Nothing
                    Next nIndex2 'Check for each marking curve
                    Set oRefCurveData = Nothing
                    Set oMarkLines = Nothing
                End If
            Next nIndex1 'Check against each reference curve
        End If
    End If
Next nIndex
Set oBoundlist = Nothing
Set oPortList = Nothing
Set oSeam = Nothing
Set oMarkLines = Nothing
Set oPartInfo = Nothing

'Depending if we have a triangle the number of butts is 1 in case of a rectangle it is 2
'Funny enough the rectangular plate returns 13 edges, need to check a triangular plate
If nEdges = 9 And nButts >= 1 Then
    ButtParallel = True
Else
    If nEdges = 12 And nButts >= 2 Then
        ButtParallel = True
    Else
        ButtParallel = False
    End If
End If

''**************************************
''* Determine FrameLine marking on/off *
''**************************************
Dim oRefPlanCol As Collection
Dim oRefCurvesCol As Collection

' Use the active coordinate system to find out if there are frame intersections

oSDPlateWrapper.ReferencePlaneRange GetParentFrameSystem(oSDPlate), "X", oRefPlanCol, oRefCurvesCol, BPT_Offset

If Not oRefCurvesCol Is Nothing Then
    If oRefCurvesCol.Count > 0 Then
        'We found more then 1 frame line intersecting the plate
         FrameLines = True
    End If
End If

'*******************************
'* Determine general or not    *
'*******************************

If PlateShape = True And FrameLines = True And ButtParallel = True Then
    General = True
End If

' * Finished getting the decision values

' * Determine unwrap algorithm to be used

'pPLH.Add "Side", 5110     'BaseSide
Dim side As Integer
side = GetConcaveSide(oPlate)
If side = 0 Then 'if GetConcaveside fails to give a side, use base side
    side = 5110 'BaseSide
End If
pPLH.Add "Side", side

If General = False Then
    If PlateCategory = "KEEL" Then
        pPLH.Add "Type", 5121     'Centerline
        pPLH.Add "Orientation", 5151    'NormalToBasePlate
        pPLH.Add "Direction", 5101   'Transversal
        pPLH.Add "PositionEven", 5160    'PositionEven
        pPLH.Add "PositionFrames", 5171      'NotUsed
        pPLH.Add "Extension", 5180    'Fixed
        pPLH.Add "BasePlane", 5192        'By System
    Else
        If AdjacentPartCenterline = True Then
            pPLH.Add "Type", 5122     'Perpendicular
            pPLH.Add "Orientation", 5152    'Perpendicular
            pPLH.Add "Direction", 5101   'Transversal
            pPLH.Add "PositionEven", 5160    'PositionEven
            pPLH.Add "PositionFrames", 5171      'NotUsed
            pPLH.Add "Extension", 5180    'Fixed
            pPLH.Add "BasePlane", 5192        'By System
        Else
            pPLH.Add "Type", 5120     'Frame
            pPLH.Add "Orientation", 5151    'NormalToBasePlate
            pPLH.Add "Direction", 5101   'Transversal
            pPLH.Add "PositionEven", 5161    'NotUsed
            pPLH.Add "PositionFrames", 5170      'PositionFrame
            pPLH.Add "Extension", 5180    'Fixed
            pPLH.Add "BasePlane", 5192        'By System
        End If
    End If
Else
    If BFL = True Then
        pPLH.Add "Type", 5120     'Frame
        pPLH.Add "Orientation", 5150    'AlongFrame
        pPLH.Add "Direction", 5101   'Transversal
        pPLH.Add "PositionEven", 5161    'NotUsed
        pPLH.Add "PositionFrames", 5170      'PositionFrame
        pPLH.Add "Extension", 5180    'Fixed
        pPLH.Add "BasePlane", 5190        'NormalToPlate
    Else
        If SPL = True Then
            pPLH.Add "Type", 5120     'Frame
            pPLH.Add "Orientation", 5150    'AlongFrame
            pPLH.Add "Direction", 5101   'Transversal
            pPLH.Add "PositionEven", 5161    'NotUsed
            pPLH.Add "PositionFrames", 5170      'PositionFrame
            pPLH.Add "Extension", 5180    'Fixed
            pPLH.Add "BasePlane", 5191        'ParallelToPlate
        Else
            pPLH.Add "Type", 5120     'Frame
            pPLH.Add "Orientation", 5151    'NormalToBasePlate
            pPLH.Add "Direction", 5101   'Transversal
            pPLH.Add "PositionEven", 5161    'NotUsed
            pPLH.Add "PositionFrames", 5170      'PositionFrame
            pPLH.Add "Extension", 5180    'Fixed
            pPLH.Add "BasePlane", 5192        'By System
        End If
    End If
End If
 

Cleanup:
    Set oSDPlateWrapper = Nothing
    Set oSDPlate = Nothing
    Set oRefCurvesCol = Nothing
    Set oRefPlanCol = Nothing
    Set SDSeams = Nothing
    Set oPlate = Nothing
    Set oReferenceCurvesCol = Nothing
    Set oCollectionOfEdges = Nothing
    Set MyPlateEdges = Nothing
    
  Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, method).Number
  GoTo Cleanup:
  'pPLH.ReportError
End Sub

' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  IJDUserSymbolServices_GetDefinitionName = m_ParameterRuleName
End Function
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
  Dim pSelectorFact As New DefinitionFactory
  pSelectorFact.InitAbstractParameterRule pSelector
 
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pSelector
  pIH.InitAs m_FamilyProgid
  ParameterRuleInputs pIH
  
 '----------- Define Output and Add it to the representation ------------------
  Dim pOH As IJDOutputsHelper
  Set pOH = New OutputHelper
  pOH.Representation = pSelector.IJDRepresentations.Item(1)

  pOH.SetOutput "Side"
  pOH.SetOutput "Type"
  pOH.SetOutput "Orientation"
  pOH.SetOutput "Direction"
  pOH.SetOutput "PositionEven"
  pOH.SetOutput "PositionFrames"
  pOH.SetOutput "Extension"
  pOH.SetOutput "BasePlane"
  DefineViewPrefix pSelector, "INGR"
  Set pOH = Nothing

End Sub
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object
  Dim pSelectorFact As New DefinitionFactory
  Set IJDUserSymbolServices_InstanciateDefinition = pSelectorFact.InstanciateParameterRule(m_ParameterRuleProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
End Function
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
End Sub
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
End Function
Public Sub CMParameterRule(pRep As IJDRepresentation)

  Dim pPLH As IJDParameterLogic
  Set pPLH = New ParameterLogic
  pPLH.Representation = pRep
  ParameterRuleLogic pPLH
End Sub



Private Function GetConcaveSide(ByRef pPlatePart As Object) As Integer
Const method = "GetConcaveSide"
On Error GoTo ErrorHandler

    'get COG of the plate
    Dim oCOGPosition As IJDPosition
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    Set oCOGPosition = oMfgGeomHelper.GetCenterOfGravityOfModelBody(pPlatePart)
  
    ''Get the distance of COG to Base surface
        'get the base surface
        'get the shortest distance from base surface
    Dim dMinDistToBase As Double
    Dim dMinDistToOffset As Double
    Dim oSurface As IJSurfaceBody
    Dim oModelBody As IJDModelBody
    Dim oPosOnSurface As IJDPosition
    
    'setup part support
    Dim oSDPartSupport As IJPartSupport
    Set oSDPartSupport = New PlatePartSupport
    Set oSDPartSupport.Part = pPlatePart
    
    Dim oSDPlatePartSupport As IJPlatePartSupport
    Set oSDPlatePartSupport = oSDPartSupport
    
    'get the base surface
    Call oSDPlatePartSupport.GetSurface(BaseSide, oSurface)
        
    'get the shortest distance from base surface
    Set oModelBody = oSurface
    Call oModelBody.GetMinimumDistanceFromPosition(oCOGPosition, oPosOnSurface, dMinDistToBase)
    
    'get the offset surface
    Call oSDPlatePartSupport.GetSurface(OffsetSide, oSurface)
    
    'get the shortest distance from base surface
    Set oModelBody = oSurface
    Call oModelBody.GetMinimumDistanceFromPosition(oCOGPosition, oPosOnSurface, dMinDistToOffset)
    
    If dMinDistToBase < dMinDistToOffset Then
        GetConcaveSide = 5110 'BaseSide
    Else
        GetConcaveSide = 5111 'OffsetSide
    End If

    
Cleanup:

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, method).Number
    GoTo Cleanup
End Function
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


'This method gets the default frame system of the Tempalte Set.Default frame
' system is the parent frame system of the plate
Private Function GetParentFrameSystem(oPlatePart As Object) As String
Const method = "GetParentFrameSystem"
On Error GoTo ErrorHandler

    Dim oFSUtils As IJDMfgFrameSysUtil
    Dim oParentFS As IJFrameSystem
    
    Set oFSUtils = New CMfgCoordSys
    Set oParentFS = oFSUtils.SystemDefaultCS(oPlatePart)
    
    If Not oParentFS Is Nothing Then
        Dim oNamedItem As IJNamedItem
        Set oNamedItem = oParentFS
        
        GetParentFrameSystem = oNamedItem.Name
    End If

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, method).Number
End Function

 
