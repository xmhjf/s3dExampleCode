Attribute VB_Name = "MfgProcessRulesHelper"
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Module: MfgProcessRulesHelper
'
' Description:  Common functions for Mfg process
'
' Author:
'
' Comments:
' 2010.06.16    Siva     Creation
'*******************************************************************************
Option Explicit

Private Const Module = "MfgProcessRulesHelper"
'********************************************************************************
'Description: This method checks if input is a Can plate
'********************************************************************************
Public Function CheckIfCanPlate(oPlateSystem As IJPlateSystem) As Boolean
    Const METHOD = "CheckIfCanPlate"
    On Error GoTo ErrorHandler
    
    CheckIfCanPlate = False
    
    If oPlateSystem.IsBuiltupPlateSystem Then
       Dim oMemberPartCommon As ISPSMemberPartCommon
       Set oMemberPartCommon = oPlateSystem.ParentBuiltup 'Parent built up

       If Not oMemberPartCommon Is Nothing Then
           Dim oMemberCrossSection As ISPSCrossSection
           Set oMemberCrossSection = oMemberPartCommon.CrossSection
           
           'Get the section type of parent built up
           If Not oMemberCrossSection Is Nothing Then
               If (oMemberCrossSection.SectionType = "BUCan") Then 'it is a can plate
                  CheckIfCanPlate = True
               End If
           End If
           
       End If
       
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'********************************************************************************
'Description: This method returns the eligible ports for can plates
'********************************************************************************
Public Sub GetEligiblePortCollectionForCan(oPlatePart As IJPlatePart, oPortColl As IJElements)
    Const METHOD = "GetEligiblePortCollectionForCan"
    On Error GoTo ErrorHandler
    
    'get the root and leaf plate systems
    Dim oRootPlateSystem    As IJPlateSystem
    Dim oLeafPlateSystem    As IJPlateSystem
    Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
        
    oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oRootPlateSystem, True
    oStructDetailHelper.IsPartDerivedFromSystem oPlatePart, oLeafPlateSystem
    
    Dim oPlateUtil As IJPlateAttributes
    Set oPlateUtil = New PlateUtils
    
    'Check if the Can was split
    Dim oLeafPlateSysColl As Collection
    Set oLeafPlateSysColl = oPlateUtil.GetSplitResults(oRootPlateSystem)
    
    If oLeafPlateSysColl.Count > 1 Then 'if can is split into two parts
         
        'get the splitters
        Dim oParent   As IUnknown
        Dim oSplitters As IEnumUnknown
        oStructDetailHelper.IsResultOfSplitWithOpr oLeafPlateSystem, oParent, oSplitters
        
        Dim oCollectionOfSplitters  As Collection
        Dim ConvertUtils            As CCollectionConversions
        Dim SplitterColl            As IJElements
        
        Set ConvertUtils = New CCollectionConversions
        ConvertUtils.CreateVBCollectionFromIEnumUnknown oSplitters, oCollectionOfSplitters
        Set SplitterColl = ConvertUtils.CreateIJElementsCollectionFromVBCollection(oCollectionOfSplitters)
        
        'Get the lateral face ports which are the result of split done with above splitters
        Set oPortColl = GetLateralFacePortsFromSplitters(oPlatePart, SplitterColl)
        
        Set oParent = Nothing
        Set oSplitters = Nothing
    End If
    
CleanUp:
    Set oRootPlateSystem = Nothing
    Set oLeafPlateSystem = Nothing
    Set oStructDetailHelper = Nothing
    Set oPlateUtil = Nothing
    Set oLeafPlateSysColl = Nothing
    Set SplitterColl = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp

End Sub
'********************************************************************************
'Description: This method returns the lateral face ports which are formed as
'             a result of input splitters
'********************************************************************************
Private Function GetLateralFacePortsFromSplitters(oPlatPart As IJPlatePart, SplitterColl As IJElements) As IJElements
    Const METHOD = "GetLateralFacePortsFromSplitters"
    On Error GoTo ErrorHandler
    
    Set GetLateralFacePortsFromSplitters = New JObjectCollection
    
    Dim oStructConnectable  As IJStructConnectable
    Set oStructConnectable = oPlatPart
    
    Dim oFacePortsColl      As IJElements
    
    'Get the all lateral face ports. OperationProgID is NULL. So, it gives the latest geometry
    oStructConnectable.enumConnectableTransientPorts oFacePortsColl, vbNullString, False, PortFace, JS_TOPOLOGY_FILTER_SOLID_LATERAL_LFACES, False
    
    Dim oStructDetailHelper As IJStructDetailHelper
    Set oStructDetailHelper = New StructDetailHelper
     
    Dim iCount As Long
    For iCount = 1 To oFacePortsColl.Count
    
       Dim oPort As IJStructPort
       Set oPort = oFacePortsColl.Item(iCount)
       
       Dim oOperator As Object
       Dim oOperation As IJStructOperation
       
       oStructDetailHelper.FindOperatorForOperationInGraphByID oPlatPart, oPort.OperationID, oPort.OperatorID, oOperation, oOperator
               
       If SplitterColl.Contains(oOperator) Then 'It is the port at the split
                
            'Donot include ports that are curved.
            If CheckIfPortIsLinear(oPort) Then
                Dim oTransientPort As IJTransientPort
                Set oTransientPort = oPort

                Dim oStructPort As IJStructPortEx
                oTransientPort.GetPersistentPort oStructPort, True
               
                GetLateralFacePortsFromSplitters.Add oStructPort.PrimaryPort
           
            End If
        
       End If
        
       Set oPort = Nothing
       Set oOperator = Nothing
       Set oOperation = Nothing

NextItem:
    Next
    
    
CleanUp:
    Set oStructConnectable = Nothing
    Set oFacePortsColl = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
' ***********************************************************************************
' Description: Gets the edge port from face port and checks if the port is curved or linear
' ***********************************************************************************
Private Function CheckIfPortIsLinear(oFacePort As IJStructPort) As Boolean
     Const METHOD = "CheckIfPortIsLinear"
     On Error GoTo ErrorHandler
    
     CheckIfPortIsLinear = False
     
     Dim oEntityHelper As New MfgEntityHelper
     Dim oEdgePort     As IJPort
     Set oEdgePort = oEntityHelper.GetEdgePortGivenFacePort(oFacePort, CTX_BASE) 'This gives the latest edge port geometry
    
     'Querying for IJLine on the port object was giving incorrect results for certain test cases(conical cross section)
     'Also, querying for IJLine on port geometry did not work.
     'So,using the MinBox routine to check if the port is linear.

     Dim oPortCurve As IJWireBody
     Set oPortCurve = oEdgePort.Geometry
           
     'get the MinBox to determine if the wirebody is a line or not
     Dim oMfgGeomHelper As MfgGeomHelper
     Set oMfgGeomHelper = New MfgGeomHelper
     Dim oMfgMGHelper As MfgMGHelper
     Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
     
     Dim oCurveElems As IJElements
     
     oMfgMGHelper.WireBodyToComplexStrings oPortCurve, oCurveElems
               
     Dim oBoxPoints As IJElements
     Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBox(oCurveElems)

     Dim length1 As Double, length2 As Double, length3 As Double
     Dim Points(1 To 4) As IJDPosition
     
     Set Points(1) = oBoxPoints.Item(1)
     Set Points(2) = oBoxPoints.Item(2)
     Set Points(3) = oBoxPoints.Item(3)
     Set Points(4) = oBoxPoints.Item(4)

     length1 = Points(1).DistPt(Points(2))
     length2 = Points(2).DistPt(Points(3))
     length3 = Points(1).DistPt(Points(4))
   
     'Two of the dimensions should be below the tolerance
     If length1 < 0.001 Then
        If length2 < 0.001 Or length3 < 0.001 Then
          CheckIfPortIsLinear = True
        End If
     ElseIf length2 < 0.01 Then
        If length1 < 0.01 Or length3 < 0.01 Then
          CheckIfPortIsLinear = True
        End If
     Else 'If length3 < 0.01 Then
        If length1 < 0.01 Or length2 < 0.01 Then
            CheckIfPortIsLinear = True
        End If
     End If
     
CleanUp:
     Set oEntityHelper = Nothing
     Set oMfgMGHelper = Nothing
     Set oMfgGeomHelper = Nothing
     Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
' ***********************************************************************************
' Public Function GetEligiblePortAndFlangeWidthCollection
'
' Description:  Get the eligible port and flange width collection for Flange bracket and knuckle plates
' ***********************************************************************************
Public Function GetEligiblePortAndFlangeWidthCollection(Part As Object, oPortColl As IJElements, oFlangeWidthColl As Collection)
    Const METHOD = "GetEligiblePortAndFlangeWidthCollection"

    On Error GoTo ErrorHandler
    
    Set oPortColl = New JObjectCollection
    Set oFlangeWidthColl = New Collection
        
    Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
    
    ' True below means "recursive" - Navigate past design splits
    ' to retrieve the root system
    
    Dim oRootPlateSystem    As IJPlateSystem
    oStructDetailHelper.IsPartDerivedFromSystem Part, oRootPlateSystem, True
    
    If oRootPlateSystem Is Nothing Then
        ' Part is a stand-alone plate part
        ' These cannot be flanged (yet)
        Exit Function
    End If
        
    Dim oFlangeAE       As IJPlateFlange_AE
    Dim oFlange         As Object
    
    ' Get the Flange AE
    Set oFlange = oRootPlateSystem.FlangeActiveEntity(Nothing)
    If Not oFlange Is Nothing Then

        Set oFlangeAE = oFlange
    
        Dim dFlangeWidth    As Double
        Dim oFlangeSymbol   As IJDSymbol
        
        ' Get the flange symbol
        Set oFlangeSymbol = oFlangeAE.FlangeSymbol
        
        If Not oFlangeSymbol Is Nothing Then
            
            Dim oSymHelper As IJMfgSymbolHelper
            Set oSymHelper = New MfgSymbolHelper
            
            ' Get the flange width
            dFlangeWidth = oSymHelper.GetSymbolOccParamValue(oFlangeSymbol, "FlangeWidth")
            
            Dim oStructPortCollection As IJElements
            ' Get the free lateral port i.e., port without any connections
            Set oStructPortCollection = GetFreeLateralFacePortColl(Part)
            
            oPortColl.AddElements oStructPortCollection
            oFlangeWidthColl.Add dFlangeWidth
                                            
            GoTo CleanUp
          End If
       
    Else
        
        Dim colChildren     As IJDTargetObjectCol
        Dim oKnuckleColl    As IJElements
                
        Dim oPartSupp       As IJPartSupport
        Dim oPlatePartSupp  As IJPlatePartSupport
        
        Set oPartSupp = New PlatePartSupport
        Set oPlatePartSupp = oPartSupp
        Set oPartSupp.Part = Part
            
        ' Get the collection of Knuckle Curves
        Dim oKnuckleCurveColl   As Collection
        Dim oSurfaceSideA       As Collection
        Dim oSurfaceSideB       As Collection
        
        oPlatePartSupp.GetKnuckleCurves PlateBaseSide, oKnuckleCurveColl, oSurfaceSideA, oSurfaceSideB
                    
        Dim oSurfaceBody As IJSurfaceBody
                
        Dim oStructConnectable  As IJStructConnectable
        Set oStructConnectable = Part
        
        Dim oFacePortsColl      As IJElements
        
        ' Get the all lateral face ports
        oStructConnectable.enumConnectableTransientPorts oFacePortsColl, vbNullString, False, PortFace, JS_TOPOLOGY_FILTER_SOLID_LATERAL_LFACES, False
        
        Dim oEligibleFacePortColl As IJElements
        Set oEligibleFacePortColl = New JObjectCollection
        
        Dim oPort As IJPort
        Dim iCount As Long
       
        For iCount = 1 To oKnuckleCurveColl.Count
        
            Dim oKnuckleWire    As IJDModelBody
            Set oKnuckleWire = oKnuckleCurveColl.Item(iCount)
            
            Dim dMinDist        As Double
            Dim oMinDistPort    As Object
            
            dMinDist = 9999#  'Initialize with large number
            
            Dim jCount  As Long
            For jCount = 1 To oFacePortsColl.Count
                
                Set oPort = Nothing
                Set oPort = oFacePortsColl.Item(jCount)
                
                Dim dTempDist   As Double
                Dim oPos1       As IJDPosition, oPos2   As IJDPosition
                ' Get the min distance from each lateral edge port to the knuckle
                
                oKnuckleWire.GetMinimumDistance oPort.Geometry, oPos1, oPos2, dTempDist
                
                If dTempDist > 0.001 Then
                    
                    If dTempDist < dMinDist Then
                        dMinDist = dTempDist
                        Set oMinDistPort = Nothing
                        Set oMinDistPort = oFacePortsColl.Item(jCount)
                    End If
                
                End If
            Next
            
            If Not oMinDistPort Is Nothing Then
            
                Dim oTransientPort As IJTransientPort
                Set oTransientPort = oMinDistPort
                
                Dim oStructPort As IJStructPortEx
                 oTransientPort.GetPersistentPort oStructPort, True
                
                oPortColl.Add oStructPort.PrimaryPort
                oFlangeWidthColl.Add dMinDist
                
            End If
            
        Next
        
    End If
    
CleanUp:
    
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

' ***********************************************************************************
' Public Function GetFreeLateralFacePort
'
' Description:  Get the port that doesn't have any connections
' ***********************************************************************************
Public Function GetFreeLateralFacePortColl(oPart As Object) As IJElements
    Const METHOD = "GetALateralFacePort"
    On Error GoTo ErrorHandler
    
    Dim oEligibleFaceColl   As IJElements
    Set oEligibleFaceColl = New JObjectCollection

    Dim oStructConnectable  As IJStructConnectable
    Dim colEdgePorts        As IJElements
    Dim colFacePorts        As IJElements
        
    Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
    
    Dim oLeafPlateSystem    As IJSystem
    oStructDetailHelper.IsPartDerivedFromSystem oPart, oLeafPlateSystem
    
    'Get all the Edge Face Ports of the given leaf system and Part
    Set oStructConnectable = oLeafPlateSystem
    
    oStructConnectable.enumConnectableTransientPorts colEdgePorts, vbNullString, False, PortEdge, JS_TOPOLOGY_FILTER_LCONNECT_SYS_LEDGES, False
    
    Set oStructConnectable = Nothing
    Set oStructConnectable = oPart
    oStructConnectable.enumConnectablePortsByOperationAndTopology colFacePorts, vbNullString, JS_TOPOLOGY_FILTER_SOLID_LATERAL_LFACES, False
 
    Dim iCount As Long
    For iCount = 1 To colEdgePorts.Count
        Dim oConnections    As IJElements
        Dim oPort           As IJPort
        Set oPort = colEdgePorts.Item(iCount)
        oPort.enumConnections oConnections
        If oConnections Is Nothing Then ' It is a free port
            
            Dim oEdgeStructPort     As IJStructPort
            Set oEdgeStructPort = oPort
            
            Dim EdgeType            As JS_TOPOLOGY_PROXY_TYPE
            Dim EdgeContext         As eUSER_CTX_FLAGS
            Dim EdgeOperationID     As Long
            Dim EdgeOperatorID      As Long
            Dim EdgeSectionID       As Long
            
            Dim bLightPart          As Boolean
            bLightPart = IsALightPart(oPart)
                        
            oEdgeStructPort.GetAttributes EdgeType, EdgeContext, EdgeOperationID, EdgeOperatorID, EdgeSectionID
            Dim jCount As Long
            For jCount = 1 To colFacePorts.Count
                Dim oStructFacePort     As IJStructPort
                Set oStructFacePort = colFacePorts.Item(jCount)
                
                Dim FaceType            As JS_TOPOLOGY_PROXY_TYPE
                Dim FaceContext         As eUSER_CTX_FLAGS
                Dim FaceOperationID     As Long
                Dim FaceOperatorID      As Long
                Dim FaceSectionID       As Long
                
                oStructFacePort.GetAttributes FaceType, FaceContext, FaceOperationID, FaceOperatorID, FaceSectionID
                
                Dim oStructPortEx       As IJStructPortEx
                Set oStructPortEx = oStructFacePort
                    
                'Check if Face and edge operator and operation IDs are same
                If FaceOperationID = EdgeOperationID And FaceOperatorID = EdgeOperatorID Then
                    oEligibleFaceColl.Add oStructPortEx.PrimaryPort
                Else
                    If bLightPart = True Then
                        ' This is true when Pre-nesting is enabled
                        ' If the part if light, port related to flange has FaceOperationID = 0 And FaceOperatorID = 0
                        If FaceOperationID = 0 And FaceOperatorID = 0 And FaceContext = CTX_LATERAL_LFACE Then
                            oEligibleFaceColl.Add oStructPortEx.PrimaryPort
                        End If
                    End If
                End If
            Next
            Exit For
        End If
    Next
    
    Set GetFreeLateralFacePortColl = oEligibleFaceColl
    
CleanUp:
    Set oStructConnectable = Nothing
    colEdgePorts.Clear
    Set colEdgePorts = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

'----------------------------------------------------------------'
' IsALightPart: Returns TRUE if the element is a LightPart.
'               Otherwise, returns FALSE.
'----------------------------------------------------------------'
Private Function IsALightPart(oPart As Object) As Boolean

    Dim oPartGeomState  As IJPartGeometryState
    Set oPartGeomState = oPart
    
    Dim partState       As PartGeometryStateType
    partState = oPartGeomState.PartGeometryState
    
    If partState = LightPart Then
        IsALightPart = True
    End If
    
End Function

Public Function GetAttribute(pObject As Object, strInterfaceName As String, strAttributeName As String) As IJDAttribute
Const METHOD = "GetAttribute"
On Error GoTo ErrorHandler

    Dim oAttrMetaData   As IJDAttributeMetaData
    Set oAttrMetaData = pObject
    
    Dim varOldAttribInt As Variant
    varOldAttribInt = oAttrMetaData.IID(strInterfaceName) '"IJUASMPlateTabType")
    'Set oAttrMetaData = Nothing
    
    Dim oAttributes     As IJDAttributes
    Set oAttributes = pObject
    
    Dim oAttributesCol  As IJDAttributesCol
    Set oAttributesCol = oAttributes.CollectionOfAttributes(varOldAttribInt)
    Set oAttributes = Nothing

    Dim i               As Integer
    Dim oAttribute      As IJDAttribute
    
    For i = 1 To oAttributesCol.Count
        Set oAttribute = oAttributesCol.Item(i)
        If oAttribute.AttributeInfo.Name = strAttributeName Then
            Set GetAttribute = oAttribute
            Exit For
        End If
    Next i
    
    Set oAttributesCol = Nothing
    Set oAttribute = Nothing

    Exit Function
ErrorHandler:
'    ReportUnanticipatedError Module, Method
End Function

Public Sub GetSmartOccRootClassGivenSCtype(ByVal lSCtype As Long, ByVal lSCSubType As Long, strRootClassName As String, oSORootClass As IJSmartClass, oTabSelHlpr As IJMfgTabSelectionHelper)
    Const METHOD As String = "GetSmartOccRootClassGivenSCtype"
    On Error GoTo ErrorHandler

    Dim oCatalogQuery As IJSRDQuery
    Set oCatalogQuery = New SRDQuery
    
    Dim oSOFeatureClassQuery As IJSmartQuery
    Set oSOFeatureClassQuery = oCatalogQuery
    
    Set oSORootClass = oSOFeatureClassQuery.GetClass(lSCtype, lSCSubType, strRootClassName)
    Set oTabSelHlpr = SP3DCreateObject(oSORootClass.SelectionRule)
    
CleanUp:
    Set oCatalogQuery = Nothing
    Set oSOFeatureClassQuery = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, Module, METHOD)
    GoTo CleanUp
End Sub


Public Function RemoveRelationWithMfgPlate(oTabobj As IJStrMfgTab) 'MarginTypes
    Const METHOD = "RemoveRelationWithMfgPlate"
    On Error GoTo ErrorHandler

    Dim oAssocRel As IJDAssocRelation
    Dim oUnkColl As Object
    Dim oRelationshipCol As IJDRelationshipCol
    Dim oRelationShip As IJDRelationship
    
    Const IID_IJStrMfgTab As Variant = "{6FE5BBE3-6DBE-431B-9D9A-5CCC1DD3B1AF}"
    Const IID_IJMfgDefinition As String = "{8E96B943-3F93-11D5-BFF5-00902770756B}"
    Set oAssocRel = oTabobj
    Set oUnkColl = oAssocRel.CollectionRelations(IID_IJMfgDefinition, "MfgDefInputFromRule_DEST")

    Set oRelationshipCol = oUnkColl
    
    If oRelationshipCol.Count > 0 Then
        Dim j As Integer
        For j = 1 To oRelationshipCol.Count
            Set oRelationShip = oRelationshipCol.Item(1)
            oRelationShip.Delete
        Next j
    End If
   
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD).Number
End Function

Public Function GetPlatePortGivenMfgContour(oContour2d As IJMfgGeom2d) As IJPort
    Const METHOD = "GetPlatePortGivenMfgContour"
    On Error GoTo ErrorHandler
    
    Set GetPlatePortGivenMfgContour = Nothing
    If oContour2d Is Nothing Then Exit Function

    Dim oPortMnkr As IMoniker
    Set oPortMnkr = oContour2d.GetMoniker
    If oPortMnkr Is Nothing Then Exit Function
    
    Dim oPOM As IJDObject
    Set oPOM = oContour2d
    If oPOM Is Nothing Then Exit Function
    
    Dim oUtil As New MfgMGHelper
    
    Dim oPort As IJPort
    oUtil.BindMoniker oPOM.ResourceManager, oPortMnkr, oPort
    
    Set GetPlatePortGivenMfgContour = oPort

CleanUp:
    Set oPort = Nothing
    Set oPortMnkr = Nothing
    Set oPOM = Nothing
    Set oUtil = Nothing
    
    Exit Function
ErrorHandler:
  Err.Raise LogError(Err, Module, METHOD).Number
  GoTo CleanUp
End Function

Public Function GetLengthOfGeom2dSegment(oGeom2d As IJMfgGeom2d) As Double
    Const METHOD = "GetLengthOfGeom2dSegment"
    On Error GoTo ErrorHandler

    Dim oCS As IJComplexString
    Set oCS = oGeom2d.GetGeometry
    
    Dim oCrv As IJCurve
    Set oCrv = oCS
    
    GetLengthOfGeom2dSegment = oCrv.Length
    
    Set oCS = Nothing
    Set oCrv = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD).Number
    Exit Function
End Function

Public Function IsFeatureRangeWithInMaxSize(oMfgObj As Object, oFeatureObj As Object, dMaxFeatureSize As Double) As Boolean
    Const METHOD = "IsFeatureRangeWithInMaxSize"
    On Error GoTo ErrorHandler
    
    Dim oMfgPartParent As IJMfgChild
    Set oMfgPartParent = oMfgObj
    
    Dim oPart   As Object
    Set oPart = oMfgPartParent.GetParent
    
    Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
    Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
    Set oSDPartSupport.Part = oPart
    
    Dim oContourColl As Collection
    Dim oPortMonikerColl As Collection
    
    ' Get the feature contour collection
    oSDPartSupport.GetFeatureInfo oFeatureObj, SideA, oContourColl, oPortMonikerColl
        
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper
    Dim oCurveElems As IJElements
    Set oCurveElems = New JObjectCollection
    Dim iCount      As Long
    
    If Not oContourColl Is Nothing Then
        For iCount = 1 To oContourColl.Count
            Dim oTempElems As IJElements
            oMfgMGHelper.WireBodyToComplexStrings oContourColl.Item(iCount), oTempElems
            oCurveElems.AddElements oTempElems
        Next
    Else
        Exit Function
    End If
    
    Dim oMfgGeomHelper As New MfgGeomHelper

    ' Get the Geom Min Box points for the feature contours
    Dim oBoxPoints As IJElements
    Set oBoxPoints = oMfgGeomHelper.GetGeometryMinBox(oCurveElems)
                            
    Dim oPoints(1 To 3) As IJDPosition
    
    Set oPoints(1) = oBoxPoints.Item(1)
    Set oPoints(2) = oBoxPoints.Item(2)
    Set oPoints(3) = oBoxPoints.Item(3)
    
    ' Get the length and width of the box
    Dim dLength(1 To 2) As Double
    dLength(1) = oPoints(1).DistPt(oPoints(2))
    dLength(2) = oPoints(2).DistPt(oPoints(3))
    
    ' If the length and width of the box is less than dMaxFeatureSize then mark the Feature
    If Abs(dLength(1)) < dMaxFeatureSize And Abs(dLength(2)) < dMaxFeatureSize Then
        IsFeatureRangeWithInMaxSize = True
    End If
    
    Set oSDPartSupport = Nothing
    Set oBoxPoints = Nothing
    Set oMfgMGHelper = Nothing
    Set oCurveElems = Nothing
    Set oPortMonikerColl = Nothing
    Set oContourColl = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD).Number
    Exit Function
End Function

Public Function IsTabControlledWithMfgPlate(oTabobj As IJStrMfgTab) As Boolean 'MarginTypes
    Const METHOD = "IsTabControlledWithMfgPlate"
    On Error GoTo ErrorHandler

    Dim oAssocRel As IJDAssocRelation
    Dim oUnkColl As Object
    Dim oRelationshipCol As IJDRelationshipCol
    Dim oRelationShip As IJDRelationship
    
    Const IID_IJStrMfgTab As Variant = "{6FE5BBE3-6DBE-431B-9D9A-5CCC1DD3B1AF}"
    Const IID_IJMfgDefinition As String = "{8E96B943-3F93-11D5-BFF5-00902770756B}"
    Set oAssocRel = oTabobj
    Set oUnkColl = oAssocRel.CollectionRelations(IID_IJMfgDefinition, "MfgDefInputFromRule_DEST")

    Set oRelationshipCol = oUnkColl
    
    If oRelationshipCol.Count > 0 Then 'MFGPlateProcess Controlled Tab
        IsTabControlledWithMfgPlate = True
    End If
   
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD).Number
End Function
