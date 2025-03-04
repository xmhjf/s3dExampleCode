VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenAssyConnSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'*********************************************************************************************
'  Copyright (C) 2015, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : GenAssyConnSel.cls
'
'  Description :
'
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'   16/May/2012 -hgunturu
'           TR-CP-212435 Addeda new Smart Item in the selection and also added a new method bIsWebCompletelyOnOneBounding()
'   10/Oct/2012 -skcheeka
'           TR-CP-218823 Added few more necessary checks to select Gen_2B_01 SmartItem in IsWebThicknessOverlappingPlateEdge() method
'   25/Oct/2012 -skcheeka
'           TR-CP-221857 Changes are made in selector logic so as to change the default AC Item to Generic_MbrAC_01
'   22/Nov/2012 -knukala
'           CR-CP-222548 Added new smart item Generic_LongBoxCuts and made this AC item as default for End-End cases
'   24/Aug/2014 -skcheeka
'           CR-CP-240787  Provide Std AC with bearing plate on flat surface
'   01/June/2012 -knukala
'           TR-CP-274015  Generic assembly connection fails when Member top is above ER top
'   14/July/2015 -knukala
'           TR-CP-270890  Cannot Place Generic AC when Handrail Post member is bounded to a plate
'    15/June/2016    knukala   TR-CP-295640: Generic AC is updated when changing the corner feature type.
'*********************************************************************************************

Const m_sClassName As String = "GenAssyConnSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sProjectName + "." + m_sClassName
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices
'

'*********************************************************************************************
' Method      : SelectorInputs
' Description : List any graphic Inputs that the Selector has here
'
'*********************************************************************************************
Public Sub SelectorInputs(pIH As IJDInputsHelper)

    Const METHOD = m_SelectorProgid & "::SelectorInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"
    
    Exit Sub
ErrorHandler:
    pIH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorQuestions
' Description : List any Questions/Answers that the Selector uses here
'
'*********************************************************************************************
Public Sub SelectorQuestions(pQH As IJDQuestionsHelper)

    Const METHOD = m_SelectorProgid & "::SelectorQuestions"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' TODO - Add your question below
    '===============================
    sMsg = "Defining/Initializing Selector Questions/Answers"
    
    pQH.SetQuestion "SelectByRule", "Yes", "BooleanCol"
    
    Exit Sub
ErrorHandler:
    pQH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the Basic Type of Member Assembly Connection here
'
'*********************************************************************************************
Public Sub SelectorLogic(pSelectorLogic As IJDSelectorLogic)

    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    ' ---------------------------------
    ' Get the assembly connection ports
    ' ---------------------------------
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    
    Set oAppConnection = pSelectorLogic.SmartOccurrence
    
    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If lStatus <> 0 Then
        Exit Sub
    End If
    
    ' -----------------------------------------------------------------------
    ' For generic ACs, both member parts from the connection must be the same
    ' -----------------------------------------------------------------------
    If Not oBoundedData.MemberPart Is oBoundingData.MemberPart Then
        Exit Sub
    End If
    
    ' -------------------------------------------------------------------------------
    ' Verify that the reference collection contains valid supporting/bounding objects
    ' -------------------------------------------------------------------------------
    Dim oReferencesCollection As IJDReferencesCollection
    Dim oEditJDArgument As IJDEditJDArgument
    
    Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
    Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
    
    Dim nRefArguments As Long
    nRefArguments = oEditJDArgument.GetCount
    If nRefArguments < 1 Then
        Exit Sub
    End If
    
    Dim iIndex As Long
    Dim oArgObject As Object
    
    For iIndex = 1 To nRefArguments
        Set oArgObject = oEditJDArgument.GetEntityByIndex(iIndex)
        If TypeOf oArgObject Is IJPort Then
        ElseIf TypeOf oArgObject Is IJPlane Then
        ElseIf TypeOf oArgObject Is IJPoint Then
        ElseIf TypeOf oArgObject Is IJSurfaceBody Then
        Else
        End If
    Next iIndex
    
    Dim bIsTubeBounded As Boolean
    Dim bIsTubeBounding As Boolean
    Dim bIsTubeByTube As Boolean
    
    bIsTubeBounded = False: bIsTubeBounding = False: bIsTubeByTube = False
    
    bIsTubeBounded = IsTubularMember(oBoundedData.MemberPart)
    
    ' -------------------------------------------------------------------------------------
    ' Find all the unique objects from among the ports included in the reference collection
    ' -------------------------------------------------------------------------------------
    Dim oBoundingObjectColl As IJElements
    Set oBoundingObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
    
    ' --------------------------------------------------------------------
    ' Collect information on the relationship between bounding and bounded
    ' --------------------------------------------------------------------
    Dim i As Long
    
    Dim oBoundingObject As Object
    Dim oPortOnBounding As IJStructPort
    
    Dim oPlateObjects As New Collection
    Dim oIsPlateEdge As New Collection
    
    Dim oMemberObjects As New Collection
    Dim oMemberPenetratesWeb As New Collection
    Dim oMemberAlias As New Collection
    
    Dim sectionAlias As Long
    Dim penetratesWeb As Boolean
    Dim oEdgeMap As JCmnShp_CollectionAlias
    
    Dim bEndToEnd As Boolean, oEndToEndPos As IJDPosition, oTan1 As IJDVector, oTan2 As IJDVector
    
    Dim oTopORWL As ConnectedEdgeInfo
    Dim oBottomOrWR As ConnectedEdgeInfo
    Dim oInsideTopFlgOrFL As ConnectedEdgeInfo
    Dim oInsideBtmFlgOrFR As ConnectedEdgeInfo
    
    Dim bIsBuiltup As Boolean, oBUMember As ISPSDesignedMember
    
    For i = 1 To oBoundingObjectColl.Count
        
        Set oBoundingObject = oBoundingObjectColl.Item(i)
        
        bIsTubeBounding = IsTubularMember(oBoundingObject)
        
        If bIsTubeBounded And bIsTubeBounding Then
            bIsTubeByTube = True
            Exit For
        End If
                
        If TypeOf oBoundingObject Is IJConnectable Then
            
            Set oPortOnBounding = GetPortsFromBoundingObject(oBoundingObject, oEditJDArgument).Item(1)
            
            If TypeOf oBoundingObject Is IJPlate Then
                oPlateObjects.Add oBoundingObject
            
                ' presuming here that only one port on the plate has been selected
                If (oPortOnBounding.ContextID And CTX_LATERAL) Then
                    oIsPlateEdge.Add True
                Else
                    oIsPlateEdge.Add False
                End If
            
            ElseIf (TypeOf oBoundingObject Is ISPSMemberPartPrismatic) Or _
                (TypeOf oBoundingObject Is IJProfile) Then 'Treat a profile as a member and add it to the oMemberObjects collection

                If (oBoundingObjectColl.Count = 1) Then
                    GetConnectedEdgeInfo oAppConnection, oBoundedData.AxisPort, oPortOnBounding, oTopORWL, oBottomOrWR, oInsideTopFlgOrFL, oInsideBtmFlgOrFR, , , , , , True
                End If
                Set oEdgeMap = New Collection
                Set oEdgeMap = GetEdgeMap(oAppConnection, oPortOnBounding, oBoundedData.AxisPort, sectionAlias, penetratesWeb)
                oMemberObjects.Add oBoundingObject
                oMemberAlias.Add sectionAlias
                oMemberPenetratesWeb.Add penetratesWeb
                
            End If
        End If
    Next i
    
    
    If bIsTubeByTube Then
        Set oBoundingObject = oBoundingObjectColl.Item(1)
        If TypeOf oBoundingObject Is ISPSMemberPartCommon Then
        
        Else
            IsFromBuiltUpMember oBoundingObject, bIsBuiltup, oBUMember
            If bIsBuiltup Then
                Set oBoundingObject = oBUMember
            End If
        End If
        
        If TypeOf oBoundingObject Is ISPSMemberPartCommon Then
            AreMembersEndConnected oBoundedData.MemberPart, oBoundingObject, bEndToEnd, oEndToEndPos, oTan1, oTan2
        End If
        
        If bEndToEnd Then
            pSelectorLogic.Add "Generic_LongBox"
            pSelectorLogic.Add "Generic_MbrAC_01"
        Else
            pSelectorLogic.Add "Generic_MbrAC_01"
        End If
        Exit Sub
    End If
        
    ' -------------------------------------------------------------------
    ' Special case for:
    ' - bounded by Plate face(or Member) and Plate Edge
    ' - bounding Plate has edge which intersects Web
    ' -------------------------------------------------------------------
    
    Dim bIsWebThicknessOverlappingPlateEdge As Boolean
    'When Bounded Web is Intersecting with Plate Edge along the thickness then only we need only one WebCut.
    If Not bIsTubeBounded And oBoundingObjectColl.Count > 1 Then

        'Check Web is hitting to one Bounding and Flange is hitting to other Bounding
        bIsWebThicknessOverlappingPlateEdge = IsWebThicknessOverlappingPlateEdge(oBoundingObjectColl, oEditJDArgument, oAppConnection, oBoundedData)
        If bIsWebThicknessOverlappingPlateEdge Then
            pSelectorLogic.Add "Generic_2B_01"
        End If
    End If
    

    ' -------------------------------------------------------------------
    ' Special case for:
    ' - bounded by plate face and member
    ' - web penetrated
    ' - bounding member has single edge
    ' - position is to edge and outside
    ' -------------------------------------------------------------------
    Dim alias As Long
    Dim topFlangeBoundingEdge As eBounding_Edge
    Dim btmFlangeBoundingEdge As eBounding_Edge
    Dim mappedEdge As JXSEC_CODE
    
    ' One plate and one member
    If oPlateObjects.Count = 1 And oMemberObjects.Count = 1 Then
        ' Plate face and pentrates web
        If (oIsPlateEdge.Item(1) = False) And (oMemberPenetratesWeb.Item(1) = True) Then
            
            alias = oMemberAlias.Item(1)
            
            Set oPortOnBounding = GetPortsFromBoundingObject(oMemberObjects.Item(1), oEditJDArgument).Item(1)
            
            GetNonPenetratedIntersectedEdge oAppConnection, oPortOnBounding, oBoundedData.AxisPort, btmFlangeBoundingEdge, mappedEdge, True
            GetNonPenetratedIntersectedEdge oAppConnection, oPortOnBounding, oBoundedData.AxisPort, topFlangeBoundingEdge, mappedEdge, False
            
            ' Bounded bottom to bounding bottom, and top above (edge and outside case 1)
            If (alias = 3) And (btmFlangeBoundingEdge = Bottom_Flange_Right) And (topFlangeBoundingEdge = Above) Then
                pSelectorLogic.Add "WPEF_PF_1" ' Web-penetrated member edge and face, plate face, configuration #1
            
            ' Bounded top to bounding top, and bottom below (edge and outside case 2)
            ElseIf (alias = 1) And (topFlangeBoundingEdge = Top_Flange_Right) And (btmFlangeBoundingEdge = Below) Then
                pSelectorLogic.Add "WPEF_PF_1" ' Web-penetrated member edge and face, plate face, configuration #1
            End If
        End If
    End If
    
    ' ----------------------------------
    ' Creates longbox cuts and short box cuts
    ' ----------------------------------
    If oBoundingObjectColl.Count >= 1 Then
        For Each oBoundingObject In oBoundingObjectColl
            If TypeOf oBoundingObject Is ISPSMemberPartCommon Then
                On Error Resume Next
                AreMembersEndConnected oBoundedData.MemberPart, oBoundingObject, bEndToEnd, oEndToEndPos, oTan1, oTan2
                On Error GoTo ErrorHandler
                If bEndToEnd Then
                    Exit For
                End If
            End If
        Next
        If bEndToEnd Then
            If oBoundingObjectColl.Count > 1 Then
                Dim bISBoundedMemberUnSupported As Boolean
                Dim oMemberConnectionServices As SPSMembers.ISPSMemberConnectionServices
                Dim eBoundedFCType As SPSFCBoundingType
                eBoundedFCType = -1 'Inittialize to non-unsupported
                If penetratesWeb Then
                    Dim oMem As ISPSMemberSystem
                    Dim oMemPart As ISPSMemberPartCommon
                    Set oMemPart = oBoundedData.MemberPart
                    Set oMem = oMemPart.MemberSystem
                    Dim oFC As ISPSFrameConnection
                    Set oFC = oMem.FrameConnectionAtEnd(oBoundedData.ePortId)
                    Set oMemberConnectionServices = oFC.Services
                    'A bounded member system can be split by other bounding member system
                    'In that case there is a split connection instead of frame connection at the intersection(oEndToEndPos)
                    'Need to get the frame connection at the the itersection and check its FC type.
                    Dim oPoint As IJPoint
                    Dim oEndToEndPoint As IJPoint
                    Dim dDistFC2EndToEndPos As Double
                    Set oEndToEndPoint = New Point3d
                    oEndToEndPoint.SetPoint oEndToEndPos.x, oEndToEndPos.y, oEndToEndPos.z
                    Set oPoint = oFC
                    dDistFC2EndToEndPos = oPoint.DistFromPt(oEndToEndPoint)
                    If dDistFC2EndToEndPos < 0.001 Then
                        oMemberConnectionServices.GetFCBoundingType oFC, eBoundedFCType
                    End If
                End If
                If eBoundedFCType = Unsupported Then
                    pSelectorLogic.Add "Generic_LongBoxCuts"
                    pSelectorLogic.Add "Generic_MbrAC_01"
                Else
                    pSelectorLogic.Add "Generic_MbrAC_01"
                    pSelectorLogic.Add "Generic_LongBoxCuts"
                End If
            ElseIf oBoundingObjectColl.Count = 1 And TypeOf oBoundingObjectColl.Item(1) Is ISPSMemberPartPrismatic Then
                pSelectorLogic.Add "Generic_LongBox"
                pSelectorLogic.Add "Generic_MbrAC_01"
            Else
                pSelectorLogic.Add "Generic_MbrAC_01"
            End If
        Else
            If oBoundingObjectColl.Count = 1 Then
                Dim eBoundedType As Integer
                Dim oBounded As ISPSMemberPartPrismatic
                Set oBoundingObject = oBoundingObjectColl.Item(1)
                Set oBounded = oBoundedData.MemberPart
                eBoundedType = oBounded.MemberType.TypeCategory
                
                If (eBoundedType = MemberCategoryAndType.HandRailElement Or eBoundedType = MemberCategoryAndType.LadderElement Or _
                   eBoundedType = MemberCategoryAndType.StairElement) And (TypeOf oBoundingObject Is IJPlate) Then
                        ' Igonore for traffic items This need to be handled by Std AC rules(through Place Assembly connection command)
                        Exit Sub
                 Else
                    pSelectorLogic.Add "Generic_MbrAC_01"
                    pSelectorLogic.Add "Gen_BearingPlateAC_01"
                End If
            Else
                pSelectorLogic.Add "Generic_MbrAC_01"
            End If
        End If
    Else
        pSelectorLogic.Add "Generic_MbrAC_01"
    End If
 
    ' --------------------------------------------------------------------------------
    ' Special case for any two boundaries.  Places a corner feature where the two meet
    ' --------------------------------------------------------------------------------
    
    If (oPlateObjects.Count + oMemberObjects.Count) = 2 Then
        'Only Add for Web Penetrated Case
        If oMemberObjects.Count = 1 And oPlateObjects.Count = 1 Then
            If oMemberPenetratesWeb.Item(1) = True Then
                pSelectorLogic.Add "WP_2B_CF"
            End If
        ElseIf oMemberObjects.Count = 2 Then
            If oMemberPenetratesWeb.Item(1) = True Or oMemberPenetratesWeb.Item(2) = True Then
                pSelectorLogic.Add "WP_2B_CF"
            End If
        Else
            pSelectorLogic.Add "WP_2B_CF"
        End If
    End If
    
    Dim bMBCase As Boolean
    
    'Check for Multi Bounding SingleCut Case
    If oBoundingObjectColl.Count > 1 Then

        bMBCase = IsMultiBoundingCase(oAppConnection)

        If bMBCase Then
            pSelectorLogic.Add "Gen_MultiBounding_01"
        End If
    End If
    Exit Sub
ErrorHandler:
    pSelectorLogic.ReportError sMsg, METHOD
    
End Sub
  
' ** Start CM **
'*************************************************************************************
'IsWebThicknessOverlappingPlateEdge()
'           This method determines Wheather the Web is hitting to One Bounding Port(Plate Edge) and
'           Flange is hitting to other Bounding
'Inputs : oBoundingObjColl, oEditJDArgument ----> Bounding objects
'         oAppConnection --->Connection between the bounded and bounding
'         oBoundedData ---> Bounded data
'
'Outputs : IsWebThicknessOverlappingPlateEdge ---> Returns True or False
'
'**************************************************************************************

Private Function IsWebThicknessOverlappingPlateEdge(ByVal oBoundingObjectColl As IJElements, ByVal oEditJDArgument As IJDEditJDArgument, _
                                ByVal oAppConnection As Object, oBoundedData As MemberConnectionData) As Boolean

    Const METHOD = m_SelectorProgid & "::IsWebThicknessOverlappingPlateEdge"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    IsWebThicknessOverlappingPlateEdge = False
    Dim oIntersect As IJDTopologyIntersect
    Dim oCommonGeom As Object
    Dim oPointOnCommon As IJDPosition
    Dim oDirVector As IJDVector
    Dim oWireColl As New Collection
    Dim oPort As IJPort
    Dim oPort2 As IJPort
    Dim oBoundedObj As Object
    Dim oNearestBoundingtoPort1 As Object
    Dim oNearestBoundingtoPort2 As Object
    Dim bIsTopIntersecting As Boolean
    Dim bIsBtmIntersecting As Boolean
    Dim bHasIintersection As Boolean
    Dim oSDO_Memberpart As StructDetailObjects.MemberPart
    Dim iCount As Integer
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    
    Set oIntersect = New DGeomOpsIntersect
    Set oBoundedObj = oBoundedData.MemberPart
    
    Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_TOP)
    Set oPort2 = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_BOTTOM)
    Set oNearestBoundingtoPort1 = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
    
    Dim oStructPort As IJStructPort
    
    If oNearestBoundingtoPort1 Is Nothing Then
        Exit Function
    End If
      
    If TypeOf GetPortsFromBoundingObject(oNearestBoundingtoPort1, oEditJDArgument).Item(1) Is IJStructPort Then
        Set oStructPort = GetPortsFromBoundingObject(oNearestBoundingtoPort1, oEditJDArgument).Item(1)
    Else
        Exit Function
    End If
    
    If Not (TypeOf oNearestBoundingtoPort1 Is IJPlatePart) Or Not (oStructPort.ContextID = CTX_LATERAL_LFACE) Then
        Exit Function
    End If
    
    Set oNearestBoundingtoPort2 = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort2)
    
    'Check both top and bottom intersects the same Bounding
    If Not oNearestBoundingtoPort1 Is oNearestBoundingtoPort2 Then
        Exit Function
    End If

    
    'Check Flanges
    CrossSection_Flanges oBoundedData.MemberPart, bTFL, bBFL, bTFR, bBFR

    Set oPort = Nothing
    Set oPort2 = Nothing
    
    'Check for Top Flanges
    If bTFL And bTFR Then 'Both the Top Flanges
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_TOP_FLANGE_LEFT)
        Set oPort2 = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_TOP_FLANGE_RIGHT)
        Set oNearestBoundingtoPort1 = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
        Set oNearestBoundingtoPort2 = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort2)
        
        'Check Complete Flange intersects one Bounding
        If Not oNearestBoundingtoPort1 Is oNearestBoundingtoPort2 Then
            Exit Function
        End If
        
    ElseIf bTFL Or bTFR Then 'one of the Top Flange
        If bTFL Then
            Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_TOP_FLANGE_LEFT)
            Set oPort2 = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_WEB_LEFT)
        Else
            Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_TOP_FLANGE_RIGHT)
            Set oPort2 = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_WEB_RIGHT)
        End If
        'Check intersection
        Set oNearestBoundingtoPort1 = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
        On Error Resume Next
        oIntersect.PlaceIntersectionObject Nothing, GetPortsFromBoundingObject(oNearestBoundingtoPort1, oEditJDArgument).Item(1), oPort2, Nothing, oCommonGeom
        On Error GoTo ErrorHandler
        'If not Intersecting Exit
        If oCommonGeom Is Nothing Then
            Exit Function
        End If
        
        Set oCommonGeom = Nothing
        Set oPort = Nothing
        Set oPort2 = Nothing
    End If
    
    'Check for Bottom Flanges
    If bBFL And bBFR Then 'Both the Top Flanges
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_BOTTOM_FLANGE_LEFT)
        Set oPort2 = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_BOTTOM_FLANGE_RIGHT)
        
        Set oNearestBoundingtoPort1 = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
        Set oNearestBoundingtoPort2 = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort2)
        
        'Check Complete Flange intersects one Bounding
        If Not oNearestBoundingtoPort1 Is oNearestBoundingtoPort2 Then
            Exit Function
        End If
    ElseIf bBFL Or bBFR Then 'one of the Top Flange
        If bBFL Then
            Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_BOTTOM_FLANGE_LEFT)
            Set oPort2 = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_WEB_LEFT)
        Else
            Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_BOTTOM_FLANGE_RIGHT)
            Set oPort2 = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_WEB_RIGHT)
        End If
        'Check intersection
        Set oNearestBoundingtoPort1 = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
        On Error Resume Next
        oIntersect.PlaceIntersectionObject Nothing, GetPortsFromBoundingObject(oNearestBoundingtoPort1, oEditJDArgument).Item(1), oPort2, Nothing, oCommonGeom
        On Error GoTo ErrorHandler
        'If not Intersecting Exit
        If oCommonGeom Is Nothing Then
            Exit Function
        End If
        Set oCommonGeom = Nothing
        Set oPort = Nothing
        Set oPort2 = Nothing
    End If
        
    'Get intersection wires
    'Top left
    If bTFL Then
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_TOP_FLANGE_LEFT_BOTTOM)
    Else
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_TOP)
    End If
    Set oPort2 = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_WEB_LEFT)
    
    On Error Resume Next
    oIntersect.PlaceIntersectionObject Nothing, oPort, oPort2, Nothing, oCommonGeom
    If Not oCommonGeom Is Nothing Then oWireColl.Add oCommonGeom
    Set oCommonGeom = Nothing
    
    'Bottom left
    If bBFL Then
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_BOTTOM_FLANGE_LEFT_TOP)
    Else
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_BOTTOM)
    End If
    
    On Error Resume Next
    oIntersect.PlaceIntersectionObject Nothing, oPort, oPort2, Nothing, oCommonGeom
    If Not oCommonGeom Is Nothing Then oWireColl.Add oCommonGeom
    Set oCommonGeom = Nothing
    
    'Top Right
    If bTFR Then
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
    Else
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_TOP)
    End If
    Set oPort2 = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_WEB_RIGHT)
    
    On Error Resume Next
    oIntersect.PlaceIntersectionObject Nothing, oPort, oPort2, Nothing, oCommonGeom
    If Not oCommonGeom Is Nothing Then oWireColl.Add oCommonGeom
    Set oCommonGeom = Nothing
        
    'Bottom Right
    If bBFR Then
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_BOTTOM_FLANGE_RIGHT_TOP)
    Else
        Set oPort = GetLateralSubPortBeforeTrim(oBoundedObj, JXSEC_BOTTOM)
    End If
    
    On Error Resume Next
    oIntersect.PlaceIntersectionObject Nothing, oPort, oPort2, Nothing, oCommonGeom
    If Not oCommonGeom Is Nothing Then oWireColl.Add oCommonGeom
    Set oCommonGeom = Nothing
    
    If oWireColl.Count < 4 Then
        Exit Function
    End If
    
    Dim dExtend As Double
    Dim oExtenededSurface As IJSurfaceBody
    Dim oTopologyLocate As TopologyLocate
    Dim oPositionColl As New Collection
    Dim oGeomOffset As IJGeometryOffset
    Set oGeomOffset = New DGeomOpsOffset
    
    Set oSDO_Memberpart = New StructDetailObjects.MemberPart
    Set oSDO_Memberpart.object = oBoundedData.MemberPart
    dExtend = oSDO_Memberpart.webThickness + 0.001

    Set oPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)
    Set oNearestBoundingtoPort1 = GetNearestBoundingToPort(oBoundingObjectColl, oEditJDArgument, oAppConnection, oPort)
    'Get Port on Nearest Bounding
    Set oPort = GetPortsFromBoundingObject(oNearestBoundingtoPort1, oEditJDArgument).Item(1)
    
    'Get Extended Surface
    oGeomOffset.CreateExtendedSheetBody Nothing, oPort.Geometry, Nothing, dExtend, Nothing, oExtenededSurface
    
    If oExtenededSurface Is Nothing Then
        Exit Function
    End If
    
    'Get Points on the Extended surface
    Set oTopologyLocate = New TopologyLocate
    For iCount = 1 To 4
        Dim oObject As Object
        
        Set oObject = oTopologyLocate.FindIntersectionPoint(oExtenededSurface, oWireColl.Item(iCount))
        
        If Not oObject Is Nothing Then
            oPositionColl.Add oObject
        End If
        Set oObject = Nothing
    Next
    
    'Collection of points obtained after intersecting with Exteneded Surface
    If oPositionColl.Count < 4 Then
        Exit Function
    End If
 
    Dim oLine As IJLine
    Dim oStPoint As IJDPosition
    Dim oEndPoint As IJDPosition
    Dim oGeomFactory As IJGeometryFactory
    
    Set oStPoint = oPositionColl.Item(1)
    Set oEndPoint = oPositionColl.Item(3)
    Set oGeomFactory = New GeometryFactory
    
    'Create Line Geometry between Top points
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                      oEndPoint.x, oEndPoint.y, oEndPoint.z)
    'Check Overlapping Geometry with Top
    oIntersect.HasOverlappingGeometry oLine, oPort, bIsTopIntersecting
    
    If Not bIsTopIntersecting Then
        Exit Function
    End If
    
    Set oStPoint = oPositionColl.Item(2)
    Set oEndPoint = oPositionColl.Item(4)
    Set oGeomFactory = New GeometryFactory
    
    'Create Line Geometry with Bottom points
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oStPoint.x, oStPoint.y, oStPoint.z, _
                                                      oEndPoint.x, oEndPoint.y, oEndPoint.z)
    'Check Overlapping Geometry with Bottom
    oIntersect.HasOverlappingGeometry oLine, oPort, bIsBtmIntersecting
    
    If bIsBtmIntersecting Then
        IsWebThicknessOverlappingPlateEdge = True
    End If

 Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Function

' *******************************************************************************************
' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String

    IJDUserSymbolServices_GetDefinitionName = m_SelectorProgid
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)

    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pSelector.IJDInputs.RemoveAllInput
    pSelector.IJDRepresentations.RemoveAllRepresentation
    
    Dim pDFact As New DefinitionFactory
    pDFact.InitAbstractSelector pSelector
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pSelector
    pIH.InitAs m_FamilyProgid
    SelectorInputs pIH
    
    Dim pQH As IJDQuestionsHelper
    Set pQH = New QuestionHelper
    pQH.Selector = pSelector
    SelectorQuestions pQH
    
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CB As String, ByVal DP As Variant, ByVal pRM As Object) As Object

    Dim pDFact As New DefinitionFactory
    Set IJDUserSymbolServices_InstanciateDefinition = pDFact.InstanciateSelector(m_SelectorProgid, CB, IJDUserSymbolServices_GetDefinitionName(DP), pRM)
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean

End Function

'*********************************************************************************************
' Method      : CMSelector
' Description :
'
'*********************************************************************************************
Public Sub CMSelector(pRep As IJDRepresentation)

    Dim pSL As IJDSelectorLogic
    Set pSL = New SelectorLogic
    pSL.Representation = pRep
    SelectorLogic pSL
    
End Sub
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************


