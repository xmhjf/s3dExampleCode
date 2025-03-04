VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrStdACSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*********************************************************************************************
'  Copyright (C) 2014-16, Intergraph Corporation.  All rights reserved.
'
'  File        : MbrStdACSel.cls
'
'  Description :
'
'  Author      : Alligators
'
'    03/Nov/2014 - MDT/GH
'         CR-CP-250198  Lapped AC for traffic items
'    03/Dec/14  CSK/CM CR-250022 Connections and free end cuts at ends of ladder rails
'   27/Feb/2015 MDT/hgunturu
'         CR-265236 Modified Selection logic.Considered OS and OS case when flip happened for box connections.
'   19/Feb/15   NK   CR-CP-267170  Mitered Stadnard AC to be considered even for Mbr End to End Offset case
' 24/Mar/15  MDT/RPK  TR-269306 Handrail placed on builtup member, cannot create assembly connection, errors,
'                    Added some checks related to design/Builtup member as the bounding member
' 12/May/15     GH   CR-260982 Created New method SelectEdgeAndEdgeACSmartItems() and updated the selection logic
' 11/Jun/15     svsmylav   CR-271082 Create new Std. AC for 'ToEdgecase'.
' 09/Sept/15     Knukala   CR-CP-226692  Rule-Based selection of standard member ACs.
' 27/11/2015    hgunturu    DI-282754: Updated SelectorLogic() Method to improve performance of Standard ACs
' 10/03/2016    hgunturu    TR-286505: Avoided selection of Slot ACs when penetrating though Flange portion
' 5/April/2016   prajavar    TR-290237: Avoided selection of StandardACs which are are not applicable when Flange penetrating cases.
' 5/April/2016   dsmamidi    TR-291344: Modified the selection logic for face and edge config to select the applicable ACs based on edge clearance values.
'*********************************************************************************************

Const m_sClassName As String = "MbrStdACSel"
Const m_FamilyProgid As String = ""
Const m_SelectorProgid As String = m_sStdACProjectName + "." + m_sClassName
Const m_SelectorName As String = m_SelectorProgid
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

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
    sMsg = "Defining/Initializing Selector Questions/Answers"

    ' Only applies to symmetric cases (see "Legend" tab on bulkload spreadsheet), but we don't want
    ' separate selectors, so user must ignore this for now.  There is a CR to allow the question to
    ' be hidden dynamically
    pQH.SetQuestion "Flip", gsNo, "BooleanCol", "CMSetFlipAnswer", m_SelectorProgid
    pQH.SetQuestion "Flip Primary and Secondary", gsNo, "BooleanCol", "CMSetFlipPrimaryAndSecondaryAnswer", m_SelectorProgid

    Exit Sub
    
ErrorHandler:
    pQH.ReportError sMsg, METHOD
End Sub

'*********************************************************************************************
' Method      : SelectorLogic
' Description : Select the Basic Type of Member EndCut
'
'*********************************************************************************************
Public Sub SelectorLogic(oSL As IJDSelectorLogic)
    
    Const METHOD = m_SelectorProgid & "::SelectorLogic"
    
    Dim sMsg As String
    
    On Error GoTo ErrorHandler
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim eBoundedType As Integer
    Dim eBndgTypeCategory As Integer
    Dim eBoundingType As Integer
    Dim oBounding As ISPSMemberPartPrismatic
    Dim oBounded As ISPSMemberPartPrismatic
    Dim sSectionType As String
    Dim oMember As New StructDetailObjects.MemberPart
    Dim MapCollection As New Collection
    
    Dim oAppConnection As IJAppConnection
    Set oAppConnection = oSL.SmartOccurrence
    
    Set_AttributeValue oAppConnection, "IJUAFlipAttribute", "FlipPrimaryandSecondary", False
    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort, False
    Dim oReferencesCollection As IJDReferencesCollection
    Dim oBdngObjectColl As IJElements
    Dim oEditJDArgument As IJDEditJDArgument
    'If Assembly connection is generic and if bounding objects are more than one, then no standard assembly connection is supported, so not selected.
    
    Dim eGetACType As eACType
    eGetACType = GetMbrAssemblyConnectionType(oAppConnection)
    
    If (eGetACType = ACType_Mbr_Generic) Or (eGetACType = ACType_Stiff_Generic) Then
        Set oReferencesCollection = GetRefCollFromSmartOccurrence(oAppConnection)
        Set oEditJDArgument = oReferencesCollection.IJDEditJDArgument
        Set oBdngObjectColl = GetBoundingObjectsFromPorts(oEditJDArgument)
        If oBdngObjectColl.Count > 1 Then
            Exit Sub
        End If
    End If
    
    Dim oBoundedAxisPort As ISPSSplitAxisPort
    
        
    Dim oBoundingAxisPort As ISPSSplitAxisPort
    
    If TypeOf oBoundingPort Is ISPSSplitAxisPort And Not TypeOf oBoundingPort.Connectable Is ISPSDesignedMember Then
        Set oBoundingAxisPort = oBoundingPort
        
        Set oBounding = oBoundingAxisPort.Part
        eBoundingType = oBounding.MemberType.Type
        eBndgTypeCategory = oBounding.MemberType.TypeCategory
        
    End If
    
    Set oBounded = oBoundedPort.Connectable
    Set oMember.object = oBounded
    sSectionType = oMember.sectionType
    eBoundedType = oBounded.MemberType.Type

    Dim StartPosition As IJDPosition
    Dim EndPosition As IJDPosition
    Dim x1 As Double
    Dim y1 As Double
    Dim z1 As Double
    Dim x2 As Double
    Dim y2 As Double
    Dim z2 As Double
    
    Dim eCase As eMemberBoundingCase
    Dim bEndToEnd As Boolean
    Dim bCollinear As Boolean
    eCase = GetMemberBoundingCase(oAppConnection, bEndToEnd, bCollinear, True)
    Dim oWebLeftBdgPort As IJPort
    Dim oBoundedLateralPort As IJPort
    If Not bEndToEnd Then
        If GetFrameConnectionType(oBoundedPort) = AxisEnd Then
            If eBoundedType = MemberCategoryAndType.StrStringer And eBoundingType = MemberCategoryAndType.StrStringer Then
                bEndToEnd = True
            End If
        End If
    End If
    
    Dim bIsBoundingTube As Boolean
    Dim bIsBoundedTube As Boolean
    bIsBoundingTube = IsTubularMember(oBoundingPort.Connectable)
    bIsBoundedTube = IsTubularMember(oBoundedPort.Connectable)
    
    Dim oStructPort As IJStructPort
    Dim ePortId As JXSEC_CODE
    
    Dim dEdgeOverlap As Double
    dEdgeOverlap = 0#
    
    Dim bPenetratesWeb As Boolean
    
    If Not TypeOf oBoundingPort.Connectable Is ISPSDesignedMember Then
        Set MapCollection = GetEdgeMap(oAppConnection, oBoundingPort, oBoundedPort, , bPenetratesWeb)
        Set oWebLeftBdgPort = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, ReverseMap(JXSEC_WEB_LEFT, MapCollection))
        If Not oWebLeftBdgPort Is Nothing Then
            Set oStructPort = oWebLeftBdgPort
            ePortId = oStructPort.SectionID
        End If
    End If
    
    If (Not bEndToEnd) And bIsBoundedTube And bIsBoundingTube Then
        oSL.Add "F0013"
    End If

   
    Set oBoundedLateralPort = oMember.BasePortBeforeTrim(BPT_Lateral)
    
    If bEndToEnd Then
        ' Checks for Miter Cases
            Dim bIsIdentical As Boolean
            Dim bIsMitered As Boolean
            AreMembersIdentical oBoundedPort, oBoundingPort, bIsIdentical
            If bIsIdentical And Not bCollinear Then
                IsMiterConnectionPossible oBoundedPort, oBoundingPort, bIsMitered
            End If
            Select Case eCase
            
                Case Center
                
                    ' If the bounded object is a Rung
                    If eBoundedType = MemberCategoryAndType.LRung Then

                        Dim BoundingAxisPort As ISPSSplitAxisPort
                        ' Checking for Top and Bottom of the member
                        oBounding.PointAtEnd(SPSMemberAxisStart).GetPoint x1, y1, z1
                        Set StartPosition = New DPosition
                        StartPosition.Set x1, y1, z1
                        oBounding.PointAtEnd(SPSMemberAxisEnd).GetPoint x2, y2, z2
                        Set EndPosition = New DPosition
                        EndPosition.Set x2, y2, z2
                        
                        GetBoundingEndPort oBoundedPort, oBoundingPort, BoundingAxisPort
                        
                        If ((BoundingAxisPort.PortIndex = SPSMemberAxisStart) And GreaterThan(StartPosition.z, EndPosition.z)) Or _
                        ((BoundingAxisPort.PortIndex = SPSMemberAxisEnd) And GreaterThan(EndPosition.z, StartPosition.z)) Then
                            
                            'As of now Slots cutouts are supported only on Web part but not Flange Part, TR-268353 to fix this
                            'So select Slot ACs when penetrating through web
                            If ePortId = JXSEC_WEB_LEFT Or ePortId = JXSEC_WEB_RIGHT Then
                                If Not (IsPortIntersectingObj(oBoundedLateralPort, oWebLeftBdgPort)) Then
                                    oSL.Add "F0001-U0001" ' Creates SnipeFreeEnd(U0001) on LongBox and ShortBox(F0001) follows bounding
                                    oSL.Add "F0010-U0001" ' Created SnipeFreeEnd(U0001) on LongBox and creates slot with EndCut(F0010)
                                Else
                                    oSL.Add "F0011-U0001" ' Created SnipeFreeEnd(U0001) on LongBox and creates ONLY slot(F0010)
                                    oSL.Add "F0010-U0001" ' Created SnipeFreeEnd(U0001) on LongBox and creates slot with EndCut(F0010)
                                    oSL.Add "F0001-U0001" ' Created SnipeFreeEnd(U0001) on LongBox and ShortBox(F0001) follows bounding
                                End If
                            Else
                                oSL.Add "F0001-U0001" ' Creates SnipeFreeEnd(U0001) on LongBox and ShortBox(F0001) follows bounding
                            End If
                        Else
                            'As of now Slots cutouts are supported only on Web part but not Flange Part, TR-268353 to fix this
                            'So select Slot ACs when penetrating through web
                            If ePortId = JXSEC_WEB_LEFT Or ePortId = JXSEC_WEB_RIGHT Then
                                If Not (IsPortIntersectingObj(oBoundedLateralPort, oWebLeftBdgPort)) Then
                                    oSL.Add "F0001-U0002" ' Created SquareFreeEnd(U0002) on LongBox and ShortBox(F0001) follows bounding
                                    oSL.Add "F0010-U0002" ' Created SquareFreeEnd(U0002) on LongBox and creates slot with EndCut(F0010)
                                Else
                                    oSL.Add "F0011-U0002" ' Created SquareFreeEnd(U0002) on LongBox and creates ONLY slot(F0010)
                                    oSL.Add "F0010-U0002" ' Created SquareFreeEnd(U0002) on LongBox and creates slot with EndCut(F0010)
                                    oSL.Add "F0001-U0002" ' Created SquareFreeEnd(U0002) on LongBox and ShortBox(F0001) follows bounding
                                End If
                            Else
                                oSL.Add "F0001-U0002" ' Created SquareFreeEnd(U0002) on LongBox and ShortBox(F0001) follows bounding
                            End If
                        End If
                    ' If the bounded object is a handrail post
                    ElseIf eBoundingType = MemberCategoryAndType.HRPost Then
                    
                        oSL.Add "F0001-U0003" ' Created SemiCircularFreeEnd(U0002) on LongBox and ShortBox(F0001) follows bounding
                    Else
                        If bIsIdentical And bIsMitered Then
                            oSL.Add "M0001"
                            oSL.Add "F0001-U0002"
                        Else
                            oSL.Add "F0001-U0002"
                        End If
                    End If
                Case OnMemberTop, OnMemberBottom
                    If bIsIdentical And bIsBoundingTube And bIsBoundedTube Then
                        oSL.Add "M0001"
                    Else
                        oSL.Add "OM0001"
                    End If

                Case OSBottomAndOSTop
                    
                    Set_AttributeValue oAppConnection, "IJUAFlipAttribute", "FlipPrimaryandSecondary", True
                    
                    Dim eReverseCase As eMemberBoundingCase
                    eReverseCase = GetMemberBoundingCase(oAppConnection, , , True)
                    
                    'Get Mapp Collection
                    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort
                    
                    If Not IsTubularMember(oBoundingPort.Connectable) Then
                        Set MapCollection = GetEdgeMap(oAppConnection, oBoundingPort, oBoundedPort)
                    End If
                    
                    'Set back the attribute value to False
                    Set_AttributeValue oAppConnection, "IJUAFlipAttribute", "FlipPrimaryandSecondary", False
                    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort
                    
                    Select Case eReverseCase
                    
                        Case Center
                            If eBoundingType = MemberCategoryAndType.LRung Then ' for box connections
                            
                                ' Checking for Top and Bottom of the member
                                oBounded.PointAtEnd(SPSMemberAxisStart).GetPoint x1, y1, z1
                                Set StartPosition = New DPosition
                                StartPosition.Set x1, y1, z1
                                oBounded.PointAtEnd(SPSMemberAxisEnd).GetPoint x2, y2, z2
                                Set EndPosition = New DPosition
                                EndPosition.Set x2, y2, z2
                                
                                Set BoundingAxisPort = oBoundedPort
                                
                                Dim oBoundingLateralPort  As IJPort
                                Dim oBoundedWLPort As IJPort
                                Dim oBngMemebr As New StructDetailObjects.MemberPart
                                Set oBngMemebr.object = oBounding
                                Set oBoundingLateralPort = oBngMemebr.BasePortBeforeTrim(BPT_Lateral)
                                Set oBoundedWLPort = GetLateralSubPortBeforeTrim(oBoundedPort.Connectable, ReverseMap(JXSEC_WEB_LEFT, MapCollection))
                                Set oStructPort = oBoundedWLPort
                                ePortId = oStructPort.SectionID
                                
                                If ((BoundingAxisPort.PortIndex = SPSMemberAxisStart) And GreaterThan(StartPosition.z, EndPosition.z)) Or _
                                ((BoundingAxisPort.PortIndex = SPSMemberAxisEnd) And GreaterThan(EndPosition.z, StartPosition.z)) Then
                                                                        
                                    'As of now Slots cutouts are supported only on Web part but not Flange Part, TR-268353 to fix this
                                    'So select Slot ACs when penetrating through web
                                    If ePortId = JXSEC_WEB_LEFT Or ePortId = JXSEC_WEB_RIGHT Then
                                        If Not (IsPortIntersectingObj(oBoundingLateralPort, oBoundedWLPort)) Then
                                            oSL.Add "F0001-U0001" ' Creates SnipeFreeEnd(U0001) on LongBox and ShortBox(F0001) follows bounding
                                            oSL.Add "F0010-U0001" ' Created SnipeFreeEnd(U0001) on LongBox and creates slot with EndCut(F0010)
                                        Else
                                            oSL.Add "F0011-U0001" ' Created SnipeFreeEnd(U0001) on LongBox and creates ONLY slot(F0010)
                                            oSL.Add "F0010-U0001" ' Created SnipeFreeEnd(U0001) on LongBox and creates slot with EndCut(F0010)
                                            oSL.Add "F0001-U0001" ' Created SnipeFreeEnd(U0001) on LongBox and ShortBox(F0001) follows bounding
                                        End If
                                    Else
                                        oSL.Add "F0001-U0001" ' Creates SnipeFreeEnd(U0001) on LongBox and ShortBox(F0001) follows bounding
                                    End If
                                Else
                                    'As of now Slots cutouts are supported only on Web part but not Flange Part, TR-268353 to fix this
                                    'So select Slot ACs when penetrating through web
                                    If ePortId = JXSEC_WEB_LEFT Or ePortId = JXSEC_WEB_RIGHT Then
                                        If Not (IsPortIntersectingObj(oBoundingLateralPort, oBoundedWLPort)) Then
                                            oSL.Add "F0001-U0002" ' Created SquareFreeEnd(U0002) on LongBox and ShortBox(F0001) follows bounding
                                            oSL.Add "F0010-U0002" ' Created SquareFreeEnd(U0002) on LongBox and creates slot with EndCut(F0010)
                                        Else
                                            oSL.Add "F0011-U0002" ' Created SquareFreeEnd(U0002) on LongBox and creates ONLY slot(F0010)
                                            oSL.Add "F0010-U0002" ' Created SquareFreeEnd(U0002) on LongBox and creates slot with EndCut(F0010)
                                            oSL.Add "F0001-U0002" ' Created SquareFreeEnd(U0002) on LongBox and ShortBox(F0001) follows bounding
                                        End If
                                    Else
                                        oSL.Add "F0001-U0002" ' Created SquareFreeEnd(U0002) on LongBox and ShortBox(F0001) follows bounding
                                    End If
                                End If

                            ElseIf eBoundedType = MemberCategoryAndType.HRPost Then
                                oSL.Add "F0001-U0003" ' Created SemiCircularFreeEnd(U0002) on LongBox and ShortBox(F0001) follows bounding
                            End If
                        Case Else
                            If bIsIdentical And bIsBoundingTube And bIsBoundedTube Then
                                oSL.Add "M0001"
                            End If
                End Select
                        
            Case Else
                If bIsIdentical And bCollinear Then
                    oSL.Add "S0001"
                    If eBndgTypeCategory = MemberCategoryAndType.HandRailElement Then ' handrail
                        oSL.Add "S0002"
                    End If
                ElseIf bIsIdentical And bIsMitered Then
                    oSL.Add "M0001"
                End If
        End Select
    Else
        
        Select Case eCase
        
            Case Center
                               
                If TypeOf oBoundingPort.Connectable Is ISPSMemberPartPrismatic Then
                    
               ' F0010 is given with slot on penetrated member and endcuts on penetrating member.
               ' F0011 is given with Slot on penetrated member. This is given by default if Rung is completelu penetrates through Rail
                    If (eBoundingType = MemberCategoryAndType.LRail) And (eBoundedType = MemberCategoryAndType.LRung) Then
                        
                        'As of now Slots cutouts are supported only on Web part but not Flange Part, TR-268353 to fix this
                        'So select Slot ACs when penetrating through web
                        If ePortId = JXSEC_WEB_LEFT Or ePortId = JXSEC_WEB_RIGHT Then
                            If Not (IsPortIntersectingObj(oBoundedLateralPort, oWebLeftBdgPort)) Then
                                oSL.Add "F0001"
                                oSL.Add "F0010"
                            Else
                                oSL.Add "F0011"
                                oSL.Add "F0010"
                                oSL.Add "F0001"
                            End If
                        Else
                            oSL.Add "F0001"
                        End If
                    
                    ElseIf ((eBoundedType = MemberCategoryAndType.HRPost) Or (eBoundedType = MemberCategoryAndType.HRAttachment) Or _
                          (eBoundedType = MemberCategoryAndType.StrSupport) Or (eBoundedType = MemberCategoryAndType.StrAttachment) Or _
                          (eBoundedType = MemberCategoryAndType.LAttachment)) And Not ((eBndgTypeCategory = MemberCategoryAndType.HandRailElement) Or _
                          (eBndgTypeCategory = MemberCategoryAndType.StairElement) Or (eBndgTypeCategory = MemberCategoryAndType.LadderElement)) Then
                            If StrComp(sSectionType, "RS", vbTextCompare) = 0 Then
                                oSL.Add "F0008"
                                oSL.Add "F0005"
                            ElseIf StrComp(sSectionType, "L", vbTextCompare) = 0 Then
                                oSL.Add "F0007"
                                oSL.Add "F0006"
                            ElseIf IsTubularMember(oBounded) Then
                                oSL.Add "F0006"
                                oSL.Add "F0005"
                            Else
                                oSL.Add "F0005"
                                oSL.Add "F0006"
                                oSL.Add "F0007"
                            End If
                                oSL.Add "F0001"
                                oSL.Add "F0001A"
                                oSL.Add "F0001B"
                                If bPenetratesWeb Then
                                    oSL.Add "F0002"
                                    oSL.Add "F0003"
                                    oSL.Add "F0003A"
                                    oSL.Add "F0004"
                                End If
                    ElseIf bIsBoundingTube And (eBoundedType = MemberCategoryAndType.HRPost) And (eBoundingType = MemberCategoryAndType.HRTopRail) Then
                        If bPenetratesWeb Then
                           oSL.Add "XX0001"
                        End If
                    Else
                        oSL.Add "F0001"
                        oSL.Add "F0001A"
                        oSL.Add "F0001B"
                        If bPenetratesWeb Then
                            oSL.Add "F0002"
                            oSL.Add "F0003"
                            oSL.Add "F0003A"
                            oSL.Add "F0004"
                        End If
                    End If
                End If
            
            Case TopEdge, BottomEdge
                oSL.Add "E0001"
            Case BottomEdgeAndTopEdge
                Dim dBottomInsideOverlap As Double
                Dim dTopInsideOverlap As Double
                
                If bPenetratesWeb Then   ' Applicable only for Selection for the WP cases.
                    'Get Top and Bottom Edge Clearance
                    GetEdgeOverlapAndClearance oAppConnection, True, True, dBottomInsideOverlap
                    GetEdgeOverlapAndClearance oAppConnection, False, False, dTopInsideOverlap
                    
                    'Selection
                    SelectEdgeAndEdgeACSmartItems oSL, dTopInsideOverlap, dBottomInsideOverlap
                End If

            
            Case BottomEdgeAndOSTop, TopEdgeAndOSBottom
            
            Case BottomEdgeAndOSTopEdge, TopEdgeAndOSBottomEdge
            
            Case FCAndBottomEdge, FCAndTopEdge

                If eCase = FCAndBottomEdge Then
                    GetEdgeOverlapAndClearance oAppConnection, True, True, dEdgeOverlap
                Else
                    GetEdgeOverlapAndClearance oAppConnection, False, False, dEdgeOverlap
                End If
                
                SelectFaceAndEdgeACSmartItems oSL, dEdgeOverlap
                
            Case FCAndOSBottomEdge, FCAndOSTopEdge
                
            
            Case FCAndOSBottom, FCAndOSTop
            
                If (Not IsTubularMember(oBounded)) Then
                    oSL.Add "FX0001"
                End If
            
            Case OnMemberTop, OnMemberBottom
            
                If TypeOf oBoundedPort Is ISPSSplitAxisAlongPort And TypeOf oBoundingPort Is ISPSSplitAxisAlongPort Then
                    oSL.Add "SN0001"
                Else
                    oSL.Add "OM0001"
                End If
                        
            Case OSBottomAndOSTopEdge, OSTopAndOSBottomEdge
            
            Case OSBottomEdgeAndOSTopEdge
                oSL.Add "XX0002"
                
            Case OSBottomAndOSTop
            
                    If bIsBoundingTube Then
                        If (eBoundedType = MemberCategoryAndType.HRPost) And (eBoundingType = MemberCategoryAndType.HRTopRail) Then
                            If bPenetratesWeb Then
                                oSL.Add "XX0001"
                            End If
                        End If
                    Else
                        If bPenetratesWeb Then
                            oSL.Add "XX0003"
                        End If
                    End If
                
            Case OnTubeMember
                If TypeOf oBoundedPort Is ISPSSplitAxisAlongPort And TypeOf oBoundingPort Is ISPSSplitAxisAlongPort Then
                    oSL.Add "SN0001"
                End If
                
            Case Else
            
                If (TypeOf oBoundingPort.Connectable Is SPSSlabEntity) Or (TypeOf oBoundingPort.Connectable Is SPSWallPart) _
                    Or (TypeOf oBoundingPort.Connectable Is IJPlate) Or (TypeOf oBoundingPort.Connectable Is ISPSDesignedMember) Then
                      'Added the check for the design member
                    Set oBounded = oBoundedPort.Connectable
                    eBoundedType = oBounded.MemberType.Type
                    Set oMember.object = oBounded
                    sSectionType = oMember.sectionType
                    If ((eBoundedType = MemberCategoryAndType.HRPost) Or (eBoundedType = MemberCategoryAndType.HRAttachment) Or _
                        (eBoundedType = MemberCategoryAndType.StrSupport) Or (eBoundedType = MemberCategoryAndType.StrAttachment) Or _
                        (eBoundedType = MemberCategoryAndType.LAttachment) Or (eBoundedType = MemberCategoryAndType.StrStringer)) Then
                        If StrComp(sSectionType, "RS", vbTextCompare) = 0 Then
                            oSL.Add "F0008"
                            oSL.Add "F0005"
                            oSL.Add "F0009"
                        ElseIf StrComp(sSectionType, "L", vbTextCompare) = 0 Then
                            oSL.Add "F0007"
                            oSL.Add "F0006"
                            oSL.Add "F0009"
                        ElseIf IsTubularMember(oBounded) Then
                            oSL.Add "F0006"
                            oSL.Add "F0005"
                            oSL.Add "F0009"
                        Else
                        
                            If eBoundedType = MemberCategoryAndType.StrStringer Then
                            
                                oBounded.PointAtEnd(SPSMemberAxisStart).GetPoint x1, y1, z1
                                Set StartPosition = New DPosition
                                StartPosition.Set x1, y1, z1
                                oBounded.PointAtEnd(SPSMemberAxisEnd).GetPoint x2, y2, z2
                                Set EndPosition = New DPosition
                                EndPosition.Set x2, y2, z2

                                Set oBoundedAxisPort = oBoundedPort
                                
                                If (oBoundedAxisPort.PortIndex = SPSMemberAxisStart And StartPosition.z < EndPosition.z) Or _
                                   (oBoundedAxisPort.PortIndex = SPSMemberAxisEnd And EndPosition.z < StartPosition.z) Then
                                    
                                    oSL.Add "F0012"
                                    
                                End If
                                
                            End If
                            
                            oSL.Add "F0005"
                            oSL.Add "F0006"
                            oSL.Add "F0007"
                            oSL.Add "F0009"
                        End If
                    Else
                        oSL.Add "F0009"
                    End If
                End If
        End Select
    End If
    
    
    Exit Sub
    
ErrorHandler:
    oSL.ReportError sMsg, METHOD
End Sub

'*********************************************************************************************
' ** End CM **

'*********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
'*********************************************************************************************

'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = m_SelectorName
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSelector As IJDSymbolDefinition)
    
    ' Remove all existing defined Input and Output (Representations) before defining the current Inputs and Outputs
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

'*********************************************************************************************
'         !!!!! End Private Code !!!!!
'*********************************************************************************************

'*********************************************************************************************
' Method      : CMSetFlipAnswer
' Description :
'*********************************************************************************************
Public Sub CMSetFlipAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                           ByRef ppArgument As Object)
                           
    Const METHOD = m_SelectorProgid & "::CMSetFlipAnswer"
    
    On Error GoTo ErrorHandler
    
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    pSL.Answer("Flip") = gsNo
    
    Exit Sub
ErrorHandler:
   Err.Raise LogError(Err, MODULE, METHOD, "").Number
   
End Sub



'*********************************************************************************************
' Method      : CMSetFlipAnswer
' Description :
'*********************************************************************************************
Public Sub CMSetFlipPrimaryAndSecondaryAnswer(ByVal pInput As IMSSymbolEntities.IJDInputStdCustomMethod, _
                           ByRef ppArgument As Object)
                           
    Const METHOD = m_SelectorProgid & "::CMSetFlipAnswer"
    
    On Error GoTo ErrorHandler
    
    Dim pSL As IJDSelectorLogic
    Set pSL = GetSelectorLogicForCustomMethod(pInput)
    
    pSL.Answer("Flip Primary and Secondary") = gsNo
    
    Exit Sub
ErrorHandler:
   Err.Raise LogError(Err, MODULE, METHOD, "").Number
   
End Sub

Private Sub SelectEdgeAndEdgeACSmartItems(oSL As IJDSelectorLogic, dTopInsideOverlap As Double, dBottomInsideOverlap As Double)

  Const METHOD = "SelectEdgeAndEdgeACSmartItems"
  On Error GoTo ErrorHandler

    oSL.Add "EE0010"
    oSL.Add "EE0009"
    oSL.Add "EE0009A"
    
    Dim dLimit As Double
    dLimit = 0.025 '25mm
    
    Dim sFlipAnswer As String
    GetSelectorAnswer oSL, "Flip", sFlipAnswer
    
    Dim bBothBetweenZeroAnd25 As Boolean
    bBothBetweenZeroAnd25 = False
    
    'Both top and bottom corner clearances fall between 0 and 25mm range
    If IsValueBetween(dTopInsideOverlap, 0, dLimit) And IsValueBetween(dBottomInsideOverlap, 0, dLimit) Then
        bBothBetweenZeroAnd25 = True
    End If
    
    If bBothBetweenZeroAnd25 Then
        oSL.Add "EE0005"
    End If
        
    oSL.Add "EE0006"
    
    If bBothBetweenZeroAnd25 Then 'Both top and bottom corner clearances fall between 0 and 25mm
        oSL.Add "EE0001"
    End If
    
    'Both Top and Bottom clearance should be positive value to place top and bottom cuts
    If GreaterThanOrEqualTo(dTopInsideOverlap, 0) And GreaterThanOrEqualTo(dBottomInsideOverlap, 0) Then
        oSL.Add "EE0002"
    End If
           
    If bBothBetweenZeroAnd25 Then 'Both top and bottom corner clearances fall between 0 and 25mm
        oSL.Add "EE0003"
    End If
           
    If GreaterThanOrEqualTo(dTopInsideOverlap, 0) Then 'TopClearance is more than zero
        If LessThanZero(dBottomInsideOverlap) Then 'BottomClearance is less than Zero
            If sFlipAnswer = gsNo Then 'Check Flip case
                oSL.Add "EE0004"
            Else
                oSL.Add "EE0008"
            End If
        Else 'BottomClearance is greater than Zero
            oSL.Add "EE0004"
            oSL.Add "EE0008"
        End If
        
    ElseIf GreaterThanOrEqualTo(dBottomInsideOverlap, 0) Then   'BottomClearance is greater than Zero
        If sFlipAnswer = gsYes Then 'Check Flip case
            oSL.Add "EE0004"
        Else
            oSL.Add "EE0008"
        End If
    End If
    
    'TopClearance is between 0 and 25mm
    If IsValueBetween(dTopInsideOverlap, 0, dLimit) Then
        'Bottom Clearance is between 0 and 25mm
        If IsValueBetween(dBottomInsideOverlap, 0, dLimit) Then
            oSL.Add "EE0007"
        Else
            If sFlipAnswer = gsYes Then
                oSL.Add "EE0007"
            End If
        End If
    ElseIf IsValueBetween(dBottomInsideOverlap, 0, dLimit) Then  'Bottom Clearance is between 0 and 25mm
        If sFlipAnswer = gsNo Then
            oSL.Add "EE0007"
        End If
    End If
    
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Sub
'*************************************************************************
'Sub-routine
'Method     :   SelectFaceAndEdgeACSmartItems
'Abstract   :   Given the edge overlap value, this method will select the appropriate selections by rule.
'input      :   Connection or its child struct feature object
'***************************************************************************
Private Sub SelectFaceAndEdgeACSmartItems(oSL As IJDSelectorLogic, edgeOverlap As Double)

    Const METHOD = "SelectFaceAndEdgeACSmartItems"
    On Error GoTo ErrorHandler
    ' Get some information about the bounding condition
    Dim bIsWebPenetrated As Boolean
    
    bIsWebPenetrated = IsConnectionWebPenetrated(oSL.SmartOccurrence)
    
    If bIsWebPenetrated Then
        oSL.Add "FE0006"
        oSL.Add "FE0008"
        oSL.Add "FE0008A"
        oSL.Add "FE0008B"
        oSL.Add "FE0010"
        oSL.Add "FE0012"
        oSL.Add "FE0012A"
        oSL.Add "FE0012B"
        oSL.Add "FE0012C"
        
        'Clearance falls between 0 and 25mm range
        If IsValueBetween(edgeOverlap, 0, 0.025) Then
            oSL.Add "FE0001"
            oSL.Add "FE0003"
            oSL.Add "FE0003A"
            oSL.Add "FE0003B"
            oSL.Add "FE0005"
            oSL.Add "FE0007"
            oSL.Add "FE0007A"
            oSL.Add "FE0007B"
            oSL.Add "FE0009"
            oSL.Add "FE0009A"
            oSL.Add "FE0009B"
            oSL.Add "FE0011"
            oSL.Add "FE0013"
        End If
        
        If GreaterThanOrEqualTo(edgeOverlap, 0) Then
            oSL.Add "FE0002"
            oSL.Add "FE0004"
            oSL.Add "FE0004A"
            oSL.Add "FE0004B"
        End If
    Else
        ' Not currently supported for flange-penetrated cases
    End If

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Sub
