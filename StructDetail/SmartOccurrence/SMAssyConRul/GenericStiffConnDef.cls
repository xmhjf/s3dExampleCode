VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericStiffConnDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'*********************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMAssyConRul
'  File        : GenericStiffConnDef.cls
'
'  Description : Control creation of the Member Items for the GenAssyConnDef Selector Class
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'
'*********************************************************************************************

Const m_ItemProgid As String = CUSTOMERID + "AssyConRul.GenericStiffConnDef"
Const m_ItemName As String = CUSTOMERID + "AssyConRul.GenericStiffConnDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\GenericStiffConnDef.cls"

Implements IJDUserSymbolServices


'*********************************************************************************************
' Method      : ItemInputs
' Description : ' List any graphic Inputs that the Definition has here
'
'*********************************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)

    Const METHOD = m_ItemProgid & "::ItemInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining Selector Inputs"
    
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT

    
    Exit Sub
ErrorHandler:
    pIH.ReportError sMsg, METHOD
    
End Sub

'*********************************************************************************************
' Method      : ItemAggregator
' Description :
'
'*********************************************************************************************
Public Sub ItemAggregator(pAD As IJDAggregatorDescription)

    Const METHOD = m_ItemProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ItemAggregator"
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

'*********************************************************************************************
' Method      : ItemMembers
' Description : ' List all the Item members that this SmartClass can create
'               lDispId 1 through 5 are for web cuts
'               lDispId 6 through 10 are for top flange cuts
'               lDispId 11 through 15 are for bottom flange cuts
'
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)

    Const METHOD = m_ItemProgid & "::ItemMembers"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sDispId As String

    Dim lDispId As Long
    Dim oMemDesc As IJDMemberDescription

    'lDispId 1 through 5 are for web cuts. A web cut may be created for each bounding member,
    'up to a total of 5 bounding members
    For lDispId = 1 To 5
        sDispId = "Generic_Stiff_Web_" & Trim(Str(lDispId))
        Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_CreateGenericWebCut", CUSTOMERID + "MbrAC.MbrACDefCM")
        oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsWebCutNeeded"
        Set oMemDesc = Nothing
    Next lDispId

    'lDispId 6 through 10 are for top flange cuts. A flange cut may be created on the top flange
    'for each bounding member, up to a total of 5 bounding members
    For lDispId = 6 To 10
        sDispId = "Generic_Stiff_TopFlange_" & Trim(Str(lDispId))
        Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_ConstructTopFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
        oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFlangeCutNeeded"
        oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CM_CopyQuestionsToTopFlangeCut"
        Set oMemDesc = Nothing
    Next lDispId

    'lDispId 11 through 15 are for bottom flange cuts. A flange cut may be created on the bottom flange
    'for each bounding member, up to a total of 5 bounding members
    For lDispId = 11 To 15
        sDispId = "Generic_Stiff_BottomFlange_" & Trim(Str(lDispId))
        Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CM_ConstructBottomFlangeCut", CUSTOMERID + "MbrAC.MbrACDefCM")
        oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBottomFlangeCutNeeded"
        oMemDesc.SetCMSetInputs CUSTOMERID + "MbrAC.MbrACDefCM", "CM_CopyQuestionsToBottomFlangeCut"
        Set oMemDesc = Nothing
    Next lDispId

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
    
End Sub

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Smart Definition as a regular symbol definition
' ********************************************************************************************

'*********************************************************************************************
' Method      : IJDUserSymbolServices_GetDefinitionName
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique

    IJDUserSymbolServices_GetDefinitionName = m_ItemName

End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InitializeSymbolDefinition
' Description :
'
'*********************************************************************************************
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)

    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pDefinition.IJDInputs.RemoveAllInput
    pDefinition.IJDRepresentations.RemoveAllRepresentation

    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
      
    ' define the inputs
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Dim pAPDs As IJDPropertyDescriptions
    Set pAD = pDefinition
    Set pAPDs = pDefinition
    pAPDs.RemoveAll ' Removes all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Removes all the previous Member descriptions
    ItemMembers pMDs
    
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InstanciateDefinition
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                            ByVal defParams As Variant, _
                                                            ByVal ActiveConnection As Object) As Object
                                                            
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_ItemProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
    
End Function

'*********************************************************************************************
' Method      : IJDUserSymbolServices_InvokeRepresentation
' Description :
'
'*********************************************************************************************
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, _
                                                      ByVal repName As String, _
                                                      ByVal outputcoll As Object, _
                                                      ByRef arrayOfInputs())
                                                      
End Sub

'*********************************************************************************************
' Method      : IJDUserSymbolServices_EditOccurence
' Description :
'
'*********************************************************************************************
Public Function IJDUserSymbolServices_EditOccurence(pSymbolOccurence As Object, _
                                                    ByVal transactionMgr As Object) As Boolean
                                                    
End Function

' *******************************************************************************************
' Method:
' CM_IsWebCutNeeded
'
' Description: Conditional method to determine if a Web cut is needed
' *******************************************************************************************
Public Sub CM_IsWebCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "MbrACDefCM.CM_IsWebCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim lDispId As Long
    Dim lStatus As Long
    Dim i As Long
    Dim oAppConnection As IJAppConnection
    Dim oAssyConn As AssemblyConn
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If

    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    bIsNeeded = False
    lDispId = pMemberDescription.dispid
    Set oAppConnection = pMemberDescription.CAO
    Set oConnAttrbs = New SDConnectionUtils

    Dim oBoundingObjectColl As JCmnShp_CollectionAlias
    Set oBoundingObjectColl = New Collection
    
    For i = 1 To oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count
        oBoundingObjectColl.Add oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Item(i)
    Next
    
    Set oAssyConn = New AssemblyConn
    Set oAssyConn.object = pMemberDescription.CAO
    
    oBoundingObjectColl.Add oAssyConn.Port2

    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg
    
    If lStatus <> 0 Then
        Exit Sub
    End If

    Dim oWebLeftPort As IJPort
    Dim oWebRightPort As IJPort

    Dim bIsWebLeftIntersecting As Boolean
    Dim bIsWebRightIntersecting As Boolean

    Dim bIsBoundedTube As Boolean
    bIsBoundedTube = IsTubularMember(oBoundedData.MemberPart)

    'If the Web Doesn't intersect any of the Bounding Objects then we will create a web cut to the
    'back port of the first bounding object
    If lDispId = 1 Then
        Dim lIndex As Long
        Dim bWebIntersecting As Boolean
        bWebIntersecting = False
        'For each bounding object, check if the web intersects the bounding object. If it doesn't intersect
        'any then create the web cut on the first bounding object.
        For lIndex = 1 To oBoundingObjectColl.Count
            If bIsBoundedTube Then
                Set oWebLeftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_OUTER_TUBE)
            Else
                Set oWebLeftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
                Set oWebRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)
            End If

            If Not oWebLeftPort Is Nothing Then
                bIsWebLeftIntersecting = IsPortIntersectingObj(oWebLeftPort, oBoundingObjectColl.Item(lIndex))
            End If
            If Not oWebRightPort Is Nothing Then
                bIsWebRightIntersecting = IsPortIntersectingObj(oWebRightPort, oBoundingObjectColl.Item(lIndex))
            End If

            If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then
                bWebIntersecting = True
                If lIndex = 1 Then bIsNeeded = True
                    Exit For
            Else
                'Check if the extended ports intersect the bounding object
                If Not oWebLeftPort Is Nothing Then
                    bIsWebLeftIntersecting = IsExtendedPortIntersectingBoundingObj(oWebLeftPort, oBoundingObjectColl.Item(lIndex))
                End If
                If Not oWebRightPort Is Nothing Then
                    bIsWebRightIntersecting = IsExtendedPortIntersectingBoundingObj(oWebRightPort, oBoundingObjectColl.Item(lIndex))
                End If
                If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then
                    bWebIntersecting = True
                    If lIndex = 1 Then bIsNeeded = True
                    Exit For
                End If
            End If
        Next lIndex

        If bWebIntersecting = False Then
            'Web doesn't intersect any objects. However, we need a web cut as input to the flange cuts.
            'Create the Web Cut If there isn't allready a web cut created by a sibling AC
            Dim oDesignParent As IJDesignParent
            Dim oDesignChild As IJDesignChild
            Dim oAC As Object
            Set oAC = pMemberDescription.CAO
            Dim bWebCutExists As Boolean
            bWebCutExists = False

            If TypeOf oAC Is IJDesignChild Then
                Set oDesignChild = oAC
                Set oDesignParent = oDesignChild.GetParent

                Dim oChildren As IJDObjectCollection
                oDesignParent.GetChildren oChildren

                For Each oAC In oChildren
                    'Check each sibling Assembly Connection on the Bounded Member to see if it is for the same
                    'end port on the bounded member
                    If TypeOf oAC Is IJAssemblyConnection And Not oAC Is pMemberDescription.CAO Then
                        Dim oACBoundedData As MemberConnectionData
                        Dim oACBoundingData As MemberConnectionData
                        InitMemberConnectionData oAC, oACBoundedData, oACBoundingData, lStatus, sMsg
                        If oACBoundedData.ePortId = oBoundedData.ePortId And oACBoundedData.MemberPart = oACBoundingData.MemberPart Then
                            'AC is for the same end port and it is a generic AC
                            'check to see if any Web Cuts exist on the AC
                            Dim oACChildren As IJDObjectCollection
                            If TypeOf oAC Is IJDesignParent Then
                                Set oDesignParent = oAC
                                oDesignParent.GetChildren oACChildren
                                Dim oEndCut As Object
                                For Each oEndCut In oACChildren
                                    If TypeOf oEndCut Is IJStructFeature Then
                                        Dim oStructFeature As IJStructFeature
                                        Set oStructFeature = oEndCut
                                        If oStructFeature.get_StructFeatureType = SF_WebCut Then
                                            'Struct Feature is a Web Cut.  Therefore, we don't need to
                                            'create another web cut.
                                            bWebCutExists = True
                                            Exit For
                                        End If
                                    End If
                                Next oEndCut
                            End If
                            If bWebCutExists Then Exit For
                        End If
                    End If
                Next oAC
                If bWebCutExists Then
                    bIsNeeded = False
                Else
                    bIsNeeded = True
                End If
            End If
        End If
    End If

    If lDispId > 1 And lDispId <= oBoundingObjectColl.Count Then
        'A web cut should be created if the web of the bounded object intersects
        'the bounding object.
        If bIsBoundedTube Then
            Set oWebLeftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_OUTER_TUBE)
        Else
            Set oWebLeftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
            Set oWebRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)
        End If
        If Not oWebLeftPort Is Nothing Then
            bIsWebLeftIntersecting = IsPortIntersectingObj(oWebLeftPort, oBoundingObjectColl.Item(lDispId))
        End If
        If Not oWebRightPort Is Nothing Then
            bIsWebRightIntersecting = IsPortIntersectingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId))
        End If

        If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then
            bIsNeeded = True
        Else
            'Check if the extended ports intersect the bounding object and create the cut if the
            'bounding object is a member.
            bIsWebLeftIntersecting = IsExtendedPortIntersectingBoundingObj(oWebLeftPort, oBoundingObjectColl.Item(lDispId))
            bIsWebRightIntersecting = IsExtendedPortIntersectingBoundingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId))

            If bIsWebLeftIntersecting Or bIsWebRightIntersecting Then bIsNeeded = True
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub



' *******************************************************************************************
' Method:
' CM_IsTopFlangeCutNeeded
'
' Description: Conditional method to determine if a TopFlange cut is needed
' *******************************************************************************************
Public Sub CM_IsTopFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "MbrACDefCM.CM_IsTopFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim lDispId As Long
    Dim lStatus As Long
    Dim i As Long
    Dim oAppConnection As IJAppConnection

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes

    bIsNeeded = False
    lDispId = pMemberDescription.dispid
    Set oAppConnection = pMemberDescription.CAO
    Set oConnAttrbs = New SDConnectionUtils

    Dim oBoundingObjectColl As JCmnShp_CollectionAlias
    Set oBoundingObjectColl = New Collection
    
    For i = 1 To oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count
        oBoundingObjectColl.Add oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Item(i)
    Next
    
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Set oAssyConn = New AssemblyConn
    Set oAssyConn.object = pMemberDescription.CAO
    
    oBoundingObjectColl.Add oAssyConn.Port2

    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg


    Dim bTopFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bBottomFlangeRight As Boolean

    Dim oWebLeftPort As IJPort
    Dim oWebRightPort As IJPort

    Dim bIsWebLeftIntersecting As Boolean
    Dim bIsWebRightIntersecting As Boolean

    Dim oTopFlangeRightTop As IJPort
    Dim oTopFlangeLeftTop As IJPort
    Dim oTopFlangeRightBottom As IJPort
    Dim oTopFlangeLeftBottom As IJPort

    'Booleans to specify if the ports are intersecting the object
    Dim bTopFlangeRightTop As Boolean
    Dim bTopFlangeLeftTop As Boolean
    Dim bTopFlangeRightBottom As Boolean
    Dim bTopFlangeLeftBottom As Boolean

    CrossSection_Flanges oBoundedData.MemberPart, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight

    'A flange cut should be created if the Top Flange is intersecting the bounding object and the web is not intersecting
    'the bounding object. If the web is intersecting the bounding object, it will be responsible for creating the flange cut.
    If bTopFlangeLeft Or bTopFlangeRight Then
        If (lDispId - 5) <= oBoundingObjectColl.Count Then
            If bTopFlangeRight Then
                Set oTopFlangeRightTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP_FLANGE_RIGHT_TOP)
                Set oTopFlangeRightBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
                If oTopFlangeRightTop Is Nothing Then Set oTopFlangeRightTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)

                bTopFlangeRightTop = IsPortIntersectingObj(oTopFlangeRightTop, oBoundingObjectColl.Item(lDispId - 5))
                bTopFlangeRightBottom = IsPortIntersectingObj(oTopFlangeRightBottom, oBoundingObjectColl.Item(lDispId - 5))
            Else
                bTopFlangeRightTop = False
                bTopFlangeRightBottom = False
            End If

            If bTopFlangeLeft Then
                Set oTopFlangeLeftTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP_FLANGE_LEFT_TOP)
                Set oTopFlangeLeftBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP_FLANGE_LEFT_BOTTOM)
                If oTopFlangeLeftTop Is Nothing Then Set oTopFlangeLeftTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_TOP)

                bTopFlangeLeftTop = IsPortIntersectingObj(oTopFlangeLeftTop, oBoundingObjectColl.Item(lDispId - 5))
                bTopFlangeLeftBottom = IsPortIntersectingObj(oTopFlangeLeftBottom, oBoundingObjectColl.Item(lDispId - 5))
            Else
                bTopFlangeLeftTop = False
                bTopFlangeLeftBottom = False
            End If

            'Top Flange is intersecting if any of the top flange ports are intersecting
            If bTopFlangeRightTop Or bTopFlangeLeftTop Or bTopFlangeRightBottom Or bTopFlangeLeftBottom Then
                'Create the flange cut if the web is not intersecting
                Set oWebLeftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
                Set oWebRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)

                bIsWebLeftIntersecting = IsPortIntersectingObj(oWebLeftPort, oBoundingObjectColl.Item(lDispId - 5))
                bIsWebRightIntersecting = IsPortIntersectingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId - 5))

                If Not bIsWebLeftIntersecting And Not bIsWebRightIntersecting Then
                    If TypeOf oBoundingObjectColl.Item(lDispId - 5) Is ISPSMemberPartPrismatic Then
                        bIsWebLeftIntersecting = IsExtendedPortIntersectingBoundingObj(oWebLeftPort, oBoundingObjectColl.Item(lDispId - 5))
                        bIsWebRightIntersecting = IsExtendedPortIntersectingBoundingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId - 5))

                        If Not bIsWebLeftIntersecting And Not bIsWebRightIntersecting Then bIsNeeded = True
                    Else
                        bIsNeeded = True
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub



' *******************************************************************************************
' Method:
' CM_IsBottomFlangeCutNeeded
'
' Description: Conditional method to determine if a Bottom Flange cut is needed
' *******************************************************************************************
Public Sub CM_IsBottomFlangeCutNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = "MbrACDefCM.CM_IsBottomFlangeCutNeeded"
    On Error GoTo ErrorHandler
    Dim sMsg As String

    Dim lDispId As Long
    Dim lStatus As Long
    Dim i As Long
    Dim oAppConnection As IJAppConnection
    Dim oConnAttrbs As GSCADSDCreateModifyUtilities.IJSDConnectionAttributes

    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData

    bIsNeeded = False
    lDispId = pMemberDescription.dispid
    Set oAppConnection = pMemberDescription.CAO
    Set oConnAttrbs = New SDConnectionUtils

    Dim oBoundingObjectColl As JCmnShp_CollectionAlias
    Set oBoundingObjectColl = New Collection
    
    For i = 1 To oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Count
        oBoundingObjectColl.Add oConnAttrbs.get_AuxiliaryPorts(oAppConnection).Item(i)
    Next
    
    Dim oAssyConn As StructDetailObjects.AssemblyConn
    Set oAssyConn = New AssemblyConn
    Set oAssyConn.object = pMemberDescription.CAO
    
    oBoundingObjectColl.Add oAssyConn.Port2

    InitMemberConnectionData oAppConnection, oBoundedData, oBoundingData, lStatus, sMsg


    Dim bTopFlangeLeft As Boolean
    Dim bTopFlangeRight As Boolean
    Dim bBottomFlangeLeft As Boolean
    Dim bBottomFlangeRight As Boolean

    Dim oWebLeftPort As IJPort
    Dim oWebRightPort As IJPort

    Dim bIsWebLeftIntersecting As Boolean
    Dim bIsWebRightIntersecting As Boolean

    Dim oBottomFlangeRightBottom As IJPort
    Dim oBottomFlangeLeftBottom As IJPort
    Dim oBottomFlangeRightTop As IJPort
    Dim oBottomFlangeLeftTop As IJPort

    'Booleans to specify if the ports are intersecting the object
    Dim bBottomFlangeRightBottom As Boolean
    Dim bBottomFlangeLeftBottom As Boolean
    Dim bBottomFlangeRightTop As Boolean
    Dim bBottomFlangeLeftTop As Boolean

    CrossSection_Flanges oBoundedData.MemberPart, bTopFlangeLeft, bBottomFlangeLeft, bTopFlangeRight, bBottomFlangeRight

    'A flange cut should be created if the Bottom Flange is intersecting the bounding object and the web is not intersecting
    'the bounding object. If the web is intersecting the bounding object, it will be responsible for creating the flange cut.
    If bBottomFlangeLeft Or bBottomFlangeRight Then
        If (lDispId - 10) <= oBoundingObjectColl.Count Then
            If bBottomFlangeRight Then
                Set oBottomFlangeRightBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM)
                Set oBottomFlangeRightTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM_FLANGE_RIGHT_TOP)
                If oBottomFlangeRightBottom Is Nothing Then Set oBottomFlangeRightBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM)

                bBottomFlangeRightBottom = IsPortIntersectingObj(oBottomFlangeRightBottom, oBoundingObjectColl.Item(lDispId - 10))
                bBottomFlangeRightTop = IsPortIntersectingObj(oBottomFlangeRightTop, oBoundingObjectColl.Item(lDispId - 10))
            Else
                bBottomFlangeRightBottom = False
                bBottomFlangeRightTop = False
            End If

            If bBottomFlangeLeft Then
                Set oBottomFlangeLeftBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM)
                Set oBottomFlangeLeftTop = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM_FLANGE_LEFT_TOP)
                If oBottomFlangeLeftBottom Is Nothing Then Set oBottomFlangeLeftBottom = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_BOTTOM)

                bBottomFlangeLeftBottom = IsPortIntersectingObj(oBottomFlangeLeftBottom, oBoundingObjectColl.Item(lDispId - 10))
                bBottomFlangeLeftTop = IsPortIntersectingObj(oBottomFlangeLeftTop, oBoundingObjectColl.Item(lDispId - 10))
            Else
                bBottomFlangeLeftBottom = False
                bBottomFlangeLeftTop = False
            End If

            'Bottom Flange is intersecting if any of the bottom flange ports are intersecting
            If bBottomFlangeRightBottom Or bBottomFlangeLeftBottom Or bBottomFlangeRightTop Or bBottomFlangeLeftTop Then
                'Create the flange cut if the web is not intersecting
                Set oWebLeftPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_LEFT)
                Set oWebRightPort = GetLateralSubPortBeforeTrim(oBoundedData.MemberPart, JXSEC_WEB_RIGHT)

                bIsWebLeftIntersecting = IsPortIntersectingObj(oWebLeftPort, oBoundingObjectColl.Item(lDispId - 10))
                bIsWebRightIntersecting = IsPortIntersectingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId - 10))

                If Not bIsWebLeftIntersecting And Not bIsWebRightIntersecting Then
                    If TypeOf oBoundingObjectColl.Item(lDispId - 10) Is ISPSMemberPartPrismatic Then
                        bIsWebLeftIntersecting = IsExtendedPortIntersectingBoundingObj(oWebLeftPort, oBoundingObjectColl.Item(lDispId - 10))
                        bIsWebRightIntersecting = IsExtendedPortIntersectingBoundingObj(oWebRightPort, oBoundingObjectColl.Item(lDispId - 10))

                        If Not bIsWebLeftIntersecting And Not bIsWebRightIntersecting Then bIsNeeded = True
                    Else
                        bIsNeeded = True
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

