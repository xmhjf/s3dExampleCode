VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CustomMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**********************************************************************************************************************
'  Copyright (C) 2014, Intergraph Corporation.  All rights reserved.
'
'  File        : CustomMethods.cls
'
'  Description :
'
'  Author      : Alligators

' 03/Dec/14  CSK/CM CR-250022 Connections and free end cuts at ends of ladder rails
' 12/May/15  GH     CR-260982 Added new method GetPCAttributeForCenterCut() and updated CMConstructCenterCutPC() with it.
'***********************************************************************************************************************
Const m_sClassName As String = "CustomMethods"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sStdACProjectName + "." + m_sClassName
Const MODULE = m_sStdACProjectPath + m_sClassName + ".cls"

'*********************************************************************************************
' Method      : IsBoundingCaseA
' Description :
'
'*********************************************************************************************
Public Function IsBoundingCaseA(oACorEC As Object, Optional ByRef strCaseSuffix As String) As Boolean
                                      
    Const METHOD = m_DefinitionProgid & "::IsBoundingCaseA"

    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Determining case A or case B for standard AC"

    IsBoundingCaseA = True
    
    ' ------------------
    ' Get the connection
    ' ------------------
    Dim oAppConnection As IJAppConnection
    Dim oBoundedData As MemberConnectionData
    Dim oBoundingData As MemberConnectionData
    Dim lStatus As Long
    Dim bPenetratesWeb As Boolean
    
    If TypeOf oACorEC Is IJStructFeature Then
        Dim sACItemName As String
        Dim oACObj As Object
        AssemblyConnection_SmartItemName oACorEC, sACItemName, oAppConnection
       
    Else
        Set oAppConnection = oACorEC
    End If
    
    ' --------------------------------
    ' Determine the bounding condition
    ' --------------------------------
    Dim eCase As eMemberBoundingCase
    eCase = GetMemberBoundingCase(oAppConnection, , , False)
    
    Select Case eCase
        
        ' ----------------------------------------
        ' For symmetric cases, check the flip flag
        ' ----------------------------------------
        Case Center, BottomEdgeAndTopEdge, OSBottomEdgeAndOSTopEdge, OSBottomAndOSTop
            Dim sFlipped As String
            GetSelectorAnswer oAppConnection, "Flip", sFlipped
            
            If sFlipped = gsYes Then
                IsBoundingCaseA = False
            End If
            
        ' -----------------------------------------------------------------------------------------
        ' For asymmetric cases, use the definition of "A" and "B" found in the bulkload spreadsheet
        ' -----------------------------------------------------------------------------------------
        Case TopEdge, BottomEdgeAndOSTop, BottomEdgeAndOSTopEdge, FCAndTopEdge, FCAndOSTopEdge, FCAndOSTop, OnMemberTop, OSTopAndOSBottomEdge

            IsBoundingCaseA = True
    
        Case Else
            IsBoundingCaseA = False
    End Select
    
    If IsBoundingCaseA Then
        strCaseSuffix = "_A"
    Else
        strCaseSuffix = "_B"
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

'*********************************************************************************************
' Method      : CM_CenterCutPCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_CenterCutPCNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsFacePCNeeded"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Determining is PC is needed for center cut" & "...oMD.dispid = " & Str(oMD.dispid)
    
    bIsNeeded = False

    ' ---------------------------
    ' Get the assembly connection
    ' ---------------------------
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName oMD.CAO, , oAppConnection
    
    ' -------------------------------------
    ' Determine if this is Case A or Case B
    ' -------------------------------------
    Dim isCaseA As Boolean
    Dim oCustomMethods As New CustomMethods
    Dim strCaseSuffix As String
    
    isCaseA = oCustomMethods.IsBoundingCaseA(oMD.CAO, strCaseSuffix)
        
    ' ---------------------------------------------
    ' Get the parent AC and ask if the PC is needed
    ' ---------------------------------------------
    Dim strPCAttribute As String
    Dim bIsBottomEdge As Boolean
    bIsBottomEdge = False
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    Dim lStatus As Long
    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort
    
    Dim bPenetratesWeb As Boolean
    bPenetratesWeb = IsWebPenetrated(oBoundingPort, oBoundedPort)


    Dim bIsBottomFlange As Boolean
    Dim strIsBottom As String
    
    GetSelectorAnswer oMD.CAO, "BottomFlange", strIsBottom
    
    If strIsBottom = gsYes Then
        bIsBottomFlange = True
    Else
        bIsBottomFlange = False
    End If
    
    
    Dim oACDef As New ACDef
    
    Select Case oMD.dispid
        Case 1
            If (Not bPenetratesWeb And bIsBottomFlange) Then
                Exit Sub
            End If
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "TopPC_LT" & strCaseSuffix)
        Case 2
            If Not bPenetratesWeb And bIsBottomFlange Then
                Exit Sub
            End If
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "TopFacePC_LT" & strCaseSuffix)
        Case 3
            If Not bPenetratesWeb And bIsBottomFlange Then
                Exit Sub
            End If
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "TopEdgeOutsidePC_LT" & strCaseSuffix)
        Case 4
            If Not bPenetratesWeb And bIsBottomFlange Then
                Exit Sub
            End If
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "TopEdgePC_LT" & strCaseSuffix)
        Case 5
            If Not bPenetratesWeb And bIsBottomFlange Then
                Exit Sub
            End If
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "TopEdgeInsidePC_LT" & strCaseSuffix)
        Case 6
            If Not bPenetratesWeb And bIsBottomFlange Then
                Exit Sub
            End If
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "FacePC_LT" & strCaseSuffix)
        Case 7
            If Not bPenetratesWeb And bIsBottomFlange Then
                Exit Sub
            End If
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "BtmEdgeInsidePC_LT" & strCaseSuffix)
        Case 8
            If Not bPenetratesWeb And bIsBottomFlange Then
                Exit Sub
            End If
            
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "BtmEdgePC_LT" & strCaseSuffix)
           ' bIsBottomEdge = True
        Case 9
            If Not bPenetratesWeb And bIsBottomFlange Then
            
                Exit Sub
            End If
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "BtmEdgeOutsidePC_LT" & strCaseSuffix)
        Case 10
            If Not bPenetratesWeb And bIsBottomFlange Then
                Exit Sub
            End If
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "BtmFacePC_LT" & strCaseSuffix)
        Case 11
            If Not bPenetratesWeb And bIsBottomFlange Then
                Exit Sub
            End If
            strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "BtmPC_LT" & strCaseSuffix)
        Case 12
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "TopPC_RB" & strCaseSuffix)
        Case 13
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "TopFacePC_RB" & strCaseSuffix)
        Case 14
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "TopEdgeOutsidePC_RB" & strCaseSuffix)
        Case 15
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "TopEdgePC_RB" & strCaseSuffix)
        Case 16
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "TopEdgeInsidePC_RB" & strCaseSuffix)
        Case 17
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "FacePC_RB" & strCaseSuffix)
        Case 18
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "BtmEdgeInsidePC_RB" & strCaseSuffix)
        Case 19
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "BtmEdgePC_RB" & strCaseSuffix)
           ' bIsBottomEdge = True
        Case 20
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "BtmEdgeOutsidePC_RB" & strCaseSuffix)
        Case 21
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "BtmFacePC_RB" & strCaseSuffix)
        Case 22
            If (bPenetratesWeb And Not HasRightWeb(oBoundedPort.Connectable)) Or _
                (Not bPenetratesWeb And Not bIsBottomFlange) Then
                Exit Sub
            ElseIf bPenetratesWeb And HasRightWeb(oBoundedPort.Connectable) And Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                Exit Sub
            End If
                strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, "BtmPC_RB" & strCaseSuffix)
        
    End Select

    If strPCAttribute <> vbNullString Then
        bIsNeeded = True
    End If
    
    If oMD.dispid = 4 Or oMD.dispid = 8 Or oMD.dispid = 15 Or oMD.dispid = 19 Then
        
        Dim dInsideOverlap As Double
        Dim dOutsideOverlap As Double
        
        GetEdgeOverlapAndClearance oAppConnection, bIsBottomEdge, bIsBottomFlange, dInsideOverlap, dOutsideOverlap
        
        If dInsideOverlap < 0.001 And dOutsideOverlap < 0.001 Then
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CMConstructCenterCutPC
' Description :
'
'*********************************************************************************************
Public Function CMConstructCenterCutPC(ByVal oMD As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef oObject As Object)
    
    Const METHOD = m_DefinitionProgid & "::CMConstructCenterCutPC"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Creating Physical Connection " & "...oMD.dispid = " & Str(oMD.dispid)

    ' -------------------------------------------------------------------------
    ' For PC being created (by dispID), identify the bounded and bounding ports
    ' Get the AC attribute that specifies how the PC is to be defined
    ' -------------------------------------------------------------------------
    Dim eBoundingSubPort As JXSEC_CODE
    Dim eBoundedSubPort As JXSEC_CODE
    Dim strACAttribute As String

    ' -------------------------------
    ' Get the PC selection and filter
    ' -------------------------------
    Dim strSelection As String
    Dim strFilter As String


    strACAttribute = GetPCAttributeForCenterCut(oMD, eBoundingSubPort, eBoundedSubPort)
    GetPCSelection oMD, strACAttribute, strSelection, strFilter

    If strSelection = vbNullString Then
        strSelection = "TeeWeld"
    End If

    ' -------------
    ' Create the PC
    ' -------------
    Dim oEndCutDefCM As Object
    Set oEndCutDefCM = SP3DCreateObject(CUSTOMERID & "MbrEndCut.EndCutDefCM")
    Set oObject = oEndCutDefCM.CreateModify_PhysConns(oMD, pResourceManager, strSelection, eBoundingSubPort, eBoundedSubPort)
    
    ' ------------------------------
    ' Set filter progID on PC object
    ' ------------------------------
    AddPCFilter oObject, strFilter

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Sub GetCornerToInsideParameters(oPRL As IJDParameterLogic, _
                                       bIsBottomCut As Boolean, _
                                       dCutDepth As Double, _
                                       dFlangeThicknessOrOffset As Double, _
                                       dRadius As Double, _
                                       Optional bIsBottomFlange As Boolean)

    Const METHOD = m_DefinitionProgid & "::GetCornerToInsideParameters"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    
    sMsg = "Calculating parameters for corner-to-inside end cut feature"
        
    ' --------------------------------------
    ' Determine if this is the bottom flange
    ' --------------------------------------
    Dim oParmCM As Object
    Set oParmCM = SP3DCreateObject(CUSTOMERID & "MbrEndCut.EndCutParmCM")
    
    bIsBottomFlange = oParmCM.IsBottomFlange(oPRL)
            
    ' ------------------------
    ' Define the cutting depth
    ' ------------------------
    dCutDepth = EndCut_GetCutDepth(oPRL)
    
    ' ---------------------------------------------------------
    ' Set the flange thickness (web cut) or offset (flange cut)
    ' ---------------------------------------------------------
    Dim oFeature As IJStructFeature
    Dim eFeatureType As StructFeatureTypes
    Set oFeature = oPRL.SmartOccurrence
    eFeatureType = oFeature.get_StructFeatureType
    
    If eFeatureType = SF_WebCut Then
        dFlangeThicknessOrOffset = oParmCM.GetWebOrFlangeThickness(oPRL, bIsBottomCut)
    Else
        dFlangeThicknessOrOffset = oParmCM.GetFlangeOffset(oPRL)
    End If
        
    ' -----------------------------
    ' Get the overlap and clearance
    ' -----------------------------
    Dim dOverlap As Double
    Dim dOutsideClearance As Double
    Dim dInsideClearance As Double
    Dim dOverall As Double
    
    GetEdgeOverlapAndClearance oPRL.SmartOccurrence, bIsBottomCut, bIsBottomFlange, dRadius
         
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
                                          
Public Function GetFlangeWidth(oPart As Object) As Double

    Const METHOD = m_DefinitionProgid & "::GetFlangeWidth"
    
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Calculating flange width"

    If TypeOf oPart Is ISPSMemberPartCommon Then
        Dim oBoundedMemberPart As New StructDetailObjects.MemberPart
        Set oBoundedMemberPart.object = oPart
        GetFlangeWidth = oBoundedMemberPart.FlangeLength
    ElseIf TypeOf oPart Is IJProfile Then
        Dim oBoundedProfilePart As New StructDetailObjects.ProfilePart
        Set oBoundedProfilePart.object = oPart
        GetFlangeWidth = oBoundedProfilePart.FlangeLength
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Sub GetDefaultPCClassAndBoundingEdge(ByVal oMD As IJDMemberDescription, strClass As String, eBoundingEdgeID As JXSEC_CODE)
                                      
    Const METHOD = m_DefinitionProgid & "::GetDefaultPCClass"

    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Determining default PC class for non-penetrated connection"

    ' --------------------------------------------------------
    ' Find out if this is web cut, top flange or bottom flange
    ' --------------------------------------------------------
    ' Once capability is added, check between left and right webs
    Dim sBottomFlange As String
    Dim isBottom As Boolean
    isBottom = False
    
    GetSelectorAnswer oMD, "BottomFlange", sBottomFlange
    If sBottomFlange = "Yes" Then
        isBottom = True
    End If
    
    ' ----------------------------------------------------------------------
    ' If the bounding object is a member, find out which edge is intersected
    ' ----------------------------------------------------------------------
    Dim oFeature As IJStructFeature
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
        
    If Not TypeOf oMD.CAO Is IJStructFeature Then
        Exit Sub
    End If
    
    Set oFeature = oMD.CAO

    If oFeature.get_StructFeatureType = SF_FlangeCut Then
        Dim oSDO_FlangeCut As New StructDetailObjects.FlangeCut
        Set oSDO_FlangeCut.object = oMD.CAO
        Set oSDO_WebCut.object = oSDO_FlangeCut.WebCut
    Else
        Set oSDO_WebCut.object = oFeature
    End If
    
    Dim eBoundingEdge As eBounding_Edge
    
    eBoundingEdge = None
    eBoundingEdgeID = -1
    
    If TypeOf oSDO_WebCut.Bounding Is ISPSMemberPartPrismatic Then
        If Not IsTubularMember(oSDO_WebCut.Bounding) Then
            GetNonPenetratedIntersectedEdge oSDO_WebCut.object, _
                                            oSDO_WebCut.BoundingPort, _
                                            oSDO_WebCut.BoundedPort, _
                                            eBoundingEdge, _
                                            eBoundingEdgeID, _
                                            isBottom
        End If
    End If
    
    ' -----------------------------------------------------------
    ' Create a Tee or Butt weld, depending on bounding edge width
    ' -----------------------------------------------------------
    If eBoundingEdge = Bottom_Flange_Right Or eBoundingEdge = Top_Flange_Right Then
        strClass = "ButtWeld"
    ElseIf Not (eBoundingEdge = Above Or eBoundingEdge = Below Or eBoundingEdge = None) Then
        strClass = "TeeWeld"
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Sub AddPCFilter(oPC As Object, strFilter As String)

    Const METHOD = m_DefinitionProgid & "::AddPCFilter"

    On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "Adding PC filter"

    ' ------------------------------------------
    ' If defined, set filter progID on PC object
    ' ------------------------------------------
    If Not strFilter = vbNullString Then
        
        Dim oAttributes As IJDAttributes
        Set oAttributes = oPC
    
        Dim oAttributesCol As IJDAttributesCol
        
        On Error Resume Next
        Set oAttributesCol = oAttributes.CollectionOfAttributes("IJUASelectionFilter")
        Err.Clear
        On Error GoTo ErrorHandler
        
        If Not oAttributesCol Is Nothing Then
            oAttributesCol.Item("FilterProgID").value = strFilter
        End If
    
        Set oAttributesCol = Nothing
        Set oAttributes = Nothing
    
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Function GetWebNonPenPCAttribute(oMD As IJDMemberDescription, bSplitOrMiter As Boolean) As String

    Const METHOD = m_DefinitionProgid & "::GetWebNonPenPCAttribute"

    On Error GoTo ErrorHandler

    Dim sMsg As String
    sMsg = "Getting the name of the attribute that contains a string specifying the type of PC to be created"
    
    ' -----------------------------------------
    ' Determine if this is the let or right web
    ' -----------------------------------------
    ' Future use
    Dim bIsRightWeb As Boolean
    bIsRightWeb = False
    
    ' -------------------------------------
    ' Determine if this is Case A or Case B
    ' -------------------------------------
    Dim isCaseA As Boolean
    Dim oCustomMethods As New CustomMethods
    Dim strCaseSuffix As String
    
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName oMD.CAO, , oAppConnection
    
    isCaseA = oCustomMethods.IsBoundingCaseA(oMD.CAO, strCaseSuffix)
    
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    GetAssemblyConnectionInputs oAppConnection, oBoundedPort, oBoundingPort
    
    Dim oBdedPort As IJPort
    Dim oBdingPort As IJPort
    GetBoundingAndBounded oMD.CAO, oBdingPort, oBdedPort
    
    ' -------------------------------
    ' Determine appropriate attribute
    ' -------------------------------
    Dim sPCAttr As String
    
    If bSplitOrMiter Then
        If IsEndToEndPrimary(oMD) Then
            If bIsRightWeb And ConsiderRightWebAttributes(oBdedPort.Connectable) Then
                sPCAttr = "RightWebPC"
            Else
                sPCAttr = "LeftWebPC"
            End If
        End If
    ElseIf (TypeOf oBdingPort.Connectable Is IJPlate) Or (TypeOf oBdingPort.Connectable Is SPSSlabEntity) Or (TypeOf oBdingPort.Connectable Is SPSWallPart) Then
            If bIsRightWeb And ConsiderRightWebAttributes(oBdedPort.Connectable) Then
                sPCAttr = "RightWebPC"
            Else
                sPCAttr = "LeftWebPC"
            End If
    Else
        If bIsRightWeb Then
            If Not ConsiderRightWebAttributes(oBoundedPort.Connectable) Then
                sPCAttr = "NPRightWebPC" & strCaseSuffix
            Else
                'Not supported for now
            End If
            
        Else
            sPCAttr = "NPLeftWebPC" & strCaseSuffix
        End If
    End If
    
    ' ----------------------------------------
    ' Check if the AC indicates a PC is needed
    ' ----------------------------------------
    GetWebNonPenPCAttribute = sPCAttr
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Function GetFlangeNonPenPCAttribute(oMD As IJDMemberDescription, bSplitOrMiter As Boolean) As String

    Const METHOD = m_DefinitionProgid & "::GetFlangeNonPenPCAttribute"

    Dim sMsg As String
    sMsg = "Getting the name of the attribute that contains a string specifying the type of PC to be created"

    On Error GoTo ErrorHandler

    ' ---------------------------------------------
    ' Determine if this is the top or bottom flange
    ' ---------------------------------------------
    Dim sBottomFlange As String
    GetSelectorAnswer oMD.CAO, "BottomFlange", sBottomFlange

    ' -------------------------------------
    ' Determine if this is Case A or Case B
    ' -------------------------------------
    Dim isCaseA As Boolean
    Dim oCustomMethods As New CustomMethods
    Dim strCaseSuffix As String
    
    isCaseA = oCustomMethods.IsBoundingCaseA(oMD.CAO, strCaseSuffix)

    Dim oBdedPort As IJPort
    Dim oBdingPort As IJPort
    GetBoundingAndBounded oMD.CAO, oBdingPort, oBdedPort
    ' -------------------------------
    ' Determine appropriate attribute
    ' -------------------------------
    Dim sPCAttr As String
    If bSplitOrMiter Then
        If IsEndToEndShortBox(oMD) Then
            If sBottomFlange = gsYes Then
                sPCAttr = "NPBottomFlangePC" & strCaseSuffix
            Else
                sPCAttr = "NPTopFlangePC" & strCaseSuffix
            End If
        ElseIf IsEndToEndPrimary(oMD) Then
            If sBottomFlange = gsYes Then
                sPCAttr = "BottomFlangePC"
            Else
                sPCAttr = "TopFlangePC"
            End If
        End If
    ElseIf (TypeOf oBdingPort.Connectable Is IJPlate) Or (TypeOf oBdingPort.Connectable Is SPSSlabEntity) Or (TypeOf oBdingPort.Connectable Is SPSWallPart) Then
            If sBottomFlange = gsYes Then
                sPCAttr = "BottomFlangePC"
            Else
                sPCAttr = "TopFlangePC"
            End If
    Else
        If sBottomFlange = gsYes Then
            sPCAttr = "NPBottomFlangePC" & strCaseSuffix
        Else
            sPCAttr = "NPTopFlangePC" & strCaseSuffix
        End If
    End If
    
    ' ---------------------------------------
    ' Check if the AC indicate a PC is needed
    ' ---------------------------------------
    GetFlangeNonPenPCAttribute = sPCAttr
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Function CMConstructNonPenPC(ByVal oMD As IJDMemberDescription, _
                                    ByVal oResourceManager As IUnknown, _
                                    bIsWeb As Boolean, _
                                    bSplitOrMiter As Boolean) As Object

    Const METHOD = m_DefinitionProgid & "::CMConstructNonPenPC"
    
    On Error GoTo ErrorHandler
    
    Set CMConstructNonPenPC = Nothing
    
    Dim sMsg As String
    sMsg = "Creating non-penetrated physical connection"
    
    Dim eBoundingEdgeID As JXSEC_CODE
    Dim strSelection As String
    
    ' ------------------------
    ' If a split or miter case
    ' ------------------------
    If bSplitOrMiter Then
        ' -----------------
        ' Set default class
        ' -----------------
        strSelection = "ButtWeld"
        
        ' ---------------------
        ' Get the bounding port
        ' ---------------------
        If Not TypeOf oMD.CAO Is IJStructFeature Then
            sMsg = "Unexpected object"
            GoTo ErrorHandler
        End If
        
        Dim oBoundingPort As IJPort
        Dim oFeature As IJStructFeature
        Set oFeature = oMD.CAO
        
        If oFeature.get_StructFeatureType = SF_FlangeCut Then
            Dim oSDOFlange As New StructDetailObjects.FlangeCut
            Set oSDOFlange.object = oMD.CAO
            Set oBoundingPort = oSDOFlange.BoundingPort
        ElseIf oFeature.get_StructFeatureType = SF_WebCut Then
            Dim oSDOWeb As New StructDetailObjects.WebCut
            Set oSDOWeb.object = oMD.CAO
            Set oBoundingPort = oSDOWeb.BoundingPort
        Else
            sMsg = "Unexpected feature type"
            GoTo ErrorHandler
        End If

        ' ------------------------------------------------------
        ' Determine if at base or offset end of connected member
        ' ------------------------------------------------------
        Dim oBounding_SplitAxisPort As ISPSSplitAxisPort
        Dim eAxisPortIndex As SPSMemberAxisPortIndex
        
        If TypeOf oBoundingPort Is ISPSSplitAxisPort Then
            Set oBounding_SplitAxisPort = oBoundingPort
            eAxisPortIndex = oBounding_SplitAxisPort.PortIndex
                
            If eAxisPortIndex = SPSMemberAxisStart Then
               eBoundingEdgeID = CTX_BASE
            ElseIf eAxisPortIndex = SPSMemberAxisEnd Then
               eBoundingEdgeID = CTX_OFFSET
            Else
                sMsg = "Unexpected port type"
                GoTo ErrorHandler
            End If
        Else
            sMsg = "Unexpected port type"
            GoTo ErrorHandler
        End If
    ' ------------------------
    ' If axis-along or box cut
    ' ------------------------
    Else
        ' ------------------------------------------------------------
        ' Determine weld class based on end cut type and bounding edge
        ' ------------------------------------------------------------
        GetDefaultPCClassAndBoundingEdge oMD, strSelection, eBoundingEdgeID
    End If

    ' -----------------------------------------------
    ' Override with selection specified by AC, if set
    ' -----------------------------------------------
    Dim strPCAttribute As String
    Dim oCM As New CustomMethods
    
    Dim strACAttribute As String
    Dim strPCSelection As String
    Dim strFilter As String
    
    If bIsWeb Then
        strACAttribute = oCM.GetWebNonPenPCAttribute(oMD, bSplitOrMiter)
    Else
        strACAttribute = oCM.GetFlangeNonPenPCAttribute(oMD, bSplitOrMiter)
    End If
    
    GetPCSelection oMD, strACAttribute, strPCSelection, strFilter

    If strPCSelection <> vbNullString Then
        strSelection = strPCSelection
    End If
    
    ' -------------
    ' Create the PC
    ' -------------
    If bSplitOrMiter Then
        Set CMConstructNonPenPC = Create_PhysConn(oMD, oResourceManager, strSelection, eBoundingEdgeID, 8193)
    Else
        Dim oEndCutDefCM As Object
        Set oEndCutDefCM = SP3DCreateObject(CUSTOMERID & "MbrEndCut.EndCutDefCM")

        Set CMConstructNonPenPC = oEndCutDefCM.CreateModify_PhysConns(oMD, oResourceManager, strSelection, eBoundingEdgeID, 8193)
    End If
    
    ' ------------------------------
    ' Set filter progID on PC object
    ' ------------------------------
    AddPCFilter CMConstructNonPenPC, strFilter

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

' Returns specific PC selection specified by an AC attribute
' attribute = 1 --> strACSelection = vbNullString ; strFilter = vbNullString
' attribute = "class name" --> strACSelection = "class name" ; strFilter = vbNullString
' attribute = "class name|filter name" --> strACSelection = "class name" ; strFilter = "filter name"

Public Sub GetPCSelection(oMD As IJDMemberDescription, strACAttribute As String, strPCSelection As String, strFilter As String)

    Const METHOD = m_DefinitionProgid & "::GetPCSelection"

    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    sMsg = "Decoding the PC string into (optional) class and filter"
    
    ' ----------------------------------
    ' Get the attribute on the parent AC
    ' ----------------------------------
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName oMD.CAO, , oAppConnection
    
    Dim oACDef As New ACDef
    
    Dim strPCAttribute As String
    strPCAttribute = oACDef.GetStandardACAttribute(oAppConnection, strACAttribute)
    
    ' ------------------------------------------------------------------------
    ' If the string includes a pipe, decompose into class name and filter name
    ' ------------------------------------------------------------------------
    Dim pipePosition As Long
    pipePosition = InStr(1, strPCAttribute, "|", vbTextCompare)

    If pipePosition > 0 And pipePosition < Len(strPCAttribute) Then
        strPCSelection = Left(strPCAttribute, pipePosition - 1)
        strFilter = Right(strPCAttribute, Len(strPCAttribute) - pipePosition)
    ElseIf strPCAttribute <> "1" And strPCAttribute <> vbNullString Then
        strPCSelection = strPCAttribute
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Function IsEndToEndPrimary(oMD As IJDMemberDescription) As Boolean

    Const METHOD = m_DefinitionProgid & "::IsEndToEndPrimary"

    Dim sMsg As String
    sMsg = "Determining if member part is primary in an end-to-end connection"

    IsEndToEndPrimary = False
    
    ' ----------------------------
    ' Get the end cut bounded port
    ' ----------------------------
    If Not TypeOf oMD.CAO Is IJStructFeature Then
        sMsg = "Unexpected object"
        GoTo ErrorHandler
    End If
    
    Dim oBoundedPort As IJPort
    Dim oFeature As IJStructFeature
    Set oFeature = oMD.CAO
    
    If oFeature.get_StructFeatureType = SF_FlangeCut Then
        Dim oSDOFlange As New StructDetailObjects.FlangeCut
        Set oSDOFlange.object = oMD.CAO
        Set oBoundedPort = oSDOFlange.BoundedPort
    ElseIf oFeature.get_StructFeatureType = SF_WebCut Then
        Dim oSDOWeb As New StructDetailObjects.WebCut
        Set oSDOWeb.object = oMD.CAO
        Set oBoundedPort = oSDOWeb.BoundedPort
    Else
        sMsg = "Unexpected feature type"
        GoTo ErrorHandler
    End If

    ' -----------------------
    ' Get the AC bounded port
    ' -----------------------
    Dim oACObject As IJAppConnection
    Dim sACItemName As String
    
    Parent_SmartItemName oMD.CAO, sACItemName, oACObject
     
    Dim oBddPort As IJPort
    Dim oBdgPort As IJPort
    GetAssemblyConnectionInputs oACObject, oBddPort, oBdgPort
    
    If oBddPort Is oBoundedPort Then
        IsEndToEndPrimary = True
    End If
    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Public Sub CopyOccurrenceAttributes(oMD As IJDMemberDescription, strInterface As String)

    Const METHOD = "CopyOccurrenceAttributes"
    
    Dim sMsg As String

    Dim oOccAttribs As IJDAttributes
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = oMD.CAO
    Set oOccAttribs = oSmartOcc
    Set oAttribs = oSmartOcc.ItemObject
    
    If Not IsSOOverridden(oOccAttribs.CollectionOfAttributes(strInterface)) Then
        Dim oOccurraceAttr As IJDAttributesCol
        Dim oItemAttr As IJDAttributesCol
        
        Set oOccurraceAttr = oOccAttribs.CollectionOfAttributes(strInterface)
        Set oItemAttr = oAttribs.CollectionOfAttributes(strInterface)
        CopyValuesToSOFromItem oOccurraceAttr, oItemAttr
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Private Function IsSOOverridden(attrCol As CollectionProxy) As Boolean

    Const METHOD = "IsSOOverridden"
    
    Dim sMsg As String
    
    On Error GoTo ErrorHandler

    IsSOOverridden = False

    Dim i As Integer
    Dim oAttr As IJDAttribute
  
    For i = 1 To attrCol.Count
        Set oAttr = attrCol.Item(i)
        If Not IsEmpty(oAttr.value) Then
            IsSOOverridden = True
            Set oAttr = Nothing
            Exit For
        End If
        Set oAttr = Nothing
    Next i
 
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

Private Sub CopyValuesToSOFromItem(soCol As CollectionProxy, itmCol As CollectionProxy)

    Const METHOD = "CopyValuesToSOFromItem"
    
    Dim sMsg As String
    
    On Error GoTo ErrorHandler
    
    Dim i As Integer

    If soCol.Count <> itmCol.Count Then
        Exit Sub
    End If
    
    For i = 1 To soCol.Count
        If soCol.Item(i).AttributeInfo.Name = itmCol.Item(i).AttributeInfo.Name Then
            soCol.Item(i).value = itmCol.Item(i).value
        End If
    Next i
     
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

Public Function IsEndToEndShortBox(oMD As IJDMemberDescription) As Boolean

    Const METHOD = m_DefinitionProgid & "::IsEndToEndShortBox"

    Dim sMsg As String
    sMsg = "Determining if member part is primary in an end-to-end connection"

    IsEndToEndShortBox = False
    
    ' ----------------------------
    ' Get the end cut bounded port
    ' ----------------------------
    If Not TypeOf oMD.CAO Is IJStructFeature Then
        sMsg = "Unexpected object"
        GoTo ErrorHandler
    End If
    
    Dim oBoundedPort As IJPort
    Dim oFeature As IJStructFeature
    Set oFeature = oMD.CAO
    
    If oFeature.get_StructFeatureType = SF_FlangeCut Then
        Dim oSDOFlange As New StructDetailObjects.FlangeCut
        Set oSDOFlange.object = oMD.CAO
        Set oBoundedPort = oSDOFlange.BoundedPort
    ElseIf oFeature.get_StructFeatureType = SF_WebCut Then
        Dim oSDOWeb As New StructDetailObjects.WebCut
        Set oSDOWeb.object = oMD.CAO
        Set oBoundedPort = oSDOWeb.BoundedPort
    Else
        sMsg = "Unexpected feature type"
        GoTo ErrorHandler
    End If

    ' -----------------------
    ' Get the AC bounded port
    ' -----------------------
    Dim oACObject As IJAppConnection
    Dim sACItemName As String
    Dim IsboxCutAC As Boolean
    Dim oACDef As New ACDef
    
    Parent_SmartItemName oMD.CAO, sACItemName, oACObject
    IsboxCutAC = oACDef.IsBoxCut(oACObject)
    
    Dim oBddPort As IJPort
    Dim oBdgPort As IJPort

    GetAssemblyConnectionInputs oACObject, oBddPort, oBdgPort
    
    If (oBddPort Is oBoundedPort) And IsboxCutAC Then
        IsEndToEndShortBox = True
    End If
    
Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function
'*********************************************************************************************
' Method      : CM_ConstructPhysConns
' Description :
' Bounding Port - lateral , Bounded port as 8193
'*********************************************************************************************
Public Function CM_ConstructTubePhysConn(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
    Const METHOD = m_DefinitionProgid & "::CM_ConstructPhysConns"
    On Error GoTo ErrorHandler

    Dim sMsg As String
   
    ' Get Bounding Part's Port
    Dim eBoundingPort As eUSER_CTX_FLAGS
    Dim eBoundedSubPort As JXSEC_CODE
    
    Dim oSDO_WebCut As New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberDescription.CAO

    Dim oBUMember As ISPSDesignedMember
    Dim bBuiltup As Boolean
    
    IsFromBuiltUpMember oSDO_WebCut.Bounding, bBuiltup, oBUMember
    If bBuiltup Then
        eBoundingPort = CTX_BASE
    Else
        eBoundingPort = CTX_LATERAL
    End If
    
    eBoundedSubPort = 8193

    Dim boundingEdgeId As eBounding_Edge
 
    boundingEdgeId = None
    
    Dim strPCSelection As String
    Dim strSelection As String
    Dim strFilter As String

    GetPCSelection pMemberDescription, "NPLeftWebPC_A", strPCSelection, strFilter
    
    If strSelection = vbNullString Then
        strSelection = "TeeWeld"
    End If

    If strPCSelection <> vbNullString Then
        strSelection = strPCSelection
    End If
    
    Set pObject = Create_PhysConn(pMemberDescription, pResourceManager, _
                                   strSelection, eBoundingPort, eBoundedSubPort)
         
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function


Public Function GetPCAttributeForCenterCut(oMD As IJDMemberDescription, eBoundingSubPort As JXSEC_CODE, eBoundedSubPort As JXSEC_CODE) As String

  Const METHOD = m_DefinitionProgid & "::GetPCAttributeForCenterCut"
  On Error GoTo ErrorHandler

  Dim sMsg As String

    ' ------------------------------
    ' Get bounding and bounded ports
    ' ------------------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    
    GetBoundingAndBounded oMD.CAO, oBoundingPort, oBoundedPort

    ' --------------------
    ' Get the edge mapping
    ' --------------------
    Dim oAppConnection As Object
    AssemblyConnection_SmartItemName oMD.CAO, , oAppConnection
    
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    Set oEdgeMapColl = GetEdgeMap(oAppConnection, oBoundingPort, oBoundedPort)
    
    ' -------------------------------------
    ' Determine if this is Case A or Case B
    ' -------------------------------------
    Dim isCaseA As Boolean
    Dim oCustomMethods As New CustomMethods
    Dim strCaseSuffix As String
    
    isCaseA = oCustomMethods.IsBoundingCaseA(oAppConnection, strCaseSuffix)
    
    ' -------------------------------------------------------------------------
    ' For PC being created (by dispID), identify the bounded and bounding ports
    ' Get the AC attribute that specifies how the PC is to be defined
    ' -------------------------------------------------------------------------
    Dim strACAttribute As String
    
    Dim lDispId As Long
    lDispId = oMD.dispid
    
    Select Case lDispId
    
    Case 1
        eBoundedSubPort = JXSEC_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP))
        strACAttribute = "TopPC_LT" & strCaseSuffix
    Case 2
        eBoundedSubPort = JXSEC_WEB_RIGHT_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT_TOP))
        strACAttribute = "TopFacePC_LT" & strCaseSuffix
    Case 3
        eBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_TOP))
        strACAttribute = "TopEdgeOutsidePC_LT" & strCaseSuffix
    Case 4
        eBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT))
        strACAttribute = "TopEdgePC_LT" & strCaseSuffix
    Case 5
        eBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_BOTTOM))
        strACAttribute = "TopEdgeInsidePC_LT" & strCaseSuffix
    Case 6
        eBoundedSubPort = JXSEC_WEB_RIGHT
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT))
        strACAttribute = "FacePC_LT" & strCaseSuffix
    Case 7
        eBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_TOP))
        strACAttribute = "BtmEdgeInsidePC_LT" & strCaseSuffix
    Case 8
        eBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT))
        strACAttribute = "BtmEdgePC_LT" & strCaseSuffix
    Case 9
        eBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM))
        strACAttribute = "BtmEdgeOutsidePC_LT" & strCaseSuffix
    Case 10
        eBoundedSubPort = JXSEC_WEB_RIGHT_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT_BOTTOM))
        strACAttribute = "BtmFacePC_LT" & strCaseSuffix
    Case 11
        eBoundedSubPort = JXSEC_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM))
        strACAttribute = "BtmPC_LT" & strCaseSuffix
    Case 12
        eBoundedSubPort = JXSEC_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP))
        strACAttribute = "TopPC_RB" & strCaseSuffix
    Case 13
        eBoundedSubPort = JXSEC_WEB_RIGHT_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT_TOP))
        strACAttribute = "TopFacePC_RB" & strCaseSuffix
    Case 14
        eBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_TOP))
        strACAttribute = "TopEdgeOutsidePC_RB" & strCaseSuffix
    Case 15
        eBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT))
        strACAttribute = "TopEdgePC_RB" & strCaseSuffix
    Case 16
        eBoundedSubPort = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_TOP_FLANGE_RIGHT_BOTTOM))
        strACAttribute = "TopEdgeInsidePC_RB" & strCaseSuffix
    Case 17
        eBoundedSubPort = JXSEC_WEB_RIGHT
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT))
        strACAttribute = "FacePC_RB" & strCaseSuffix
    Case 18
        eBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_TOP))
        strACAttribute = "BtmEdgeInsidePC_RB" & strCaseSuffix
    Case 19
        eBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT))
        strACAttribute = "BtmEdgePC_RB" & strCaseSuffix
    Case 20
        eBoundedSubPort = JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM))
        strACAttribute = "BtmEdgeOutsidePC_RB" & strCaseSuffix
    Case 21
        eBoundedSubPort = JXSEC_WEB_RIGHT_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT_BOTTOM))
        strACAttribute = "BtmFacePC_RB" & strCaseSuffix
    Case 22
        eBoundedSubPort = JXSEC_BOTTOM
        eBoundingSubPort = oEdgeMapColl.Item(CStr(JXSEC_BOTTOM))
        strACAttribute = "BtmPC_RB" & strCaseSuffix
    End Select
        
    GetPCAttributeForCenterCut = strACAttribute
    
  Exit Function
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function
