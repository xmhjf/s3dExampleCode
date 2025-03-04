VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MbrAxisCenterDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2011-12, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMMbrAC
'  File        : MbrAxisCenterDef.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    08/APR/2011 - Created
'    12/Sep/2011 - mpulikol
'           DI-CP-200263 Improve performance by caching measurement symbol results
'    22/Sep/2011 - svsmylav TR-202526:
'                  Top/Bottom PC conditionals are corrected to check for InsideCornerToOutside option (it was missing).
'   22/Nov/2011 - pnalugol
'           CR-CP-205624 - Removing functionality to defer corner feature from webcut to flange cut
'   27/APr/2012 - hkvijaya

'    27/Jul/2012 - svsmylav
'      CR-216426: Removed instances of class 'SMMbrAC.MbrACDefCM' and allied method calls.
'           Instead, a new method call 'IsCornerFeatureNeededForShapeAtEdge' is used now.
'    7/4/ 2013 - gkharrin\dsmamidi
'               DI-CP-235071 Improve performance by caching edge mapping data
'               Reduced the number of calls to edgemapping rule
'               Replaced GetEdgeMapping with GetEdgeMap, From GetEdgeMap,getting edgemap data if already cached, if not call the edgemapping rule to get the same
'*********************************************************************************************
Const m_sClassName As String = "MbrAxisCenterDef"
Const m_FamilyProgid As String = ""
Const m_DefinitionProgid As String = m_sProjectName + "." + m_sClassName
Const m_DefinitionName As String = m_DefinitionProgid
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

Implements IJDUserSymbolServices

'*********************************************************************************************
' Method      : ItemInputs
' Description : List any graphic Inputs that the Definition has here
'
'*********************************************************************************************
Public Sub ItemInputs(pIH As IJDInputsHelper)
    Const METHOD = m_DefinitionProgid & "::ItemInputs"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining Definition Inputs"

    pIH.SetInput INPUT_BOUNDING
    pIH.SetInput INPUT_BOUNDED
    
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
    Const METHOD = m_DefinitionProgid & "::ItemAggregator"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Defining ItemAggregator"
    
    pAD.UserTypeClsid = CA_WEBCUT
    pAD.AggregatorClsid = CA_AGGREGATE     ' CStructFeature
    pAD.SetCMFinalConstruct CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_FinalConstructEndCut"
    pAD.SetCMMigrate CUSTOMERID & "MbrEndCut.EndCutDefCM", "CM_MigrateAggregator"
    
    Dim pPDs As IJDPropertyDescriptions
    Set pPDs = pAD
    pPDs.RemoveAll
    
    ' Set the FlangeCutting Behaviour
    sMsg = "Setting Flange Cutting Behaviour property."
    pPDs.AddProperty "CuttingBehavior", 1, "IJUAMbrEndCuts", "CM_SetFlangeCuttingBehaviour", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : ItemMembers
' Description : List all the Item members that this SmartClass can create
'
'*********************************************************************************************
Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    Const METHOD = m_DefinitionProgid & "::ItemMembers"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sDispId As String
    
    Dim lDispId As Long
    Dim oMemDesc As IJDMemberDescription
    Dim oPropDesc As IJDPropertyDescriptions
    
    pMDs.RemoveAll  ' Remove all the previous cached member descriptions(best practice)
        
    '******************************************************************************
    'if any more item members are needed to be added please add only at the bottom
    '******************************************************************************
        
    'Creates a Physical Connection with the bounding Mapped Web Right Port(Face)
    lDispId = 1
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsFacePCNeeded"
    Set oMemDesc = Nothing
    
    'Creates a Physical Connection with the bounding Mapped TopFlangeRight Port(Top Edge)
    lDispId = 2
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopEdgePCNeeded"
    Set oMemDesc = Nothing
    
    'Creates a Physical Connection with the bounding Mapped BottomFlangeRight Port(Bottom Edge)
    lDispId = 3
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBottomEdgePCNeeded"
    Set oMemDesc = Nothing
    
    'Creates a Physical Connection with the bounding Mapped TopFlangeRightBottom Port(Top Edge Inside)
    lDispId = 4
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopEdgeInsidePCNeeded"
    Set oMemDesc = Nothing
    
    'Creates a Physical Connection with the bounding Mapped BottomFlangeRightTop Port(Bottom Edge Inside)
    lDispId = 5
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBottomEdgeInsidePCNeeded"
    Set oMemDesc = Nothing
    
    'Creates a Physical Connection with the bounding Mapped Top Port(Top)
    lDispId = 6
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopPCNeeded"
    Set oMemDesc = Nothing
    
    'Creates a Physical Connection with the bounding Mapped Bottom Port(Bottom)
    lDispId = 7
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBottomPCNeeded"
    Set oMemDesc = Nothing
    
    'Creates a Physical Connection with the bounding Mapped TopFlangeRightTop Port(Top Edge Outside)
    lDispId = 8
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopEdgeOutsidePCNeeded"
    Set oMemDesc = Nothing
    
    'Creates a Physical Connection with the bounding Mapped BottomFlangeRightBottom Port(Bottom Edge Outside)
    lDispId = 9
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBottomEdgeOutsidePCNeeded"
    Set oMemDesc = Nothing
    
    'Creates a Physical Connection with the bounding Mapped WebRightTop Port(Face Top)
    lDispId = 10
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopFacePCNeeded"
    Set oMemDesc = Nothing
    
    'Creates a Physical Connection with the bounding Mapped WebRightBottom Port(Face Bottom)
    lDispId = 11
    sDispId = "MbrAxis_EndCut_PC_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructPhysConns", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBottomFacePCNeeded"
    Set oMemDesc = Nothing
    
    'ItemMembers for the Corner Features
    'Creates a face top corner
    lDispId = 12
    sDispId = "CF_FaceTopCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFaceTopCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsFaceTopCornerNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputeCF1", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID & "MbrEndCut.EndCutDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing

    'Creates a face bottom corner
    lDispId = 13
    sDispId = "CF_FaceBottomCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructFaceBottomCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsFaceBottomCornerNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputeCF2", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID & "MbrEndCut.EndCutDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Creates a top edge inside corner
    lDispId = 14
    sDispId = "CF_TopEdgeInsideCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructTopEdgeInsideCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsTopEdgeInsideCornerNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputeCF3", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID & "MbrEndCut.EndCutDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    'Creates a bottom edge inside corner
    lDispId = 15
    sDispId = "CF_BottomEdgeInsideCorner_" & Trim(Str(lDispId))
    Set oMemDesc = pMDs.AddMember(sDispId, lDispId, "CMConstructBottomEdgeInsideCorner", CUSTOMERID & "MbrEndCut.EndCutDefCM")
    oMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CM_IsBottomEdgeInsideCornerNeeded"
    Set oPropDesc = oMemDesc
    oPropDesc.AddProperty "NeedToComputeCF4", 1, "IJSmartOccurrence", "CMNeedToCompute", CUSTOMERID & "MbrEndCut.EndCutDefCM"
    Set oPropDesc = Nothing
    Set oMemDesc = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Methods HERE
' *******************************************************************************************
'*********************************************************************************************
' Method      : CM_IsFacePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFacePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFacePCNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    Dim sMsg As String
    
    Dim sParentItemName As String
    Dim oParentObj As Object
    
    ' Get the Current Selection
    bIsNeeded = False
    
    Dim sTopShapeAtFace As String
    Dim sBtmShpAtFace As String
    Dim sShapeAtFace As String
    
    'For WebRight exists in all aliases, this PC is independent of bounding alias
    'check to see what the relative position is
    
     Parent_SmartItemName pMemberDescription.CAO, sParentItemName, oParentObj
     
    'ShapeAtFace question doesnt have any effect on PC creation.
        Select Case sParentItemName
            Case gsMbrAxisToEdge, gsMbrAxisToOnMember, gsStiffEndToMbrEdge, gsStiffEndToOnMember
                bIsNeeded = False
            Case gsMbrAxisToCenter, gsMbrAxisToFaceAndOutSideNoEdge, gsStiffEndToMbrCenter, gsStiffEndToMbrFaceAndOutSideNoEdge
                GetSelectorAnswer oParentObj, "TopShapeAtFace", sTopShapeAtFace
                GetSelectorAnswer oParentObj, "BottomShapeAtFace", sBtmShpAtFace
                If InStr(1, sTopShapeAtFace, "Offset", vbTextCompare) > 0 And InStr(1, sBtmShpAtFace, "Offset", vbTextCompare) > 0 Then
                    bIsNeeded = False
                Else
                    bIsNeeded = True
                End If
            Case gsMbrAxisToEdgeAndOutSide1Edge, gsMbrAxisToFaceAndEdge, gsMbrAxisToFaceAndOutSide1Edge, gsMbrAxisToOutSideAndOutSide1Edge, _
                 gsStiffEndToMbrEdgeAndOutSide1Edge, gsStiffEndToMbrFaceAndEdge, gsStiffEndToMbrFaceAndOutSide1Edge, gsStiffEndToMbrOutSideAndOutSide1Edge
                GetSelectorAnswer oParentObj, "ShapeAtFace", sShapeAtFace
                If InStr(1, sShapeAtFace, "Offset", vbTextCompare) > 0 Then
                    bIsNeeded = False
                Else
                    bIsNeeded = True
                End If
            Case gsMbrAxisToEdgeAndEdge, gsMbrAxisToEdgeAndOutSide2Edge, gsMbrAxisToOutSideAndOutSide2Edge, gsMbrAxisToOutSideAndOutSideNoEdge, _
                 gsStiffEndToMbrEdgeAndEdge, gsStiffEndToMbrEdgeAndOutSide2Edge, gsStiffEndToMbrOutSideAndOutSide2Edge, gsStiffEndToMbrOutSideAndOutSideNoEdge
                bIsNeeded = True
            Case Else
'                bIsNeeded = True
        End Select
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub


'*********************************************************************************************
' Method      : CM_IsTopEdgePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopEdgePCNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)

    Const METHOD = m_DefinitionProgid & "::CM_IsTopEdgePCNeeded"
    
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    
    bIsNeeded = IsEdgePCNeeded(oMD, False)
    Dim sParentItemName As String
    Dim oParentObj As Object
    
    Parent_SmartItemName oMD.CAO, sParentItemName, oParentObj
    Dim bCutBTWBraceAndBounded As Boolean
    Dim bCutBTWBraceAndBounding As Boolean
    'check whether it is struct assembly connection or not, because stiffenerAC is not of this type
    'braces only applicable for memberAC of type structassembly connection
    If Not oParentObj Is Nothing Then
        If TypeOf oParentObj Is StructAssemblyConnection Then
        'check whether the AC is between Brace and bounding or Bounded
        IsCutOnInsetBrace oParentObj, bCutBTWBraceAndBounded, bCutBTWBraceAndBounding
            If bCutBTWBraceAndBounding Then
                'if the AC is between brace  and bounding then PC is not needed
                bIsNeeded = False
                Exit Sub
            End If
        End If
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomEdgePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomEdgePCNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomEdgePCNeeded"
    
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
        
    Dim sParentItemName As String
    Dim oParentObj As Object
    
    Parent_SmartItemName oMD.CAO, sParentItemName, oParentObj
    Dim bCutBTWBraceAndBounded As Boolean
    Dim bCutBTWBraceAndBounding As Boolean
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As IJPort
    'check whether it is struct assembly connection or not, because stiffenerAC is not of this type
    'braces only applicable for memberAC of type structassembly connection
    If Not oParentObj Is Nothing Then
        If TypeOf oParentObj Is StructAssemblyConnection Then
        'check whether the AC is between Brace and bounding or Bounded
        IsCutOnInsetBrace oParentObj, bCutBTWBraceAndBounded, bCutBTWBraceAndBounding
        GetBoundingAndBounded oMD.CAO, oBoundingPort, oBoundedPort
            If bCutBTWBraceAndBounding Then
                Select Case LCase(sParentItemName)

                Case LCase(gsMbrAxisToFaceAndOutSide1Edge), LCase(gsMbrAxisToCenter), _
                            LCase(gsMbrAxisToFaceAndOutSideNoEdge)
                    bIsNeeded = False
                    Exit Sub
                Case LCase(gsMbrAxisToFaceAndEdge)
                    Dim lSectionAlias As Long
                    Dim bPenetratesWeb As Boolean
                    Dim oEdgeMapColl As JCmnShp_CollectionAlias
                    'Edge Mapping
                    Set oEdgeMapColl = New Collection
                    Set oEdgeMapColl = GetEdgeMap(oMD.CAO, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
                    If lSectionAlias = 1 Or lSectionAlias = 6 Then
                        bIsNeeded = False
                        Exit Sub
                    Else
                        'need to continue
                    End If
                End Select
            
            End If
        End If
    End If
    bIsNeeded = IsEdgePCNeeded(oMD, True)
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsTopEdgeInsidePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopEdgeInsidePCNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopEdgeInsidePCNeeded"
    
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
   
    bIsNeeded = True
        
    ' -----------------------------
    ' Get the shape type and answer
    ' -----------------------------
    Dim sTopAnswerCol As String
    Dim sTopShape As String

    GetMemberACTopAndBottomShape oMD.CAO, , , sTopAnswerCol, sTopShape
    
    ' ------------------------------------------------------
    ' If the Top shape is "At Face", then no PC is needed
    ' ------------------------------------------------------
    If sTopAnswerCol = gsShapeAtFaceCol Then
        bIsNeeded = False
        Exit Sub
    End If
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oMD.CAO, sParentItemName, oParentObj

    Dim bCutBTWBraceAndBounded As Boolean
    Dim bCutBTWBraceAndBounding As Boolean
    If Not oParentObj Is Nothing Then
        'checking whether the AC is for Member or stiffener
        If TypeOf oParentObj Is StructAssemblyConnection Then
        ' Checking whether the ac is between brace and bounding\bounded
        IsCutOnInsetBrace oParentObj, bCutBTWBraceAndBounded, bCutBTWBraceAndBounding
            If bCutBTWBraceAndBounding Then
                'anytime the top edge insidePC is not needed because topflange will cutoff when we create braces
                bIsNeeded = False
                Exit Sub
            End If
        End If
        
    End If
    ' ------------------------------------------------------------------------
    ' Otherwise, base answer on collection type, answer, and amount of overlap
    ' ------------------------------------------------------------------------
    If sTopAnswerCol = gsShapeAtEdgeCol Or sTopAnswerCol = gsShapeAtEdgeOverlapCol Then
        
        Select Case sTopShape
           
            Case gsFaceToCorner, gsFaceToEdge, gsFaceToFlange, gsFaceToOutside, gsFaceToInsideCorner, gsFaceToOutsideCorner
                bIsNeeded = False
        
        End Select
    Else
        bIsNeeded = False
    End If
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomEdgeInsidePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomEdgeInsidePCNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomEdgeInsidePCNeeded"
    
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
   
    bIsNeeded = True
        
    ' -----------------------------
    ' Get the shape type and answer
    ' -----------------------------
    Dim sBottomAnswerCol As String
    Dim sBottomShape As String

    GetMemberACTopAndBottomShape oMD.CAO, sBottomAnswerCol, sBottomShape
    Dim sParentItemName As String
    Dim oParentObj As Object
    Dim oBoundingPort As IJPort
    Dim oBoundedPort As IJPort
    Parent_SmartItemName oMD.CAO, sParentItemName, oParentObj
    Dim bCutBTWBraceAndBounded As Boolean
    Dim bCutBTWBraceAndBounding As Boolean
    
    If Not oParentObj Is Nothing Then
        GetBoundingAndBounded oMD.CAO, oBoundingPort, oBoundedPort
        'check whether it is struct assembly connection or not, because stiffenerAC is not of this type
        'braces only applicable for memberAC of type structassembly connection
        If TypeOf oParentObj Is StructAssemblyConnection Then
            'check whether it is AC between BRACE and bounding\bounded
            IsCutOnInsetBrace oParentObj, bCutBTWBraceAndBounded, bCutBTWBraceAndBounding
            If bCutBTWBraceAndBounding Then
                Select Case LCase(sParentItemName)

                    Case LCase(gsMbrAxisToFaceAndOutSide1Edge), LCase(gsMbrAxisToCenter), _
                                LCase(gsMbrAxisToFaceAndOutSideNoEdge)
                        bIsNeeded = False
                        Exit Sub
                    Case LCase(gsMbrAxisToFaceAndEdge)
                        Dim lSectionAlias As Long
                        Dim bPenetratesWeb As Boolean
                        Dim oEdgeMapColl As JCmnShp_CollectionAlias
                        'Edge Mapping
                        Set oEdgeMapColl = New Collection
                        Set oEdgeMapColl = GetEdgeMap(oMD.CAO, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
                        If lSectionAlias = 1 Or lSectionAlias = 6 Then
                            bIsNeeded = False
                            Exit Sub
                        Else
                            'need to continue
                        End If

                End Select
            End If
        End If
    End If
      
    
    ' ------------------------------------------------------
    ' If the bottom shape is "At Face", then no PC is needed
    ' ------------------------------------------------------
    If sBottomAnswerCol = gsShapeAtFaceCol Then
        bIsNeeded = False
        Exit Sub
    End If
    
    ' ------------------------------------------------------------------------
    ' Otherwise, base answer on collection type, answer, and amount of overlap
    ' ------------------------------------------------------------------------
    If sBottomAnswerCol = gsShapeAtEdgeCol Or sBottomAnswerCol = gsShapeAtEdgeOverlapCol Then
        
        Select Case sBottomShape
           
            Case gsFaceToCorner, gsFaceToEdge, gsFaceToFlange, gsFaceToOutside, gsFaceToInsideCorner, gsFaceToOutsideCorner
                bIsNeeded = False
        
        End Select
    Else
        bIsNeeded = False
    End If

Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsTopEdgeInsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopEdgeInsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsTopEdgeInsideCornerNeeded"
    On Error GoTo ErrorHandler
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
        
    ' ----------
    ' Get the AC
    ' ----------
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oMD.CAO, sParentItemName, oParentObj
    
    ' ---------------------------------------------
    ' Determine if this cut is on the bottom flange
    ' ---------------------------------------------
    Dim sBottomFlange As String
    Dim bIsBottomFlange As Boolean
    
    GetSelectorAnswer oMD.CAO, "BottomFlange", sBottomFlange
        
    bIsBottomFlange = False
    
    If LCase(sBottomFlange) = LCase(gsYes) Then
        bIsBottomFlange = True
    End If
    
    ' ---------------------------------------------------
    ' Use method to determine if the feature is needed
    ' ---------------------------------------------------
    IsCornerFeatureNeededForShapeAtEdge oParentObj, False, bIsBottomFlange, bIsNeeded
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number

End Sub

'*********************************************************************************************
' Method      : CM_IsBottomEdgeInsideCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomEdgeInsideCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomEdgeInsideCornerNeeded"
    On Error GoTo ErrorHandler
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    
    ' ----------
    ' Get the AC
    ' ----------
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oMD.CAO, sParentItemName, oParentObj
    
    ' ---------------------------------------------
    ' Determine if this cut is on the bottom flange
    ' ---------------------------------------------
    Dim sBottomFlange As String
    Dim bIsBottomFlange As Boolean
    
    GetSelectorAnswer oMD.CAO, "BottomFlange", sBottomFlange
        
    bIsBottomFlange = False
    
    If LCase(sBottomFlange) = LCase(gsYes) Then
        bIsBottomFlange = True
    End If
    
    ' ---------------------------------------------------
    ' Use method to determine if the feature is needed
    ' ---------------------------------------------------
    IsCornerFeatureNeededForShapeAtEdge oParentObj, True, bIsBottomFlange, bIsNeeded
        
                
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsTopFacePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopFacePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopFacePCNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    
    bIsNeeded = False  'Initial value, make it to True only for valid cases
    
'1st check:- does web-right-top exist? exit otherwise
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    'Get bounding and bounded ports
    GetBoundingAndBounded pMemberDescription.CAO, oBoundingPort, oBoundedPort

    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    'Edge Mapping
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
    'Only one valid case i.e. WebBuiltUpTopFlangeRight (lSectionAlias = 2)
    If lSectionAlias <> 2 Then Exit Sub      '*** Exit ***
    
'2nd check:- findout cases for which this PC is not applicable and exit
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName pMemberDescription.CAO, sParentItemName, oParentObj
    Select Case LCase(sParentItemName)
        'Atleast one edge is needed to define web-right-top
        'No built up exits with two flanges and similar to WebBuiltUpTopFlangeRight
        Case LCase(gsMbrAxisToCenter), LCase(gsMbrAxisToEdgeAndEdge), _
            LCase(gsMbrAxisToFaceAndEdge), LCase(gsMbrAxisToOnMember), _
            LCase(gsMbrAxisToFaceAndOutSideNoEdge), _
            LCase(gsMbrAxisToOutSideAndOutSideNoEdge), _
            LCase(gsMbrAxisToOutSideAndOutSide2Edge), _
            LCase(gsMbrAxisToEdgeAndOutSide2Edge), _
            LCase(gsStiffEndToMbrCenter), LCase(gsStiffEndToMbrEdgeAndEdge), _
            LCase(gsStiffEndToMbrFaceAndEdge), LCase(gsStiffEndToOnMember), _
            LCase(gsStiffEndToMbrFaceAndOutSideNoEdge), _
            LCase(gsStiffEndToMbrOutSideAndOutSideNoEdge), _
            LCase(gsStiffEndToMbrOutSideAndOutSide2Edge), _
            LCase(gsStiffEndToMbrEdgeAndOutSide2Edge)
            Exit Sub                         '*** Exit ***
            '
    End Select
    
'3rd check:-
    Dim cTopOrWL As ConnectedEdgeInfo
    Dim cBottomOrWR As ConnectedEdgeInfo
    Dim cInsideTForTFL As ConnectedEdgeInfo
    Dim cInsideBForTFR As ConnectedEdgeInfo
     
    GetConnectedEdgeInfo oParentObj, oBoundedPort, oBoundingPort, cTopOrWL, cBottomOrWR, _
        cInsideTForTFL, cInsideBForTFR
    'Check that Web_Right_Top is the intersecting edge
    If cTopOrWL.IntersectingEdge = eBounding_Edge.Web_Right_Top Or _
        cBottomOrWR.IntersectingEdge = eBounding_Edge.Web_Right_Top Then
        
        Dim sShapeAtEdgeAnswer As String
        GetSelectorAnswer oParentObj, "ShapeAtEdge", sShapeAtEdgeAnswer
        Select Case LCase(sParentItemName)
            Case LCase(gsMbrAxisToEdgeAndOutSide1Edge), LCase(gsStiffEndToMbrEdgeAndOutSide1Edge)
                Select Case LCase(sShapeAtEdgeAnswer)
                    Case LCase(gsFaceToOutside), LCase(gsInsideToOutside), _
                         LCase(gsCornerToOutside), LCase(gsEdgeToOutside)
                        'bIsNeeded = False ; invalid cases
                    Case Else
                        bIsNeeded = True
                End Select
            Case LCase(gsMbrAxisToFaceAndOutSide1Edge), LCase(gsStiffEndToMbrFaceAndOutSide1Edge)
                Dim sExtendOrSetBackAnswer As String
                Dim sOffsetAnswer As String
                GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBackAnswer
                GetSelectorAnswer oParentObj, "Offset", sOffsetAnswer
                If (InStr(1, sExtendOrSetBackAnswer, "OffsetNearCorner", vbTextCompare) = 0 And _
                    CDbl(sOffsetAnswer) > 0.00002) Then
                    Select Case LCase(sShapeAtEdgeAnswer) 'uses ShapeAtEdgeOutside codelist
                        'Check for valid cases
                        Case LCase(gsFaceToOutsideCorner), LCase(gsFaceToOutside), _
                            LCase(gsInsideToOutsideCorner), LCase(gsInsideToOutside), _
                            LCase(gsCornerToOutside), LCase(gsEdgeToOutside)
                            'bIsNeeded = False ; invalid cases
                        Case Else
                            bIsNeeded = True
                    End Select
                End If
            Case LCase(gsMbrAxisToOutSideAndOutSide1Edge), LCase(gsStiffEndToMbrOutSideAndOutSide1Edge)
                    Select Case LCase(sShapeAtEdgeAnswer) 'uses ShapeAtEdgeOutside codelist
                        'Check for valid cases
                        Case LCase(gsFaceToOutsideCorner), LCase(gsFaceToOutside), _
                            LCase(gsInsideToOutsideCorner), LCase(gsInsideToOutside), _
                            LCase(gsCornerToOutside), LCase(gsEdgeToOutside)
                            'bIsNeeded = False ; invalid cases
                        Case Else
                            bIsNeeded = True
                    End Select
            Case Else
                '
        End Select
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomFacePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomFacePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomFacePCNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    
    bIsNeeded = False  'Initial value, make it to True only for valid cases
    
'1st check:- does web-right-bottom exist? exit otherwise
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    'Get bounding and bounded ports
    GetBoundingAndBounded pMemberDescription.CAO, oBoundingPort, oBoundedPort

    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    'Edge Mapping
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
    'Only one valid case i.e. WebBuiltUpBottomFlangeRight (lSectionAlias = 4)
    If lSectionAlias <> 4 Then Exit Sub      '*** Exit ***
    
'2nd check:- findout cases for which this PC is not applicable and exit
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName pMemberDescription.CAO, sParentItemName, oParentObj
    Select Case LCase(sParentItemName)
        'Atleast one edge is needed to define web-right-top
        'No built up exits with two flanges and similar to WebBuiltUpTopFlangeRight
        Case LCase(gsMbrAxisToCenter), LCase(gsMbrAxisToEdgeAndEdge), _
            LCase(gsMbrAxisToFaceAndEdge), LCase(gsMbrAxisToOnMember), _
            LCase(gsMbrAxisToFaceAndOutSideNoEdge), _
            LCase(gsMbrAxisToOutSideAndOutSideNoEdge), _
            LCase(gsMbrAxisToOutSideAndOutSide2Edge), _
            LCase(gsMbrAxisToEdgeAndOutSide2Edge), _
            LCase(gsStiffEndToMbrCenter), LCase(gsStiffEndToMbrEdgeAndEdge), _
            LCase(gsStiffEndToMbrFaceAndEdge), LCase(gsStiffEndToOnMember), _
            LCase(gsStiffEndToMbrFaceAndOutSideNoEdge), _
            LCase(gsStiffEndToMbrOutSideAndOutSideNoEdge), _
            LCase(gsStiffEndToMbrOutSideAndOutSide2Edge), _
            LCase(gsStiffEndToMbrEdgeAndOutSide2Edge)
            Exit Sub                         '*** Exit ***
            '
    End Select
    
'3rd check:-
    Dim cTopOrWL As ConnectedEdgeInfo
    Dim cBottomOrWR As ConnectedEdgeInfo
    Dim cInsideTForTFL As ConnectedEdgeInfo
    Dim cInsideBForTFR As ConnectedEdgeInfo
     
    GetConnectedEdgeInfo oParentObj, oBoundedPort, oBoundingPort, cTopOrWL, cBottomOrWR, _
        cInsideTForTFL, cInsideBForTFR
    'Check that Web_Right_Bottom is the intersecting edge
    If cTopOrWL.IntersectingEdge = eBounding_Edge.Web_Right_Bottom Or _
        cBottomOrWR.IntersectingEdge = eBounding_Edge.Web_Right_Bottom Then
        Dim sShapeAtEdgeAnswer As String
        GetSelectorAnswer oParentObj, "ShapeAtEdge", sShapeAtEdgeAnswer
        Select Case LCase(sParentItemName)
            Case LCase(gsMbrAxisToEdgeAndOutSide1Edge), LCase(gsStiffEndToMbrEdgeAndOutSide1Edge)
                Select Case LCase(sShapeAtEdgeAnswer)
                    Case LCase(gsFaceToOutside), LCase(gsInsideToOutside), _
                         LCase(gsCornerToOutside), LCase(gsEdgeToOutside)
                        'bIsNeeded = False ; invalid cases
                    Case Else
                        bIsNeeded = True
                End Select
            Case LCase(gsMbrAxisToFaceAndOutSide1Edge), LCase(gsStiffEndToMbrFaceAndOutSide1Edge)
                Dim sExtendOrSetBackAnswer As String
                Dim sOffsetAnswer As String
                GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBackAnswer
                GetSelectorAnswer oParentObj, "Offset", sOffsetAnswer
                If (InStr(1, sExtendOrSetBackAnswer, "OffsetNearCorner", vbTextCompare) = 0 And _
                    CDbl(sOffsetAnswer) > 0.00002) Then
                    Select Case LCase(sShapeAtEdgeAnswer) 'uses ShapeAtEdgeOutside codelist
                        Case LCase(gsFaceToOutsideCorner), LCase(gsFaceToOutside), _
                            LCase(gsInsideToOutsideCorner), LCase(gsInsideToOutside), _
                            LCase(gsCornerToOutside), LCase(gsEdgeToOutside)
                            'bIsNeeded = False ; invalid cases
                        Case Else
                            bIsNeeded = True
                    End Select
                End If
            Case LCase(gsMbrAxisToOutSideAndOutSide1Edge), LCase(gsStiffEndToMbrOutSideAndOutSide1Edge)
                    Select Case LCase(sShapeAtEdgeAnswer) 'uses ShapeAtEdgeOutside codelist
                        Case LCase(gsFaceToOutsideCorner), LCase(gsFaceToOutside), _
                            LCase(gsInsideToOutsideCorner), LCase(gsInsideToOutside), _
                            LCase(gsCornerToOutside), LCase(gsEdgeToOutside)
                            'bIsNeeded = False ; invalid cases
                        Case Else
                            bIsNeeded = True
                    End Select
            Case Else
                '
        End Select
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub
'*********************************************************************************************
' Method      : CM_IsTopPCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopPCNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsTopPCNeeded"
    
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    Dim sMsg As String
    
    bIsNeeded = False
    
    ' ------------------------------
    ' Get bounded and bounding ports
    ' ------------------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    GetBoundingAndBounded oMD.CAO, oBoundingPort, oBoundedPort

    ' -------------------------------------------------------------------
    ' If it is not a AC type for which a PC is possible, exit immediately
    ' -------------------------------------------------------------------
    Dim sParentItemName As String
    Dim oParentObj As Object
    
    Parent_SmartItemName oMD.CAO, sParentItemName, oParentObj
    
    Select Case LCase(sParentItemName)
        Case LCase(gsMbrAxisToCenter), LCase(gsMbrAxisToEdgeAndEdge), LCase(gsMbrAxisToFaceAndEdge)
            Exit Sub
    End Select
    
    Dim bCutBTWBraceAndBounded As Boolean
    Dim bCutBTWBraceAndBounding As Boolean
    If Not oParentObj Is Nothing Then
        If TypeOf oParentObj Is StructAssemblyConnection Then
            'check whether the AC between Brace and Bounding/Bounded
            IsCutOnInsetBrace oParentObj, bCutBTWBraceAndBounded, bCutBTWBraceAndBounding
            If bCutBTWBraceAndBounding Then
                'if AC between brace and bouning then PC is not needed
                bIsNeeded = False
                Exit Sub
            End If
        End If
    End If
    ' -------------------------------------------------------------------------
    ' If the top of the bounded is not above the bounding, then no PC is needed
    ' -------------------------------------------------------------------------
    ' This eliminates ToFace cases that do not go outside
    ' Use inside of web or flange, as appropriate
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBtmOrWR As ConnectedEdgeInfo
    Dim eITFOrFL As ConnectedEdgeInfo
    Dim eIBFOrFR As ConnectedEdgeInfo
     
    GetConnectedEdgeInfo oParentObj, oBoundedPort, oBoundingPort, eTopOrWL, eBtmOrWR, eITFOrFL, eIBFOrFR

    If IsWebPenetrated(oBoundingPort, oBoundedPort) Then
        If Not (eITFOrFL.IntersectingEdge = eBounding_Edge.Top Or eITFOrFL.IntersectingEdge = eBounding_Edge.Above) Then
            Exit Sub
        End If
    Else
        Dim bTFL As Boolean
        Dim bTFR As Boolean
        Dim bBFL As Boolean
        Dim bBFR As Boolean
        CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR
                    
        If (bTFL) Or (bBFL) Then
            If Not (eITFOrFL.IntersectingEdge = eBounding_Edge.Top Or eITFOrFL.IntersectingEdge = eBounding_Edge.Above) Then
                Exit Sub
            End If
        Else
            If Not (eBtmOrWR.IntersectingEdge = eBounding_Edge.Top Or eBtmOrWR.IntersectingEdge = eBounding_Edge.Above) Then
                Exit Sub
            End If
        End If
    End If
        
    ' -------------------------------------------------------------------------------
    ' If there is an extension, and it is offset from the near corner, there is no PC
    ' -------------------------------------------------------------------------------
    Dim sExtendOrSetBackAnswer As String
    
    GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBackAnswer
                
    If sExtendOrSetBackAnswer = gsOffsetNearCorner Then
        Exit Sub
    End If
    
    ' -----------------------------------
    ' Get the shape at the top and bottom
    ' -----------------------------------
    Dim sBottomAnswerCol As String
    Dim sBottomShape As String
    Dim sTopAnswerCol As String
    Dim sTopShape As String

    GetMemberACTopAndBottomShape oMD.CAO, , , sTopAnswerCol, sTopShape
            
    ' ----------------------------------------------------------------
    ' If the top shape collection is ShapeAtFace and no shape is added
    ' ----------------------------------------------------------------
    If sTopAnswerCol = gsShapeAtFaceCol And sTopShape = gsNone Then
        
        ' -------------------------------------------------------------
        ' A PC is needed if the top is angled toward the bounded object
        ' -------------------------------------------------------------
        Dim lSectionAlias As Long
        Dim bPenetratesWeb As Boolean
        Dim oEdgeMapColl As JCmnShp_CollectionAlias
                
        Set oEdgeMapColl = New Collection
        
        Set oEdgeMapColl = GetEdgeMap(oMD.CAO, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
        Dim oBoundingWebRightPort As IJPort
        Dim oBoundedTopPort As IJPort
        
        Set oBoundingWebRightPort = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT)))
        Set oBoundedTopPort = GetLateralSubPortBeforeTrim(oBoundedPort.Connectable, JXSEC_TOP)
            
        Dim oStructFeature As IJStructFeature
        Dim oBoundedLocation As IJDPosition
        Dim oSDOWebCut As New StructDetailObjects.WebCut
        
        Set oStructFeature = oMD.CAO
                    
        If oStructFeature.get_StructFeatureType = SF_WebCut Then
            Set oSDOWebCut.object = oStructFeature
            Set oBoundedLocation = oSDOWebCut.BoundedLocation
        Else
            Dim oSDOFlangeCut As New StructDetailObjects.FlangeCut
            Set oSDOFlangeCut.object = oStructFeature
            Set oSDOWebCut.object = oSDOFlangeCut.WebCut
            Set oBoundedLocation = oSDOWebCut.BoundedLocation
        End If
        
        Dim oTopSurface As IJSurfaceBody
        Dim oWebRightSurface As IJSurfaceBody
        
        If oBoundedTopPort Is Nothing Then
            'Unable to get Top Port (Tunular Member)
            Exit Sub
        End If
        
        Set oTopSurface = oBoundedTopPort.Geometry
        Set oWebRightSurface = oBoundingWebRightPort.Geometry
        
        Dim oPosition As IJDPosition
        Dim dDistance As Double
        
        Dim oTopNormal As IJDVector
        Dim oWebRightNormal As IJDVector
        Dim oModelBodyUtils As IJSGOModelBodyUtilities
        Set oModelBodyUtils = New SGOModelBodyUtilities
    
        oModelBodyUtils.GetClosestPointOnBody oTopSurface, oBoundedLocation, oPosition, dDistance
        oTopSurface.GetNormalFromPosition oPosition, oTopNormal
        
        oModelBodyUtils.GetClosestPointOnBody oWebRightSurface, oBoundedLocation, oPosition, dDistance
        oWebRightSurface.GetNormalFromPosition oPosition, oWebRightNormal

        If oTopNormal.Dot(oWebRightNormal) < -0.0001 Then
            bIsNeeded = True
        End If
        
    ' -----------------------------------------------------------------
    ' If the top shape collection is ShapeAtEdgeOutside, a PC is needed
    ' -----------------------------------------------------------------
    ElseIf sTopAnswerCol = gsShapeAtEdgeOutsideCol Then
        bIsNeeded = True
    ' ----------------------------------------------------------------------------------------
    ' If the top shape collection is ShapeAtEdgeOverlap a PC is needed depending on the answer
    ' ----------------------------------------------------------------------------------------
    ElseIf sTopAnswerCol = gsShapeAtEdgeOverlapCol Then
        ' We checked for the extension being set back from the near edge above
        ' So a connection is needed unless the extension is not defined
        If Not sExtendOrSetBackAnswer = vbNullString Then
            bIsNeeded = True
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomPCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomPCNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomPCNeeded"
    
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    Dim sMsg As String
    
    bIsNeeded = False
    
    ' ------------------------------
    ' Get bounded and bounding ports
    ' ------------------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    GetBoundingAndBounded oMD.CAO, oBoundingPort, oBoundedPort

    ' -------------------------------------------------------------------
    ' If it is not a AC type for which a PC is possible, exit immediately
    ' -------------------------------------------------------------------
    Dim sParentItemName As String
    Dim oParentObj As Object
    
    Parent_SmartItemName oMD.CAO, sParentItemName, oParentObj
    
    Select Case LCase(sParentItemName)
        Case LCase(gsMbrAxisToCenter), LCase(gsMbrAxisToEdgeAndEdge), LCase(gsMbrAxisToFaceAndEdge)
            Exit Sub
    End Select
    Dim bCutBTWBraceAndBounded As Boolean
    Dim bCutBTWBraceAndBounding As Boolean
    If Not oParentObj Is Nothing Then
        If TypeOf oParentObj Is StructAssemblyConnection Then
            'check whether the AC is between brace and bounding or bounded
            IsCutOnInsetBrace oParentObj, bCutBTWBraceAndBounded, bCutBTWBraceAndBounding
            If bCutBTWBraceAndBounding Then
                'if it is Brace AC then PC is not needed
                bIsNeeded = False
                Exit Sub
            End If
        End If
    End If
    ' ----------------------------------------------------------------------------
    ' If the bottom of the bounded is not below the bounding, then no PC is needed
    ' ----------------------------------------------------------------------------
    ' This eliminates ToFace cases that do not go outside
    ' Use inside of web or flange, as appropriate
    Dim eTopOrWL As ConnectedEdgeInfo
    Dim eBtmOrWR As ConnectedEdgeInfo
    Dim eITFOrFL As ConnectedEdgeInfo
    Dim eIBFOrFR As ConnectedEdgeInfo
     
    GetConnectedEdgeInfo oParentObj, oBoundedPort, oBoundingPort, eTopOrWL, eBtmOrWR, eITFOrFL, eIBFOrFR

    If IsWebPenetrated(oBoundingPort, oBoundedPort) Then
        If Not (eIBFOrFR.IntersectingEdge = eBounding_Edge.Bottom Or eIBFOrFR.IntersectingEdge = eBounding_Edge.Below) Then
            Exit Sub
        End If
    Else
        Dim bTFL As Boolean
        Dim bTFR As Boolean
        Dim bBFL As Boolean
        Dim bBFR As Boolean
        CrossSection_Flanges oBoundedPort.Connectable, bTFL, bBFL, bTFR, bBFR
                    
        If (bTFL) Or (bBFL) Then
            If Not (eIBFOrFR.IntersectingEdge = eBounding_Edge.Bottom Or eIBFOrFR.IntersectingEdge = eBounding_Edge.Below) Then
                Exit Sub
            End If
        Else
            If Not (eTopOrWL.IntersectingEdge = eBounding_Edge.Bottom Or eTopOrWL.IntersectingEdge = eBounding_Edge.Below) Then
                Exit Sub
            End If
        End If
    End If
        
    ' -------------------------------------------------------------------------------
    ' If there is an extension, and it is offset from the near corner, there is no PC
    ' -------------------------------------------------------------------------------
    Dim sExtendOrSetBackAnswer As String
    
    GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBackAnswer
                
    If sExtendOrSetBackAnswer = gsOffsetNearCorner Then
        Exit Sub
    End If
    
    ' ---------------------------
    ' Get the shape at the bottom
    ' ---------------------------
    Dim sBottomAnswerCol As String
    Dim sBottomShape As String

    GetMemberACTopAndBottomShape oMD.CAO, sBottomAnswerCol, sBottomShape
            
    ' -------------------------------------------------------------------
    ' If the bottom shape collection is ShapeAtFace and no shape is added
    ' -------------------------------------------------------------------
    If sBottomAnswerCol = gsShapeAtFaceCol And sBottomShape = gsNone Then
        
        ' ----------------------------------------------------------------
        ' A PC is needed if the bottom is angled toward the bounded object
        ' ----------------------------------------------------------------
        Dim lSectionAlias As Long
        Dim bPenetratesWeb As Boolean
        Dim oEdgeMapColl As JCmnShp_CollectionAlias
        
        Set oEdgeMapColl = New Collection
        
        Set oEdgeMapColl = GetEdgeMap(oMD.CAO, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
        Dim oBoundingWebRightPort As IJPort
        Dim oBoundedBottomPort As IJPort
        
        Set oBoundingWebRightPort = GetLateralSubPortBeforeTrim(oBoundingPort.Connectable, oEdgeMapColl.Item(CStr(JXSEC_WEB_RIGHT)))
        Set oBoundedBottomPort = GetLateralSubPortBeforeTrim(oBoundedPort.Connectable, JXSEC_BOTTOM)
            
        Dim oStructFeature As IJStructFeature
        Dim oBoundedLocation As IJDPosition
        Dim oSDOWebCut As New StructDetailObjects.WebCut
        
        Set oStructFeature = oMD.CAO
                    
        If oStructFeature.get_StructFeatureType = SF_WebCut Then
            Set oSDOWebCut.object = oStructFeature
            Set oBoundedLocation = oSDOWebCut.BoundedLocation
        Else
            Dim oSDOFlangeCut As New StructDetailObjects.FlangeCut
            Set oSDOFlangeCut.object = oStructFeature
            Set oSDOWebCut.object = oSDOFlangeCut.WebCut
            Set oBoundedLocation = oSDOWebCut.BoundedLocation
        End If
            
        Dim oBottomSurface As IJSurfaceBody
        Dim oWebRightSurface As IJSurfaceBody
        
        If oBoundedBottomPort Is Nothing Then
            'Unable to get Bounded Bottom Port (Tubular Member)
            Exit Sub
        End If
        
        Set oBottomSurface = oBoundedBottomPort.Geometry
        Set oWebRightSurface = oBoundingWebRightPort.Geometry
        
        Dim oPosition As IJDPosition
        Dim dDistance As Double
        
        Dim oBottomNormal As IJDVector
        Dim oWebRightNormal As IJDVector
        Dim oModelBodyUtils As IJSGOModelBodyUtilities
        Set oModelBodyUtils = New SGOModelBodyUtilities
    
        oModelBodyUtils.GetClosestPointOnBody oBottomSurface, oBoundedLocation, oPosition, dDistance
        oBottomSurface.GetNormalFromPosition oPosition, oBottomNormal
        
        oModelBodyUtils.GetClosestPointOnBody oWebRightSurface, oBoundedLocation, oPosition, dDistance
        oWebRightSurface.GetNormalFromPosition oPosition, oWebRightNormal

        If oBottomNormal.Dot(oWebRightNormal) < -0.0001 Then
            bIsNeeded = True
        End If
        
    ' -----------------------------------------------------------------
    ' If the top shape collection is ShapeAtEdgeOutside, a PC is needed
    ' -----------------------------------------------------------------
    ElseIf sBottomAnswerCol = gsShapeAtEdgeOutsideCol Then
        bIsNeeded = True
    ' ------------------------------------------------------------------------------------------------------
    ' If the top shape collection is ShapeAtEdgeOverlap a PC is needed depending on the answer and extension
    ' ------------------------------------------------------------------------------------------------------
    ElseIf sBottomAnswerCol = gsShapeAtEdgeOverlapCol Then
        ' We checked for the extension being set back from the near edge above
        ' So a connection is needed unless the extension is not defined
        If Not sExtendOrSetBackAnswer = vbNullString Then
            bIsNeeded = True
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub
'*********************************************************************************************
' Method      : CM_IsTopEdgeOutsidePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsTopEdgeOutsidePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsTopEdgeOutsidePCNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    
    bIsNeeded = False  'Initial value, make it to True only for valid cases

'1st check:- does flange-right-top exist? exit otherwise
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    'Get bounding and bounded ports
    GetBoundingAndBounded pMemberDescription.CAO, oBoundingPort, oBoundedPort

    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    'Edge Mapping
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
    'Only one valid case i.e. WebBuiltUpTopFlangeRight (lSectionAlias = 2)
    If lSectionAlias <> 2 Then Exit Sub      '*** Exit ***


    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName pMemberDescription.CAO, sParentItemName, oParentObj

'2nd check:- confirm that bounding-top/above is the intersecting edge, otherwise exit
    Dim cTopOrWL As ConnectedEdgeInfo
    Dim cBottomOrWR As ConnectedEdgeInfo
    Dim cInsideTForTFL As ConnectedEdgeInfo
    Dim cInsideBForTFR As ConnectedEdgeInfo
     
    GetConnectedEdgeInfo oParentObj, oBoundedPort, oBoundingPort, cTopOrWL, cBottomOrWR, _
        cInsideTForTFL, cInsideBForTFR

    If Not (cTopOrWL.IntersectingEdge = eBounding_Edge.Top Or _
       cTopOrWL.IntersectingEdge = eBounding_Edge.Above Or _
       cInsideTForTFL.IntersectingEdge = eBounding_Edge.Top Or _
       cInsideTForTFL.IntersectingEdge = eBounding_Edge.Above) Then
        Exit Sub                            '*** Exit ***
    End If
        
'3rd check:- selector answers need to be checked

    Dim sExtendOrSetBackAnswer As String
    Dim sOffsetAnswer As String
    Dim sShapeAtEdgeAnswer As String
        Select Case LCase(sParentItemName)
            Case LCase(gsMbrAxisToOnMember), LCase(gsMbrAxisToEdge), LCase(gsMbrAxisToFaceAndOutSide1Edge), _
                 LCase(gsStiffEndToOnMember), LCase(gsStiffEndToMbrEdge), LCase(gsStiffEndToMbrFaceAndOutSide1Edge)
               GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBackAnswer
               GetSelectorAnswer oParentObj, "Offset", sOffsetAnswer
                'Otherthan OffsetNearCorner cases are needed
                If (InStr(1, sExtendOrSetBackAnswer, "OffsetNearCorner", vbTextCompare) = 0 And _
                    CDbl(sOffsetAnswer) > 0.00002) Then
                    If StrComp(sParentItemName, gsMbrAxisToOnMember, vbTextCompare) = 0 Then
                        'On member
                        bIsNeeded = True
                    ElseIf StrComp(sParentItemName, gsMbrAxisToEdge, vbTextCompare) = 0 Then
                        GetSelectorAnswer oParentObj, "ShapeAtEdge", sShapeAtEdgeAnswer
                        'Axis to edge
                        Select Case LCase(sShapeAtEdgeAnswer)
                            Case LCase(gsFaceToOutside), LCase(gsInsideToOutside), _
                                 LCase(gsCornerToOutside), LCase(gsEdgeToOutside)
                                'bIsNeeded = False ; invalid cases
                            Case Else
                                bIsNeeded = True
                        End Select
                    ElseIf StrComp(sParentItemName, gsMbrAxisToFaceAndOutSide1Edge, vbTextCompare) = 0 Then
                        GetSelectorAnswer oParentObj, "ShapeAtEdge", sShapeAtEdgeAnswer
                        'Face and out side 1 edge
                        Select Case LCase(sShapeAtEdgeAnswer) 'uses ShapeAtEdgeOutside codelist
                            'Check for valid cases
                            Case LCase(gsFaceToOutsideCorner), LCase(gsFaceToOutside), _
                                LCase(gsInsideToOutsideCorner), LCase(gsInsideToOutside), _
                                LCase(gsCornerToOutside), LCase(gsEdgeToOutside)
                                'bIsNeeded = False ; invalid cases
                            Case Else
                                bIsNeeded = True
                        End Select
                    End If
                End If
            Case Else
                '
        End Select
        
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsBottomEdgeOutsidePCNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsBottomEdgeOutsidePCNeeded(pMemberDescription As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsBottomEdgeOutsidePCNeeded"
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    
    bIsNeeded = False  'Initial value, make it to True only for valid cases

'1st check:- does Flange-right-bottom exist? exit otherwise
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort

    'Get bounding and bounded ports
    GetBoundingAndBounded pMemberDescription.CAO, oBoundingPort, oBoundedPort

    Dim lSectionAlias As Long
    Dim bPenetratesWeb As Boolean
    Dim oEdgeMapColl As JCmnShp_CollectionAlias
    
    'Edge Mapping
    Set oEdgeMapColl = New Collection
    Set oEdgeMapColl = GetEdgeMap(pMemberDescription.CAO, oBoundingPort, oBoundedPort, lSectionAlias, bPenetratesWeb)
    'Only one valid case i.e. WebBuiltUpBottomFlangeRight (lSectionAlias = 4)
    If lSectionAlias <> 4 Then Exit Sub      '*** Exit ***

'2nd check:- confirm that bounding-top/above is the intersecting edge, otherwise exit
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName pMemberDescription.CAO, sParentItemName, oParentObj
    Dim cTopOrWL As ConnectedEdgeInfo
    Dim cBottomOrWR As ConnectedEdgeInfo
    Dim cInsideTForTFL As ConnectedEdgeInfo
    Dim cInsideBForTFR As ConnectedEdgeInfo
    
    
    GetConnectedEdgeInfo oParentObj, oBoundedPort, oBoundingPort, cTopOrWL, cBottomOrWR, _
        cInsideTForTFL, cInsideBForTFR

    If Not (cBottomOrWR.IntersectingEdge = eBounding_Edge.Below Or _
            cBottomOrWR.IntersectingEdge = eBounding_Edge.Bottom Or _
            cInsideBForTFR.IntersectingEdge = eBounding_Edge.Below Or _
            cInsideBForTFR.IntersectingEdge = eBounding_Edge.Bottom) Then
        Exit Sub                            '*** Exit ***
    End If
        
'3rd check:- selector answers need to be checked
    Dim sExtendOrSetBackAnswer As String
    Dim sOffsetAnswer As String
    Dim sShapeAtEdgeAnswer As String
    Select Case LCase(sParentItemName)
        Case LCase(gsMbrAxisToOnMember), LCase(gsMbrAxisToEdge), LCase(gsMbrAxisToFaceAndOutSide1Edge), _
             LCase(gsStiffEndToOnMember), LCase(gsStiffEndToMbrEdge), LCase(gsStiffEndToMbrFaceAndOutSide1Edge)
           GetSelectorAnswer oParentObj, "ExtendOrSetBack", sExtendOrSetBackAnswer
           GetSelectorAnswer oParentObj, "Offset", sOffsetAnswer
            If (InStr(1, sExtendOrSetBackAnswer, "OffsetNearCorner", vbTextCompare) = 0 And _
                CDbl(sOffsetAnswer) > 0.00002) Then
                If StrComp(sParentItemName, gsMbrAxisToOnMember, vbTextCompare) = 0 Then
                    'On member
                    bIsNeeded = True
                ElseIf StrComp(sParentItemName, gsMbrAxisToEdge, vbTextCompare) = 0 Then
                    GetSelectorAnswer oParentObj, "ShapeAtEdge", sShapeAtEdgeAnswer
                    'Axis to edge
                    Select Case LCase(sShapeAtEdgeAnswer)
                        Case LCase(gsFaceToOutside), LCase(gsInsideToOutside), _
                             LCase(gsCornerToOutside), LCase(gsEdgeToOutside)
                            'bIsNeeded = False ; invalid cases
                        Case Else
                            bIsNeeded = True
                    End Select
                ElseIf StrComp(sParentItemName, gsMbrAxisToFaceAndOutSide1Edge, vbTextCompare) = 0 Then
                    GetSelectorAnswer oParentObj, "ShapeAtEdge", sShapeAtEdgeAnswer
                    'Face and out side 1 edge
                    Select Case LCase(sShapeAtEdgeAnswer) 'uses ShapeAtEdgeOutside codelist
                        'Check for valid cases
                        Case LCase(gsFaceToOutsideCorner), LCase(gsFaceToOutside), _
                            LCase(gsInsideToOutsideCorner), LCase(gsInsideToOutside), _
                            LCase(gsCornerToOutside), LCase(gsEdgeToOutside)
                            'bIsNeeded = False ; invalid cases
                        Case Else
                            bIsNeeded = True
                    End Select
                End If
            End If
        Case Else
            '
    End Select
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub
'*********************************************************************************************
' Method      : CM_IsFaceTopCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFaceTopCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFaceTopCornerNeeded"
    On Error GoTo ErrorHandler
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    
    Dim oDefCM As New EndCutDefCM
    oDefCM.CM_IsAxisFaceTopCornerRequested oMD.CAO, bIsNeeded
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Sub

'*********************************************************************************************
' Method      : CM_IsFaceBottomCornerNeeded
' Description :
'
'*********************************************************************************************
Public Sub CM_IsFaceBottomCornerNeeded(oMD As IJDMemberDescription, bIsNeeded As Boolean)
    Const METHOD = m_DefinitionProgid & "::CM_IsFaceBottomCornerNeeded"
    On Error GoTo ErrorHandler
    
    Dim oBounded As Object
    Dim eItemObjectType As eObjectType
    
    GetBoundingAndBounded oMD.CAO, Nothing, oBounded
    
    If TypeOf oBounded Is ISPSMemberPartCommon Then
        eItemObjectType = eObjectType.e_MemberEndCutCornerFeature
    Else
        eItemObjectType = eObjectType.e_StiffenerEndCutCornerFeature
    End If
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eItemObjectType) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sMsg As String
    
    Dim oDefCM As New EndCutDefCM
    oDefCM.CM_IsAxisFaceBottomCornerRequested oMD.CAO, bIsNeeded
        
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
    IJDUserSymbolServices_GetDefinitionName = m_DefinitionName
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
    pDefinition.ProgId = m_DefinitionProgid
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

'*********************************************************************************************
' Method      : CM_IsTopEdgePCNeeded
' Description :
'
'*********************************************************************************************
Public Function IsEdgePCNeeded(oMD As IJDMemberDescription, bIsBottomEdge As Boolean) As Boolean

    Const METHOD = m_DefinitionProgid & "::CM_IsEdgePCNeeded"
    
    On Error GoTo ErrorHandler

    Dim sMsg As String
    
    IsEdgePCNeeded = False
    
    ' ----------------------------------------------------------
    ' Determine if this is a flange cut, and if so, which flange
    ' ----------------------------------------------------------
    Dim sBottomFlange As String
    GetSelectorAnswer oMD.CAO, "BottomFlange", sBottomFlange
    
    Dim bIsBottomFlange As Boolean
    If LCase(sBottomFlange) = LCase(gsYes) Then
        bIsBottomFlange = True
    End If
    
    ' --------------------------------------
    ' Determine how much overlap at the edge
    ' --------------------------------------
    Dim dInsideOverlap As Double
    Dim dOutsideOverlap As Double
    Dim bIsEdgeToEdge As Boolean
    
    GetEdgeOverlapAndClearance oMD.CAO, bIsBottomEdge, bIsBottomFlange, dInsideOverlap, dOutsideOverlap, , , , bIsEdgeToEdge
    
    ' -----------------------------
    ' Get the shape type and answer
    ' -----------------------------
    Dim sBottomAnswerCol As String
    Dim sBottomShape As String
    Dim sTopAnswerCol As String
    Dim sTopShape As String
    
    GetMemberACTopAndBottomShape oMD.CAO, sBottomAnswerCol, sBottomShape, sTopAnswerCol, sTopShape
    
    Dim bWebIsOutside As Boolean
    Dim sAnswerCol As String
    Dim sShape As String
    Dim dOverlap As Double
    
    If bIsBottomEdge Then
        If sBottomAnswerCol = vbNullString Then
            bWebIsOutside = True
            sAnswerCol = sTopAnswerCol
            sShape = sTopShape
            dOverlap = dOutsideOverlap
        ElseIf sTopAnswerCol = vbNullString Then
            Exit Function
        Else
            bWebIsOutside = False
            sAnswerCol = sBottomAnswerCol
            sShape = sBottomShape
            dOverlap = dInsideOverlap
        End If
    Else
        If sBottomAnswerCol = vbNullString Then
            Exit Function
        ElseIf sTopAnswerCol = vbNullString Then
            bWebIsOutside = True
            sAnswerCol = sBottomAnswerCol
            sShape = sBottomShape
            dOverlap = dOutsideOverlap
        Else
            bWebIsOutside = False
            sAnswerCol = sTopAnswerCol
            sShape = sTopShape
            dOverlap = dInsideOverlap
        End If
    End If
    
    ' -------------------------
    ' Adjust overlap for fillet
    ' -------------------------
    Dim oBoundedPort As IJPort
    Dim oBoundingPort As IJPort
    GetBoundingAndBounded oMD.CAO, oBoundingPort, oBoundedPort

    Dim oSDOMember As StructDetailObjects.MemberPart
    Dim oSDOStiffener As StructDetailObjects.ProfilePart
    Dim dBoundedEdgeThickness As Double

    dBoundedEdgeThickness = -1#

    Dim oFeature As IJStructFeature
    Set oFeature = oMD.CAO
    
    If oFeature.get_StructFeatureType = SF_WebCut Then

        If (Not bIsBottomEdge And Not bWebIsOutside And HasTopFlange(oBoundedPort.Connectable)) Or _
           (Not bIsBottomEdge And bWebIsOutside And HasBottomFlange(oBoundedPort.Connectable)) Or _
           (bIsBottomEdge And Not bWebIsOutside And HasBottomFlange(oBoundedPort.Connectable)) Or _
           (bIsBottomEdge And bWebIsOutside And HasTopFlange(oBoundedPort.Connectable)) Then
           
            If ProfileHasDetailedRepresentation(oBoundedPort.Connectable) Then
                    
                If TypeOf oBoundedPort.Connectable Is ISPSMemberPartCommon Then
                    Set oSDOMember = New MemberPart
                    Set oSDOMember.object = oBoundedPort.Connectable
                    dBoundedEdgeThickness = oSDOMember.flangeThickness
                ElseIf TypeOf oBoundedPort.Connectable Is IJProfile Then
                    Set oSDOStiffener = New ProfilePart
                    Set oSDOStiffener.object = oBoundedPort.Connectable
                    dBoundedEdgeThickness = oSDOStiffener.flangeThickness
                End If
            End If
            
            If dBoundedEdgeThickness > 0# Then
                Dim bGetBottomDistance As Boolean
                If bIsBottomEdge Then
                    bGetBottomDistance = Not bWebIsOutside
                Else
                    bGetBottomDistance = bWebIsOutside
                End If
                
                Dim dDistFromOutsideToInsideOfFillet As Double
                dDistFromOutsideToInsideOfFillet = GetDistanceFromTopOrBottomToWebRight(oBoundedPort.Connectable, bWebIsOutside)
                dOverlap = dOverlap + dBoundedEdgeThickness - dDistFromOutsideToInsideOfFillet
            End If
            
        End If
    Else
        ' ToDo
    End If
    
    ' -----------------------------------------------
    ' If the shape is "At Face", then no PC is needed
    ' -----------------------------------------------
    If sAnswerCol = gsShapeAtFaceCol Then
        Exit Function
    End If
    
    ' ------------------------------------------------------------------------
    ' Otherwise, base answer on collection type, answer, and amount of overlap
    ' ------------------------------------------------------------------------
    ' We must also account for corner features that can be replaced with web cuts that remove all the material from the edge
    ' For shapes where this is possible, we check to see if a corner feature should be used.  If not, we know the web cut
    ' will remove the material necessary for a PC
    Dim bCornerFeatureIsUsedForShape As Boolean
    Dim sParentItemName As String
    Dim oParentObj As Object
    Parent_SmartItemName oMD.CAO, sParentItemName, oParentObj
    
    IsCornerFeatureNeededForShapeAtEdge oParentObj, bIsBottomEdge, bIsBottomFlange, bCornerFeatureIsUsedForShape
    
    If sAnswerCol = gsShapeAtEdgeCol Or sAnswerCol = gsShapeAtEdgeOverlapCol Then
        
        Select Case sShape
           
            Case gsNone
                
                If dOverlap > GetMinPCLength() Then
                    IsEdgePCNeeded = True
                End If
                
            Case gsFaceToCorner, gsFaceToInsideCorner
                
                If dOverlap > GetMinPCLength() And bCornerFeatureIsUsedForShape Then
                    IsEdgePCNeeded = True
                End If
                
            Case gsEdgeToOutside, gsEdgeToFlange, gsEdgeToOutside
                
                If dOverlap > 2 * GetMinPCLength() Then
                    IsEdgePCNeeded = True
                End If
                
            Case gsFaceToEdge, gsInsideToEdge
                
                If dOverlap > 2 * GetMinPCLength() And bCornerFeatureIsUsedForShape Then
                    IsEdgePCNeeded = True
                End If
                
            Case gsCornerToFlange
            
                If dOverlap > GetMinPCLength() And Not bCornerFeatureIsUsedForShape And bIsEdgeToEdge Then
                    IsEdgePCNeeded = True
                End If
                
        End Select
    
    ElseIf sAnswerCol = gsShapeAtEdgeOutsideCol Then
        
        Select Case sShape
            
            Case gsNone
                
                If dOverlap > GetMinPCLength() Then
                    IsEdgePCNeeded = True
                End If
            
            Case gsEdgeToOutside
                
                If dOverlap > 2 * GetMinPCLength() Then
                    IsEdgePCNeeded = True
                End If
            
            Case gsEdgeToFlange
                
                If dOverlap > 2 * GetMinPCLength() And bIsEdgeToEdge Then
                    IsEdgePCNeeded = True
                End If
            
            Case gsOutsideToEdge
                
                If dOverlap > 2 * GetMinPCLength() And bCornerFeatureIsUsedForShape Then
                    IsEdgePCNeeded = True
                End If
                
            Case gsCornerToFlange
                
                If dOverlap > GetMinPCLength() And Not bCornerFeatureIsUsedForShape And bIsEdgeToEdge Then
                    IsEdgePCNeeded = True
                End If
                
        End Select
    
    End If
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
 
End Function

'******************************************************************************************
' Method:
' CM_SetFlangeCuttingBehaviour
'
' *******************************************************************************************
Public Sub CM_SetFlangeCuttingBehaviour(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CM_SetFlangeCuttingBehaviour"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    sMsg = "Updating a Flange Cutting Behaviour"
            
    Dim oMemberDescriptions As IJDMemberDescriptions
    Set oMemberDescriptions = pPropertyDescriptions.definition
    
    Dim oMemberDescription As IJDMemberDescription
    Dim iMember As Integer
    Dim bIsNeeded As Boolean
    bIsNeeded = True
    
    Dim IsProExists As Boolean
    Dim oStructFeature As IJStructFeature
    Set oStructFeature = pPropertyDescriptions.CAO
        
    If Not oStructFeature.get_StructFeatureType = SF_FlangeCut Then Exit Sub
    
    'Check if the attribute is exists on flange cut
    IsProExists = Has_Attribute(pPropertyDescriptions.CAO, "CuttingBehavior")
    
    'Exit if attribute is not bulkloaded
    If Not IsProExists Then Exit Sub
        
    'Looping for only 1 to 11 items as they are only the PC item members
    'Call each Item member conditional and check weather it is True or False
    'Based on this Set FlangeCut Cutting Behaviour
    For iMember = 1 To 11
        
        'Get each Member Description
        Set oMemberDescription = oMemberDescriptions.ItemByDispid(iMember)
        oMemberDescription.CAO = pPropertyDescriptions.CAO
    
        Select Case iMember
        
            Case 1
                CM_IsFacePCNeeded oMemberDescription, bIsNeeded
            
            Case 2
                CM_IsTopEdgePCNeeded oMemberDescription, bIsNeeded
                
            Case 3
                CM_IsBottomEdgePCNeeded oMemberDescription, bIsNeeded
            
            Case 4
                CM_IsTopEdgeInsidePCNeeded oMemberDescription, bIsNeeded
            
            Case 5
                CM_IsBottomEdgeInsidePCNeeded oMemberDescription, bIsNeeded
                
            Case 6
                CM_IsTopPCNeeded oMemberDescription, bIsNeeded
            
            Case 7
                CM_IsBottomPCNeeded oMemberDescription, bIsNeeded
                
            Case 8
                CM_IsTopEdgeOutsidePCNeeded oMemberDescription, bIsNeeded
            
            Case 9
                CM_IsBottomEdgeOutsidePCNeeded oMemberDescription, bIsNeeded
            
            Case 10
                CM_IsTopFacePCNeeded oMemberDescription, bIsNeeded
                
            Case 11
                CM_IsBottomFacePCNeeded oMemberDescription, bIsNeeded
                
        End Select
        
        'If any of the PC conditional is True, then Exit
        If bIsNeeded Then Exit For
    
    Next
    
    'Set Proper Flange Cutting Behaviour
    Set_FlangeCuttingBehavior oMemberDescription.CAO, bIsNeeded
    
    oMemberDescription.CAO = Nothing
  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub


