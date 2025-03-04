VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeEndCutDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*********************************************************************************************
'  Copyright (C) 2012, Intergraph Corporation.  All rights reserved.
'
'  Project     : SMFreeEndCutRules
'  File        : FreeEndCutDefCM.cls
'
'  Description :
'
'  Author      : Alligators
'
'  History     :
'    15/Feb/2012 - svsmylav
'       CR-CP-174918: Handled 'Extend' Mfg. option: (i) Upddated WebCut and FlangeCut
'       construct methods use plate's port (ii) Used Box-WebCut item and accordingly
'       BoxCuts progid is used in CMCopyWCQuestionsToMembers for this case.
'*********************************************************************************************
Private sError As String
Private sMETHOD As String
Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "FreeEndCutRules\FreeEndCutSelCM.cls"
'

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function CMConstruct_FreeEndWebCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    On Error GoTo ErrorHandler
    sMETHOD = "CMConstruct_FreeEndWebCut"
    
    Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
    Set oFreeEndCut = pMemberDescription.CAO
    
    ' Retrieve the Bounded and Bounding Port objects from the Free End Cut
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
    
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oBoundedObject
            
    'Get EndCutType answer
    Dim strEndCutTypeAns As String
    GetSelectorAnswer pMemberDescription.CAO, "EndCutType", strEndCutTypeAns
    
    'Create webcut
    Dim oWebCut As StructDetailObjects.WebCut
    Set oWebCut = New StructDetailObjects.WebCut
    sError = "Creating web cut."
    'if the connectable is MemeberPart call MbrFreeWebSel rule
    If TypeOf oBoundedPort.Connectable Is ISPSMemberPartPrismatic Then
        'Check if newly Added Part Class exist
        'thereby we can ensure that End user has updated CATALOG
        'with new PartClass needed for new functionality, otherwise
        'point to old rules
        If CheckPartClassExist("MbrFreeWebSel") Then
            oWebCut.Create pResourceManager, _
                        oBoundingObject, oBoundedObject, _
                        "MbrFreeWebSel", oFreeEndCut
        Else
            oWebCut.Create pResourceManager, _
                        oBoundingObject, oBoundedObject, _
                        "MBR_FreeWebCut", oFreeEndCut ' MP_Web_Weld_SS
        End If
    Else 'Stiffener or Beam Case
    
        'If Free EndCut is from profile pseudo knuckle, need to use
        Dim bIsFromExtendAndConvex As Boolean
        Dim oTempPort As IJPort
    
        bIsFromExtendAndConvex = IsFreeEndCutFromExtend_ConvexKnuckle(oBoundedObject, oTempPort)
        If bIsFromExtendAndConvex Then Set oBoundingObject = oTempPort  'Use plate's port as bounding
    
        If bIsFromExtendAndConvex And StrComp(strEndCutTypeAns, "Welded", vbTextCompare) = 0 Then
            'WebCut is because of 'Extend' Mfg. option and convex knuckle, also this is welded endcut
            'so use Box EndCuts selection
            oWebCut.Create pResourceManager, _
                       oBoundingObject, oBoundedObject, _
                       "WebCutsLongBox", oFreeEndCut
        Else
            'Regular FreeEndcut
            oWebCut.Create pResourceManager, _
                       oBoundingObject, oBoundedObject, _
                       "WebCuts", oFreeEndCut
        End If
    End If
    
'Clean up
    Set oTempPort = Nothing
    
    sError = "Return the webcut"
    Set pObject = oWebCut.object
    Set oBoundedPort = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function CMConstruct_FreeEndTopFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
    On Error GoTo ErrorHandler
    sMETHOD = "CMConstruct_FreeEndTopFlangeCut"
     
    CreateFlangeCut pMemberDescription, pResourceManager, pObject
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function CMConstruct_FreeEndBottomFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
    On Error GoTo ErrorHandler
    sMETHOD = "CMConstruct_FreeEndBottomFlangeCut"
     
    CreateFlangeCut pMemberDescription, pResourceManager, pObject
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Private Function CreateFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObject As Object)
    On Error GoTo ErrorHandler
    sMETHOD = "CreateFlangeCut"
    
    Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
    Set oFreeEndCut = pMemberDescription.CAO
    
    ' Retrieve the Bounded and Bounding Port objects from the Free End Cut
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
           
    ' Create flange cut if one is required
    Dim oFlangeCut As New FlangeCut
    sError = "Creating flange cut."
    
    ' create the flange when the web cut is found
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oWebCut As IJStructFeature
    
    ' all Item Members are not IJStructFeature objects,
    Dim iIndex As Long
    Dim iCount As Long
    Dim bFound As Boolean
    bFound = False
    iCount = pMemberObjects.Count
    
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oBoundedObject
    
    'If Free EndCut is from profile pseudo knuckle, need to use
    ' plate port as bounding
    Dim bIsFromExtendAndConvex As Boolean
    Dim oTempPort As IJPort
    bIsFromExtendAndConvex = IsFreeEndCutFromExtend_ConvexKnuckle(oBoundedObject, oTempPort)
    If bIsFromExtendAndConvex Then Set oBoundingObject = oTempPort 'Use plate's port as bounding
    
    For iIndex = 1 To iCount
        If TypeOf pMemberObjects.Item(iIndex) Is IJStructFeature Then
        
            Set oWebCut = pMemberObjects.Item(iIndex)
            If (oWebCut.get_StructFeatureType = SF_WebCut) Then
                'need to call MbrEndFlangeSel rule in case of memeber
                If TypeOf oBoundedPort.Connectable Is ISPSMemberPartPrismatic Then
                    'Check if newly Added Part Class exist
                    'thereby we can ensure that End user has updated CATALOG
                    'with new PartClass needed for new functionality, otherwise
                    'point to old rules
                    If CheckPartClassExist("MbrEndFlangeSel") Then
                        oFlangeCut.Create pResourceManager, _
                                    oBoundingObject, oBoundedObject, oWebCut, _
                                      "MbrEndFlangeSel", oFreeEndCut  'FlangeCuts
                    Else
                        oFlangeCut.Create pResourceManager, _
                                      oBoundingObject, oBoundedObject, oWebCut, _
                                      "MBR_EndFlangeSel", oFreeEndCut 'FlangeCuts
                    End If
                Else 'for Beam and Stiffeners, use FlangeCuts rule
                    oFlangeCut.Create pResourceManager, _
                                  oBoundingObject, oBoundedObject, oWebCut, _
                                  "FlangeCuts", oFreeEndCut
                End If
                Set oWebCut = Nothing
                
                sError = "Return the flange cut "
                Set pObject = oFlangeCut.object
                Exit For
            End If
            
            Set oWebCut = Nothing
        End If
    Next iIndex
    
    Set oBoundedPort = Nothing
    Set oTempPort = Nothing
    Exit Function
    
    
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'   Determines if a Web Cut is required
'   currently all cross-sections require a Web Cut
Public Sub CMConditional_FreeEndWebCut(ByRef pMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    sMETHOD = "CMConditional_FreeEndWebCut"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'   Determines if a top Flange Cut is required
'   all cross-sections without a flange return false
'   all other cross-sections return true
Public Sub CMConditional_FreeEndTopFlangeCut(ByRef pMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    sMETHOD = "CMConditional_FreeEndTopFlangeCut"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim strSectionType As String
    strSectionType = GetCrossSectionType(pMD)
        
    bIsNeeded = True
    If ((strSectionType = "FB") Or (strSectionType = "HalfR") Or _
        (strSectionType = "P") Or (strSectionType = "R") Or _
        (strSectionType = "SB") Or (strSectionType = "SqTu") Or _
        (strSectionType = "RT") Or (strSectionType = "L") Or _
        (strSectionType = "2L") Or (strSectionType = "HSSC") Or _
        (strSectionType = "HSSR") Or (strSectionType = "PIPE") Or _
        (strSectionType = "CS") Or (strSectionType = "RS")) Then
        bIsNeeded = False
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'   Determines if a bottom Flange Cut is required
'   only I cross-ssections return true
'   all other cross-sections return false
Public Sub CMConditional_FreeEndBottomFlangeCut(ByRef pMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    sMETHOD = "CMConditional_FreeEndBottomFlangeCut"
   
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
       
    Dim strSectionType As String
    strSectionType = GetCrossSectionType(pMD)
        
    bIsNeeded = False
    If ((strSectionType = "I") Or (strSectionType = "ISType") Or _
        (strSectionType = "C_SS") Or (strSectionType = "CSType") Or _
        (strSectionType = "H") Or (strSectionType = "MC") Or (strSectionType = "HP") Or _
        (strSectionType = "L" Or strSectionType = "2L" Or strSectionType = "W") Or _
         (strSectionType = "C") Or (strSectionType = "M") Or (strSectionType = "S")) Then
        bIsNeeded = True
    End If
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub CMSetTopFlange(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    sMETHOD = "CMSetTopFlange"
    Dim strError As String
    strError = "CMCopyFCQuestionsToMembers"
    CMCopyFCQuestionsToMembers pMemberDescription
    
    Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
    Set oFreeEndCut = pMemberDescription.CAO
    
    ' Retrieve the Bounded and Bounding Port objects from the Free End Cut
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oBoundedObject
    
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set bottom flange"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    'pass the correct root selector for members and stiffenrs
    If TypeOf oBoundedPort.Connectable Is ISPSMemberPartPrismatic Then
        'Check if newly Added Part Class exist
        'thereby we can ensure that End user has updated CATALOG
        'with new PartClass needed for new functionality, otherwise
        'point to old rules
        If CheckPartClassExist("MbrEndFlangeSel") Then
            pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.MbrEndFlangeSel", _
                                    "BottomFlange", "No"
        Else
            pCopyAnswerHelper.PutAnswer CUSTOMERID + "Mbr_EndCuts.MbrEndFlangeSel", _
                                    "BottomFlange", "No"
        End If
                                    
    Else 'For stiffeners and beams
        pCopyAnswerHelper.PutAnswer CUSTOMERID + "EndCutRules.FlangeCutSel", _
                                "TheBottomFlange", "No"
    End If
    Set oBoundedPort = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetBottomFlange", strError).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub CMSetBottomFlange(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    sMETHOD = "CMSetBottomFlange"
    Dim strError As String
    strError = "CMCopyFCQuestionsToMembers"
    CMCopyFCQuestionsToMembers pMemberDescription
         
    Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
    Set oFreeEndCut = pMemberDescription.CAO
    
    ' Retrieve the Bounded and Bounding Port objects from the Free End Cut
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oBoundedObject
    
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set bottom flange"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    'pass the correct root selector for members and stiffenrs
    If TypeOf oBoundedPort.Connectable Is ISPSMemberPartPrismatic Then
        'Check if newly Added Part Class exist
        'thereby we can ensure that End user has updated CATALOG
        'with new PartClass needed for new functionality, otherwise
        'point to old rules
        If CheckPartClassExist("MbrEndFlangeSel") Then
            pCopyAnswerHelper.PutAnswer CUSTOMERID + "MbrEndCut.MbrEndFlangeSel", _
                                    "BottomFlange", "Yes"
        Else
            pCopyAnswerHelper.PutAnswer CUSTOMERID + "Mbr_EndCuts.MbrEndFlangeSel", _
                                    "BottomFlange", "Yes"
        End If
    Else
        pCopyAnswerHelper.PutAnswer CUSTOMERID + "EndCutRules.FlangeCutSel", _
                                "TheBottomFlange", "Yes"
    End If
    Set oBoundedPort = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMSetBottomFlange", strError).Number
End Sub
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub CMCopyWCQuestionsToMembers(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    sMETHOD = "CMCopyWCQuestionsToMembers"
    Dim strError As String
    Dim pCopyAnswerHelper As CopyAnswerHelper
    Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
    Set oFreeEndCut = pMemberDescription.CAO
    
    ' Retrieve the Bounded and Bounding Port objects from the Free End Cut
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oBoundedObject
    
    strError = "Calling CopyAnswer to copy assembly questions to endcut selector."
    Set pCopyAnswerHelper = New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    
    'Check if Profile is on Convex side and it intersects base plate
    Dim bIsFromExtendAndConvex As Boolean
    bIsFromExtendAndConvex = IsFreeEndCutFromExtend_ConvexKnuckle(oBoundedObject)
    
    'Get EndCut type
    Dim strEndCutTypeAns As String
    GetSelectorAnswer pMemberDescription.CAO, "EndCutType", strEndCutTypeAns
    
    'for all parts the parent root selector is same
    'pass the appropriate root selector for members , stiffenrs and beams
    If TypeOf oBoundedPort.Connectable Is ISPSMemberPartPrismatic Then
        'Check if newly Added Part Class exist
        'thereby we can ensure that End user has updated CATALOG
        'with new PartClass needed for new functionality, otherwise
        'point to old rules
        If CheckPartClassExist("MbrFreeWebSel") Then
            'No need to push any answers currently
        Else
            pCopyAnswerHelper.CopyAnswer CUSTOMERID + "Mbr_EndCuts.MbrAxisWebSel", QUES_ENDCUTTYPE, _
                                     CUSTOMERID + "FreeEndCutRules.FreeEndCutSel", QUES_ENDCUTTYPE
        End If
    Else 'for Beams and Stiffeners
        If StrComp(strEndCutTypeAns, "Welded", vbTextCompare) = 0 And bIsFromExtendAndConvex Then
            'WebCut is because of 'Extend' Mfg. option, use Box EndCuts selection
            pCopyAnswerHelper.CopyAnswer CUSTOMERID + "EndCutRules.EndToLongBoxWebCutSel", QUES_ENDCUTTYPE, _
                             CUSTOMERID + "FreeEndCutRules.FreeEndCutSel", QUES_ENDCUTTYPE
        Else
            'Regular FreeEndcut
            pCopyAnswerHelper.CopyAnswer CUSTOMERID + "EndCutRules.WebCutSel", QUES_ENDCUTTYPE, _
                             CUSTOMERID + "FreeEndCutRules.FreeEndCutSel", QUES_ENDCUTTYPE
        End If
    End If
    Set oBoundedPort = Nothing
    
    ' Force an Update on the WebCut using the same interface,IJStructGeometry,
    ' as is used when placing the WebCut as an input to the FlangeCuts
    ' This appears to allow Assoc to always recompute the WebCut before FlangeCuts
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Dim pMemberObjects As IJDMemberObjects
    
    strError = "Calling Structdetailobjects.WebCut::ForceUpdateForFlangeCuts"
    Set pMemberObjects = pMemberDescription.CAO
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMemberObjects.Item(1)
    oSDO_WebCut.ForceUpdateForFlangeCuts
    Set oSDO_WebCut = Nothing
    Set pMemberObjects = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyWCQuestionsToMembers", strError).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub CMCopyFCQuestionsToMembers(pMemberDescription As IJDMemberDescription)
    On Error GoTo ErrorHandler
    sMETHOD = "CMCopyFCQuestionsToMembers"
    Dim strError As String
    Dim pCopyAnswerHelper As CopyAnswerHelper
    
    Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
    Set oFreeEndCut = pMemberDescription.CAO
    
    ' Retrieve the Bounded and Bounding Port objects from the Free End Cut
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
    Dim oBoundedPort As IJPort
    Set oBoundedPort = oBoundedObject
    
    strError = "Calling CopyAnswer to copy assembly questions to endcut selector."
    Set pCopyAnswerHelper = New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
     
    'for all parts the parent root selector is same
    'pass the appropriate root selector for members , stiffenrs and beams
    If TypeOf oBoundedPort.Connectable Is ISPSMemberPartPrismatic Then
        'Check if newly Added Part Class exist
        'thereby we can ensure that End user has updated CATALOG
        'with new PartClass needed for new functionality, otherwise
        'point to old rules
        If CheckPartClassExist("MbrEndFlangeSel") Then
            'Not need to push any answers currently
        Else
            pCopyAnswerHelper.CopyAnswer CUSTOMERID + "Mbr_EndCuts.MbrEndFlangeSel", QUES_ENDCUTTYPE, _
                                 CUSTOMERID + "FreeEndCutRules.FreeEndCutSel", QUES_ENDCUTTYPE
        End If
    Else 'for Beams and Stiffeners
        pCopyAnswerHelper.CopyAnswer CUSTOMERID + "EndCutRules.FlangeCutSel", QUES_ENDCUTTYPE, _
                                 CUSTOMERID + "FreeEndCutRules.FreeEndCutSel", QUES_ENDCUTTYPE
    End If
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCopyFCQuestionsToMembers", strError).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Private Function GetCrossSectionType(ByRef pMD As IJDMemberDescription) As String
    On Error GoTo ErrorHandler
    sMETHOD = "GetCrossSectionType"
    
    GetCrossSectionType = ""
    
    ' Initialize wrapper class and get the bounded profile
    sError = "Setting Assembly Connection "
    Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
    Set oFreeEndCut = pMD.CAO
     
    ' Retrieve the Bounded and Bounding Port objects from the Free End Cut
    Dim oBoundedObject As Object
    Dim oBoundingObject As Object
    oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
    
    ' From the Bounded Port, retrieve the Port's IJConnectable object
    If TypeOf oBoundedObject Is IJPort Then
        Dim oPort As IJPort
        Set oPort = oBoundedObject
    
        Dim oBoundedProfile As IJConnectable
        Set oBoundedProfile = oPort.Connectable
        
        ' Return Cross Section type from the Connectable object
        If TypeOf oBoundedProfile Is IJStiffener Then
            Dim oProfile As StructDetailObjects.ProfilePart
            Set oProfile = New StructDetailObjects.ProfilePart
            Set oProfile.object = oBoundedProfile
            
            GetCrossSectionType = oProfile.sectionType
            
            Set oProfile = Nothing
            Set oBoundedProfile = Nothing
        ElseIf TypeOf oBoundedProfile Is IJBeam Then
            Dim oBeam As StructDetailObjects.BeamPart
            Set oBeam = New StructDetailObjects.BeamPart
            Set oBeam.object = oBoundedProfile
            
            GetCrossSectionType = oBeam.sectionType
            
            Set oBeam = Nothing
            Set oBoundedProfile = Nothing
        ElseIf TypeOf oBoundedProfile Is ISPSMemberPartPrismatic Then
            Dim oMemberPart As StructDetailObjects.MemberPart
            Set oMemberPart = New StructDetailObjects.MemberPart
            Set oMemberPart.object = oBoundedProfile
            
            GetCrossSectionType = oMemberPart.sectionType
        
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

'********************************************************************
' Routine: CheckPartClassExist
' Description:  Check if the given PartClass exist in Catalog
'
'********************************************************************
Private Function CheckPartClassExist(sPartClassName As String) As Boolean
    Const sMETHOD = "CheckPartClassExist"
    On Error GoTo ErrorHandler
    
    Dim oCatalogQuery As IJSRDQuery
    Dim oSmartQuery As IJSmartQuery
    Dim oPartClass As IJSmartClass
    
    Set oCatalogQuery = New SRDQuery
    Set oSmartQuery = oCatalogQuery
    
    ' Query for SmartClass... Check if exist
    On Error Resume Next
    Set oPartClass = oSmartQuery.GetClassByName(sPartClassName)
    On Error GoTo ErrorHandler
    
    If Not oPartClass Is Nothing Then
        CheckPartClassExist = True
    End If
    
    Set oPartClass = Nothing
    Set oSmartQuery = Nothing
    Set oCatalogQuery = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
    
End Function
