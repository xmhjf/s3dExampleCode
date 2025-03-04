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

Private sERROR As String
Private sMETHOD As String
Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\FreeEndCutRules\FreeEndCutSelCM.cls"
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
           
    ' Create webcut
    Dim oWebcut As Structdetailobjects.WebCut
    Set oWebcut = New Structdetailobjects.WebCut
    sERROR = "Creating web cut."
    oWebcut.Create pResourceManager, _
                    oBoundingObject, oBoundedObject, _
                   "WebCuts", oFreeEndCut
                   
                               
    sERROR = "Return the webcut"
    Set pObject = oWebcut.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number

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
    sERROR = "Creating flange cut."
    
    ' create the flange when the web cut is found
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oWebcut As IJStructFeature
    
    ' all Item Members are not IJStructFeature objects,
    Dim iIndex As Long
    Dim iCount As Long
    Dim bFound As Boolean
    bFound = False
    iCount = pMemberObjects.Count
    
    For iIndex = 1 To iCount
        If TypeOf pMemberObjects.Item(iIndex) Is IJStructFeature Then
            Set oWebcut = pMemberObjects.Item(iIndex)
            If (oWebcut.get_StructFeatureType = SF_WebCut) Then
                oFlangeCut.Create pResourceManager, _
                                  oBoundingObject, oBoundedObject, oWebcut, _
                                  "FlangeCuts", oFreeEndCut
                Set oWebcut = Nothing
                
                sERROR = "Return the flange cut "
                Set pObject = oFlangeCut.object
                Exit For
            End If
            
            Set oWebcut = Nothing
        End If
    Next iIndex
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    
    Dim strSectionType As String
    strSectionType = GetCrossSectionType(pMD)
        
    bIsNeeded = True
    If ((strSectionType = "FB") Or (strSectionType = "HalfR") Or _
        (strSectionType = "P") Or (strSectionType = "R") Or _
        (strSectionType = "SB") Or (strSectionType = "SqTu") Or _
        (strSectionType = "RT")) Then bIsNeeded = False
            
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
   
    Dim strSectionType As String
    strSectionType = GetCrossSectionType(pMD)
        
    bIsNeeded = False
    If ((strSectionType = "I") Or (strSectionType = "ISType") Or _
        (strSectionType = "C_SS") Or (strSectionType = "CSType")) Then
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
          
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set bottom flange"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.PutAnswer "EndCutRules.FlangeCutSel", _
                                "TheBottomFlange", "No"
    
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
          
    ' Pass the collar order of secondary for the creation of this collar
    strError = "Set bottom flange"
    
    Dim pCopyAnswerHelper As New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.PutAnswer "EndCutRules.FlangeCutSel", _
                                "TheBottomFlange", "Yes"
    
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
    
    strError = "Calling CopyAnswer to copy assembly questions to endcut selector."
    Set pCopyAnswerHelper = New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.CopyAnswer "EndCutRules.WebCutSel", QUES_ENDCUTTYPE, _
                                 "FreeEndCutRules.FreeEndCutSel", QUES_ENDCUTTYPE

    ' Force an Update on the WebCut using the same interface,IJStructGeometry,
    ' as is used when placing the WebCut as an input to the FlangeCuts
    ' This appears to allow Assoc to always recompute the WebCut before FlangeCuts
    Dim oSDO_WebCut As Structdetailobjects.WebCut
    Dim pMemberObjects As IJDMemberObjects
    
    strError = "Calling Structdetailobjects.WebCut::ForceUpdateForFlangeCuts"
    Set pMemberObjects = pMemberDescription.CAO
    Set oSDO_WebCut = New Structdetailobjects.WebCut
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
    
    strError = "Calling CopyAnswer to copy assembly questions to endcut selector."
    Set pCopyAnswerHelper = New CopyAnswerHelper
    Set pCopyAnswerHelper.MemberDescription = pMemberDescription
    pCopyAnswerHelper.CopyAnswer "EndCutRules.FlangeCutSel", QUES_ENDCUTTYPE, _
                                 "FreeEndCutRules.FreeEndCutSel", QUES_ENDCUTTYPE

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
    sERROR = "Setting Assembly Connection "
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
            Dim oProfile As Structdetailobjects.ProfilePart
            Set oProfile = New Structdetailobjects.ProfilePart
            Set oProfile.object = oBoundedProfile
            
            GetCrossSectionType = oProfile.SectionType
            
            Set oProfile = Nothing
            Set oBoundedProfile = Nothing
        Else
            Dim oBeam As Structdetailobjects.BeamPart
            Set oBeam = New Structdetailobjects.BeamPart
            Set oBeam.object = oBoundedProfile
            
            GetCrossSectionType = oBeam.SectionType
            
            Set oBeam = Nothing
            Set oBoundedProfile = Nothing
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Function

