VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultGroupNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2010, Intergraph Corporation.  All rights reserved.
'
'  Project: CmnPartNaming
'
'  Abstract:
'
'  Author:
'******************************************************************

Option Explicit
Implements IJNameRule
                                                        
Private Const Module = "DefaultGroupNameRule: "
Private Const MODELDATABASE = "Model"

Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors       ' To collect and propagate the errors.
Private m_oError As IJEditError         ' The error to raise.

Private Sub Class_Initialize()
   Set m_oErrors = CreateObject(ERRORPROGID)
End Sub

Private Sub Class_Terminate()
   Set m_oErrors = Nothing
   Set m_oError = Nothing
End Sub

Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
Const METHOD = "IJNameRule_ComputeName"
On Error GoTo ErrorHandler
    
    Dim oCommonPartGroup As IJCommonPartGroup
    Dim oCPM             As IJCommonPartManager
    Dim oFolderColl      As IJElements
    Dim lnumOfGoups      As Long
    Dim lProcessControl  As Long
    Dim oCommonParts     As IJElements
    Dim oCommonPart      As Object
    Dim strName          As String
    Dim oNamedItem       As IJNamedItem
    
    Dim oMembPart        As SPSMembers.ISPSMemberPartPrismatic
    Dim strXsectionName  As String
    Dim lCount           As Long
    
    Set oCommonPartGroup = pEntity

    If oCommonPartGroup Is Nothing Then GoTo ErrorHandler

    'Get the first object under CPG
    lProcessControl = oCommonPartGroup.ProcessControl
    oCommonPartGroup.GetCommonParts oCommonParts
    
     If lProcessControl = 2 Then
        lnumOfGoups = 0
        oCommonPartGroup.GetManager oCPM
        oCPM.GetCommonPartFolders oFolderColl, 1
    End If
    
    If Not oCommonParts Is Nothing Then
        If oCommonParts.Count > 0 Then
                    
            Set oCommonPart = oCommonParts.Item(1)
            
            'StandardByXML or 'StandardByModel
            If (oCommonPartGroup.ProcessPurpose = 2 Or _
               oCommonPartGroup.ProcessPurpose = 3) And _
               oCommonPartGroup.StandardReferenceEntityName <> vbNullString Then
               
               Set oNamedItem = oCommonPartGroup
               oNamedItem.Name = oCommonPartGroup.StandardReferenceEntityName
                        
            ElseIf TypeOf oCommonPart Is IJPlatePart And _
                Not TypeOf oCommonPart Is IJCollarPart And _
                Not IsBracket(oCommonPart) Then
                
                    strName = GetGroupNameFromHierarchy(oCommonPart)
            
                    Set oNamedItem = oCommonPartGroup
                    'if CPG is Manually Created add the index
                    If lProcessControl = 2 Then
                         
                        If Not oFolderColl Is Nothing Then
                            lnumOfGoups = lNumberofGroups(oFolderColl.Count, oFolderColl)
                            oNamedItem.Name = strName & "_" & "UserDefined_" & lnumOfGoups
                        End If
                         
                    Else
                        oNamedItem.Name = strName
                     
                    End If
            
            ElseIf TypeOf oCommonPart Is IJProfilePart Then
                'Using the same logic as of the PlatePart Group
                strName = GetGroupNameFromHierarchy(oCommonPart)
                
                Set oNamedItem = oCommonPartGroup
                'if CPG is Manually Created add the index
                If lProcessControl = 2 Then
                        
                    If Not oFolderColl Is Nothing Then
                        lnumOfGoups = lNumberofGroups(oFolderColl.Count, oFolderColl)
                        oNamedItem.Name = strName & "_" & "UserDefined_" & lnumOfGoups
                    End If
                        
                Else
                    oNamedItem.Name = strName
                    
                End If
                
            ElseIf TypeOf oCommonPart Is IJCollarPart Then
                
                If IsStandardCollar(oCommonPart) Then
                    strName = GetCollarGroupName(oCommonPart)
                Else
                    strName = GetGroupNameFromHierarchy(oCommonPart)
                End If
                
                Set oNamedItem = oCommonPartGroup
                'if CPG is Manually Created add the index
                If lProcessControl = 2 Then
                        
                    If Not oFolderColl Is Nothing Then
                        lnumOfGoups = lNumberofGroups(oFolderColl.Count, oFolderColl)
                        oNamedItem.Name = strName & "_" & "UserDefined_" & lnumOfGoups
                    End If
                        
                Else
                    oNamedItem.Name = strName
                    
                End If
            ElseIf IsBracket(oCommonPart) Then
                            
                If IsStandardBracket(oCommonPart) Then
                    strName = GetBracketGroupName(oCommonPart)
                Else
                    strName = GetGroupNameFromHierarchy(oCommonPart)
                End If
                
                Set oNamedItem = oCommonPartGroup
                'if CPG is Manually Created add the index
                If lProcessControl = 2 Then
                        
                    If Not oFolderColl Is Nothing Then
                        lnumOfGoups = lNumberofGroups(oFolderColl.Count, oFolderColl)
                        oNamedItem.Name = strName & "_" & "UserDefined_" & lnumOfGoups
                    End If
                        
                Else
                    oNamedItem.Name = strName
                    
                End If
            
            ElseIf TypeOf oCommonPart Is ISPSMemberPartPrismatic Then
                
                Set oMembPart = oCommonPart
                strXsectionName = oMembPart.CrossSection.SectionName
                                
                lCount = GetNameCount(oCommonPart)
                strName = strXsectionName & "-0" & lCount
                
                Set oNamedItem = oCommonPartGroup
                'if CPG is Manually Created add the index
                If lProcessControl = 2 Then
                        
                    If Not oFolderColl Is Nothing Then
                        lnumOfGoups = lNumberofGroups(oFolderColl.Count, oFolderColl)
                        oNamedItem.Name = strName & "_" & "UserDefined_" & lnumOfGoups
                    End If
                        
                Else
                    oNamedItem.Name = strName
                    
                End If
                
            ElseIf TypeOf oCommonPart Is IJAssemblyBase Then
            
                Set oNamedItem = oCommonPart
            
                lCount = GetNameCount(oCommonPart)
                strName = oNamedItem.Name & "-0" & lCount
                
                Set oNamedItem = oCommonPartGroup
                'if CPG is Manually Created add the index
               If lProcessControl = 2 Then
                        
                    If Not oFolderColl Is Nothing Then
                        lnumOfGoups = lNumberofGroups(oFolderColl.Count, oFolderColl)
                        oNamedItem.Name = strName & "_" & "UserDefined_" & lnumOfGoups
                    End If
                        
                Else
                    oNamedItem.Name = strName
                    
                End If

                    
            End If
            
        End If
    End If
    
CleanUp:
    Set oNamedItem = Nothing
    Set oCommonPart = Nothing
    Set oCommonPartGroup = Nothing
    Set oCommonParts = Nothing
    Set oFolderColl = Nothing
    Set oCPM = Nothing
    
    Exit Sub
    
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    GoTo CleanUp
End Sub


Private Function GetGroupNameFromHierarchy(oCommonPart As Object) As String
Const METHOD = "GetGroupNameFromHierarchy"
On Error GoTo ErrorHandler

    Dim oAssemblychild      As IJAssemblyChild
    Dim oParent(0 To 3)     As IJAssembly
    Dim i                   As Integer
    Dim oNamedItem          As IJNamedItem
    Dim j                   As Integer
    Dim strName             As String
    Dim lCount              As Long
    
    Set oAssemblychild = oCommonPart
        
    If oAssemblychild Is Nothing Then
        Exit Function
    End If
    
    'Get the parent Assemblies upto 4 hierarchy levels
    For i = 0 To 3
        
        If Not oAssemblychild.Parent Is Nothing Then
            If TypeOf oAssemblychild.Parent Is IJAssemblyBase Then
                Set oParent(i) = oAssemblychild.Parent
                
                If TypeOf oParent(i) Is IJBlock Then
                    Exit For
                End If
            Else
                'parent is config root
                Exit For
            End If
            
            If i = 3 Then Exit For
            
        End If
        
        Set oAssemblychild = oAssemblychild.Parent
             
    Next
    
    ' Add the name of parent objects
    For j = i To 0 Step -1
    
        If Not oParent(j) Is Nothing Then
        
            Set oNamedItem = oParent(j)
            If strName = vbNullString Then
                strName = oNamedItem.Name
            Else
                strName = strName & "-" & oNamedItem.Name
            End If
        End If
        
        Set oNamedItem = Nothing
        
    Next
    
    For j = 1 To (3 - i)
        strName = strName & "-0"
    Next
    
    lCount = GetNameCount(oCommonPart)
    
    strName = strName & "-0" & lCount
    
    GetGroupNameFromHierarchy = strName

CleanUp:
    Set oAssemblychild = Nothing
    
    Exit Function
    
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    GoTo CleanUp
    
End Function


Private Function GetNameCount(oSrcObject As Object) As Long
Const METHOD = "GetNameCount"
On Error GoTo ErrorHandler
    
    Dim jContext            As IJContext
    Dim oDBTypeConfig       As IJDBTypeConfiguration
    Dim oConnectMiddle      As IJDAccessMiddle
    Dim oModelResourceMgr   As IUnknown
    Dim oChildName          As IJNamedItem
    Dim oNameCounter        As IJNameCounter
    Dim strPartname         As String
    Dim strModelDBID        As String
    Dim strLocationID       As String
    
    'Get the middle context
    Set jContext = GetJContext()
    
    'First get resourcemanager frem modeldatabase where objects are to be stored
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
        
    ' Get Typestring from IJNamedItem
    Set oChildName = oSrcObject
    strPartname = oChildName.TypeString
    
    ' Use part Name for basename, remove blanks:
    strPartname = Join(Split(strPartname), "")
    
    ' Get a string from the namecounter
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    GetNameCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocationID)
    
    ' To avoid '1'
    GetNameCount = GetNameCount + 1

CleanUp:
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oModelResourceMgr = Nothing
    Set oNameCounter = Nothing
    
    Exit Function
    
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    GoTo CleanUp

End Function


Private Function GetBracketGroupName(oBracketPart As Object) As String
Const METHOD = "GetBracketGroupName"
On Error GoTo ErrorHandler
    
    Dim oSmartPlate                 As IJSmartPlate
    Dim oSDOBracket                 As Object
    Dim sObjectType                 As String
    Dim strBracketItemName          As String
    
    Set oSDOBracket = CreateObject("StructDetailObjects.Bracket")

    If TypeOf oBracketPart Is IJSmartPlate Then
        Set oSmartPlate = oBracketPart
    
        If Not oSmartPlate Is Nothing Then ' For StructDetailing Brackets

            Set oSDOBracket.object = oSmartPlate
            GetBracketGroupName = oSDOBracket.ItemName
        End If
    Else  'This is class name and not part name
        GetBracketGroupName = GetBracketItemName(oBracketPart)
    End If
    
    
CleanUp:
    Set oSmartPlate = Nothing
    Set oSDOBracket = Nothing
    
    Exit Function
    
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    GoTo CleanUp

End Function


Private Function GetCollarGroupName(oCollarPart As Object) As String
Const METHOD = "GetCollarGroupName"
On Error GoTo ErrorHandler

    Dim oSO                     As IJSmartOccurrence
    Dim oSmartItem              As IJSmartItem
    Dim strCmnPartName          As String
    Dim oPlate                  As IJPlate
    Dim oStructMat              As IJStructureMaterial
    Dim pIJCollarOps            As IJSDCollarAttributes
    Dim strRootClassName        As String
    Dim oBasePlate              As Object
    Dim oPenetratingProfile     As IJNamedItem
    Dim oSlot                   As Object
    Dim strGrade                As String
    
    Set oSO = oCollarPart
    Set oSmartItem = oSO.ItemObject
    
    strCmnPartName = oSmartItem.Name
    
    Set oPlate = oCollarPart
    strCmnPartName = strCmnPartName + "_" + CStr(Round(oPlate.thickness * 1000, 1))
    
    Set oStructMat = oCollarPart
    strGrade = oStructMat.MaterialGrade
    
    If strGrade <> "A" Then
        strCmnPartName = strCmnPartName + "'" + oStructMat.MaterialGrade + "'"
    End If
    
    Set pIJCollarOps = New SDCollarUtils
    pIJCollarOps.GetInput_Collar oCollarPart, oBasePlate, oPenetratingProfile, strRootClassName, oSlot
    
    strCmnPartName = strCmnPartName + "(" + oPenetratingProfile.Name + ")"
    
    GetCollarGroupName = strCmnPartName

CleanUp:
    Set oSO = Nothing
    Set oSmartItem = Nothing
    Set oPlate = Nothing
    Set oStructMat = Nothing
    Set pIJCollarOps = Nothing
    Set oPenetratingProfile = Nothing
    
    Exit Function
    
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    GoTo CleanUp

End Function

Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
Const METHOD = "IJNameRule_GetNamingParents"
On Error GoTo ErrorHandler

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

Exit Function
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise

End Function

Private Function lNumberofGroups(lFolderCount As Long, oFolderColl As IJElements) As Long
Const METHOD = "lNumberofGroups"
On Error GoTo ErrorHandler
    Dim oGroupColl       As IJElements
    Dim i                As Long
    Dim oCpFolder        As IJCommonPartFolder
    Dim lnumOfGoups      As Long
     
     
     For i = 1 To lFolderCount
        Set oCpFolder = oFolderColl.Item(i)
        oCpFolder.GetCommonPartGroups oGroupColl
        If oGroupColl.Count > 0 Then
            lnumOfGoups = lnumOfGoups + oGroupColl.Count
        End If
        
        Set oCpFolder = Nothing
        Set oGroupColl = Nothing
    Next
    
    lNumberofGroups = lnumOfGoups
    
Exit Function
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise

End Function

