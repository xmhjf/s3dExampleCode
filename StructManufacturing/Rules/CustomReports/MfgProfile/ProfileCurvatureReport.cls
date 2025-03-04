VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurvatureReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : MfgProfileCustomReports
''  File        : ProfileCurvatureReport.cls
''
''  Description :
''
''
''  Author      : Intergraph
''
''  History     :
''               Initial Creation   -
''
''
''**************************************************************************************

Implements IJDCustomReport

Private Const MODULE = "MfgProfileCustomReports.CurvatureReport"
Private Const IJMFGPROFILEPARTGUID = "{1BEB9DD4-3B5D-4571-AEFA-4DC8B9C21434}"
Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
    On Error GoTo ErrorHandler
       
    If pElements.Count > 0 Then

        Dim oFSO As FileSystemObject
        Set oFSO = New FileSystemObject
        
        Dim oTextStream As TextStream
        
        Set oTextStream = oFSO.OpenTextFile(strFileName, ForAppending, True)
        oTextStream.WriteLine "Profile Name,Assembly Path,Manufacturability,Section Type,Section Name,Start Point X,Start Point Y,Start Point Z,End Point X,End Point Y,End Point Z,Curvature Type,Depth Of BowString(mm),Radius Of BowString(mm),Maximum Depth Of BowString(mm),Maximum Radius Of BowString(mm)"
                
        Dim oSelectedObj As Object
        For Each oSelectedObj In pElements
            If TypeOf oSelectedObj Is IJProfilePart Then
              ReportProfilePartInformation oSelectedObj, oTextStream
            ElseIf TypeOf oSelectedObj Is IJAssembly Then
                Dim oSelectedProfileObj As Object
                Dim oAllProfiles As IJElements
                Set oAllProfiles = GetAllProfilePartsFromAssembly(oSelectedObj)
                For Each oSelectedProfileObj In oAllProfiles
                    If TypeOf oSelectedProfileObj Is IJProfilePart Then
                      ReportProfilePartInformation oSelectedProfileObj, oTextStream
                    End If
                Next
                Set oAllProfiles = Nothing
            End If
        Next
        eCustomReportStatus = StrMfgProcessFinished
    End If
    
Wrapup:
    If Not oTextStream Is Nothing Then
        oTextStream.Close
        Set oTextStream = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not oTextStream Is Nothing Then
        oTextStream.Close
        Set oTextStream = Nothing
    End If
    
    eCustomReportStatus = StrMfgErrorUnknown
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub ReportProfilePartInformation(oProfileObj As Object, oTextStream As TextStream)
 Const METHOD = "ReportProfilePartInformation"
    On Error GoTo ErrorHandler

    Dim nNumberOfSections As Long
    Dim dPointCoordinates() As Double
    Dim lConvexConcaveInfo() As Long
    Dim dDepthOfBowString() As Double
    Dim dRadiusOfBowString() As Double
    Dim dMaxDepthOfBowString() As Double
    Dim dMaxRadiusOfBowString() As Double
    Dim indSection As Long
    
    On Error Resume Next
    Dim oNamedItem As IJNamedItem
    Set oNamedItem = oProfileObj
    
    Dim oProfilePartHlpr As New ProfilePartHlpr
    Set oProfilePartHlpr.object = oProfileObj
    ' Use the profile part as input
    Dim oProfileClass As StructDetailObjects.ProfilePart
    
    Set oProfileClass = New StructDetailObjects.ProfilePart
    
    Set oProfileClass.object = oProfileObj
    
    Dim bstrSectionType As String
    Dim bstrViewName As String
    Dim bstrMin As String
    Dim bstrMax As String
    Dim bstrAssemblyPath As String
    Dim bstrSectionName As String
    
    Dim oAssemblyChild As IJAssemblyChild
    
    Set oAssemblyChild = oProfileObj
    
StartAgain:
    If Not oAssemblyChild Is Nothing Then
        Dim oAssemblyParentUnk As IUnknown
        Set oAssemblyParentUnk = oAssemblyChild.Parent
        If TypeOf oAssemblyParentUnk Is IJAssembly Then
            Dim oAssemblyParentNamedItem As IJNamedItem
            Set oAssemblyParentNamedItem = oAssemblyParentUnk
            If Len(bstrAssemblyPath) = 0 Then
                bstrAssemblyPath = oAssemblyParentNamedItem.Name
            Else
                bstrAssemblyPath = oAssemblyParentNamedItem.Name & " | " & bstrAssemblyPath
            End If
            If TypeOf oAssemblyParentUnk Is IJAssemblyChild Then
                Set oAssemblyChild = oAssemblyParentUnk
                GoTo StartAgain
            End If
            Set oAssemblyParentNamedItem = Nothing
        End If
        Set oAssemblyParentUnk = Nothing
    End If
    
    bstrSectionName = oProfileClass.SectionName
    
    bstrSectionType = oProfileClass.SectionType
    If oProfileClass.IsCrossSectionABuiltUp = True Then
        bstrViewName = "JUAMfgProfileCurvatureLimBU"
        bstrMin = "BuiltUpSizeMin"
        bstrMax = "BuiltUpSizeMax"
    ElseIf bstrSectionType = "UA" Or bstrSectionType = "B" Then
        bstrViewName = "JUAMfgProfileCurvatureLimAB"
        bstrMin = "WebSizeMin"
        bstrMax = "WebSizeMax"
    End If
    
    oProfilePartHlpr.GetConvexConcaveCurvatureInfo 1#, Nothing, nNumberOfSections, _
                dPointCoordinates, lConvexConcaveInfo, dDepthOfBowString, dRadiusOfBowString, dMaxDepthOfBowString, dMaxRadiusOfBowString
    
    Dim sProfileCurvatureInfo As String
       
    For indSection = 0 To (nNumberOfSections - 1)
        Dim sManufacturable As String
         
        If ((lConvexConcaveInfo(indSection) <> 1) And (Len(bstrViewName) > 0)) Then
            Dim dWebLength As Double
            Dim oQueryOutputValues() As Variant
            Dim strQuery As String
            Dim oMfgCatalogQueryHelper As IJMfgCatalogQueryHelper
            
            dWebLength = oProfileClass.WebLength
            
            Set oMfgCatalogQueryHelper = New MfgCatalogQueryHelper
            
            strQuery = "SELECT BowStringDepth FROM " & bstrViewName & " WHERE ( (CurvatureType = " & CStr(lConvexConcaveInfo(indSection)) & ") and (" & CStr(dWebLength) & " > " & bstrMin & ") and (" & CStr(dWebLength) & " < " & bstrMax & " ) and (" & CStr(dMaxRadiusOfBowString(indSection)) & " < RadiusOfCurvature ) )"
            'MsgBox "strQuery = " & strQuery
            oQueryOutputValues = oMfgCatalogQueryHelper.GetValuesFromDBQuery(strQuery)
            'MsgBox "UBound(oQueryOutputValues ) = " & UBound(oQueryOutputValues)
            If (UBound(oQueryOutputValues) >= 0) Then
                'MsgBox oQueryOutputValues(0)
                If (dMaxRadiusOfBowString(indSection) > oQueryOutputValues(0)) Then
                    sManufacturable = ",NO"
                End If
            End If
        End If
        
        If (Len(sManufacturable) = 0) Then
            sManufacturable = ",YES"
        End If
                
        sProfileCurvatureInfo = oNamedItem.Name & "," & bstrAssemblyPath & sManufacturable & "," & bstrSectionType & "," & bstrSectionName & "," & Round((dPointCoordinates(3 * indSection)), 6) _
                    & "," & Round((dPointCoordinates(3 * indSection + 1)), 6) & "," & Round((dPointCoordinates(3 * indSection + 2)), 6) _
                    & "," & Round((dPointCoordinates(3 * indSection + 3)), 6) & "," & Round((dPointCoordinates(3 * indSection + 4)), 6) & "," & Round((dPointCoordinates(3 * indSection + 5)), 6)
                            
        If lConvexConcaveInfo(indSection) = 1 Then
            sProfileCurvatureInfo = sProfileCurvatureInfo + ",STRAIGHT,"
        ElseIf lConvexConcaveInfo(indSection) = 2 Then
            sProfileCurvatureInfo = sProfileCurvatureInfo + ",CONVEX,"
        ElseIf lConvexConcaveInfo(indSection) = 3 Then
            sProfileCurvatureInfo = sProfileCurvatureInfo + ",CONCAVE,"
        End If
        
        sProfileCurvatureInfo = sProfileCurvatureInfo & Round((dDepthOfBowString(indSection) * 1000), 4) & "," _
                    & Round((dRadiusOfBowString(indSection) * 1000), 4) & "," & Round((dMaxDepthOfBowString(indSection) * 1000), 4) & "," _
                    & Round((dMaxRadiusOfBowString(indSection) * 1000), 4)
                   
        oTextStream.WriteLine sProfileCurvatureInfo
    Next
    
    Set oProfilePartHlpr = Nothing
    
CleanUp:
  

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function GetAllProfilePartsFromAssembly(ByVal oElem As GSCADAsmHlpers.IJAssembly) As IJElements
    Const METHOD = "GetAllProfilePartsFromAssembly"
    
    On Error GoTo ErrorHandler
    
    Set GetAllProfilePartsFromAssembly = New JObjectCollection
    
    'Get children of assembly
    Dim oAssembly As IJAssembly
    Set oAssembly = oElem
    
    Dim oChildren As IJDTargetObjectCol
    Set oChildren = oAssembly.GetChildren
    
    'Only consider assemblies with assembly children
    If oChildren.Count > 1 Then
        Dim Index As Long
        Dim oItem As Object
                    
        'Get each assembly child and add child to elements list
        For Index = 1 To oChildren.Count
            'Get next item
            Set oItem = oChildren.Item(Index)
                                           
            'Add item to elements list if type of item
            'is IJPlatePart and not MfgPlatePart
            If TypeOf oItem Is IJProfilePart Then
                GetAllProfilePartsFromAssembly.Add oItem
            ElseIf TypeOf oItem Is IJAssembly Then 'Assembly or block is child
                'get parts in this assembly
                GetAllProfilePartsFromAssembly.AddElements GetAllProfilePartsFromAssembly(oItem)
            End If
        Next
    End If

CleanUp:
    'Clean up
    Set oAssembly = Nothing
    Set oChildren = Nothing
    Set oItem = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
    GoTo CleanUp
End Function
