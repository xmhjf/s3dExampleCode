VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KnuckleReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : MfgProfileCustomReports
''  File        : KnuckleReport.cls
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

Private Const MODULE = "MfgProfileCustomReports.KnuckleReport"
Private Const IJMFGPROFILEPARTGUID = "{1BEB9DD4-3B5D-4571-AEFA-4DC8B9C21434}"
Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
    On Error GoTo ErrorHandler
       
    If pElements.Count > 0 Then

        Dim oFSO As FileSystemObject
        Set oFSO = New FileSystemObject
        
        Dim oTextStream As TextStream
        
        Set oTextStream = oFSO.OpenTextFile(strFileName, ForAppending, True)
        oTextStream.WriteLine "Profile Name,Assembly Path,Manufacturability,Section Type,Section Name,Side,Knuckle Angle"
                
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

    Dim indKnuckle As Long
    
    On Error Resume Next
    Dim oNamedItem As IJNamedItem
    Set oNamedItem = oProfileObj
    
    ' Use the profile part as input
    Dim oProfileClass As StructDetailObjects.ProfilePart
    
    Set oProfileClass = New StructDetailObjects.ProfilePart
    
    Set oProfileClass.object = oProfileObj
    
    Dim vKnuckleAngleValue As Variant
    Dim bstrSectionType As String
    Dim bstrAssemblyPath As String
    Dim bstrSectionName As String
    Dim bstrSide As String
    
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
    
    Dim oProfilePartSupport As IJProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Dim oBendProfileKnuckles As JCmnShp_CollectionAlias, oKnuckleLines As JCmnShp_CollectionAlias, oKnuckleAnglesInDegrees As JCmnShp_CollectionAlias
    Dim sProfileKnuckleInfo As String, sManufacturable As String
    
    Set oProfilePartSupport = New ProfilePartSupport
    Set oPartSupport = oProfilePartSupport
    Set oPartSupport.Part = oProfileObj
    
    oProfilePartSupport.GetBendProfileKnuckleMfgData JXSEC_WEB_LEFT, oBendProfileKnuckles, oKnuckleLines, oKnuckleAnglesInDegrees
    
    indKnuckle = 0
    
    If Not oBendProfileKnuckles Is Nothing Then
        If oBendProfileKnuckles.Count > 0 Then
            For indKnuckle = 1 To oBendProfileKnuckles.Count
                vKnuckleAngleValue = oKnuckleAnglesInDegrees.Item(indKnuckle)
                sManufacturable = ",NO,"
                
                If vKnuckleAngleValue < 180# Then
                    bstrSide = "WEB_RIGHT"
                    If ((vKnuckleAngleValue > 140) And (vKnuckleAngleValue < 176)) Then
                        sManufacturable = ",YES,"
                    End If
                Else
                    bstrSide = "WEB_LEFT"
                    vKnuckleAngleValue = 360# - vKnuckleAngleValue
                    If vKnuckleAngleValue > 176# Then
                        sManufacturable = ",YES,"
                    End If
                End If
                sProfileKnuckleInfo = oNamedItem.Name & "," & bstrAssemblyPath & sManufacturable & bstrSectionType & "," & bstrSectionName & "," & bstrSide & "," & CStr(vKnuckleAngleValue)
            Next
        End If
        oTextStream.WriteLine sProfileKnuckleInfo
    End If

    Set oBendProfileKnuckles = Nothing
    Set oKnuckleLines = Nothing
    Set oKnuckleAnglesInDegrees = Nothing
    
    oProfilePartSupport.GetBendProfileKnuckleMfgData JXSEC_TOP_FLANGE_TOP, oBendProfileKnuckles, oKnuckleLines, oKnuckleAnglesInDegrees
    
    indKnuckle = 0
    
    If Not oBendProfileKnuckles Is Nothing Then
        If oBendProfileKnuckles.Count > 0 Then
            For indKnuckle = 1 To oBendProfileKnuckles.Count
                vKnuckleAngleValue = oKnuckleAnglesInDegrees.Item(indKnuckle)
                sManufacturable = ",NO,"
                
                If vKnuckleAngleValue < 180# Then
                    bstrSide = "FLANGE BOTTOM"
                    If ((vKnuckleAngleValue > 1177)) Then
                        sManufacturable = ",YES,"
                    End If
                Else
                    bstrSide = "FLANGE TOP"
                    vKnuckleAngleValue = 360# - vKnuckleAngleValue
                    If vKnuckleAngleValue > 179# Then
                        sManufacturable = ",YES,"
                    End If
                End If
                sProfileKnuckleInfo = oNamedItem.Name & "," & bstrAssemblyPath & sManufacturable & bstrSectionType & "," & bstrSectionName & "," & bstrSide & "," & CStr(vKnuckleAngleValue)
            Next
        End If
        oTextStream.WriteLine sProfileKnuckleInfo
    End If
    
    Set oProfilePartSupport = Nothing
    
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


