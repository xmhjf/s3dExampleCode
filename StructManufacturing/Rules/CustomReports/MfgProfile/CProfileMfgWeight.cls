VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrMfgWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : CustomReports
''  File        : CProfileMfgWeight.cls
''
''  Description : Populates a csv file with Profile detailing weight and mfg weight
''
''
''  Author      : Intergraph Development.
''
''  History     :
''               Initial Creation   -
''
''
''**************************************************************************************
Implements IJDCustomReport

Private Const Module = "MfgProfileCustomReports.CPrMfgWeight"
Private Const TOLERANCE = 2

'- MfgProfile ----------------------------------------------------'
Private Const IJMfgProfilePart = "{1BEB9DD4-3B5D-4571-AEFA-4DC8B9C21434}"
Private Declare Function GetTempPath Lib "kernel32" _
         Alias "GetTempPathA" (ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long
         
Private m_iCount As Long

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
            
    If pElements.Count > 0 Then
        Dim objPinJig As Object
        Dim oSelectedObj As Object
        
        'Open log file
        Dim oStream As TextStream
        Dim oFSO As FileSystemObject
        Set oFSO = New FileSystemObject
        Set oStream = oFSO.OpenTextFile(strFileName, ForAppending, True)
        Set oFSO = Nothing
        oStream.WriteLine "UNIT,kg"
        oStream.WriteLine "Margin is applied (if any)."
        'oStream.WriteLine "Part Name,Assembly,Dry Weight,Raw Mfg Weight,Bevels Applied,Chamfers Applied,Bevels + Chamfers Applied,Holes Applied,Final Mfg Weight"
        oStream.WriteLine "Part Name,Assembly,Cross Section,Dry Weight,Final Mfg Weight,Web Weight,Flange Weight,Non Trimmed Margin Applied"

        For Each oSelectedObj In pElements
            If TypeOf oSelectedObj Is IJMfgProfilePart Then
                Dim oMfgProfile           As IJMfgProfilePart
                Dim oProfilePart          As IJProfilePart
                
                ReportWeightInformation oSelectedObj, oStream
                
            ElseIf TypeOf oSelectedObj Is IJProfilePart Then
                'This is Profile. Get Mfg Profile Part from Profile
                
                Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
                Dim oColl As IJElements
                Set oColl = oStructMfgGlobals.GetMfgPart(oSelectedObj, IJMfgProfilePart)
                
                If Not oColl Is Nothing Then
                    If oColl.Count > 0 Then
                        Dim oMfgProfilePart As Object
                        Set oMfgProfilePart = oColl.Item(1) 'count will be one
                        ReportWeightInformation oMfgProfilePart, oStream
                    End If
                End If
                
            ElseIf TypeOf oSelectedObj Is IJAssembly Then
                Dim oSelectedProfileObj As Object
                Dim oAllProfiles As IJElements
                Set oAllProfiles = GetAllProfilePartsFromAssembly(oSelectedObj)
                For Each oSelectedProfileObj In oAllProfiles

                    If TypeOf oSelectedProfileObj Is IJProfilePart Then
                        Dim oColl2 As IJElements
                        Dim oStrc
                        Dim oStructMfgGlobals2 As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
                        Set oColl2 = oStructMfgGlobals2.GetMfgPart(oSelectedProfileObj, IJMfgProfilePart)
                        
                        Dim oNamedItem As IJNamedItem
                        Set oNamedItem = oSelectedProfileObj
                        
                        If Not oColl2 Is Nothing Then
                            If oColl2.Count > 0 Then
                                ReportWeightInformation oColl2.Item(1), oStream
                            End If
                        End If
                    End If
                Next
                Set oAllProfiles = Nothing
            End If
        Next
        
        oStream.Close
        eCustomReportStatus = StrMfgProcessFinished
    End If
            
End Sub
Private Sub ReportWeightInformation(oMfgProfile As Object, oStream As TextStream)
 Const METHOD = "ReportWeightInformation"
    On Error GoTo ErrorHandler

    Dim oMfgProfilePart As IJMfgProfilePart
    Dim oProfilePart As IJProfilePart
    Dim oNamedItem As IJNamedItem
    Dim strPartName As String
    Dim dPartTotalMfgWeight As Double, dPartFaceWeight As Double
    
    Set oNamedItem = oMfgProfile
    strPartName = oNamedItem.Name
    
    Set oMfgProfilePart = oMfgProfile
    Set oProfilePart = oMfgProfilePart.GetDetailedPart
    
    Dim oIJProfile As IJProfile
    Set oIJProfile = oProfilePart
    
    
    'Get cross section info
    Dim oProfileSection As IJDProfileSection
    Set oProfileSection = oProfilePart
    
    Dim oCrossSection As IJCrossSection
    Set oCrossSection = oProfileSection.crossSection
    
    Dim oEntityHelper As New GSCADStrMfgUtilities.MfgEntityHelper
    Dim oGeomCol2d As IJMfgGeomCol2d
    Set oGeomCol2d = oMfgProfilePart.FinalGeometriesAfterProcess2D
    
    Dim lFaces() As Long
    ReDim lFaces(0 To 0) As Long
    Dim oGeom2d As IJMfgGeom2d
    'Get distinct face ids
    'For Each oGeom2d In oGeomCol2d
    
    Dim k As Integer
    For k = 1 To oGeomCol2d.Getcount
        Set oGeom2d = oGeomCol2d.GetGeometry(k)
        Dim i As Integer
        Dim length As Long
        
        length = UBound(lFaces)
        For i = 0 To length - 1
            If lFaces(i) = oGeom2d.FaceId Then
                GoTo NextK
            End If
        Next i
        ReDim Preserve lFaces(0 To UBound(lFaces) + 1) As Long
        lFaces(i) = oGeom2d.FaceId
        
NextK:
    Next
    
    Dim oSDProfilePart As StructDetailObjects.ProfilePart
    Set oSDProfilePart = New StructDetailObjects.ProfilePart
    Set oSDProfilePart.object = oProfilePart
    
    Dim dFlangeWeight As Double
    Dim dWebWeight As Double
    Dim dTotalMarginWeight As Double
    
    Dim j As Integer
    For i = 0 To UBound(lFaces) - 1
    
        If lFaces(i) = JXSEC_TOP Or lFaces(i) = JXSEC_BOTTOM Then
            dPartFaceWeight = oEntityHelper.GetManufacturingPartWeight(oMfgProfile, 31, lFaces(i), oSDProfilePart.FlangeThickness)
            dFlangeWeight = dFlangeWeight + dPartFaceWeight
            
        'ElseIf lFaces(i) = JXSEC_WEB_LEFT Or lFaces(i) = JXSEC_WEB_RIGHT then
        ElseIf lFaces(i) = oMfgProfilePart.Upside Then
            dPartFaceWeight = oEntityHelper.GetManufacturingPartWeight(oMfgProfile, 31, lFaces(i), oSDProfilePart.WebThickness)
            dWebWeight = dWebWeight + dPartFaceWeight
            
        End If
        ' Without considering bevels, chamfers and hole
        'dPartFaceWeight = oEntityHelper.GetManufacturingPartWeight(oMfgProfile, 15, lFaces(i), oIJProfile.thickness)
        
        dPartTotalMfgWeight = dPartTotalMfgWeight + dPartFaceWeight
    
    Next
    
    dTotalMarginWeight = oEntityHelper.GetManufacturingPartWeight(oMfgProfile, 16, 0, oSDProfilePart.WebThickness)
    
        
'''''    ' Without considering bevels, chamfers and hole
'''''    dPartAllMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgProfile, 0, 0, oIJProfile.thickness)
'''''
'''''    ' With considering bevels, chamfers and hole
'''''    dPartTotalMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgProfile, 15, 0, oIJProfile.thickness)
'''''
'''''    ' Without considering bevels
'''''    dPartNoBevelsMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgProfile, 1, 0, oIJProfile.thickness)
'''''
'''''    ' Without considering chamfers
'''''    dPartNoChamfersMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgProfile, 2, 0, oIJProfile.thickness)
'''''
'''''    ' Without considering bevels and chamfers
'''''    dPartNoBevelsChamfersMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgProfile, 3, 0, oIJProfile.thickness)
'''''
'''''    ' Without considering hole
'''''    dPartNoHolesMfgWeight = oEntityHelper.GetManufacturingPartWeight(oMfgProfile, 4, 0, oIJProfile.thickness)
    
    
    Dim oWightCG As IJWeightCG
    Set oWightCG = oProfilePart
    
    Dim sAssemblyPath As String
    sAssemblyPath = GetAssemblyType(oProfilePart)
    
    'MsgBox "Profile.MfgWeight = " & oMfgProfilePart.MfgWeight(15)
    
    'Get Nesting info of Mfg Profile Part
    oStream.WriteLine strPartName & "," & sAssemblyPath & "," _
                                & oProfileSection.CrossSectionName & "," _
                                & oWightCG.DryWeight & "," _
                                & oMfgProfilePart.MfgWeight(15) & "," & dWebWeight & "," & dFlangeWeight & "," & dTotalMarginWeight
'                                & dPartNoBevelsMfgWeight & "," _
'                                & dPartNoChamfersMfgWeight & "," _
'                                & dPartNoBevelsChamfersMfgWeight & "," _
'                                & dPartNoHolesMfgWeight & "," _
'                                & dPartTotalMfgWeight
                                
    
    Set oMfgProfilePart = Nothing
    Set oNamedItem = Nothing
    
    'oStream.Close

    Exit Sub
ErrorHandler:
    'Instead of Erroring out, dump the surfaces which are causing error with ERROR_ as suffix
'    m_iCount = m_iCount + 1
'    strPath = GetTempFolder & "ERROR_" & Format(Date, "yyyy-mm-dd") & "_SurfaceGeometry_" & m_iCount & ".sat"
'    DumpSurfaceSatFile oObject, strPath
    'Err.Raise Err.Number, , Err.Description
End Sub


' ***********************************************************************************
' Public Function GetAssemblyType
'
' Description:  Method will give the Type of Assembly
'
' ***********************************************************************************
Private Function GetAssemblyType(oPart As Object) As String
    Const METHOD = "GetAssemblyType"
    On Error GoTo ErrorHandler

    Dim bstrAssemblyPath As String
    
    ' Get Assembly hierarchy
    Dim oAssemblyChild As IJAssemblyChild
    Set oAssemblyChild = oPart
StartAgain:
    If Not oAssemblyChild Is Nothing Then
        Dim oAssemblyParentUnk As IUnknown
        Set oAssemblyParentUnk = oAssemblyChild.Parent
        If Not oAssemblyParentUnk Is Nothing Then
        If TypeOf oAssemblyParentUnk Is IJAssembly Then
            Dim oNamed As IJNamedItem
            Set oNamed = oAssemblyParentUnk
            If Len(bstrAssemblyPath) = 0 Then
                bstrAssemblyPath = oNamed.Name
            Else
                bstrAssemblyPath = oNamed.Name & " | " & bstrAssemblyPath
            End If
            If TypeOf oAssemblyParentUnk Is IJAssemblyChild Then
                Set oAssemblyChild = oAssemblyParentUnk
                GoTo StartAgain
            End If
            Set oNamed = Nothing
        End If
        End If
        Set oAssemblyParentUnk = Nothing
    End If

    GetAssemblyType = bstrAssemblyPath
    
    Exit Function
    
ErrorHandler:
   Err.Raise Err.Description
End Function



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
            'is IJProfilePart and not MfgProfilePart
            If TypeOf oItem Is IJProfilePart Then
                GetAllProfilePartsFromAssembly.Add oItem
            ElseIf TypeOf oItem Is IJAssembly Then 'Assembly or block is child
                'get parts in this assembly
                On Error Resume Next
                GetAllProfilePartsFromAssembly.AddElements GetAllProfilePartsFromAssembly(oItem)
            End If
        Next
    End If

Cleanup:
    'Clean up
    Set oAssembly = Nothing
    Set oChildren = Nothing
    Set oItem = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
    GoTo Cleanup
End Function

Private Function GetTempFolder() As String

    Const METHOD = "GetTempFolder"
    
    On Error GoTo ErrorHandler
   ' Returns the path to the user's Temp folder. To boot, Windows
   ' requires that a temporary folder exist, so this should always
   ' safely return a path to one. Just in case, though, check the
   ' return value of GetTempPath.
   Dim strTempPath As String
   Dim lngTempPath As Long
   
   ' Fill string with null characters.
   strTempPath = String(144, vbNullChar)
   ' Get length of string.
   lngTempPath = Len(strTempPath)
   ' Call GetTempPath, passing in string length and string.
   If (GetTempPath(lngTempPath, strTempPath) > 0) Then
      ' GetTempPath returns path into string.
      ' Truncate string at first null character.
      GetTempFolder = Left(strTempPath, _
         InStr(1, strTempPath, vbNullChar) - 1)
   Else
      GetTempFolder = ""
   End If
   
   Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function



