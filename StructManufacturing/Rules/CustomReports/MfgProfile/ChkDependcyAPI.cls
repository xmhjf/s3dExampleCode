VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChkDpndcyAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : MfgProfileCustomReports
''  File        : ChkDependcyAPI.cls
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

Private Const MODULE = "MfgProfileCustomReports.ChkDpndcyAPI::"

Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, _
                                     strFileName As String, _
                                     eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
    On Error GoTo ErrorHandler

    If pElements.Count < 1 Then
        eCustomReportStatus = StrMfgNoObjectFound
        Exit Sub
    End If

    Dim oFSO As FileSystemObject
    Set oFSO = New FileSystemObject

    Dim oTextStream As TextStream

    Set oTextStream = oFSO.OpenTextFile(strFileName, ForAppending, True)
    oTextStream.WriteLine "Assembly Path," & _
                          "Profile Name," & _
                          "Status," & _
                          "Details," & _
                          "Section Type," & _
                          "Section Name," & _
                          "Width," & _
                          "Height," & _
                          "Web Depth," & _
                          "Web Thickness," & _
                          "Flange Width," & _
                          "Flange Thickness," & _
                          "Cross section Area," & _
                          "Molded Length," & _
                          "Material Type," & _
                          "Approximate Length," & _
                          "Is Linear," & _
                          "Is Twisted,"

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

Wrapup:
    oTextStream.Close
    Set oTextStream = Nothing

    Exit Sub

ErrorHandler:
    If Not oTextStream Is Nothing Then
        oTextStream.Close
        Set oTextStream = Nothing
    End If

    eCustomReportStatus = StrMfgErrorUnknown
    Err.Raise Err.Number, MODULE & METHOD, Err.Description
End Sub

Private Sub ReportProfilePartInformation(oProfileObj As Object, oTextStream As TextStream)
 Const METHOD = "ReportProfilePartInformation"
    On Error GoTo ErrorHandler

    Dim bstrAssemblyPath As String

    ' Get Assembly hierarchy
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

    ' Use the profile part as input to StructDetailObjects, PartSupport

    Dim oProfileSDO As StructDetailObjects.ProfilePart
    Set oProfileSDO = New StructDetailObjects.ProfilePart
    Set oProfileSDO.object = oProfileObj

    Dim SDOlibResultStr As String
    SDOlibResultStr = oProfileSDO.SectionType & "," & _
                      oProfileSDO.SectionName & "," & _
                      oProfileSDO.Width & "," & _
                      oProfileSDO.Height & "," & _
                      oProfileSDO.WebLength & "," & _
                      oProfileSDO.WebThickness & "," & _
                      oProfileSDO.FlangeLength & "," & _
                      oProfileSDO.FlangeThickness & "," & _
                      oProfileSDO.CrossSectionArea & "," & _
                      oProfileSDO.MoldedLength & "," & _
                      oProfileSDO.MaterialType

    Dim oProfilePPS As IJProfilePartSupport
    Set oProfilePPS = New GSCADSDPartSupport.ProfilePartSupport
    Dim oProfilePS As IJPartSupport
    Set oProfilePS = oProfilePPS
    Set oProfilePS.Part = oProfileObj

    Dim PartSupPropResultStr As String
    PartSupPropResultStr = oProfilePPS.ApproximateLength & "," & _
                           oProfilePPS.IsLinear & "," & _
                           oProfilePPS.IsTwisted & ","

    Dim APIresultSummary As String

    On Error Resume Next

    Dim oWB As IJWireBody
    Dim oTD As IJDVector
    Dim IsCenter As Boolean
    oProfilePPS.GetProfilePartLandingCurve oWB, oTD, IsCenter, SideA

    If Err.Number <> 0 Then
        APIresultSummary = "GetProfilePartLandingCurve failed for SideA"
        Err.Clear
    Else
        Dim oSP As IJDPosition, oEP As IJDPosition
        oWB.GetEndPoints oSP, oEP

        Dim oXV As IJDVector, oYV As IJDVector
        Dim oOP As IJDPosition

        oProfilePPS.GetOrientation oSP, oXV, oYV, oOP

        If Err.Number <> 0 Then
            If Len(APIresultSummary) > 0 Then
                APIresultSummary = APIresultSummary & vbLf & "GetOrientation failed at start point"
            Else
                APIresultSummary = "GetOrientation failed at start point"
            End If
            Err.Clear
        End If

        oProfilePPS.GetOrientation oEP, oXV, oYV, oOP

        If Err.Number <> 0 Then
            If Len(APIresultSummary) > 0 Then
                APIresultSummary = APIresultSummary & vbLf & "GetOrientation failed at end point"
            Else
                APIresultSummary = "GetOrientation failed at end point"
            End If
            Err.Clear
        End If
    End If

    oProfilePPS.GetProfilePartLandingCurve oWB, oTD, IsCenter, SideB

    If Err.Number <> 0 Then
        If Len(APIresultSummary) > 0 Then
            APIresultSummary = APIresultSummary & vbLf & "GetProfilePartLandingCurve failed for SideB"
        Else
            APIresultSummary = "GetProfilePartLandingCurve failed for SideB"
        End If
        Err.Clear
    End If

    oProfilePPS.GetExtendedStiffPartLandingCurve oWB, oTD, IsCenter, SideA

    If Err.Number <> 0 Then
        If Len(APIresultSummary) > 0 Then
            APIresultSummary = APIresultSummary & vbLf & "GetExtendedStiffPartLandingCurve failed for SideA"
        Else
            APIresultSummary = "GetExtendedStiffPartLandingCurve failed for SideA"
        End If
        Err.Clear
    End If

    oProfilePPS.GetExtendedStiffPartLandingCurve oWB, oTD, IsCenter, SideB

    If Err.Number <> 0 Then
        If Len(APIresultSummary) > 0 Then
            APIresultSummary = APIresultSummary & vbLf & "GetExtendedStiffPartLandingCurve failed for SideB"
        Else
            APIresultSummary = "GetExtendedStiffPartLandingCurve failed for SideB"
        End If
        Err.Clear
    End If

    oProfilePPS.GetMoldedStiffPartLandingCurve oWB

    If Err.Number <> 0 Then
        If Len(APIresultSummary) > 0 Then
            APIresultSummary = APIresultSummary & vbLf & "GetMoldedStiffPartLandingCurve failed"
        Else
            APIresultSummary = "GetMoldedStiffPartLandingCurve failed"
        End If
        Err.Clear
    End If

    Dim oSF As IJSurfaceBody
    Dim oContourColl As Collection, oPortColl As Collection
    oProfilePPS.GetProfileContours WebLeftFace, oSF, oContourColl, oPortColl

    If Err.Number <> 0 Then
        If Len(APIresultSummary) > 0 Then
            APIresultSummary = APIresultSummary & vbLf & "GetProfileContours failed for Web Left"
        Else
            APIresultSummary = "GetProfileContours failed for Web Left"
        End If
        Err.Clear
    End If

    oProfilePPS.GetProfileContours WebRightFace, oSF, oContourColl, oPortColl

    If Err.Number <> 0 Then
        If Len(APIresultSummary) > 0 Then
            APIresultSummary = APIresultSummary & vbLf & "GetProfileContours failed for Web Right"
        Else
            APIresultSummary = "GetProfileContours failed for Web Right"
        End If
        Err.Clear
    End If

    On Error GoTo ErrorHandler

    Dim IsMfgStr As String
    If Len(APIresultSummary) > 0 Then
        IsMfgStr = "BAD"
    Else
        IsMfgStr = "Okay"
    End If

    oTextStream.WriteLine bstrAssemblyPath & "," & _
                          oProfileSDO.Name & "," & _
                          IsMfgStr & "," & _
                          """" & APIresultSummary & """," & _
                          SDOlibResultStr & "," & _
                          PartSupPropResultStr

CleanUp:
    Set oProfileSDO = Nothing
    Set oProfilePPS = Nothing
    Set oProfilePS = Nothing
    Exit Sub

ErrorHandler:
    Set oProfileSDO = Nothing
    Set oProfilePPS = Nothing
    Set oProfilePS = Nothing
    Err.Raise Err.Number, MODULE & METHOD, Err.Description
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
    Set oAssembly = Nothing
    Set oChildren = Nothing
    Set oItem = Nothing

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, MODULE & METHOD, Err.Description
    GoTo CleanUp
End Function
