Attribute VB_Name = "NRModule"
'--------------------------------------------------------------------------------------------'
'    Copyright (C) 2004 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'    This module contains useful functions for Grids Naming Rules.
'
'Notes
'
'History
'    Eric Faivre         07/21/04                Creation.
'    Eric Faivre         05/12/05                Fix TR-CP·76912: Alphanumeric & Percent Name Rule should round to first decimal place.
'    Eric Faivre         05/12/05                Fix TR-CP·74959: Name of radial plane and cylinder disappears on changing name rule.
'                                                Handle Radial Plane and Cylinder entities.
'--------------------------------------------------------------------------------------------'

Option Explicit

Public mg_oErrors As IJEditErrors

Public Const E_FAIL = -2147467259

Public Function GetIndexOfPrimaryPlane(planePosition As Double, planeAxis As AxisType, pCS As Object, planeInclude As Collection, planeExclude As Collection) As Long
    On Error GoTo ErrorHandler
    
    Dim oPlanes As IJElements
    
    Dim oMiddleHelper As SPGMiddleHelper
    Set oMiddleHelper = New SPGMiddleHelper
    
    oMiddleHelper.EnumPlanes pCS, planeAxis, -100000 - 10 ^ -6, planePosition - 10 ^ -6, Primary, oPlanes

    Dim i As Integer
    
    Dim includeCount As Long
    includeCount = 0
    
    Dim oNRData As NRData
    
    If Not planeInclude Is Nothing Then
        For i = 1 To planeInclude.Count
            Set oNRData = planeInclude.Item(i)
            If oNRData.NestingLevel = Primary And oNRData.Position < planePosition Then includeCount = includeCount + 1
        Next
    End If
    
    Dim excludeCount As Long
    excludeCount = 0
    
    If Not planeExclude Is Nothing Then
        For i = 1 To planeExclude.Count
            Set oNRData = planeExclude.Item(i)
            If oNRData.NestingLevel = Primary And oNRData.Position < planePosition Then excludeCount = excludeCount + 1
        Next
    End If

    GetIndexOfPrimaryPlane = 1 + oPlanes.Count + includeCount - excludeCount
    
    Exit Function
    
ErrorHandler:
    mg_oErrors.Add Err.Number, "GSNamingRules::GetIndexOfPrimaryPlane", Err.Description
    Err.Raise Err.Number
End Function

Public Function GetAlphanumericNameFromIndex(ByVal planeIndex As Long) As String
    On Error GoTo ErrorHandler
    
    Dim alphanumericName As String
    
    Dim exp26 As Long
    exp26 = 0
    
    While Int(planeIndex / 26 ^ exp26) > 26
        exp26 = exp26 + 1
    Wend
    
    Dim unit As Long
    
    For exp26 = exp26 To 0 Step -1
        unit = Int(planeIndex / 26 ^ exp26)
        alphanumericName = alphanumericName + Chr$(64 + unit)
        planeIndex = planeIndex - (unit * 26 ^ exp26)
    Next

    Dim i As Integer
    For i = Len(alphanumericName) To 2 Step -1
        If Asc(Mid$(alphanumericName, i, 1)) <= 64 Then
            alphanumericName = Left$(alphanumericName, i - 2) & Chr$(Asc(Mid$(alphanumericName, i - 1, 1)) - 1) & Chr$(Asc(Mid$(alphanumericName, i, 1)) + 26) & Right$(alphanumericName, Len(alphanumericName) - i)
        End If
    Next
    If Asc(Left$(alphanumericName, 1)) <= 64 Then
        alphanumericName = Right$(alphanumericName, Len(alphanumericName) - 1)
    End If
    
    GetAlphanumericNameFromIndex = alphanumericName

    Exit Function
    
ErrorHandler:
    mg_oErrors.Add Err.Number, "GSNamingRules::GetAlphanumericNameFromIndex", Err.Description
    Err.Raise Err.Number
End Function

Private Function IsPlaneInExclude(planePosition As Double, planeNL As NestingLevelType, planeExclude As Collection) As Boolean
    On Error GoTo ErrorHandler

    Dim oNRData As NRData
    
    IsPlaneInExclude = False
    
    If Not planeExclude Is Nothing Then
        For Each oNRData In planeExclude
            If oNRData.Position = planePosition And oNRData.NestingLevel = planeNL Then
                IsPlaneInExclude = True
                Exit Function
            End If
        Next
    End If

    Exit Function
    
ErrorHandler:
    mg_oErrors.Add Err.Number, "GSNamingRules::IsPlaneInExclude", Err.Description
    Err.Raise Err.Number
End Function

Private Function GetNextOrPreviousPlanePosition(planePosition As Double, planeAxis As AxisType, pCS As Object, planeInclude As Collection, planeExclude As Collection, bNext As Boolean, nestingLevelPN As NestingLevelType) As Double
    On Error GoTo ErrorHandler
    
    Dim oPlanes As IJElements
    
    Dim oMiddleHelper As SPGMiddleHelper
    Set oMiddleHelper = New SPGMiddleHelper
    
    oMiddleHelper.EnumPlanes pCS, planeAxis, IIf(bNext = True, planePosition, -100000), IIf(bNext = True, 100000, planePosition), nestingLevelPN, oPlanes
    
    Dim NPplanePosition As Double
    NPplanePosition = IIf(bNext = True, 100000, -100000)
    
    Dim oPlaneData As ISPGGridData
    
    For Each oPlaneData In oPlanes
        If IsPlaneInExclude(oPlaneData.Position, oPlaneData.NestingLevel, planeExclude) = False Then
            If ((bNext = True And oPlaneData.Position > planePosition) Or (bNext = False And oPlaneData.Position < planePosition)) And ((bNext = True And oPlaneData.Position < NPplanePosition) Or (bNext = False And oPlaneData.Position > NPplanePosition)) Then
                NPplanePosition = oPlaneData.Position
            End If
        End If
    Next
    
    Dim oNRDataI As NRData
    
    For Each oNRDataI In planeInclude
        If IsPlaneInExclude(oNRDataI.Position, oNRDataI.NestingLevel, planeExclude) = False Then
            If oNRDataI.NestingLevel = nestingLevelPN And ((bNext = True And oNRDataI.Position > planePosition) Or (bNext = False And oNRDataI.Position < planePosition)) And ((bNext = True And oNRDataI.Position < NPplanePosition) Or (bNext = False And oNRDataI.Position > NPplanePosition)) Then
                NPplanePosition = oNRDataI.Position
            End If
        End If
    Next
  
    GetNextOrPreviousPlanePosition = NPplanePosition
    
    Exit Function
    
ErrorHandler:
    mg_oErrors.Add Err.Number, "GSNamingRules::GetNextOrPreviousPlane", Err.Description
    Err.Raise Err.Number
End Function

Public Function GetNameFromNPercentNR(planePosition As Double, planeAxis As AxisType, NestingLevel As NestingLevelType, pCS As Object, planeInclude As Collection, planeExclude As Collection, bAlphanumericNR As Boolean) As String
    On Error GoTo ErrorHandler
    
    Dim planeName As String
    
    If planeAxis = z Then
        Err.Raise E_FAIL, , "Invalid plane axis for " & IIf(bAlphanumericNR = True, "'Alphanumeric and Percent' Grids Naming Rule.", "'Index and Percent' Grids Naming Rule.")
    End If
    
    If NestingLevel = Primary Then
        Dim planeIndex As Long
        planeIndex = GetIndexOfPrimaryPlane(planePosition, planeAxis, pCS, planeInclude, planeExclude)
        
        If bAlphanumericNR = True Then
            Select Case planeAxis
                Case x, R, C
                    planeName = GetAlphanumericNameFromIndex(planeIndex)
                Case y
                    planeName = CStr(planeIndex)
            End Select
        Else
            Select Case planeAxis
                Case x
                    planeName = "GPX" & CStr(planeIndex)
                Case y
                    planeName = "GPY" & CStr(planeIndex)
                Case R
                    planeName = "R" & CStr(planeIndex)
                Case C
                    planeName = "C" & CStr(planeIndex)
            End Select
        End If
    Else
        Dim planePositionPrev As Double
        planePositionPrev = GetNextOrPreviousPlanePosition(planePosition, planeAxis, pCS, planeInclude, planeExclude, False, IIf(NestingLevel = Secondary, Primary, Secondary))
        
        Dim planePositionNext As Double
        planePositionNext = GetNextOrPreviousPlanePosition(planePosition, planeAxis, pCS, planeInclude, planeExclude, True, IIf(NestingLevel = Secondary, Primary, Secondary))
        
        planeName = GetNameFromNPercentNR(planePositionPrev, planeAxis, IIf(NestingLevel = Secondary, Primary, Secondary), pCS, planeInclude, planeExclude, bAlphanumericNR)
        
        planeName = planeName & Right(CStr(Round((planePosition - planePositionPrev) / (planePositionNext - planePositionPrev), 1)), 2)
    End If
    
    GetNameFromNPercentNR = planeName
    
    Exit Function
    
ErrorHandler:
    mg_oErrors.Add Err.Number, "GSNamingRules::GetNameFromNPercentNR", Err.Description
    Err.Raise Err.Number
End Function

Public Sub ApplyNameFromNPercentNR(pPlane As Object, planeInclude As Collection, planeExclude As Collection, bAlphanumericNR As Boolean)
    On Error GoTo ErrorHandler
    
    Dim oPlaneData As ISPGGridData
    Set oPlaneData = pPlane
    
    Dim oPlaneNavigate As ISPGNavigate
    Set oPlaneNavigate = pPlane
    
    Dim oCS As Object
    oPlaneNavigate.GetParent oCS
    
    Dim oNamedItem As IJNamedItem
    Set oNamedItem = pPlane
    
    oNamedItem.Name = GetNameFromNPercentNR(oPlaneData.Position, oPlaneData.Axis, oPlaneData.NestingLevel, oCS, planeInclude, planeExclude, bAlphanumericNR)
    
    Exit Sub
    
ErrorHandler:
    mg_oErrors.Add Err.Number, "GSNamingRules::ApplyNameFromNPercentNR", Err.Description
    Err.Raise Err.Number
End Sub

Public Function GetPlaneType(pPlane As Object) As String
    On Error GoTo ErrorHandler

    Dim jContext As IJContext
    Set jContext = GetJContext()

    Dim oDBTypeConfiguration As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim oModelResourceMgr As IUnknown
    
    Set oDBTypeConfiguration = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(oDBTypeConfiguration.get_DataBaseFromDBType("Model"))

    Dim oPlaneData As ISPGGridData
    Set oPlaneData = pPlane

    Dim oCodeListMetaData As IJDCodeListMetaData
    Set oCodeListMetaData = oModelResourceMgr
    
    Dim oCodeListCollection As IJDInfosCol
    Set oCodeListCollection = oCodeListMetaData.CodelistValueCollection(IIf(oPlaneData.Axis = z, "ElevPlaneType", "GridPlaneType"))

    Dim oCodeListValue As IJDCodeListValue
    Dim i As Long
    
    For i = 1 To oCodeListCollection.Count
        Set oCodeListValue = oCodeListCollection.Item(i)
        If oPlaneData.Type = oCodeListValue.ValueID Then
            GetPlaneType = oCodeListValue.ShortValue
            Exit For
        End If
    Next

    Exit Function
    
ErrorHandler:
    mg_oErrors.Add Err.Number, "GSNamingRules::GetPlaneType", Err.Description
    Err.Raise Err.Number
End Function
