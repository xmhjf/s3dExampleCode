VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IndexNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'
'  Abstract:
'   The file contains  implementation of a naming rule for Elevation
'   Planes
'
'  Author:
'   Celine Faure    10/22/02    Creation
'   Celine Faure    03/28/03    Change Naming Rule Index:When position of Plane is negative then the
'                               Index in Plane Name is negative too
'   Celine Faure    04/03/03    Change Prefix for Ship Naming Rule
'   Celine Faure    04/25/03    Re fix Need better default name rule for grid planes.  Change method RetrieveIndex
'                               to assume NL in NamingRule
'   Eric Faivre     06/23/03    Set starting Frame index to 0 for Naming Rule used in Intelliship.
'                               Fix TR#42831: Default axis naming rule changed to start index at 1 instead of 0.
'   Eric Faivre     01/04/05    Support of Radial Plane and Cylinder.
'   Eric Faivre     06/05/06    Fix DI-CP·50027: Should mark modified Plane on IJNamedItem when needed.
'                               IJNamedItem.Name set property updates the entity on IJNamedItem if the name is
'                               really changed. To avoid multiple update calls, compute the name and then set
'                               it at the end instead of calling the property many times.
'   Apparao         11/12/08    TR-CP-147330  Performance isssues observed while creating Grids is very high in SM3D
'******************************************************************

Option Explicit

Implements IJNameRule
Implements IJNRAuxiliaryPosition

Const vbInvalidArg = &H80070057

Private Const Module = "GSNamingRules: "
Private Const MODELDATABASE = "Model"

Private m_IncludeList As New Collection
Private m_ExcludeList As New Collection

Private Sub Class_Initialize()
    Set mg_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set mg_oErrors = Nothing
End Sub

Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IMSICDPInterfacesLib.IJElements, ByVal pActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo label

    Dim oChildNamedItem As IJNamedItem
    Dim oGridData As ISPGGridData
    Dim Axis As AxisType
    Dim primaryindex As Long
    Dim secondaryindex As Long
    Dim tertiaryindex As Long
    Dim oNavigate As ISPGNavigate
    Dim oCS As ISPGCoordinateSystemProperties
    Dim CSType As CoordinateSystemType
    Dim frameEntityName As String

    If pEntity Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
        Exit Sub
    End If

    Set oChildNamedItem = pEntity
    Set oGridData = pEntity
    Set oNavigate = pEntity

    oNavigate.GetParent oCS
    CSType = oCS.Type
    Axis = oGridData.Axis
    RetrieveIndex pEntity, primaryindex, secondaryindex, tertiaryindex
    
    If CSType = Ship And primaryindex > 0 Then
        primaryindex = primaryindex - 1
    End If
    
    If CSType = Ship Then
        If Axis = z Then
            frameEntityName = "D" + CStr(primaryindex)
        ElseIf Axis = x Then
            frameEntityName = "F" + CStr(primaryindex)
        ElseIf Axis = y Then
            frameEntityName = "L" + CStr(primaryindex)
        ElseIf Axis = R Then
            frameEntityName = "R" + CStr(primaryindex)
        Else
            frameEntityName = "C" + CStr(primaryindex)
        End If
    Else
        If Axis = z Then
            frameEntityName = "ElevPlane" + CStr(primaryindex)
        ElseIf Axis = x Then
            frameEntityName = "GPX" + CStr(primaryindex)
        ElseIf Axis = y Then
            frameEntityName = "GPY" + CStr(primaryindex)
        ElseIf Axis = R Then
            frameEntityName = "R" + CStr(primaryindex)
        Else
            frameEntityName = "C" + CStr(primaryindex)
        End If
    End If
    If secondaryindex <> 0 Then
        frameEntityName = frameEntityName + "." + CStr(secondaryindex)
    End If
    If tertiaryindex <> 0 Then
        frameEntityName = frameEntityName + "." + CStr(tertiaryindex)
    End If
    
    oChildNamedItem.Name = frameEntityName
    
Exit Sub

label:
    ' log the error in middle tier and propagate the error code to the caller
    mg_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
    
End Sub


Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Set IJNameRule_GetNamingParents = Nothing
End Function

Private Sub RetrieveIndex(ByVal pEntity As Object, primaryindex As Long, secondaryindex As Long, tertiaryindex As Long)
    Dim oNavigate As ISPGNavigate
    Dim oMiddleHelper As New SPGMiddleHelper
    Dim oCS As SPGCoordinateSystem
    Dim Axis As AxisType
    Dim oGridData As ISPGGridData
    Dim oGridDataEnt As ISPGGridData
    Dim ppPlanes As IJElements
    Dim i As Long
    Dim excludeCount As Long
    Dim includeCount As Long
    Dim pos As Double
    Dim posinlist As Double
    Dim indexexcludelist As Long
    Dim indexincludelist As Long
    Dim index(0 To 2) As Long
    Dim fPositive As Boolean
    Dim k As Long
    Dim NL As NestingLevelType
    Dim oNRData As New NRData
    Dim posforenum As Double
    Dim posforexclude As Double
    Dim posforinclude As Double
    
    'Retrieve Coordinate System
    Set oNavigate = pEntity
    oNavigate.GetParent oCS
    
    'Retrieve planes
    Set oGridDataEnt = pEntity
    Axis = oGridDataEnt.Axis
    pos = oGridDataEnt.position
    NL = oGridDataEnt.NestingLevel
    index(0) = 0
    index(1) = 0
    index(2) = 0
    
    If pos < 0 Then
        posforenum = -0.0000001
        posforinclude = 0
        posforexclude = 0
    Else
        posforenum = 0
        posforinclude = 0
        posforexclude = 0
    End If
    For k = 1 To NL
        indexexcludelist = 0
        indexincludelist = 0
        
        If pos < 0 Then
            oMiddleHelper.PlaneCollection pEntity, Axis, pos, posforenum, k, ppPlanes
            fPositive = False
        Else
            oMiddleHelper.PlaneCollection pEntity, Axis, posforenum, pos, k, ppPlanes
            fPositive = True
        End If
        
        If ppPlanes.Count > 0 Then
            Set oGridData = ppPlanes.Item(ppPlanes.Count)
            posforenum = oGridData.position
        End If
        
        'search in ExcludeList the index for NL Primary
        If Not m_ExcludeList Is Nothing Then
            excludeCount = m_ExcludeList.Count
            If fPositive And excludeCount > 0 Then
                posinlist = -100000
                i = 1
                indexexcludelist = 0
                Do While i <= excludeCount And posinlist < pos
                    Set oNRData = New NRData
                    Set oNRData = m_ExcludeList.Item(i)
                    posinlist = oNRData.position
                    If posinlist >= posforexclude And oNRData.NestingLevel = k Then
                        indexexcludelist = indexexcludelist + 1
                        posforexclude = posinlist
                    End If
                    i = i + 1
                Loop
               
            ElseIf excludeCount > 0 Then
                posinlist = 100000
                i = excludeCount - 1
                indexexcludelist = 0
                Do While i >= 1 And posinlist > pos
                    Set oNRData = New NRData
                    Set oNRData = m_ExcludeList.Item(i)
                    posinlist = oNRData.position
                    If posinlist < posforexclude And oNRData.NestingLevel = k Then
                        indexexcludelist = indexexcludelist + 1
                        posforexclude = posinlist
                    End If
                    i = i - 1
                Loop
                
            End If
        End If
        
        
        If Not m_IncludeList Is Nothing Then
            Set oNRData = New NRData
            'search in the include list the index for NL Primary
            includeCount = m_IncludeList.Count
            If fPositive And includeCount > 0 Then
                posinlist = -100000
                i = 1
                indexincludelist = 0
                
                Do While i <= includeCount And posinlist < pos
                    Set oNRData = New NRData
                    Set oNRData = m_IncludeList.Item(i)
                    posinlist = oNRData.position
                    If posinlist >= posforinclude And oNRData.NestingLevel = k Then
                        indexincludelist = indexincludelist + 1
                        posforinclude = posinlist
                    End If
                    i = i + 1
                Loop
                
            ElseIf includeCount > 0 Then
                posinlist = 100000
                i = includeCount
                indexincludelist = 0
                Do While i >= 1 And posinlist > pos
                    Set oNRData = New NRData
                    Set oNRData = m_IncludeList.Item(i)
                    posinlist = oNRData.position
                    If posinlist < posforinclude And oNRData.NestingLevel = k Then
                        indexincludelist = indexincludelist + 1
                        posforinclude = posinlist
                    End If
                    i = i - 1
                Loop
                
            End If
        End If
        
        'First step: Primary Nesting Level step
        
        Dim lCount As Long
                
        If (indexincludelist = 0) Then
            lCount = ppPlanes.Count
        Else
            If ppPlanes.Count = 1 Then
                lCount = 0
            Else
                lCount = ppPlanes.Count
            End If
        End If

        If (m_ExcludeList.Count > 0) Then
            indexexcludelist = lCount
        End If
        
        index(k - 1) = lCount - indexexcludelist + indexincludelist
        
        If ((k = Primary And NL = Secondary) Or (k = Primary And NL = Tertiary) Or (k = Secondary And NL = Tertiary)) And index(k - 1) = 0 Then
            index(k - 1) = -1
        End If
        
    Next k

    If fPositive Then
        primaryindex = index(0)
        secondaryindex = index(1)
        tertiaryindex = index(2)
    Else
        primaryindex = -index(0)
        secondaryindex = index(1)
        tertiaryindex = index(2)
    End If
    
End Sub

Private Sub IJNRAuxiliaryPosition_ExcludeAuxiliaryPos(ByVal pExcludeList As Object)
    Set m_ExcludeList = pExcludeList
End Sub

Private Sub IJNRAuxiliaryPosition_IncludeAuxiliaryPos(ByVal pIncludeList As Object)
    Set m_IncludeList = pIncludeList
End Sub


