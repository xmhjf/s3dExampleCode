Attribute VB_Name = "modHrchyWlkr"
'********************************************************************
' Copyright (C) 1998-2000 Intergraph Corporation.  All Rights Reserved.
'
' File: modHrchyWlkr.bas
'
' Author: sypark@ship.samsung.co.kr
'
' Abstract: modHrchyWlkr function
'
' Description:
'********************************************************************
Option Explicit

Public m_oActiveConn As String

Const Module = "GSCADHMRules.modHrchyWlkr:"

'******************************************************************************
' Routine: GetChildOfThePlate
'
' Abstract: Using the hrchywlkr, to get the all of child of plate
'
' Description:
'******************************************************************************
Public Function GetChildOfThePlate(oPlate As IJPlate) As IMSCoreCollections.IJElements
    Const METHOD As String = "GetChildOfThePlate"
    On Error GoTo ErrorHandler
      
    Dim oChildElem As IMSCoreCollections.IJElements
    Set oChildElem = New IMSCoreCollections.JObjectCollection
    
    Dim strIIDs(0 To 2) As String
    Dim oHierarchyFilter As IJHierarchyFilter
    Dim oFilter As IJSimpleFilter
    Dim oCompoundFilter As IJCompoundFilter
    Dim oFactory As IJMiddleFiltersFactory
    
    Dim oObjColl As IJDObjectCollection
    Dim strQuery As String
    
    Set oObjColl = New IMSCoreCollections.JObjectCollection
    
    oObjColl.Add oPlate
    
    Set oFactory = New MiddleFiltersFactory
    Set oHierarchyFilter = oFactory.CreateEntity(HierarchyFilter, Nothing)
    
    oHierarchyFilter.AdapterProgID = "GSCADHierarchyHelper.SystemAdapter"
    oHierarchyFilter.Depth = 15
    oHierarchyFilter.IncludeNested = True
    oHierarchyFilter.HierarchyObjects = oObjColl
    
    Set oFilter = oHierarchyFilter
  
    strQuery = oFilter.BuildQuery(vbNullString, "", LANGUAGE_SQL)
    
    Dim oCommand As MIDDLECOMMANDLib.JMiddleCommand
    Set oCommand = New MIDDLECOMMANDLib.JMiddleCommand
    
    oCommand.ActiveConnection = m_oActiveConn
    oCommand.QueryLanguage = LANGUAGE_SQL
    oCommand.CommandText = strQuery
    
    Dim oEnum As IEnumMoniker
    Set oEnum = oCommand.SelectObjects
    Dim varMoniker As Variant
    Dim oObjUnk As Object
    Dim colPlateChildren As IMSCoreCollections.IJElements
    Set colPlateChildren = New IMSCoreCollections.JObjectCollection
    Set oCommand = Nothing
    
    Dim oPOM As IJDPOM
    Set oPOM = GetPOMConnection("Model")
    
    If Not oEnum Is Nothing Then
        oEnum.Reset
        For Each varMoniker In oEnum
            Set oObjUnk = oPOM.GetObject(varMoniker)
            colPlateChildren.Add oObjUnk
            Set oObjUnk = Nothing
        Next

    End If
    
    Set GetChildOfThePlate = colPlateChildren

    
Cleanup:
    Set varMoniker = Nothing
    Set oEnum = Nothing
    
    Set colPlateChildren = Nothing
    Set oHierarchyFilter = Nothing
    Set oFilter = Nothing
    
    Set oFactory = Nothing
    Set oObjColl = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetHolesFromChild
'
' Abstract: Using the hrchywlkr, to get the all of holes of plate
'
' Description:
'******************************************************************************
Public Function GetHolesFromChild(oChildElem As IMSCoreCollections.IJElements) As IMSCoreCollections.IJElements
    Const METHOD = "GetHolesFromChild"
    On Error GoTo ErrorHandler
    
    Dim oElement As IMSCoreCollections.IJElements
    Dim oObject As Object
    
    Set oElement = New IMSCoreCollections.JObjectCollection
    For Each oObject In oChildElem
        If TypeOf oObject Is IJHoleTraceAE Then
            oElement.Add oObject
        End If
    Next oObject
    
    Set oObject = Nothing
    Set GetHolesFromChild = oElement

Cleanup:
    Set oElement = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetSeamLineFromChild
'
' Abstract: Using the hrchywlkr, to get all of seamline of plate
'
' Description:
'******************************************************************************
Public Function GetSeamLineFromChild(oChildElem As IMSCoreCollections.IJElements) As IMSCoreCollections.IJElements
    Const METHOD = "GetSeamLineFromChild"
    On Error GoTo ErrorHandler
    
    Dim oElement As IMSCoreCollections.IJElements
    Dim oObject As Object
    
    Set oElement = New IMSCoreCollections.JObjectCollection
    For Each oObject In oChildElem
        If TypeOf oObject Is IJSeam Then
            Dim oSeam As IJSeam
            Set oSeam = oObject
            oElement.Add oSeam
            Set oSeam = Nothing
        End If
    Next oObject
    
    Set oObject = Nothing
    Set GetSeamLineFromChild = oElement

Cleanup:
    Set oElement = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetOpeningFromChild
'
' Abstract: Using the hrchywlkr, to get all of opening of plate
'
' Description:
'******************************************************************************
Public Function GetOpeningFromChild(oChildElem As IMSCoreCollections.IJElements) As IMSCoreCollections.IJElements
    Const METHOD = "GetOpeningFromChild"
    On Error GoTo ErrorHandler
    
    Dim oElement As IMSCoreCollections.IJElements
    Dim oObject As Object
    
    Set oElement = New IMSCoreCollections.JObjectCollection
    For Each oObject In oChildElem
        If TypeOf oObject Is IJDStructOpeningOccurrence Then
            Dim oOpening As IJDStructOpeningOccurrence
            Set oOpening = oObject
            oElement.Add oOpening
            Set oOpening = Nothing
        End If
    Next oObject
    
    Set oObject = Nothing
    Set GetOpeningFromChild = oElement

Cleanup:
    Set oElement = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetProfileSystemFromChild
'
' Abstract: Using the hrchywlkr, to get all of profile of plate
'
' Description:
'******************************************************************************
Public Function GetProfileSystemFromChild(oChildElem As IMSCoreCollections.IJElements) As IMSCoreCollections.IJElements
    Const METHOD = "GetProfileSystemFromChild"
    On Error GoTo ErrorHandler
    
    Dim oElement As IMSCoreCollections.IJElements
    Dim oObject As Object
    
    Set oElement = New IMSCoreCollections.JObjectCollection
    For Each oObject In oChildElem
        If TypeOf oObject Is IJProfileSystemAE Then
            Dim oProfilesystem As IJProfileSystemAE
            Set oProfilesystem = oObject
            oElement.Add oProfilesystem
            Set oProfilesystem = Nothing
        End If
    Next oObject
    
    Set oObject = Nothing
    Set GetProfileSystemFromChild = oElement

Cleanup:
    Set oElement = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function
 
