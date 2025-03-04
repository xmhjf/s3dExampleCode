VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCheckvsParts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2000 Intergraph Corporation.  All Rights Reserved.
'
' File: CCheckvsParts.cls
'
' Author: sypark@ship.samsung.co.kr
'
' Abstract: Check the minmun distance between Holes and Part
'
' Description:
' IJCheckMfcty is implemented by GrobalCheck.tlb.
'********************************************************************
Option Explicit

Private Const Module = "GSCADHMRules.CCheckvsParts:"
Private Const PROG_ID = "GSCADHMRules.CCheckvsParts"

Implements IJCheckMfcty

Dim m_MDPartEdge As Double     'For minimum distance between hole and Edge of structure.
Dim m_MDOutfitting As Double   'For minimum distance between hole and Edge of structure.
Dim m_MDBracket As Double      'For minimum distance between hole and Edge of structure.

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: InitializeUserDefined
'
'Interface: Private function
'
'Abstract: This method is to set the minimum distance between hole and hole
'          User can set the minimum distance
'
'Attention : The unit of measure should be mm
'
'   For instance
'   m_RuleMinDistance = 0.05 is that the minimum distance is 50 mm
'
'********************************************************************
Private Sub InitializeUserDefined()
    Const METHOD = "InitializeUserDefined"
    On Error GoTo ErrorHandler

    m_MDPartEdge = 0.05             'If the nearest distance between hole and Edge of part of structure is less than 50 mm, should be checked.
    m_MDOutfitting = 0.05           'If the nearest distance between hole and Outfitting is less than 50 mm, should be checked.
    m_MDBracket = 0.05              'If the nearest distance between hole and bracket is less than 50 mm, should be checked.
        
Cleanup:
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Sub

'******************************************************************************
' Routine: IJCheckMfcty_ExecuteCheck
'
' Abstract: Implemented by IJCheckMfcty interface
'
' Description: 1. Set the rule minimum distance, Call "InitializeUserDefined"
'              2. Get the distance between hole and compared object, Call "GetMinimumDistance"
'              3. Compare the minimum distance and distance between holes, call "CheckRules"
'******************************************************************************
Private Sub IJCheckMfcty_ExecuteCheck(ByVal pCollection As IMSCoreCollections.IJElements, _
                                      ByVal bStopOnError As Boolean, _
                                      pOptionCodes() As Variant, _
                                      ByVal pCallingObj As Object)
    Const METHOD = "IJCheckMfcty_Execute"
    On Error GoTo ErrorHandler

    Dim oCallBack As IJCheckMfctyCallback
    Dim oCheckingTrace As IJHoleTraceAE
    Dim oCheckingCurves As IMSCoreCollections.IJElements
    Dim oColHoles As IMSCoreCollections.IJElements
    
    Dim lngDummy As Long
    Dim lngCancel As Long

    Dim oStrMsg As String
    
    'Set the m_RuleMinDistance
    InitializeUserDefined
    
    'If workingset is nothing, exits.
    If pCollection.Count = 0 Then GoTo Cleanup
    
    'Get the holes from pCollection. This collection is workingset or selectset
    SetCollectionHoles pCollection, oColHoles

    Set oCallBack = pCallingObj
    oCallBack.OnProgress lngDummy, lngCancel

    For Each oCheckingTrace In oColHoles
        'Get the curves from HoleTrace
        Set oCheckingCurves = CanRetriveCurvesFromHole(PROG_ID, oCheckingTrace, oCallBack, 2)
        
        'if the hole is wrong, We don't have curves, display this hole information on the list.
        If Not oCheckingCurves Is Nothing Then
            If oCheckingCurves.Count <> 0 Then
                 'Get the parent system from HoleTrace
                Dim oObject As Object
                Dim oPlate As IJPlate
                Set oObject = GetPlateSystemFromHoleTrace(oCheckingTrace)
                If TypeOf oObject Is IJPlate Then
                    Set oPlate = oObject
                End If
                Set oObject = Nothing

                'If the plate is root, the hole cannot be compared with the system root.
                If oPlate Is Nothing Then
                    oStrMsg = "The parents of this hole is a standalone part or Profile. Skip the checking between hole and parent."
                    PassStringToCallBack PROG_ID, oCheckingTrace, oStrMsg, oCallBack, 2
                Else
                    'Check the hole with Plate part
                    CheckForPlatePart oCheckingCurves, oPlate, oCheckingTrace, oCallBack
                    
                    'Check the hole with the Coaming
                    CheckForCoamingOnPlate oCheckingCurves, oPlate, oCheckingTrace, oCallBack
                End If
                
                Set oPlate = Nothing
            End If
        End If
    Next oCheckingTrace
    
Cleanup:
    Set oCallBack = Nothing
    Set oCheckingTrace = Nothing
    Set oCheckingCurves = Nothing
    Set oColHoles = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Sub

'******************************************************************************
' Routine: CheckForEdgeLineOnPlatePart
'
' Abstract: Check hole with Edge Line of the plate
'
' Description:
'******************************************************************************
Public Sub CheckForPlatePart(oCheckingCurves As IMSCoreCollections.IJElements, _
                             oPlate As IJPlate, oHoleTrace As IJHoleTraceAE, _
                             oCallBack As IJCheckMfctyCallback)
    Const METHOD = "CheckForPlatePart"
    On Error GoTo ErrorHandler
    
    Dim oComparedCurves As IMSCoreCollections.IJElements
    Dim dblMinDistance As Double       'Get real minimum distance between hole and object what you want to compare.
    Dim oStrMsg As String              'display the wrong hole or plate
    
''    'Get the parent system from HoleTrace
''    Dim oPlate As IJPlate
''    Dim oObject As Object
''
''    Set oObject = GetPlateSystemFromHoleTrace(oHoleTrace)
''    If TypeOf oObject Is IJPlate Then
''        Set oPlate = oObject
''    End If
''    Set oObject = Nothing
''
''    'If the plate is root, the hole cannot be compared with the system root.
''    If oPlate Is Nothing Then
''        oStrMsg = "The parents of standalone part can be root. Skip the checking between hole and root."
''        PassStringToCallBack PROG_ID, oHoleTrace, oStrMsg, oCallBack, 2
''    Else
        
        Dim oChildStructs As IMSCoreCollections.IJElements
        Dim oPlateParts As IMSCoreCollections.IJElements
        Dim oPlatePart As IJPlatePart
        Dim oPlateThickness As Double

        'Get the thickness of plate
        oPlateThickness = oPlate.thickness
        
        'Get the curves from plate
        Set oChildStructs = GetChildOfThePlate(oPlate)
        Set oPlateParts = GetPlatePartFromStruct(oChildStructs)

        For Each oPlatePart In oPlateParts
            Set oComparedCurves = GetGeometryFromPlatePart(oPlatePart)
            
            If oComparedCurves.Count > 0 Then
                'Get the minimun distance
                dblMinDistance = GetMDOfPartBTWCurveAndObject(oCheckingCurves, _
                                                             oComparedCurves, _
                                                             oPlateThickness)
                If dblMinDistance >= 0 Then
                    'Check the distance and return callback
                    CheckRulesAndCallBack PROG_ID, oHoleTrace, oPlatePart, _
                                          dblMinDistance, m_MDPartEdge, oCallBack, 2
                End If
            End If
            Set oPlatePart = Nothing
        Next

        Set oPlateParts = Nothing
        Set oChildStructs = Nothing

''    End If

Cleanup:
''    Set oPlate = Nothing
    Set oComparedCurves = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Sub

'******************************************************************************
' Routine: CheckForCoamingOnPlate
'
' Abstract: Check hole with Seam line on the plate
'
' Description:
'******************************************************************************
Public Sub CheckForCoamingOnPlate(oCheckingCurves As IMSCoreCollections.IJElements, _
                                  oPlate As IJPlate, oHoleTrace As IJHoleTraceAE, _
                                  oCallBack As IJCheckMfctyCallback)
    Const METHOD = "CheckForCoamingOnPlate"
    On Error GoTo ErrorHandler
    
    Dim oComparedCurves As IMSCoreCollections.IJElements
    Dim dblMinDistance As Double       'Get real minimum distance between hole and object what you want to compare.
    Dim oStrMsg As String              'display the wrong hole or plate
    
''    'Get the parent system from HoleTrace
''    Dim oPlate As IJPlate
''    Dim oObject As Object
''
''    Set oObject = GetPlateSystemFromHoleTrace(oHoleTrace)
''    If TypeOf oObject Is IJPlate Then
''        Set oPlate = oObject
''    End If
''    Set oObject = Nothing
''
''    'If the plate is root, the hole cannot be compared with the system root.
''    If oPlate Is Nothing Then
''        oStrMsg = "The parents of this hole is a standalone part or Profile. Skip the checking between hole and root."
''        PassStringToCallBack PROG_ID, oHoleTrace, oStrMsg, oCallBack, 2
''    Else
        
        Dim oChildStructs As IMSCoreCollections.IJElements
        Dim oHoleTraceAEs As IMSCoreCollections.IJElements
        Dim oHoleTraceAE As IJHoleTraceAE
        Dim oCoamings As IMSCoreCollections.IJElements
        'Dim oCoaming As IJDBracketPart
        Dim oCoaming As Object
        'Get the curves from plate
        Set oChildStructs = GetChildOfThePlate(oPlate)
        Set oHoleTraceAEs = GetHoleTraceFromFromStruct(oChildStructs)
        
        For Each oHoleTraceAE In oHoleTraceAEs
            Set oCoamings = GetCoamingFromStruct(oHoleTraceAE)
                
            For Each oCoaming In oCoamings
                Set oComparedCurves = New IMSCoreCollections.JObjectCollection
                Set oComparedCurves = GetGeometryFromCoaming(oCoaming)
                                
                If oComparedCurves.Count > 0 Then
                    'Get the minimun distance
                    dblMinDistance = GetMDBTWCurveAndObject(oCheckingCurves, _
                                                        oComparedCurves)
    
                    If dblMinDistance >= 0 Then
                        'Check the distance and return callback
                        CheckRulesAndCallBack PROG_ID, oHoleTrace, oCoaming, _
                                              dblMinDistance, m_MDOutfitting, oCallBack, 2
                    End If
                End If
                Set oCoaming = Nothing
            Next
            
            Set oCoamings = Nothing
            Set oHoleTraceAE = Nothing
        Next
        
        Set oHoleTraceAEs = Nothing
        
''    End If

Cleanup:
''    Set oPlate = Nothing
    Set oComparedCurves = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Sub

'******************************************************************************
' Routine: GetPlatePartFromStruct
'
' Abstract: Get Edge Lines from plate part
'
' Description:
'******************************************************************************
Public Function GetPlatePartFromStruct(ByVal oStruct As IMSCoreCollections.IJElements) As IMSCoreCollections.IJElements
    Const METHOD As String = "GetPlatePartFromStruct"
    On Error GoTo ErrorHandler
    
    Dim oElem As IMSCoreCollections.IJElements
    Set oElem = New IMSCoreCollections.JObjectCollection
    
    Dim oObject As Object
    For Each oObject In oStruct
        If TypeOf oObject Is IJPlatePart Then
            oElem.Add oObject
        End If
    Next oObject

    Set oObject = Nothing
    Set GetPlatePartFromStruct = oElem
    
Cleanup:
    Set oElem = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetCoamingFromStruct
'
' Abstract: Get Openings from child of plate
'
' Description:
'******************************************************************************
Public Function GetCoamingFromStruct(ByVal oHoleTrace As IJHoleTraceAE) As IMSCoreCollections.IJElements
    Const METHOD As String = "GetCoamingFromStruct"
    On Error GoTo ErrorHandler
    
    Dim oElem As IMSCoreCollections.IJElements
    Set oElem = New IMSCoreCollections.JObjectCollection
    
    Dim oSysParent As IJSystem
    'Dim oSysChildren As IMSCoreCollections.IJElements
    Dim oSysChildren As IJDTargetObjectCol
    
    Set oSysParent = oHoleTrace
    Set oSysChildren = oSysParent.GetChildren
        
    Dim oObject As Object
    Dim iIndex As Integer
    
    'For Each oObject In oSysChildren
    For iIndex = 1 To oSysChildren.Count
        Set oObject = oSysChildren.Item(iIndex)
        If TypeOf oObject Is IJProfilePart Or _
           TypeOf oObject Is IJPlatePart Or _
           TypeOf oObject Is IJEquipment Then
            
           oElem.Add oObject
        End If
    Next 'oObject

    Set oObject = Nothing
    Set GetCoamingFromStruct = oElem
    
Cleanup:
    Set oElem = Nothing
    Set oSysParent = Nothing
    Set oSysChildren = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetHoleTraceFromFromStruct
'
' Abstract: Get HoleTrace from child of plate
'
' Description:
'******************************************************************************
Public Function GetHoleTraceFromFromStruct(ByVal oStruct As IMSCoreCollections.IJElements) As IMSCoreCollections.IJElements
    Const METHOD As String = "GetHoleTraceFromFromStruct"
    On Error GoTo ErrorHandler
    
    Dim oElem As IMSCoreCollections.IJElements
    Set oElem = New IMSCoreCollections.JObjectCollection
    
    Dim oObject As Object
    For Each oObject In oStruct
        If TypeOf oObject Is IJHoleTraceAE Then
            oElem.Add oObject
        End If
    Next oObject

    Set oObject = Nothing
    Set GetHoleTraceFromFromStruct = oElem
    
Cleanup:
    Set oElem = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function
 
