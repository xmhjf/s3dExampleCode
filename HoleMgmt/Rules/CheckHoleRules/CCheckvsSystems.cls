VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCheckvsSystems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
' Copyright (C) 1998-2000 Intergraph Corporation.  All Rights Reserved.
'
' File: CCheckvsSystems.cls
'
' Author: sypark@ship.samsung.co.kr
'
' Abstract: Check the minmun distance between Holes and System
'           The list of system which should be checked between hole are following.
'           1. Edge line of the plate
'           2. Profile line on the plate
'           3. Seam line on the plate
'           4. Opening on the plate.
'           If you want to add more system what you want to check, You can add the code.
' Description:
' IJCheckMfcty is implemented by GrobalCheck.tlb.
'********************************************************************
Option Explicit

Private Const Module = "GSCADHMRules.CCheckvsSystems:"
Private Const PROG_ID = "GSCADHMRules.CCheckvsSystems"

Implements IJCheckMfcty

Dim m_MDEdge As Double       'For minimum distance defined by user between hole and Edge of structure.
Dim m_MDProfile As Double    'For minimum distance defined by user between hole and Profile line on the plate.
Dim m_MDSeam As Double       'For minimum distance defined by user between hole and Seam line on the plate.
Dim m_MDOpening As Double    'For minimum distance defined by user between hole and Opening on the plate.

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
    
    m_MDEdge = 0.05             'If the nearest distance between hole and Edge of structure is less than 50 mm, should be checked.
    m_MDProfile = 0.05          'If the nearest distance between hole and Profile is less than 50 mm, should be checked.
    m_MDSeam = 0.05             'If the nearest distance between hole and seamline is less than 50 mm, should be checked.
    m_MDOpening = 0.05          'If the nearest distance between hole and Opening is less than 50 mm, should be checked.
    
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
            
                    'Check the hole with the edge line of plate
                    CheckForEdgeLineOnPlate oCheckingCurves, oPlate, oCheckingTrace, oCallBack
                    
            '''     'Check the hole with the edge line of plate
                    CheckForProfileOnPlate oCheckingCurves, oPlate, oCheckingTrace, oCallBack
            '''
                    'Check the hole with the edge line of plate
                    CheckForSeamOnPlate oCheckingCurves, oPlate, oCheckingTrace, oCallBack
            '''
                    'Check the hole with the edge line of plate
                    CheckForOpeningOnPlate oCheckingCurves, oPlate, oCheckingTrace, oCallBack
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
' Routine: CheckForEdgeLineOnPlate
'
' Abstract: Check hole with Edge Line of the plate
'
' Description:
'******************************************************************************
Public Sub CheckForEdgeLineOnPlate(oCheckingCurves As IMSCoreCollections.IJElements, _
                                   oPlate As IJPlate, oHoleTrace As IJHoleTraceAE, _
                                   oCallBack As IJCheckMfctyCallback)
    Const METHOD = "CheckForEdgeLineOnPlate"
    On Error GoTo ErrorHandler
    
    Dim oComparedCurves As IMSCoreCollections.IJElements
    Dim dblMinDistance As Double       'Get real minimum distance between hole and object what you want to compare.

''    Dim oPlate As IJPlate
''    Dim oStrMsg As String              'display the wrong hole or plate
    
''    'Get the parent system from HoleTrace
''    Dim oObject As Object
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
        
        'Get the curves from plate
        Set oComparedCurves = GetEdgeLineFromPlateSystem(oPlate)
    
        If oComparedCurves.Count > 0 Then
            'Get the minimun distance
            dblMinDistance = GetMDBTWCurveAndObject(oCheckingCurves, _
                                            oComparedCurves)
    
            If dblMinDistance >= 0 Then
                'Check the distance and return callback
                CheckRulesAndCallBack PROG_ID, oHoleTrace, oPlate, _
                                           dblMinDistance, m_MDEdge, oCallBack, 2
            End If
        End If
        
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
' Routine: CheckForProfileOnPlate
'
' Abstract: Check hole with Profile
'
' Description:
'******************************************************************************
Public Sub CheckForProfileOnPlate(oCheckingCurves As IMSCoreCollections.IJElements, _
                                  oPlate As IJPlate, oHoleTrace As IJHoleTraceAE, _
                                  oCallBack As IJCheckMfctyCallback)
    Const METHOD = "CheckForProfileOnPlate"
    On Error GoTo ErrorHandler
    
    Dim oComparedCurves As IMSCoreCollections.IJElements
    Dim dblMinDistance As Double       'Get real minimum distance between hole and object what you want to compare.

''    Dim oPlate As IJPlate
''    Dim oStrMsg As String              'display the wrong hole or plate
    
''    'Get the parent system from HoleTrace
''    Dim oObject As Object
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
        Dim oProfiles As IMSCoreCollections.IJElements
        Dim oProfile As IJStiffenerSystem
        
        'Get the curves from plate
        Set oChildStructs = GetChildOfThePlate(oPlate)
        Set oProfiles = GetProfileFromStruct(oChildStructs)
        
        For Each oProfile In oProfiles
            Set oComparedCurves = GetGeometryFromProfile(oProfile)
            
            If oComparedCurves.Count > 0 Then
                'Get the minimun distance
                dblMinDistance = GetMDBTWCurveAndObject(oCheckingCurves, _
                                                    oComparedCurves)
                If dblMinDistance >= 0 Then
                    'Check the distance and return callback
                    CheckRulesAndCallBack PROG_ID, oHoleTrace, oProfile, _
                                          dblMinDistance, m_MDEdge, oCallBack, 2
                End If
            End If
            
            Set oProfile = Nothing
        Next
    
        Set oProfiles = Nothing
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
' Routine: CheckForSeamOnPlate
'
' Abstract: Check hole with Seam line on the plate
'
' Description:
'******************************************************************************
Public Sub CheckForSeamOnPlate(oCheckingCurves As IMSCoreCollections.IJElements, _
                               oPlate As IJPlate, oHoleTrace As IJHoleTraceAE, _
                               oCallBack As IJCheckMfctyCallback)
    Const METHOD = "CheckForSeamOnPlate"
    On Error GoTo ErrorHandler
    
    Dim oComparedCurves As IMSCoreCollections.IJElements
    Dim dblMinDistance As Double       'Get real minimum distance between hole and object what you want to compare.

''    Dim oPlate As IJPlate
''    Dim oStrMsg As String              'display the wrong hole or plate
''
''    'Get the parent system from HoleTrace
''    Dim oObject As Object
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
        Dim oSeams As IMSCoreCollections.IJElements
        Dim oSeam As IJSeam
        
        'Get the curves from plate
        Set oChildStructs = GetChildOfThePlate(oPlate)
        Set oSeams = GetSeamLineFromStruct(oChildStructs)
            
        For Each oSeam In oSeams
            Dim oCurve As IJCurve
            Set oComparedCurves = New IMSCoreCollections.JObjectCollection
            Set oCurve = oSeam
            
            oComparedCurves.Clear
            oComparedCurves.Add oCurve
            
            If oComparedCurves.Count > 0 Then
                'Get the minimun distance
                dblMinDistance = GetMDBTWCurveAndObject(oCheckingCurves, _
                                                    oComparedCurves)

                If dblMinDistance >= 0 Then
                    'Check the distance and return callback
                    CheckRulesAndCallBack PROG_ID, oHoleTrace, oSeam, _
                                          dblMinDistance, m_MDEdge, oCallBack, 2
                End If
            End If
            Set oSeam = Nothing
            Set oCurve = Nothing
         Next

        Set oSeams = Nothing
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
' Routine: CheckForOpeningOnPlate
'
' Abstract: Check hole with Opening
'
' Description:
'******************************************************************************
Public Sub CheckForOpeningOnPlate(oCheckingCurves As IMSCoreCollections.IJElements, _
                                  oPlate As IJPlate, oHoleTrace As IJHoleTraceAE, _
                                  oCallBack As IJCheckMfctyCallback)
    Const METHOD = "CheckForOpeningOnPlate"
    On Error GoTo ErrorHandler
    
    Dim oComparedCurves As IMSCoreCollections.IJElements
    Dim dblMinDistance As Double       'Get real minimum distance between hole and object what you want to compare.

''    Dim oPlate As IJPlate
''    Dim oStrMsg As String              'display the wrong hole or plate
''
''    'Get the parent system from HoleTrace
''    Dim oObject As Object
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
        Dim oOpenings As IMSCoreCollections.IJElements
        Dim oOpening As Object
        
        'Get the curves from plate
        Set oChildStructs = GetChildOfThePlate(oPlate)
        Set oOpenings = GetOpeningsFromStruct(oChildStructs)
            
        For Each oOpening In oOpenings
            Dim oCurve As IJCurve
            Set oComparedCurves = New IMSCoreCollections.JObjectCollection
            ' Fix for TR# 181212:(TR with SR) "Check against connected system objects" check for hole mgmt fails
            ' Get the complex string from the wirebody and then save it to the oComparedCurves collection. Since
            ' GetMDBTWCurveAndObject requires oComparedCurves, having IJCurves.

            Dim oWireBody As IJWireBody
            Dim oElements As IJElements

            Set oWireBody = oOpening
            oWireBody.GetComplexStrings Nothing, oElements

            If oElements.Count > 0 Then
               Set oCurve = oElements.Item(1)
            End If
            
            oComparedCurves.Clear
            oComparedCurves.Add oCurve
            
            If oComparedCurves.Count > 0 Then
                'Get the minimun distance
                dblMinDistance = GetMDBTWCurveAndObject(oCheckingCurves, _
                                                    oComparedCurves)

                If dblMinDistance >= 0 Then
                    'Check the distance and return callback
                    CheckRulesAndCallBack PROG_ID, oHoleTrace, oOpening, _
                                          dblMinDistance, m_MDEdge, oCallBack, 2
                End If
            End If
            Set oOpening = Nothing
            
            Set oWireBody = Nothing
            Set oElements = Nothing
            Set oCurve = Nothing
            
        Next

        Set oOpenings = Nothing
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
' Routine: GetEdgeLineFromPlateSystem
'
' Abstract: Get Edge Line from port
'
' Description:
'******************************************************************************
Public Function GetEdgeLineFromPlateSystem(ByVal oPlate As IJPlate) As IMSCoreCollections.IJElements
    Const METHOD As String = "GetEdgeLineFromPlateSystem"
    On Error GoTo ErrorHandler
    
    Dim oConnectable As IJConnectable
    Set oConnectable = oPlate
    
    Dim oPlateEdgeElm As IMSCoreCollections.IJElements
    
    Dim oElem As IMSCoreCollections.IJElements
    Set oElem = New IMSCoreCollections.JObjectCollection
    
    Call oConnectable.enumPorts(oPlateEdgeElm, PortEdge)
        
    Dim oPort As IJPort
    Dim oObject As Object
    Dim oGeomObject As Object
    Dim oGeomCurve As IJCurve
    For Each oObject In oPlateEdgeElm
        Set oPort = oObject
        If Not oPort Is Nothing Then
        
           Set oGeomObject = oPort
           Set oGeomCurve = ConvertObjectToCurve(oGeomObject)
           If Not oGeomCurve Is Nothing Then
                oElem.Add oGeomCurve
           End If
           
           Set oGeomCurve = Nothing
           Set oGeomObject = Nothing
           
        End If
        
        Set oObject = Nothing
        Set oPort = Nothing
        
    Next oObject
    
    Set GetEdgeLineFromPlateSystem = oElem
    
Cleanup:
    Set oElem = Nothing
    Set oPlateEdgeElm = Nothing
    Set oConnectable = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetSeamLineFromStruct
'
' Abstract: Get Seam Line from child of plate
'
' Description:
'******************************************************************************
Public Function GetSeamLineFromStruct(ByVal oStruct As IMSCoreCollections.IJElements) As IMSCoreCollections.IJElements
    Const METHOD As String = "GetSeamLineFromStruct"
    On Error GoTo ErrorHandler
    
    Dim oElem As IMSCoreCollections.IJElements
    Set oElem = New IMSCoreCollections.JObjectCollection
    
    Dim oObject As Object
    For Each oObject In oStruct
        If TypeOf oObject Is IJSeam Then
            oElem.Add oObject
        End If
    Next oObject

    Set oObject = Nothing
    Set GetSeamLineFromStruct = oElem
    
Cleanup:
    Set oElem = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetOpeningsFromStruct
'
' Abstract: Get Openings from child of plate
'
' Description:
'******************************************************************************
Public Function GetOpeningsFromStruct(ByVal oStruct As IMSCoreCollections.IJElements) As IMSCoreCollections.IJElements
    Const METHOD As String = "GetOpeningsFromStruct"
    On Error GoTo ErrorHandler
    
    Dim oElem As IMSCoreCollections.IJElements
    Set oElem = New IMSCoreCollections.JObjectCollection
    
    Dim oObject As Object
    For Each oObject In oStruct
        If TypeOf oObject Is IJOpeningEntity Then
            oElem.Add oObject
        End If
    Next oObject

    Set oObject = Nothing
    Set GetOpeningsFromStruct = oElem
    
Cleanup:
    Set oElem = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetProfileFromStruct
'
' Abstract: Get profile from child of plate
'
' Description:
'******************************************************************************
Public Function GetProfileFromStruct(ByVal oStruct As IMSCoreCollections.IJElements) As IMSCoreCollections.IJElements
    Const METHOD As String = "GetProfileFromStruct"
    On Error GoTo ErrorHandler
    
    Dim oElem As IMSCoreCollections.IJElements
    Set oElem = New IMSCoreCollections.JObjectCollection
    
    Dim oObject As Object
    For Each oObject In oStruct
        If TypeOf oObject Is IJStiffenerSystem Then
            oElem.Add oObject
        End If
    Next oObject

    Set oObject = Nothing
    Set GetProfileFromStruct = oElem
    
Cleanup:
    Set oElem = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function
 
