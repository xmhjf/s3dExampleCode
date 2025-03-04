Attribute VB_Name = "modCommon"
'********************************************************************
' Copyright (C) 1998-2000 Intergraph Corporation.  All Rights Reserved.
'
' File: modCommon.bas
'
' Author: sypark@ship.samsung.co.kr
'
' Abstract: common function
'
' Description:
'********************************************************************
Option Explicit

Const Module = "GSCADHMRules.modCommon:"

Public Enum ESeverityIndex
    siError = 101
    siWarining = 102
End Enum

'******************************************************************************
' Routine: GetHoleTraceCurves
'
' Abstract: Get the elements which has a entry of geometry of symbol in HoleTrace
'
' Description: Get the symbol of HoleTrace and then get the geometry as complexString
'******************************************************************************
Public Function GetPOMConnection(strDbType As String) As IJDPOM
    Const METHOD = "GetPOMConnection"
    On Error GoTo ErrHandler
    
      
    Dim oAccessMiddle As IJDAccessMiddle
    Set oAccessMiddle = GetJContext().GetService("ConnectMiddle")
    Set GetPOMConnection = oAccessMiddle.GetResourceManagerFromType(strDbType)
    
    m_oActiveConn = GetPOMConnection.DatabaseID

Cleanup:
    Set oAccessMiddle = Nothing
    Exit Function
    
ErrHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetHoleTraceCurves
'
' Abstract: Get the elements which has a entry of geometry of symbol in HoleTrace
'
' Description: Get the symbol of HoleTrace and then get the geometry as complexString
'******************************************************************************
Public Function GetHoleTraceCurves(oHole As IJHoleTraceAE) As IMSCoreCollections.IJElements
    Const METHOD = "GetHoleTraceCurves"
    On Error GoTo ErrorHandler
    
    Dim oHoleTraceCurves As IMSCoreCollections.IJElements
    Set oHoleTraceCurves = New IMSCoreCollections.JObjectCollection
    
    Dim oCurves As IMSCoreCollections.IJElements
    Dim oCurve As Object
    Dim cenX As Double, cenY As Double, cenZ As Double
        
    'This should be skip when complexstring cannot be retrived.
    On Error Resume Next
    Dim oComplexStrings As IJDObjectCollection
    Dim oComplexString As IJComplexString
    Set oComplexStrings = oHole.GetComplexStrings
    
    If oComplexStrings.Count > 0 Then
        For Each oComplexString In oComplexStrings
            If Not oComplexString Is Nothing Then
               oComplexString.GetCurves oCurves
                  For Each oCurve In oCurves
                    If Not oCurve Is Nothing Then
                        oHoleTraceCurves.Add oCurve
                    End If
                  Next
            End If
        Next
    End If
    
    Set GetHoleTraceCurves = oHoleTraceCurves
    
Cleanup:
    Set oHoleTraceCurves = Nothing
    Set oCurve = Nothing
    Set oComplexStrings = Nothing
    Set oComplexString = Nothing
    Set oCurves = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetMinimumDistance
'
' Abstract: Check the minimum distance between Hole and Hole
'
' Description: Compare the geometry as line and arc to each Hole
' How to get the minimun distance.
'   /1 Get two first curves from comparing Hole and compared Hole.
'   /2 Get the distance between two curves.
'      This distance  should be default value to be compared by other distances from other curves.
'   /3 Get the distance between remained curves, compared with above distance.
'   /4 Get the minimun distance among them.
'******************************************************************************
Public Function GetMinimumDistance(oFirstCurves As IMSCoreCollections.IJElements, oSecondCurves As IMSCoreCollections.IJElements) As Double
    Const METHOD = "GetMinimumDistance"
    On Error GoTo ErrorHandler

    Dim dblDistance As Double
    Dim dblMinDist As Double
    Dim oFirstCurve As IJCurve
    Dim oSecondCurve As IJCurve
    Dim srcX As Double, srcY As Double, srcZ As Double
    Dim inX As Double, inY As Double, inZ As Double
    
    dblMinDist = -1
    
    For Each oFirstCurve In oFirstCurves
        For Each oSecondCurve In oSecondCurves
            oFirstCurve.DistanceBetween oSecondCurve, dblDistance, _
                                        srcX, srcY, srcZ, inX, inY, inZ
            If dblDistance < dblMinDist Or dblMinDist = -1 Then
                dblMinDist = dblDistance
            End If
        Next oSecondCurve
    Next oFirstCurve

    GetMinimumDistance = dblMinDist
    
Cleanup:
    Set oFirstCurve = Nothing
    Set oSecondCurve = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetMDBTWCurveAndObject
'
' Abstract: Check the minimum distance between Hole and object
'
' Description: Compare the geometry as line and arc of object to each Hole
'
' How to get the minimun distance.
'   /1 Get two first curves from comparing object and compared object.
'      For instance, get the circle from hole and the line from opening
'   /2 Get the distance between two curves.
'      This distance  should be default value to be compared by other distances from other curves.
'   /3 Get the distance between remained curves, compared with above distance.
'   /4 Get the minimun distance among them.
'******************************************************************************
Public Function GetMDBTWCurveAndObject(oFirstCurves As IMSCoreCollections.IJElements, oSecondCurves As IMSCoreCollections.IJElements) As Double
    Const METHOD = "GetMDBTWCurveAndObject"
    On Error GoTo ErrorHandler

    Dim dblDistance As Double
    Dim dblMinDist As Double
    Dim oFirstCurve As IJCurve 'They are always Curves of Hole
    Dim oObject As Object      'They are always Curves of compared object. They are various GType.
    Dim oSecondCurve As IJCurve
    Dim oIsCurve As Boolean
    Dim srcX As Double, srcY As Double, srcZ As Double
    Dim inX As Double, inY As Double, inZ As Double
    
    'It should return as -1, If there has something wrong with firstcurve and secondcurve.
    dblMinDist = -1
    
    For Each oFirstCurve In oFirstCurves
        For Each oObject In oSecondCurves
            
            'Check if the curve is object which has implemented by IJCurve, the minimun distance can be retrieved.
            'But, If this object is like projection, plane so on, the minimun distance cannot be retrieved.
            'So, if olscurve is false, try to get the IJCurve from object and compare again.
            oIsCurve = CheckObjectIsCurve(oObject)
            
            'If true, this curve has IJCurve.
            If oIsCurve = True Then
               
                Set oSecondCurve = ConvertObjectToCurve(oObject)
                oFirstCurve.DistanceBetween oSecondCurve, dblDistance, _
                                            srcX, srcY, srcZ, inX, inY, inZ
                If dblDistance < dblMinDist Or dblMinDist = -1 Then
                    dblMinDist = dblDistance
                End If
          
            'If false, this curve don't have IJCurve. try to get curves from object
            Else
                'If the curve is complexString or other GType like Projection, DistanceBetween method can't take this.
                'So, Get the curve from ComplexString.
                Dim oElemCurves As IMSCoreCollections.IJElements
                Set oElemCurves = ConvertGTypeToCurves(oObject)
                
                Dim oSecondObject As IJCurve
                For Each oSecondObject In oElemCurves
            
                   oIsCurve = CheckObjectIsCurve(oSecondObject)
                   
                   If oIsCurve = True Then
                        oFirstCurve.DistanceBetween oSecondObject, dblDistance, _
                                                    srcX, srcY, srcZ, inX, inY, inZ
                        If dblDistance < dblMinDist Or dblMinDist = -1 Then
                             dblMinDist = dblDistance
                        End If
                   Else
                        'The plane has complexstring, Complexstring also made by lots of curves.
                        'If the curve is complexString, DistanceBetween method can't take this.
                        'So, Get the curves from ComplexString again.
                        Dim oElemCStrings As IMSCoreCollections.IJElements
                        Set oElemCStrings = ConvertGTypeToCurves(oSecondObject)
                
                        Dim oSecondCString As IJCurve
                        For Each oSecondCString In oElemCStrings
                             oFirstCurve.DistanceBetween oSecondCString, dblDistance, _
                                                    srcX, srcY, srcZ, inX, inY, inZ

                            If dblDistance < dblMinDist Or dblMinDist = -1 Then
                                 dblMinDist = dblDistance
                            End If
                        Next
                        Set oElemCStrings = Nothing
                        Set oSecondCString = Nothing
                   End If
                Next
                
                Set oSecondObject = Nothing
                Set oElemCurves = Nothing
            End If
            
        Next oObject
    Next oFirstCurve

    GetMDBTWCurveAndObject = dblMinDist
    
Cleanup:
    Set oFirstCurve = Nothing
    Set oObject = Nothing
    Set oSecondCurve = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetMDOfPartBTWCurveAndObject
'
' Abstract: Check the minimum distance between Hole and object
'
' Description: Compare the geometry as line and arc of object to each Hole
'
' How to get the minimun distance.
'   /1 Get two first curves from comparing object and compared object.
'      For instance, get the circle from hole and the line from opening
'   /2 Get the distance between two curves.
'      This distance  should be default value to be compared by other distances from other curves.
'   /3 Get the distance between remained curves, compared with above distance.
'   /4 Get the minimun distance among them.
'******************************************************************************
Public Function GetMDOfPartBTWCurveAndObject(oFirstCurves As IMSCoreCollections.IJElements, _
                                             oSecondCurves As IMSCoreCollections.IJElements, _
                                             oThickness As Double) As Double
    Const METHOD = "GetMDOfPartBTWCurveAndObject"
    On Error GoTo ErrorHandler
    
    Dim oSameObject As Boolean
    Dim dblDistance As Double
    Dim dblMinDist As Double
    Dim oFirstCurve As IJCurve 'They are always Curves of Hole
    Dim oObject As Object      'They are always Curves of compared object. They are various GType.
    Dim oSecondCurve As IJCurve
    Dim oIsCurve As Boolean
    Dim srcX As Double, srcY As Double, srcZ As Double
    Dim inX As Double, inY As Double, inZ As Double
    
    'It should return as -1, If there has something wrong with firstcurve and secondcurve.
    oSameObject = False
    dblMinDist = -1

    For Each oFirstCurve In oFirstCurves
        For Each oObject In oSecondCurves
            
            'Check if the curve is object which has implemented by IJCurve, the minimun distance can be retrieved.
            'But, If this object is like projection, plane so on, the minimun distance cannot be retrieved.
            'So, if olscurve is false, try to get the IJCurve from object and compare again.
            oIsCurve = CheckObjectIsCurve(oObject)
            
            'If true, this curve has IJCurve.
            If oIsCurve = True Then
               
                Set oSecondCurve = ConvertObjectToCurve(oObject)
                'Check if they are same between Hole and part.
                oSameObject = CheckSameObject(oFirstCurve, oSecondCurve, oThickness)
                If oSameObject = False Then
                    oFirstCurve.DistanceBetween oSecondCurve, dblDistance, _
                                                    srcX, srcY, srcZ, inX, inY, inZ
                    If dblDistance < dblMinDist Or dblMinDist = -1 Then
                        dblMinDist = dblDistance
                    End If
                End If

            'If false, this curve don't have IJCurve. try to get curves from object
            Else
                'If the curve is complexString or other GType like Projection, DistanceBetween method can't take this.
                'So, Get the curve from ComplexString.
                Dim oElemCurves As IMSCoreCollections.IJElements
                Set oElemCurves = ConvertGTypeToCurves(oObject)
                
                Dim oSecondObject As IJCurve
                For Each oSecondObject In oElemCurves
            
                   oIsCurve = CheckObjectIsCurve(oSecondObject)
                   
                   If oIsCurve = True Then
                   
                        'Check if they are same between Hole and part.
                        oSameObject = CheckSameObject(oFirstCurve, oSecondObject, oThickness)
                        If oSameObject = False Then
                            oFirstCurve.DistanceBetween oSecondObject, dblDistance, _
                                                        srcX, srcY, srcZ, inX, inY, inZ
                            If dblDistance < dblMinDist Or dblMinDist = -1 Then
                                dblMinDist = dblDistance
                            End If
                        End If
                   Else
                        'The plane has complexstring, Complexstring also made by lots of curves.
                        'If the curve is complexString, DistanceBetween method can't take this.
                        'So, Get the curves from ComplexString again.
                        Dim oElemCStrings As IMSCoreCollections.IJElements
                        Set oElemCStrings = ConvertGTypeToCurves(oSecondObject)
                        
                        'Check if they are same between Hole and part.
                        oSameObject = CheckHoleAndCut(oFirstCurve, oElemCStrings, oThickness)
                        
                        If oSameObject = False Then
                            Dim oSecondCString As IJCurve
                            For Each oSecondCString In oElemCStrings
                                oFirstCurve.DistanceBetween oSecondCString, dblDistance, _
                                                        srcX, srcY, srcZ, inX, inY, inZ
                                If dblDistance <= dblMinDist Or dblMinDist = -1 Then
                                    dblMinDist = dblDistance
                                End If
                            Next
                        End If
                        
                        Set oElemCStrings = Nothing
                        Set oSecondCString = Nothing
                   End If
                Next
                
                Set oSecondObject = Nothing
                Set oElemCurves = Nothing
            End If
            
        Next oObject
    Next oFirstCurve

    GetMDOfPartBTWCurveAndObject = dblMinDist
    
Cleanup:
    Set oFirstCurve = Nothing
    Set oObject = Nothing
    Set oSecondCurve = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: CheckOverall
'
' Abstract: Check the parents of HoleTrace
'
' Description: If the parents of HoleTrace is barcket and beam, return warning
'              If the parents of holetrace is Plate, return warning.
'              This is an example for user
'******************************************************************************
Public Sub CheckOverall(strProgID As String, oHoleTraceAE As IJHoleTraceAE, _
                        oCallBack As IJCheckMfctyCallback, lngOptionCode As Long)
    Const METHOD As String = "CheckOverall"
    On Error GoTo ErrorHandler

    Dim oSysParent As IJSystem
    Dim oSysChild As IJSystemChild
    Dim oParent As Object
    Dim oNamedItem As IJNamedItem
    Dim strParentName As String
    Dim strHoleName As String
    Dim strMessage As String
    Dim bError As Boolean
    Dim eSeverity As ESeverityIndex
    
    Set oSysChild = oHoleTraceAE
    Set oSysParent = oSysChild.GetParent
    Set oParent = oSysParent
    Set oNamedItem = oHoleTraceAE
    
    strHoleName = oNamedItem.Name
    
    If TypeOf oParent Is IJBeam Then
        strParentName = "Beam"
        eSeverity = siWarining
        bError = True
    Else
        bError = False
    End If
    
    If bError Then
        strMessage = "The example for checking parent." & " : " & "The Hole (" & strHoleName & ") is placed on the " & _
                     strParentName & "."
        oCallBack.OnCheckError oHoleTraceAE, strProgID, eSeverity, _
                          lngOptionCode, strMessage, "", ""
    End If
    
Cleanup:
    Set oSysParent = Nothing
    Set oSysChild = Nothing
    Set oParent = Nothing
    Set oNamedItem = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Sub

'******************************************************************************
' Routine: GetObjectName
'
' Abstract: Get the name of object
'
' Description:
'******************************************************************************
Public Function GetObjectName(oElem As Object) As String
    Const METHOD = "GetObjectName"
    On Error GoTo ErrorHandler
    
    Dim oNamedItem As IJNamedItem
    If oElem Is Nothing Then
        GetObjectName = "<Object Not Found>"
    Else
        On Error Resume Next
        Set oNamedItem = oElem
        On Error GoTo ErrorHandler
        If oNamedItem Is Nothing Then
            GetObjectName = "<No Prod Model>"
        Else
            GetObjectName = oNamedItem.Name
        End If
    End If

Cleanup:
    Set oNamedItem = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: SetCollectionHoles
'
' Abstract: retrives the hole from elements.
'
' Description: Get the holes from the IJElements.
'******************************************************************************
Public Sub SetCollectionHoles(ByVal pCollection As IMSCoreCollections.IJElements, oColHoles As IMSCoreCollections.IJElements)
    Const METHOD As String = "SetCollectionHoles"
    On Error GoTo ErrorHandler

    Dim vSupMnkr As Variant
    Dim oEntity As Object

    Dim oPOM As IJDPOM
    Set oPOM = GetPOMConnection("model")
    
    Set oColHoles = New IMSCoreCollections.JObjectCollection

    ' get the HoleTrace from miniker
    For Each vSupMnkr In pCollection 'elements
        If oPOM.SupportsInterface(vSupMnkr, "IJHoleTraceAE") Then
            On Error Resume Next
            Set oEntity = oPOM.GetObject(vSupMnkr)
            If TypeOf oEntity Is IJHoleTraceAE Then
                oColHoles.Add oEntity
            End If
        End If
    Next vSupMnkr

Cleanup:
    Set oEntity = Nothing
    Set vSupMnkr = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Sub

'******************************************************************************
' Routine: CheckRulesAndCallBack
'
' Abstract: Compare the minimum distance and distance between objects,
'
' Description: Return the callback to display the error message on the list box
'******************************************************************************
Public Sub CheckRulesAndCallBack(strProgID As String, _
                                 oHoleTraceAE As IJHoleTraceAE, _
                                 oComparedObject As Object, _
                                 dblMinDistance As Double, _
                                 dblRuleDistance As Double, _
                                 oCallBack As IJCheckMfctyCallback, _
                                 lngOptionCode As Long)
    Const METHOD As String = "CheckRulesAndCallBack"
    On Error GoTo ErrorHandler
    
    Dim index As Long
    Dim strArea As String
    Dim strErrMsg As String
    Dim oNamedItem As IJNamedItem
    Dim oComparedObjectName As String
    Dim oComparedObjectTypeName As String
    Dim oCheckingHoleName As String
    
    'Get the name of checking hole
    Set oNamedItem = oHoleTraceAE
    oCheckingHoleName = oNamedItem.Name
    
    'Get the name of Compared object
    Set oNamedItem = oComparedObject
    oComparedObjectName = oNamedItem.Name
    
    'Get the type name of Compared object
    oComparedObjectTypeName = GetObjectTypeName(oComparedObject)
           
    'Compare the distance and return onCheckError
    If dblMinDistance < dblRuleDistance Then
        strErrMsg = "The Hole (" & oCheckingHoleName & ")" & " " & _
                    "is too close with " & "the " & oComparedObjectTypeName & " (" & oComparedObjectName & ")." & vbCrLf
        If dblMinDistance > 0 Then
            strErrMsg = strErrMsg & "The Minimun distance is " & Format$(dblMinDistance, "#0.000") & "m."
        Else
            strErrMsg = strErrMsg & "They are interfered."
        End If
        oCallBack.OnCheckError oHoleTraceAE, strProgID, _
                              ESeverityIndex.siError, lngOptionCode, strErrMsg, "", ""
    End If

Cleanup:
    Set oNamedItem = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Sub

'******************************************************************************
' Routine: PassStringToCallBack
'
' Abstract: Pass the string for display the message on the list view.
'
' Description: Return the callback to display the error message on the list box
'******************************************************************************
Public Sub PassStringToCallBack(strProgID As String, _
                                oHoleTraceAE As IJHoleTraceAE, _
                                strErrMsg As String, _
                                oCallBack As IJCheckMfctyCallback, _
                                lngOptionCode As Long)
    Const METHOD As String = "PassStringToCallBack"
    On Error GoTo ErrorHandler
              
    oCallBack.OnCheckError oHoleTraceAE, strProgID, _
                           ESeverityIndex.siError, lngOptionCode, strErrMsg, "", ""

Cleanup:
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Sub

'******************************************************************************
' Routine: ConvertObjectToCurve
'
' Abstract: Convert object to curve
'
' Description:
'    Followings are the list of geometry type.
'''     GLineTYPE
'''     GLineStringTYPE
'''     GEllipseTYPE
'''     GArcTYPE
'''     GBspCurveTYPE
'''     GComplexStringTYPE
'''     GPlaneTYPE
'''     GConeTYPE
'''     GBspSurfaceTYPE
'''     GSphereTYPE
'''     GTorusTYPE
'''     GProjectionTYPE
'''     GRevolutionTYPE
'''     GRuledTYPE
'''     GPointTYPE
'''     GPolyMeshTYPE
'******************************************************************************
Public Function ConvertObjectToCurve(oObject As Object) As IJCurve
    Const METHOD As String = "ConvertObjectToCurve"
    On Error GoTo ErrorHandler

    If TypeOf oObject Is Line3d Then
        Dim oLine3d As Line3d
        Set oLine3d = oObject
        Set ConvertObjectToCurve = oLine3d
        Set oLine3d = Nothing
    ElseIf TypeOf oObject Is LineString3d Then
        Dim oLineString3d As LineString3d
        Set oLineString3d = oObject
        Set ConvertObjectToCurve = oLineString3d
        Set oLineString3d = Nothing
    ElseIf TypeOf oObject Is Arc3d Then
        Dim oArc3d As Arc3d
        Set oArc3d = oObject
        Set ConvertObjectToCurve = oArc3d
        Set oArc3d = Nothing
    ElseIf TypeOf oObject Is Circle3d Then
        Dim oCircle3d As Circle3d
        Set oCircle3d = oObject
        Set ConvertObjectToCurve = oCircle3d
        Set oCircle3d = Nothing
    ElseIf TypeOf oObject Is Ellipse3d Then
        Dim oEllipse3d As Ellipse3d
        Set oEllipse3d = oObject
        Set ConvertObjectToCurve = oEllipse3d
        Set oEllipse3d = Nothing
    ElseIf TypeOf oObject Is EllipticalArc3d Then
        Dim oEllipticalArc3d As EllipticalArc3d
        Set oEllipticalArc3d = oObject
        Set ConvertObjectToCurve = oEllipticalArc3d
        Set oEllipticalArc3d = Nothing
    ElseIf TypeOf oObject Is BSplineCurve3d Then
        Dim oBSplineCurve3d As BSplineCurve3d
        Set oBSplineCurve3d = oObject
        Set ConvertObjectToCurve = oBSplineCurve3d
        Set oBSplineCurve3d = Nothing
    Else
        Set ConvertObjectToCurve = Nothing
    End If
    
Cleanup:
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: CheckObjectIsCurve
'
' Abstract: Check if object is curve
'
' Description:
'******************************************************************************
Public Function CheckObjectIsCurve(oObject As Object) As Boolean
    Const METHOD As String = "ConvertObjectToCurve"
    On Error GoTo ErrorHandler

    CheckObjectIsCurve = False
    If TypeOf oObject Is Line3d Or TypeOf oObject Is LineString3d Or TypeOf oObject Is Arc3d Or _
       TypeOf oObject Is Circle3d Or TypeOf oObject Is Ellipse3d Or TypeOf oObject Is EllipticalArc3d Or _
       TypeOf oObject Is BSplineCurve3d Or TypeOf oObject Is Point3d Then
        CheckObjectIsCurve = True
    End If
    
Cleanup:
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetPlateSystemFromHoleTrace
'
' Abstract: Get the plate sytem from the parents of the holetrace
'
' Description:
'******************************************************************************
Public Function GetPlateSystemFromHoleTrace(oHoleTraceAE As IJHoleTraceAE) As Object
    Const METHOD = "GetPlateSystemFromHoleTrace"
    On Error GoTo ErrorHandler
    
    'the call to the hole trace no longer returns the port but the actual structure
    Set GetPlateSystemFromHoleTrace = oHoleTraceAE.GetParentStructure
    
Cleanup:
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: CanRetriveCurvesFromHole
'
' Abstract: Hole should be return curves, But, Wrong hole can't
'
' Description:
'******************************************************************************
Public Function CanRetriveCurvesFromHole(strProgID As String, _
                                         oHoleTraceAE As IJHoleTraceAE, _
                                         oCallBack As IJCheckMfctyCallback, _
                                         lngOptionCode As Long) As IMSCoreCollections.IJElements
    Const METHOD = "CanRetriveCurvesFromHole"
    On Error GoTo ErrorHandler
    
    Dim oCheckingCurves As IMSCoreCollections.IJElements
   
    Set oCheckingCurves = GetHoleTraceCurves(oHoleTraceAE)
    If oCheckingCurves Is Nothing Then
        Dim oNamedItem As IJNamedItem
        Dim oHoleName As String
        Dim oStrMsg As String
        
        'Get the name of checking hole
        Set oNamedItem = oHoleTraceAE
        oHoleName = oNamedItem.Name
        Set oNamedItem = Nothing
    
         oStrMsg = "The Hole (" & oHoleName & ")" & " " & _
                          "doesn't have curves" & " Check the hole."
         PassStringToCallBack strProgID, oHoleTraceAE, oStrMsg, oCallBack, lngOptionCode
    End If
    
    Set CanRetriveCurvesFromHole = oCheckingCurves
    
Cleanup:
    Set oNamedItem = Nothing
    Set oCheckingCurves = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function
'******************************************************************************
' Routine: IsSameParent
'
' Abstract: Check the structure parents is same.
'
' Description: The hole which is checked should be on the same structure plate
'******************************************************************************
Public Function IsSameParent(oFirstTrace As IJHoleTraceAE, oSecondTrace As IJHoleTraceAE) As Boolean
    Const METHOD = "IsSameParent"
    On Error GoTo ErrorHandler
    
    Dim oFirstParent As Object
    Dim oSecondParent As Object
    
    Set oFirstParent = oFirstTrace.GetParentStructure
    Set oSecondParent = oSecondTrace.GetParentStructure
    
    If oFirstParent Is oSecondParent Then
        IsSameParent = True
    Else
        IsSameParent = False
    End If
    
Cleanup:
    Set oFirstParent = Nothing
    Set oSecondParent = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetGeometryFromOpening
'
' Abstract: Get the Geometry from Opening
'
' Description: Using the middlerhelper, to get the geometry of Opening
'******************************************************************************
Public Function GetGeometryFromOpening(oUnknown As Object) As Object
    Const METHOD = "GetGeometryFromOpening"
    On Error GoTo ErrorHandler
    
    'make the middle tier helper
    Dim oMiddleHelper As IJHMMiddleHelper
    Set oMiddleHelper = New CHMMiddleHelper

    'try and get the geometry if it is a Opening
    Dim oOpening As IJDStructContourOccurrence

    If TypeOf oUnknown Is IJDStructContourOccurrence Then
        Set oOpening = oUnknown
        If Not oOpening Is Nothing Then
            Set GetGeometryFromOpening = oMiddleHelper.GetOpeningGeometry(oOpening)
            Set oOpening = Nothing
        End If
    End If

Cleanup:
    Set oMiddleHelper = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function


'******************************************************************************
' Routine: GetGeometryFromSeam
'
' Abstract: Get the Geometry from Seam
'
' Description: Using the middlerhelper, to get the geometry of Seam
'******************************************************************************
Public Function GetGeometryFromSeam(oUnknown As Object) As Object
    Const METHOD = "GetGeometryFromSeam"
    On Error GoTo ErrorHandler
    
    'make the middle tier helper
    Dim oMiddleHelper As IJHMMiddleHelper
    Set oMiddleHelper = New CHMMiddleHelper
    
    'try and get the geometry if it is a seam
    Dim oSeam As IJDSeamType
    If TypeOf oUnknown Is IJDSeamType Then
        Set oSeam = oUnknown
        If Not oSeam Is Nothing Then
            Set GetGeometryFromSeam = oMiddleHelper.GetSeamGeometry(oSeam)
            Set oSeam = Nothing
        End If
    End If
    
Cleanup:
    Set oMiddleHelper = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetGeometryFromProfile
'
' Abstract: Get the Geometry from Profile
'
' Description: Using the middlehelper, to get the outer geometry of Profile
'******************************************************************************
Public Function GetGeometryFromProfile(oUnknown As Object) As IMSCoreCollections.IJElements
    Const METHOD = "GetGeometryFromProfile"
    On Error GoTo ErrorHandler
        
    'make the middle tier helper
    Dim oMiddleHelper As IJHMMiddleHelper
    Set oMiddleHelper = New CHMMiddleHelper
    
    Dim oObjectCollection As IJDObjectCollection
    'try and get the geometry if it is a profile
    Dim oProfile As IJStiffener
    
    If TypeOf oUnknown Is IJStiffener Then
        Set oProfile = oUnknown
        If Not oProfile Is Nothing Then
            Set oObjectCollection = oMiddleHelper.GetProfileGeometry(oProfile)
            Set GetGeometryFromProfile = ConvertObjectCollectionToElement(oObjectCollection)
            Set oProfile = Nothing
        End If
    End If

Cleanup:
    Set oObjectCollection = Nothing
    Set oMiddleHelper = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetGeometryFromPlatePart
'
' Abstract: Get the Geometry from Seam
'
' Description: Using the middlehelper, to get the outer geometry of Plate Part
'******************************************************************************
Public Function GetGeometryFromPlatePart(oUnknown As Object) As IMSCoreCollections.IJElements
    Const METHOD = "GetGeometryFromPlatePart"
    On Error GoTo ErrorHandler
    
    Dim oObjectCollection As IJDObjectCollection
    'make the middle tier helper
    Dim oMiddleHelper As IJHMMiddleHelper
    Set oMiddleHelper = New CHMMiddleHelper
    
    'try and get the geometry if it is a Plate Part
    Dim oPlatePart As IJPlatePart
    If TypeOf oUnknown Is IJPlatePart Then
        Set oPlatePart = oUnknown
        If Not oPlatePart Is Nothing Then
            Set oObjectCollection = oMiddleHelper.GetPlatePartGeometry(oPlatePart)
            Set GetGeometryFromPlatePart = ConvertObjectCollectionToElement(oObjectCollection)
            Set oPlatePart = Nothing
        End If
    End If

Cleanup:
    Set oObjectCollection = Nothing
    Set oMiddleHelper = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetGeometryFromEquipment
'
' Abstract: Get the Geometry from Equipment
'
' Description: Using the middlehelper, to get the outer geometry of Equipment
'******************************************************************************
Public Function GetGeometryFromEquipment(oUnknown As Object) As IMSCoreCollections.IJElements
    Const METHOD = "GetGeometryFromEquipment"
    On Error GoTo ErrorHandler
    
    Dim oObjectCollection As IJDObjectCollection
    'make the middle tier helper
    Dim oMiddleHelper As IJHMMiddleHelper
    Set oMiddleHelper = New CHMMiddleHelper
    
    'try and get the geometry if it is a Equipment
    Dim oEquipment As IJEquipment
    If TypeOf oUnknown Is IJEquipment Then
        Set oEquipment = oUnknown
        If Not oEquipment Is Nothing Then
            Set oObjectCollection = oMiddleHelper.GetEquipmentGeometry(oEquipment)
            Set GetGeometryFromEquipment = ConvertObjectCollectionToElement(oObjectCollection)
            Set oEquipment = Nothing
        End If
    End If

Cleanup:
    Set oObjectCollection = Nothing
    Set oMiddleHelper = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetGeometryFromCoaming
'
' Abstract: Get the Geometry from Coamings
'
' Description: Using the middlerhelper, to get the geometry of Seam
'******************************************************************************
Public Function GetGeometryFromCoaming(oUnknown As Object) As IMSCoreCollections.IJElements
    Const METHOD = "GetGeometryFromCoaming"
    On Error GoTo ErrorHandler
    
    Dim oObjectCollection As IJDObjectCollection
    'make the middle tier helper
    Dim oMiddleHelper As IJHMMiddleHelper
    Set oMiddleHelper = New CHMMiddleHelper
    
    'try and get the geometry if it is a Coaming what can be made by Profile/PlatePart/Equipment
    Dim oProfile As IJStiffenerPart
    Dim oPlatePart As IJPlatePart
    Dim oEquipment As IJEquipment
    
    If TypeOf oUnknown Is IJStiffenerPart Then
        Set oProfile = oUnknown
        If Not oProfile Is Nothing Then
            Set oObjectCollection = oMiddleHelper.GetProfileGeometry(oProfile)
            Set GetGeometryFromCoaming = ConvertObjectCollectionToElement(oObjectCollection)
            Set oProfile = Nothing
        End If
    ElseIf TypeOf oUnknown Is IJPlatePart Then
        Set oPlatePart = oUnknown
        If Not oPlatePart Is Nothing Then
            Set oObjectCollection = oMiddleHelper.GetPlatePartGeometry(oPlatePart)
            Set GetGeometryFromCoaming = ConvertObjectCollectionToElement(oObjectCollection)
            Set oPlatePart = Nothing
        End If
    ElseIf TypeOf oUnknown Is IJEquipment Then
        Set oEquipment = oUnknown
        If Not oEquipment Is Nothing Then
            Set oObjectCollection = oMiddleHelper.GetEquipmentGeometry(oEquipment)
            Set GetGeometryFromCoaming = ConvertObjectCollectionToElement(oObjectCollection)
            Set oEquipment = Nothing
        End If
    End If
    
Cleanup:
    If Not oProfile Is Nothing Then Set oProfile = Nothing
    If Not oPlatePart Is Nothing Then Set oPlatePart = Nothing
    If Not oEquipment Is Nothing Then Set oEquipment = Nothing
    Set oObjectCollection = Nothing
    Set oMiddleHelper = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: ConvertObjectCollectionToElement
'
' Abstract: Convert ObjectCollection To Element
'
' Description: Convert ObjectCollection to Element
'*****************************************************************************
Public Function ConvertObjectCollectionToElement(oIJDObjectCollection As IJDObjectCollection) As IMSCoreCollections.IJElements
    Const METHOD = "ConvertObjectCollectionToElement"
    On Error GoTo ErrorHandler
    
    Dim oObject As Object
    Dim oElem As IJDObjectCollection
    Set oElem = New IMSCoreCollections.JObjectCollection
    For Each oObject In oIJDObjectCollection
        oElem.Add oObject
    Next
    
    Set ConvertObjectCollectionToElement = oElem

Cleanup:
    Set oElem = Nothing
    Set oObject = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetObjectTypeName
'
' Abstract: Get the Object type Name
'
' Description: Get the object type name
'*****************************************************************************
Private Function GetObjectTypeName(oObject As Object) As String
    Const METHOD = "ConvertObjectCollectionToElement"
    On Error GoTo ErrorHandler
    
    If TypeOf oObject Is IJPlateSystem Then
        GetObjectTypeName = "Plate System"
    ElseIf TypeOf oObject Is IJPlatePart Then
        GetObjectTypeName = "Plate Part"
    ElseIf TypeOf oObject Is IJSeam Then
        GetObjectTypeName = "Seam Line"
    ElseIf TypeOf oObject Is IJStiffenerSystem Then
        GetObjectTypeName = "Profile System"
    ElseIf TypeOf oObject Is IJStiffenerPart Then
        GetObjectTypeName = "Profile Part"
    ElseIf TypeOf oObject Is IJDStructContourOccurrence Then
        GetObjectTypeName = "Opening"
    ElseIf TypeOf oObject Is IJEquipment Then
        GetObjectTypeName = "Coamings"
    Else
        GetObjectTypeName = "Unknown"
    End If
    
Cleanup:
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: ConvertGTypeToCurves
'
' Abstract: Convert GType to Curves
'
' Description: If the object dosen't have IJCurve, try to get the curves
' Plane, surface has a method named "GetBoundaries". This returns the curves.
'*****************************************************************************
Private Function ConvertGTypeToCurves(oObject As Object) As IMSCoreCollections.IJElements
    Const METHOD = "ConvertGTypeToCurves"
    On Error GoTo ErrorHandler
    
    Dim oElemCurves As IMSCoreCollections.IJElements
    Set oElemCurves = New IMSCoreCollections.JObjectCollection
    Dim oCurve As IJCurve
            
    If TypeOf oObject Is ComplexString3d Then
        Dim oComplexString3d As ComplexString3d
        Set oComplexString3d = oObject
        oComplexString3d.GetCurves oElemCurves
        Set oComplexString3d = Nothing

    ElseIf TypeOf oObject Is Plane3d Then
        Dim oPlane3D As Plane3d
        Set oPlane3D = oObject
        oPlane3D.GetBoundaries oElemCurves
        Set oPlane3D = Nothing
        
    ElseIf TypeOf oObject Is Cone3d Then
        Dim oCone3D As Cone3d
        Set oCone3D = oObject
        oCone3D.GetBoundaries oElemCurves
        Set oCone3D = Nothing
        
    ElseIf TypeOf oObject Is BSplineSurface3d Then
        Dim oBSplineSurface3d As BSplineSurface3d
        Set oBSplineSurface3d = oObject
        oBSplineSurface3d.GetBoundaries oElemCurves
        Set oBSplineSurface3d = Nothing
        
    ElseIf TypeOf oObject Is Sphere3d Then
        Dim oSphere3d As Sphere3d
        Set oSphere3d = oObject
        oSphere3d.GetBoundaries oElemCurves
        Set oSphere3d = Nothing
        
    ElseIf TypeOf oObject Is Torus3d Then
        Dim oTorus3d As Torus3d
        Set oTorus3d = oObject
        oTorus3d.GetBoundaries oElemCurves
        Set oTorus3d = Nothing

    ElseIf TypeOf oObject Is Projection3d Then
        Dim oProjection3d As Projection3d
        Set oProjection3d = oObject
        Set oCurve = oProjection3d.Curve
        oElemCurves.Add oCurve
        Set oCurve = Nothing
        Set oProjection3d = Nothing
        
    ElseIf TypeOf oObject Is Revolution3d Then
        Dim oRevolution3d As Revolution3d
        Set oRevolution3d = oObject
        Set oCurve = oRevolution3d.Curve
        oElemCurves.Add oCurve
        Set oCurve = Nothing
        Set oRevolution3d = Nothing
        
    ElseIf TypeOf oObject Is RuledSurface3d Then
        Dim oRuledSurface3d As RuledSurface3d
        Set oRuledSurface3d = oObject
        Set oCurve = oRuledSurface3d.GenCurveBase
        oElemCurves.Add oCurve
        Set oCurve = oRuledSurface3d.GenCurveTop
        oElemCurves.Add oCurve
        Set oCurve = Nothing
        Set oRuledSurface3d = Nothing
    End If
    
    Set ConvertGTypeToCurves = oElemCurves
    
Cleanup:
    Set oElemCurves = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: CheckSameObject
'
' Abstract: Compare two IJCurves and return trun if they are same.
'
' Description: Hole on the plate part has a problem. Hole and cut has a same geometry
'              So, Hole needs to be check against cut and skip to check minmun distance.
'
'
'*****************************************************************************
Public Function CheckSameObject(oComparingCurve As IJCurve, oComparedCurve As IJCurve, _
                                oThickness As Double) As Boolean
    Const METHOD = "CheckSameObject"
    On Error GoTo ErrorHandler
    
    CheckSameObject = False
        
    Dim Tolerance As Double
    'Set curve from HoleTrace
    Dim oComparingLine As IJLine, oComparingArc As IJArc
    Dim oComparingCircle As IJCircle, oComparingEllipse As IJEllipse
    
    'Set the curve from Plate part.
    Dim oComparedLine As IJLine, oComparedArc As IJArc
    Dim oComparedCircle As IJCircle, oComparedEllipse As IJEllipse
    
    'Set the start and end point of HoleTrace to compare if they are same in case of line and arc
    Dim oComparingStartX As Double, oComparingStartY As Double, oComparingStartZ As Double
    Dim oComparingEndX As Double, oComparingEndY As Double, oComparingEndZ As Double
    
    'Set the start and end point of Plate part to compare if they are same in case of line and arc
    Dim oComparedStartX As Double, oComparedStartY As Double, oComparedStartZ As Double
    Dim oComparedEndX As Double, oComparedEndY As Double, oComparedEndZ As Double
    
    'In case of Arc or circle, the center point should be needed.
    Dim oComparingCenterX As Double, oComparingCenterY As Double, oComparingCenterZ As Double
    Dim oComparedCenterX As Double, oComparedCenterY As Double, oComparedCenterZ As Double
    
    'In case of circle, the radius should be needed.
    Dim oComparingCircleRadius As Double, oComparedCircleRadius As Double
    Dim oComparingEllipsMaxRadius As Double, oComparingEllipsMinRadius As Double, oComparingEllipsMinMajorRadius As Double
    Dim oComparedEllipsMaxRadius As Double, oComparedEllipsMinRadius As Double, oComparedEllipsMinMajorRadius As Double
    
    'Tolerance should be larger then thickness of plate part.
    Tolerance = oThickness + 0.001

    'Comaring between line from hole and line from part.
    'Comaring between line from hole and arc from part.
    If TypeOf oComparingCurve Is IJLine Then
        Set oComparingLine = oComparingCurve
        
        'Get the start and end point from HoleTrace.
        oComparingLine.GetStartPoint oComparingStartX, oComparingStartY, oComparingStartZ
        oComparingLine.GetEndPoint oComparingEndX, oComparingEndY, oComparingEndZ
        
        'If the part has line, compare to start and end position from line of hole
        If TypeOf oComparedCurve Is IJLine Then
            
            'Get the line from part
            Set oComparedLine = oComparedCurve
                
                'Get the start and end point from part.
                oComparedLine.GetStartPoint oComparedStartX, oComparedStartY, oComparedStartZ
                oComparedLine.GetEndPoint oComparedEndX, oComparedEndY, oComparedEndZ
                
                'Check the point if they are same.
                If ((oComparingStartX - oComparedStartX) < Tolerance _
                     And (oComparingStartY - oComparedStartY) < Tolerance _
                     And (oComparingStartZ - oComparedStartZ) < Tolerance _
                     And (oComparingEndX - oComparedEndX) < Tolerance _
                     And (oComparingEndY - oComparedEndY) < Tolerance _
                     And (oComparingEndZ - oComparedEndZ) < Tolerance) _
                     Or ((oComparingStartX - oComparedEndX) < Tolerance _
                     And (oComparingStartY - oComparedEndY) < Tolerance _
                     And (oComparingStartZ - oComparedEndZ) < Tolerance _
                     And (oComparingEndX - oComparedStartX) < Tolerance _
                     And (oComparingEndY - oComparedStartY) < Tolerance _
                     And (oComparingEndZ - oComparedStartZ) < Tolerance) Then
                     'Return true if they share the same point among them.
                     CheckSameObject = True
                End If
                
        'If the part has arc, compare to start and end position from line of hole
        ElseIf TypeOf oComparedCurve Is IJArc Then
                
                'Get the arc from part
                Set oComparedArc = oComparedCurve
                
                'Get the start and end point from part.
                oComparedArc.GetStartPoint oComparedStartX, oComparedStartY, oComparedStartZ
                oComparedArc.GetEndPoint oComparedEndX, oComparedEndY, oComparedEndZ
                
                'Check the point if they are same.
                If ((oComparingStartX - oComparedStartX) < Tolerance _
                     And (oComparingStartY - oComparedStartY) < Tolerance _
                     And (oComparingStartZ - oComparedStartZ) < Tolerance) _
                     Or ((oComparingEndX - oComparedEndX) < Tolerance _
                     And (oComparingEndY - oComparedEndY) < Tolerance _
                     And (oComparingEndZ - oComparedEndZ) < Tolerance) _
                     Or ((oComparingStartX - oComparedEndX) < Tolerance _
                     And (oComparingStartY - oComparedEndY) < Tolerance _
                     And (oComparingStartZ - oComparedEndZ) < Tolerance) _
                     Or ((oComparingEndX - oComparedStartX) < Tolerance _
                     And (oComparingEndY - oComparedStartY) < Tolerance _
                     And (oComparingEndZ - oComparedStartZ) < Tolerance) Then
                     'Return true if they share the same point among them.
                     CheckSameObject = True
                End If
        End If
    
    'Comaring between arc from hole and line from part.
    'Comaring between arc from hole and arc from part.
    ElseIf TypeOf oComparingCurve Is IJArc Then
        'Get the arc from hole
        Set oComparingArc = oComparingCurve
        
        'Get the start , end and center point from HoleTrace.
        oComparingArc.GetStartPoint oComparingStartX, oComparingStartY, oComparingStartZ
        oComparingArc.GetEndPoint oComparingEndX, oComparingEndY, oComparingEndZ
        oComparingArc.GetCenterPoint oComparingCenterX, oComparingCenterY, oComparingCenterZ
        
        'If the part has arc, compare to start and end position from arc of hole
        If TypeOf oComparedCurve Is IJArc Then
            Set oComparedArc = oComparedCurve
                oComparedArc.GetStartPoint oComparedStartX, oComparedStartY, oComparedStartZ
                oComparedArc.GetEndPoint oComparedEndX, oComparedEndY, oComparedEndZ
                oComparedArc.GetCenterPoint oComparedCenterX, oComparedCenterY, oComparedCenterZ
                
                If Abs(oComparingCenterX - oComparedCenterX) < Tolerance _
                     And Abs(oComparingCenterY - oComparedCenterY) < Tolerance _
                     And Abs(oComparingCenterZ - oComparedCenterZ) < Tolerance Then
                        
                        If (Abs(oComparingStartX - oComparedStartX) < Tolerance _
                        And Abs(oComparingStartY - oComparedStartY) < Tolerance _
                        And Abs(oComparingStartZ - oComparedStartZ) < Tolerance _
                        And Abs(oComparingEndX - oComparedEndX) < Tolerance _
                        And Abs(oComparingEndY - oComparedEndY) < Tolerance _
                        And Abs(oComparingEndZ - oComparedEndZ) < Tolerance) _
                        Or (Abs(oComparingStartX - oComparedEndX) < Tolerance _
                        And Abs(oComparingStartY - oComparedEndY) < Tolerance _
                        And Abs(oComparingStartZ - oComparedEndZ) < Tolerance _
                        And Abs(oComparingEndX - oComparedStartX) < Tolerance _
                        And Abs(oComparingEndY - oComparedStartY) < Tolerance _
                        And Abs(oComparingEndZ - oComparedStartZ) < Tolerance) Then
                            'Return true if they share the same point among them.
                            CheckSameObject = True
                        End If
                End If
        'If the part has line, compare to start and end position from arc of hole
        ElseIf TypeOf oComparedCurve Is IJLine Then
                'Get the circle from hole
                Set oComparedLine = oComparedCurve
                oComparedLine.GetStartPoint oComparedStartX, oComparedStartY, oComparedStartZ
                oComparedLine.GetEndPoint oComparedEndX, oComparedEndY, oComparedEndZ

                If ((oComparingStartX - oComparedStartX) < Tolerance _
                     And (oComparingStartY - oComparedStartY) < Tolerance _
                     And (oComparingStartZ - oComparedStartZ) < Tolerance) _
                     Or ((oComparingEndX - oComparedEndX) < Tolerance _
                     And (oComparingEndY - oComparedEndY) < Tolerance _
                     And (oComparingEndZ - oComparedEndZ) < Tolerance) _
                     Or ((oComparingStartX - oComparedEndX) < Tolerance _
                     And (oComparingStartY - oComparedEndY) < Tolerance _
                     And (oComparingStartZ - oComparedEndZ) < Tolerance) _
                     Or ((oComparingEndX - oComparedStartX) < Tolerance _
                     And (oComparingEndY - oComparedStartY) < Tolerance _
                     And (oComparingEndZ - oComparedStartZ) < Tolerance) Then
                     'Return true if they share the same point among them.
                     CheckSameObject = True
                End If
        End If
    
    'Comaring between circle from hole and circle from part.
    ElseIf TypeOf oComparingCurve Is IJCircle Then
        'Get the circle from hole
        Set oComparingCircle = oComparingCurve
        
        'Get center point and radius from HoleTrace.
        oComparingCircle.GetCenterPoint oComparingCenterX, oComparingCenterY, oComparingCenterZ
        oComparingCircleRadius = oComparingCircle.Radius
    
            If TypeOf oComparedCurve Is IJCircle Then
                Set oComparedCircle = oComparingCurve
                oComparedCircle.GetCenterPoint oComparedCenterX, oComparedCenterY, oComparedCenterZ
                oComparedCircleRadius = oComparedCircle.Radius
                
                If (oComparingCenterX - oComparedCenterX) < Tolerance _
                     And (oComparingCenterY - oComparedCenterY) < Tolerance _
                     And (oComparingCenterZ - oComparedCenterZ) < Tolerance Then
                     
                     If (oComparingCircleRadius - oComparedCircleRadius) < Tolerance Then
                        'Return true if they share the same point among them.
                        CheckSameObject = True
                     End If
                End If
             End If
    
    'Comaring between IJEllipse from hole and IJEllipse from part.
    ElseIf TypeOf oComparingCurve Is IJEllipse Then
        'Get the ellipse from hole
        Set oComparingEllipse = oComparingCurve
        
        'Get center point and radius from HoleTrace.
        oComparingEllipse.GetCenterPoint oComparingCenterX, oComparingCenterY, oComparingCenterZ
        oComparingEllipsMaxRadius = oComparingEllipse.MajorRadius
        oComparingEllipsMinRadius = oComparingEllipse.MinorRadius
        oComparingEllipsMinMajorRadius = oComparingEllipse.MinorMajorRatio
    
            If TypeOf oComparedCurve Is IJEllipse Then
                Set oComparedEllipse = oComparedCurve
                oComparedEllipse.GetCenterPoint oComparedCenterX, oComparedCenterY, oComparedCenterZ
                oComparedEllipsMaxRadius = oComparedEllipse.MajorRadius
                oComparedEllipsMinRadius = oComparedEllipse.MinorRadius
                oComparedEllipsMinMajorRadius = oComparedEllipse.MinorMajorRatio
                
                If Abs(oComparingCenterX - oComparedCenterX) < Tolerance _
                   And Abs(oComparingCenterY - oComparedCenterY) < Tolerance _
                   And Abs(oComparingCenterZ - oComparedCenterZ) < Tolerance Then
                     
                     If Abs(oComparingEllipsMaxRadius - oComparedEllipsMaxRadius) < Tolerance _
                        And Abs(oComparingEllipsMinRadius - oComparedEllipsMinRadius) < Tolerance _
                        And Abs(oComparingEllipsMinMajorRadius - oComparedEllipsMinMajorRadius) < Tolerance Then
                            'Return true if they share the same point among them.
                            CheckSameObject = True
                     End If
                End If
             End If
    End If
    
Cleanup:
    Set oComparingLine = Nothing
    Set oComparingArc = Nothing
    Set oComparingCircle = Nothing
    Set oComparingEllipse = Nothing
    
    Set oComparedLine = Nothing
    Set oComparedArc = Nothing
    Set oComparedCircle = Nothing
    Set oComparedEllipse = Nothing

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: CheckHoleAndCut
'
' Abstract: Compare IJCurve and Complex strings return trun if one curve is same.
'
' Description: Hole on the plate part has a problem. Hole and cut has a same geometry
'              So, Hole needs to be check against cut and skip to check minmun distance.
'
'
'*****************************************************************************
Public Function CheckHoleAndCut(oComparingCurve As IJCurve, oComplexStrings As IMSCoreCollections.IJElements, _
                                oThickness As Double) As Boolean
    Const METHOD = "CheckSameObject"
    On Error GoTo ErrorHandler
    
    CheckHoleAndCut = False

    Dim oCurve As IJCurve
    For Each oCurve In oComplexStrings
        CheckHoleAndCut = CheckSameObject(oComparingCurve, oCurve, oThickness)
        If CheckHoleAndCut = True Then Exit For
    Next
    
Cleanup:
    Set oCurve = Nothing
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, "").Number
    GoTo Cleanup
End Function
 
