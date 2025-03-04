Attribute VB_Name = "PrProHelpers"
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    PrProHelpers for the MfgProfileMarking rules.
'
'  History:
'       TBH         april 8. 2002   created
'       MJV         2004.04.23      included correct error handling
'******************************************************************

Option Explicit

Const MODULE = "PrProHelpers.bas"

Public Enum CurvatureType
    Straight = 0
    InCurvature = 1
    OutCurvature = 2
    SCurved = 3
End Enum

Public Const UPSIDE_XML_LOCATION = "\StructManufacturing\SMS_SCHEMA\SMS_OUTPUT\OutputControlXML.xml"

Public Function GetRefDataAttribute(Part As IJProfilePart, itemName As String) As Variant
Const METHOD = "GetRefDataAttribute"
On Error GoTo ErrorHandler

    Dim oProfilePartSupport As IJProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Set oProfilePartSupport = New GSCADSDPartSupport.ProfilePartSupport
    Set oPartSupport = oProfilePartSupport
    Set oPartSupport.Part = Part
    
    Dim oCrossSection As IJCrossSection
    oProfilePartSupport.GetCrossSection oCrossSection
    
    Dim oAttCol As IJDAttributesCol
    Set oAttCol = oCrossSection.Attributes
    
    Dim oAttribute As IJDAttribute
    Set oAttribute = oAttCol.Item(itemName)
    
    GetRefDataAttribute = oAttribute.Value
   
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

' ***********************************************************************************
' Public Function CurveLength
'
' Description:  Helper function to get length of a curve. Input is a Wirebody and
'               output is a double which is the length of the curve.
'
' ***********************************************************************************
Public Function CurveLength(ByVal oWB As IJWireBody) As Double
Const METHOD = "CurveLength"
    On Error GoTo ErrorHandler
    Dim oCurve As IJCurve
    'Dim oHelper As New MfgRuleHelper
    
    'Waiting for cda to finish the helper
    'Set oCurve = oHelper.WireBodyToComplexString(oWB)
    
    CurveLength = oCurve.length
    Exit Function
ErrorHandler:
    CurveLength = 0#
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Function



 
Public Function CheckCurvature(pProfilePart As Object) As CurvatureType
    Const METHOD = "CheckCurvature"
    On Error GoTo ErrorHandler
    
    Dim oStartPos               As IJDPosition
    Dim oEndPos                 As IJDPosition
    Dim oStraightLineMidPos     As IJDPosition
    Dim oWebCenterPos           As IJDPosition
    Dim oLandingCurveMidPos     As IJDPosition
    Dim oStraightVector         As IJDVector
    Dim oLandingVector          As IJDVector
    Dim oNormal                 As IJDVector
    Dim oLandingCurve           As IJWireBody
    Dim oTopologyLocate         As New TopologyLocate
    Dim oHelper                 As New MfgRuleHelpers.Helper
    Dim dDot                    As Double
    Dim oWebLeftPort            As IJPort
    Dim oSDOHelper As StructDetailObjects.ProfilePart
    Dim oStraightLine           As IJLine
    Dim oMfgMGHelper            As IJMfgMGHelper
    Dim oCS                     As IJComplexString
    Dim oCurve                  As IJCurve
    
    Set oSDOHelper = New StructDetailObjects.ProfilePart
    Set oSDOHelper.object = pProfilePart
    
    Set oWebLeftPort = oSDOHelper.SubPort(JXSEC_WEB_LEFT)
    
    Set oLandingCurve = oTopologyLocate.GetProfileParentWireBody(pProfilePart)

    oLandingCurve.GetEndPoints oStartPos, oEndPos
    'circular edge reinforcements
    Set oStraightVector = oStartPos.Subtract(oEndPos)
    If oStraightVector.length < 0.0001 Then
        CheckCurvature = SCurved
        GoTo CleanUp:
    Else
        Set oStraightVector = Nothing
    End If
    
    Set oLandingCurveMidPos = oHelper.GetMiddlePoint(oLandingCurve)
    
    Set oStraightLineMidPos = New DPosition
    oStraightLineMidPos.Set (oStartPos.x + oEndPos.x) / 2, (oStartPos.y + oEndPos.y) / 2, (oStartPos.z + oEndPos.z) / 2
    
    Set oStraightLine = New Line3d
    oStraightLine.DefineBy2Points oStartPos.x, oStartPos.y, oStartPos.z, oEndPos.x, oEndPos.y, oEndPos.z
    
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    oMfgMGHelper.WireBodyToComplexString oLandingCurve, oCS
    Set oCurve = oCS
        
    Set oStraightVector = oStraightLineMidPos.Subtract(oLandingCurveMidPos)
    
    If oStraightVector.length < 0.0001 Then
        CheckCurvature = Straight
    Else
        oTopologyLocate.FindApproxCenterAndNormal oWebLeftPort.Geometry, oWebCenterPos, oNormal
        Set oLandingVector = oWebCenterPos.Subtract(oLandingCurveMidPos)
         
        dDot = oStraightVector.Dot(oLandingVector)
    
        If dDot < 0 Then
            CheckCurvature = OutCurvature
        Else
            CheckCurvature = InCurvature
        End If
    End If
    
    Dim nIntersects As Long
    Dim Points() As Double
    Dim nOverlaps As Long
    Dim eCode As Geom3dIntersectConstants
    
    oCurve.Intersect oStraightLine, nIntersects, Points, nOverlaps, eCode
    If nIntersects > 2 Then ' other than end points
        CheckCurvature = SCurved
    End If
    
    
CleanUp:
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Set oStraightLineMidPos = Nothing
    Set oWebCenterPos = Nothing
    Set oLandingCurveMidPos = Nothing
    Set oStraightVector = Nothing
    Set oLandingVector = Nothing
    Set oNormal = Nothing
    Set oLandingCurve = Nothing
    Set oWebLeftPort = Nothing
    Set oTopologyLocate = Nothing
    Set oHelper = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
    GoTo CleanUp
End Function


Public Function GetMarginValueOnThePort(oPort As IJPort) As Double

    Const METHOD = "GetMarginValueOnThePort"
    On Error GoTo ErrorHandler
    
    Dim oMfgDefCol As Collection
    
    Dim oFabMargin As IJDFabMargin
    Dim oMFGRuleHelper As MfgRuleHelpers.Helper
    
    Set oMFGRuleHelper = New MfgRuleHelpers.Helper
    
    Set oMfgDefCol = oMFGRuleHelper.GetMfgDefinitions(oPort)
    
    If oMfgDefCol.Count > 0 Then
        Dim lFabMargin As Double, lAssyMargin As Double, lCustomMargin As Double
        lFabMargin = 0
        lAssyMargin = 0
        lCustomMargin = 0
        Dim j As Integer
        For j = 1 To oMfgDefCol.Count
            If TypeOf oMfgDefCol.Item(j) Is IJDFabMargin Then
                Dim dStartValue As Double
                Dim dEndValue As Double
            
                Set oFabMargin = oMfgDefCol.Item(j)
                oFabMargin.MarginValues dStartValue, dEndValue
                
                If oFabMargin.MarginMode = AssemblyMargin Then
                    lAssyMargin = lAssyMargin + dStartValue
                ElseIf oFabMargin.MarginMode = ObliqueMargin Then
                    If dEndValue > dStartValue Then
                        lFabMargin = lFabMargin + dEndValue
                    Else
                        lFabMargin = lFabMargin + dStartValue
                    End If
                ElseIf oFabMargin.MarginMode = ConstMargin Then
                    lFabMargin = lFabMargin + dStartValue
                End If
            'ElseIf TypeOf oMfgDefCol.Item(j) Is ??? Then
            End If
            
        Next j
        If lAssyMargin <> 0 Or lFabMargin <> 0 Or lCustomMargin <> 0 Then
             Dim TotMargin As Double
             TotMargin = lAssyMargin + lFabMargin + lCustomMargin
        End If
    End If
    GetMarginValueOnThePort = TotMargin
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number

End Function

''**************************************************************************************
'' Routine      : GetSymbolSharePath
'' Abstract     : This method will return the SymbolShare Path ( like M:\CatalogData\Symbols).
''                Assumption: There will be Structmanufacturing folder under symbols directory. If doesn't exist
''                Get the LogicalSymbolPath registry key (This will be applicable for developer build)
''**************************************************************************************
Public Function GetSymbolSharePath() As String
    Const METHOD = "GetSymbolLocationPath"
    On Error GoTo ErrorHandler
    
    Dim oContext As IJContext
    Dim strContextString As String
    Dim strSymbolShare As String
    
    strContextString = "OLE_SERVER"
    
    'Get IJContext
    Set oContext = GetJContext()
    
    If Not oContext Is Nothing Then
        'Get the Symbol Share
        strSymbolShare = oContext.GetVariable(strContextString)
        GetSymbolSharePath = strSymbolShare
    End If
    
    Set oContext = Nothing
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

Public Function SetBuiltUpNames(pProfilePart As Object, bMfgAsPlate As Boolean)
    Const METHOD = "SetBuiltUpNames"
    On Error GoTo ErrorHandler
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = pProfilePart
    Dim pMfgProfilePart As IJMfgProfilePart
    Dim oProfileOutput As IJMfgProfileOutput
    
    If oProfileWrapper.ProfileHasMfgPart(pMfgProfilePart) Then
        
    Else
        Exit Function
    End If
    Set oProfileOutput = pMfgProfilePart
    Dim lCells As Long
    lCells = 0
    Dim strProgID As String
    Dim strCellNames() As Variant
    oProfileOutput.GetOutputKeys lCells, strCellNames
    Dim oNamedItem As IJNamedItem
    Set oNamedItem = pMfgProfilePart
    Dim strMfgPartName As String
    strMfgPartName = oNamedItem.Name
    Dim lCount As Long
    Dim oRootElement As IXMLDOMElement
    Dim oXMLDoc As DOMDocument
    Set oXMLDoc = New DOMDocument
    Set oRootElement = oXMLDoc.createElement("NAMES")
    oXMLDoc.appendChild oRootElement
    If lCells > 0 Then
        For lCount = LBound(strCellNames) To UBound(strCellNames)
            If lCount = 0 Then
                oRootElement.setAttribute strCellNames(lCount), strMfgPartName & "-V"
            Else
                oRootElement.setAttribute strCellNames(lCount), strMfgPartName & "-1"
            End If
        Next lCount
    Else
        oProfileOutput.GetCellKeys "", lCells, strCellNames
        If lCells > 0 Then
            For lCount = LBound(strCellNames) To UBound(strCellNames)
                If InStr(strCellNames(lCount), "WEB") > 0 Then
                    oRootElement.setAttribute strCellNames(lCount), strMfgPartName & "-V"
                Else
                    oRootElement.setAttribute strCellNames(lCount), strMfgPartName & "-1"
                End If
            Next lCount
        Else
            If bMfgAsPlate Then
                For lCount = 0 To 2
                    If lCount = 0 Then
                        oRootElement.setAttribute "WebNestData", strMfgPartName & "W"
                    ElseIf lCount = 1 Then
                        oRootElement.setAttribute "TopFlangeNestData", strMfgPartName & "F"
                    ElseIf lCount = 2 Then
                        oRootElement.setAttribute "BottomFlangeNestData", strMfgPartName & "F2"
                    End If
                Next lCount
            End If
        End If
    End If
    pMfgProfilePart.AdditionalNames = oXMLDoc
    
    Set oNamedItem = Nothing
    Set pMfgProfilePart = Nothing

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function
