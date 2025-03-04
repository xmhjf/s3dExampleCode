VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProfileReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
''**************************************************************************************
''  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
''
''  Project     : MfgProfileCustomReports
''  File        : ProfileReport.cls
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

Private Const MODULE = "MfgProfileCustomReports.ProfileReport"
Private Const IJMFGPROFILEPARTGUID = "{1BEB9DD4-3B5D-4571-AEFA-4DC8B9C21434}"
Private Const ROUND_CONST = 1
Private Const RADIANS_TO_DEGREES = 57.2957795130823


Private Sub IJDCustomReport_Generate(ByVal pElements As GSCADStructMfgGlobals.IJElements, strFileName As String, eCustomReportStatus As GSCADStructMfgGlobals.CustomReportStatus)
    Const METHOD = "IJDCustomReport_Generate"
    On Error GoTo ErrorHandler
       
    If pElements.Count > 0 Then
        Dim oSelectedObj As Object
        For Each oSelectedObj In pElements
            If Not TypeOf oSelectedObj Is IJMfgProfilePart Then
                'This is profile part. Get MFG part from profile
                Dim oStructMfgGlobals As New GSCADStructMfgGlobals.StructMfgGlobalsQuery
                Dim oColl As IJElements
                Set oColl = oStructMfgGlobals.GetMfgPart(oSelectedObj, IJMFGPROFILEPARTGUID)
                
                'For each Profile, report the information
                Dim oMfgProfilePart As Object
                For Each oMfgProfilePart In oColl
                    ReportProfilePartInformation oMfgProfilePart, strFileName
                Next
            Else
                ReportProfilePartInformation oSelectedObj, strFileName
            End If
        Next
        eCustomReportStatus = StrMfgProcessFinished
    End If
    
Wrapup:
    Exit Sub
    
ErrorHandler:
    eCustomReportStatus = StrMfgErrorUnknown
    Err.Raise Err.Number, , Err.Description
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints
' Purpose   : Given SMS_EDGE node with geometry, return new SMS_EDGE node with Points
'---------------------------------------------------------------------------------------
'
Private Function ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oInputEdgeNode As IXMLDOMNode, _
                                                               eStrokeMethod As eMfgStrokeCriterion, _
                                                               lenToUse As Double) As IXMLDOMNode
    Const METHOD = "ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints"
    On Error GoTo ErrorHandler

    On Error GoTo ErrorHandler

    Dim oReturnEdgeNode As IXMLDOMNode
    Set oReturnEdgeNode = oInputEdgeNode.cloneNode(True)

    Set ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints = oReturnEdgeNode

    Dim oMfgXMLhelper As New MfgXMLHelper

    Dim oPointColl As IJElements
    Set oPointColl = oMfgXMLhelper.StrokeEdgeNodeIntoPoints(oInputEdgeNode, _
                                                            eStrokeMethod, _
                                                            lenToUse)

    If oPointColl Is Nothing Then Exit Function
    If oPointColl.Count = 0 Then Exit Function

    Dim oChildCVGnode As IXMLDOMNode
    For Each oChildCVGnode In oReturnEdgeNode.childNodes
        If oChildCVGnode.baseName = "CVG_CURVE" Then
            oReturnEdgeNode.removeChild oChildCVGnode
        End If
    Next

    Dim oCreator As IXMLDOMDocument
    Set oCreator = oInputEdgeNode.ownerDocument

    Dim oNewElem As IXMLDOMElement
    Set oNewElem = oCreator.createElement("CVG_CURVE")

    Set oChildCVGnode = oReturnEdgeNode.appendChild(oNewElem)

    Dim i As Long
    For i = 1 To oPointColl.Count
        Dim oPos As IJDPosition
        Set oPos = oPointColl.Item(i)

        Set oNewElem = oCreator.createElement("CVG_VERTEX")
        If i = 1 Then
            oNewElem.setAttribute "POINT_CODE", "s_point"
        ElseIf i = oPointColl.Count Then
            oNewElem.setAttribute "POINT_CODE", "e_point"
        Else
            oNewElem.setAttribute "POINT_CODE", "node"
        End If
        oNewElem.setAttribute "SEG_TYPE", "point"

        oNewElem.setAttribute "SX", oPos.x * 1000#
        oNewElem.setAttribute "SY", oPos.y * 1000#

        oNewElem.setAttribute "A", 0#
        oNewElem.setAttribute "B", 0#
        oNewElem.setAttribute "C", 0#

        oChildCVGnode.appendChild oNewElem
    Next

    Set ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints = oReturnEdgeNode

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, METHOD
End Function

Private Sub ReportProfilePartInformation(oObject As Object, strLogFile As String)
 Const METHOD = "ReportProfilePartInformation"
    On Error GoTo ErrorHandler
    
    Dim oMfgProfilePart      As IJMfgProfilePart
    'Dim oElements           As IJElements
    Dim iCount               As Long
    Dim oNamedItem As IJNamedItem

    Set oMfgProfilePart = oObject
       
    If oMfgProfilePart Is Nothing Then
        Exit Sub
    End If
    
    'Open log file
    Dim oStream As TextStream
    Dim oFSO As FileSystemObject
    Set oFSO = New FileSystemObject
    Set oStream = oFSO.OpenTextFile(strLogFile, ForWriting, True)
    Set oFSO = Nothing
    
    DisplayParentObjectsInformation oMfgProfilePart, oStream
    Dim strMsg                 As String

        
    On Error GoTo ErrorHandler

    
    Dim oStrMfgProfileSupport As IJMfgProfileReport
    
    Dim lCounter As Long
    
    '' Query for IJMfgProfileReport interface:
    Set oStrMfgProfileSupport = oMfgProfilePart
    
    Dim dProfLength As Double
    dProfLength = oStrMfgProfileSupport.GetProcessingLength(STRMFG_AfterFeaturesLength, -1)
    
    '' *************************************************************************
    ''  TEST : GetProcessingLength
    '' *************************************************************************
    
    Dim dLandingCurveLength As Double
    Dim dLandingCurveStart As Double
    dLandingCurveLength = oStrMfgProfileSupport.GetProcessingLength(STRMFG_LandingCurveLength, STRMFG_PROFILE_LENGTH)
    dLandingCurveStart = oStrMfgProfileSupport.GetProcessingLength(STRMFG_LandingCurveStart, STRMFG_PROFILE_LENGTH)
    
    oStream.WriteLine ""
    oStream.WriteLine "STRMFG_LandingCurveLength:" & vbTab & "= " & Round(1000# * dLandingCurveLength, ROUND_CONST)
    oStream.WriteLine ""
    oStream.WriteLine "STRMFG_LandingCurveStart:" & vbTab & "= " & Round(1000# * dLandingCurveStart, ROUND_CONST)
    oStream.WriteLine ""
    oStream.WriteLine "STRMFG_StructDetailLength:"
    oStream.WriteLine vbTab & "Profile Length" & vbTab & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_StructDetailLength, -1), ROUND_CONST)
    oStream.WriteLine vbTab & "Web Length" & vbTab & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_StructDetailLength, 0), ROUND_CONST)
    oStream.WriteLine vbTab & "Top Flange Length" & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_StructDetailLength, JXSEC_TOP), ROUND_CONST)
    oStream.WriteLine vbTab & "Bottom Flange Length" & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_StructDetailLength, JXSEC_BOTTOM), ROUND_CONST)
    oStream.WriteLine ""
    oStream.WriteLine "STRMFG_UnfoldedLength:"
    oStream.WriteLine vbTab & "Profile Length" & vbTab & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_UnfoldedLength, -1), ROUND_CONST)
    oStream.WriteLine vbTab & "Web Length" & vbTab & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_UnfoldedLength, 0), ROUND_CONST)
    oStream.WriteLine vbTab & "Top Flange Length" & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_UnfoldedLength, JXSEC_TOP), ROUND_CONST)
    oStream.WriteLine vbTab & "Bottom Flange Length" & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_UnfoldedLength, JXSEC_BOTTOM), ROUND_CONST)
    oStream.WriteLine ""
    oStream.WriteLine "STRMFG_BeforeFeaturesLength:"
    oStream.WriteLine vbTab & "Profile Length" & vbTab & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_BeforeFeaturesLength, -1), ROUND_CONST)
    oStream.WriteLine vbTab & "Web Length" & vbTab & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_BeforeFeaturesLength, 0), ROUND_CONST)
    oStream.WriteLine vbTab & "Top Flange Length" & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_BeforeFeaturesLength, JXSEC_TOP), ROUND_CONST)
    oStream.WriteLine vbTab & "Bottom Flange Length" & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_BeforeFeaturesLength, JXSEC_BOTTOM), ROUND_CONST)
    oStream.WriteLine ""
    oStream.WriteLine "STRMFG_AfterFeaturesLength:"
    oStream.WriteLine vbTab & "Profile Length" & vbTab & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_AfterFeaturesLength, -1), ROUND_CONST)
    oStream.WriteLine vbTab & "Web Length" & vbTab & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_AfterFeaturesLength, 0), ROUND_CONST)
    oStream.WriteLine vbTab & "Top Flange Length" & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_AfterFeaturesLength, JXSEC_TOP), ROUND_CONST)
    oStream.WriteLine vbTab & "Bottom Flange Length" & vbTab & "= " & Round(1000# * oStrMfgProfileSupport.GetProcessingLength(STRMFG_AfterFeaturesLength, JXSEC_BOTTOM), ROUND_CONST)
    oStream.WriteLine ""


    '' *************************************************************************
    ''  TEST : GetNeutralAxis
    '' *************************************************************************
    oStream.WriteLine ""
    oStream.WriteLine "Testing GetNeutralAxis"
    Dim dNeutralAxisX As Double, dNeutralAxisY As Double
    dNeutralAxisX = -1
    dNeutralAxisY = -1
    oStrMfgProfileSupport.GetNeutralAxis dNeutralAxisX, dNeutralAxisY
    oStream.WriteLine "Neutral Axis X value: " & Round(1000 * dNeutralAxisX, ROUND_CONST) & vbNewLine & "Neutral Axis Y value: " & Round(1000 * dNeutralAxisY, ROUND_CONST)
    oStream.WriteLine ""
    oStream.WriteLine ""
    
    '' *************************************************************************
    ''  TEST : GetBevelShift
    '' *************************************************************************
    oStream.WriteLine "Testing GetBevelShift"
    Dim dBevelAngle As Double, dBevelGapDist As Double
    Dim strGapMsg As String
    dBevelAngle = -1
    dBevelGapDist = -1
    strMsg = "Distance:"
    For lCounter = 0 To dProfLength
        strMsg = strMsg & vbTab & 1000 * lCounter
    Next lCounter
    oStream.WriteLine strMsg
    lCounter = 0
    strMsg = "BvlAngle:"
    strGapMsg = "GapDist:"
    For lCounter = 0 To dProfLength
        oStrMfgProfileSupport.GetBevelShift lCounter, dBevelAngle, dBevelGapDist
        strMsg = strMsg & vbTab & Round(RADIANS_TO_DEGREES * dBevelAngle, ROUND_CONST)
        strGapMsg = strGapMsg & vbTab & Round(1000 * dBevelGapDist, ROUND_CONST)
    Next lCounter
    lCounter = 0
    oStream.WriteLine strMsg
    oStream.WriteLine strGapMsg
    oStream.WriteLine ""
    oStream.WriteLine ""

    '' *************************************************************************
    ''  TEST : GetConnObjDistFromStartOfUnfoldProfile
    '' *************************************************************************
    Dim oMfgChild           As IJMfgChild
    Dim oParent             As IJMfgParent
    Dim oSDPartWrapper As Object
    
    Set oMfgChild = oMfgProfilePart
    Set oParent = oMfgChild.getParent
    Set oMfgChild = Nothing
    'Create the SD part Wrapper and initialize it
    If Not oParent Is Nothing Then
        If TypeOf oParent Is IJStiffenerPart Then
            Set oSDPartWrapper = New StructDetailObjects.ProfilePart
            Set oSDPartWrapper.object = oParent
        ElseIf TypeOf oParent Is IJBeamPart Then
            Set oSDPartWrapper = New StructDetailObjects.BeamPart
            Set oSDPartWrapper.object = oParent
        End If
    End If
    Set oParent = Nothing
    
    
    '' Get all the connected objects on the Profile
    oStream.WriteLine "Testing GetConnObjDistFromStartOfUnfoldProfile"
    Dim oConnObjsColl As Collection
    Set oConnObjsColl = oSDPartWrapper.ConnectedObjects
    oStream.WriteLine "Number of connections to other objects (objects may be connected more than once): " & oConnObjsColl.Count
    If oConnObjsColl.Count > 0 Then
        oStream.WriteLine "Index  " & vbTab & " Name" & vbTab & vbTab & vbTab & vbTab & "Distance"
    End If
    
    
    For lCounter = 1 To oConnObjsColl.Count
        Dim oConnectionData As ConnectionData
        On Error Resume Next
        oConnectionData = oConnObjsColl.Item(lCounter)
        '**********************************
        Dim oConnObject As IUnknown
        Set oConnObject = oConnectionData.ToConnectable
        If oConnObject Is Nothing Then
            Set oConnObject = oConnObjsColl.Item(lCounter)
        End If
        If Not oConnObject Is Nothing Then
            Dim strConnObjName As String
            Set oNamedItem = oConnObject
            If Not oNamedItem Is Nothing Then
                strConnObjName = oNamedItem.Name
                Set oNamedItem = Nothing
            Else
                strConnObjName = "no name"
            End If
            Dim dConnObjDis As Double
            On Error Resume Next
            dConnObjDis = oStrMfgProfileSupport.GetConnObjDistFromStartOfUnfoldProfile(oConnObject)
            On Error GoTo ErrorHandler
            If Len(strConnObjName) <= 15 Then
                oStream.WriteLine lCounter & vbTab & strConnObjName & vbTab & vbTab & vbTab & Round(1000 * dConnObjDis, ROUND_CONST)
            ElseIf Len(strConnObjName) <= 23 Then
                oStream.WriteLine lCounter & vbTab & strConnObjName & vbTab & vbTab & Round(1000 * dConnObjDis, ROUND_CONST)
            Else
                oStream.WriteLine lCounter & vbTab & strConnObjName & vbTab & Round(1000 * dConnObjDis, ROUND_CONST)
            End If
            Set oConnObject = Nothing
        Else
            oStream.WriteLine lCounter & vbTab & "Could not obtain connected object information"
        End If
        '*****************************
    Next lCounter
    oStream.WriteLine ""

    
    '' ***************************************************************************
    ''  TEST : GetFeatureDistFromStartOfUnfoldProfile and GetFeatureDistAndHeight
    '' ***************************************************************************
    Dim oFeaturesColl As Collection
    '' Get all the features on the Part
    Set oFeaturesColl = oSDPartWrapper.ProfileFeatures
    'oStrMfgProfileSupport.UseMarginShrinkage = False 'is no longer supported
    
    oStream.WriteLine ""
    oStream.WriteLine "Testing GetFeatureDistFromStartOfUnfoldProfile (FD) and GetFeatureDistAndHeight (FD2, FH)"
    oStream.WriteLine "Number of Feature objects: " & oFeaturesColl.Count
    oStream.WriteLine "dLandingCurveStart (LCS) : " & Round(1000 * dLandingCurveStart, ROUND_CONST)
    oStream.WriteLine "dLandingCurveLength(LCL) : " & Round(1000 * dLandingCurveLength, ROUND_CONST)
    
    If oFeaturesColl.Count > 0 Then
        oStream.WriteLine "Index" & vbTab & "Type" & vbTab & vbTab & vbTab & "FD" & vbTab & vbTab & _
                          "(FD-LCS)" & vbTab & "(LCL-FD)" & vbTab & "(LCL-(FD-LCS))" & vbTab & _
                          "EndCutIsAtBase" & vbTab & "FD2" & vbTab & vbTab & "FDH"
    End If
    
    Dim oSlotCol As IJElements
    Set oSlotCol = New JObjectCollection
    
    '' Get each object in collection and check for required feature type
    For lCounter = 1 To oFeaturesColl.Count
        On Error Resume Next
        
        Dim oStructFeature As IJStructFeature
        Set oStructFeature = oFeaturesColl.Item(lCounter)
        If Not oStructFeature Is Nothing Then
            Dim enumType As StructFeatureTypes
            enumType = oStructFeature.get_StructFeatureType
        
            Dim enumTypeStr As String
            Select Case enumType
                Case SF_EdgeFeature
                    enumTypeStr = "SF_EdgeFeature  "
                Case SF_CornerFeature
                    enumTypeStr = "SF_CornerFeature"
                Case SF_FaceFeature
                    enumTypeStr = "SF_FaceFeature  "
                Case SF_FlangeCut
                    enumTypeStr = "SF_FlangeCut    "
                Case SF_Slot
                    enumTypeStr = "SF_Slot         "
                    oSlotCol.Add oStructFeature
                Case SF_WaterStop
                    enumTypeStr = "SF_WaterStop    "
                Case SF_WebCut
                    enumTypeStr = "SF_WebCut       "
            End Select
            
            Dim oFeatBool As Boolean
            oFeatBool = oStrMfgProfileSupport.CheckIfEndCutIsAtBase(oStructFeature)
            
            Dim dFeatureDis As Double
            dFeatureDis = oStrMfgProfileSupport.GetFeatureDistFromStartOfUnfoldProfile(oStructFeature)
            
            Dim dFD2 As Double, dFH As Double
            oStrMfgProfileSupport.GetFeatureDistAndHeight oStructFeature, dFD2, dFH
            
            oStream.WriteLine lCounter & vbTab & enumTypeStr _
                & vbTab & Round(1000 * dFeatureDis, ROUND_CONST) _
                & vbTab & vbTab & Round(1000 * (dFeatureDis - dLandingCurveStart), ROUND_CONST) _
                & vbTab & vbTab & Round(1000 * (dLandingCurveLength - (dFeatureDis)), ROUND_CONST) _
                & vbTab & vbTab & Round(1000 * (dLandingCurveLength - (dFeatureDis - dLandingCurveStart)), ROUND_CONST) _
                & vbTab & vbTab & oFeatBool _
                & vbTab & vbTab & Round(1000 * dFD2, ROUND_CONST) _
                & vbTab & vbTab & Round(1000 * dFH, ROUND_CONST)
            Set oStructFeature = Nothing
        End If
    Next
    oStream.WriteLine ""
    oStream.WriteLine ""


    
    '' ***************************************************************************
    ''  TEST : GetSlotDistFromStartOfUnfoldProfile
    '' ***************************************************************************
    For lCounter = 1 To oSlotCol.Count
        If lCounter = 1 Then
            oStream.WriteLine "Testing GetSlotDistFromStartOfUnfoldProfile"
            oStream.WriteLine "Number of Slot Features = " & oSlotCol.Count
            oStream.WriteBlankLines 1
            oStream.WriteLine "Idx,Slot Name,Distance of contour edge due to slot on Webleft from start"
        End If
        
        Dim dSlotDist As Double
        dSlotDist = oStrMfgProfileSupport.GetSlotDistFromStartOfUnfoldProfile(oSlotCol.Item(lCounter), JXSEC_WEB_LEFT)
        
        Set oNamedItem = oSlotCol.Item(lCounter)
        oStream.WriteLine lCounter & "," & oNamedItem.Name & "," & Round(1000 * dSlotDist, ROUND_CONST)
    Next
    oStream.WriteBlankLines 2
    
   
    '' *************************************************************************
    ''  TEST : GetMarkingLineDistFromStartOfUnfoldProfile
    '' *************************************************************************
        '' Collect the Marking Lines on Web & Flanges
    Dim oWebMLs As Collection
    Dim oTopFlangeMLs As Collection
    Dim oBtmFlangeMLs As Collection
    Dim oUnkMark As IUnknown
    Dim dMarkingDis As Double
    'Dim dDistDist As Double
    Dim strMLName As String
    
    strMLName = "test"
    Set oWebMLs = New Collection
    Set oTopFlangeMLs = New Collection
    Set oBtmFlangeMLs = New Collection
       
    '' Get all the Markings on the Profile
    GetMarkingLines oMfgProfilePart, oWebMLs, oTopFlangeMLs, oBtmFlangeMLs
    If oWebMLs.Count > 0 Then
        oStream.WriteLine "Testing GetMarkingLineDistFromStartOfUnfoldProfile for WEB "
        oStream.WriteLine "Number of Marking objects: " & oWebMLs.Count / 2
        oStream.WriteLine "Index" & vbTab & "GetMkgLineDist" & vbTab & vbTab & "Markingline Type"
        
        'oLogFile.ostream.writeline "UseMarginShrinkage = false"
        'oStrMfgProfileSupport.UseMarginShrinkage = false 'is no longer supported
        '' Get each object in collection and check for required feature type
'        Dim oMarkGeom2d As IJMfgGeom2d
        For lCounter = 1 To (oWebMLs.Count / 2)
            On Error Resume Next
            Set oUnkMark = oWebMLs.Item(lCounter)
            If Not oUnkMark Is Nothing Then
                strMLName = oWebMLs.Item(lCounter + (oWebMLs.Count / 2))
                dMarkingDis = oStrMfgProfileSupport.GetMarkingLineDistFromStartOfUnfoldProfile(oUnkMark)
                strMsg = lCounter & vbTab & Round(1000 * dMarkingDis, ROUND_CONST) & vbTab & vbTab & vbTab & strMLName
'                'If TypeOf oUnkMark Is IJMfgSystemMark Then
''                Set oMarkGeom2d = GetMarkingLineGeom2D(oUnkMark)
''                dDistDist = oStrMfgProfileSupport.DistanceFromStartOfUnfoldedProfile(oMarkGeom2d)
''                Set oMarkGeom2d = Nothing
'                'strMsg = strMsg & vbTab & vbTab & Round(1000 * dDistDist, ROUND_CONST) & vbTab & vbTab & strMLName
'                'Else
'                    'strMsg = strMsg & vbTab & vbTab & "N/A" & vbTab & vbTab & strMLName
'                'End If
                oStream.WriteLine strMsg
                Set oUnkMark = Nothing
            End If
        Next
    End If
    
    If oTopFlangeMLs.Count > 0 Then
        oStream.WriteLine ""
        oStream.WriteLine "Testing GetMarkingLineDistFromStartOfUnfoldProfile for TOPFLANGE"
        oStream.WriteLine "Number of Marking objects: " & oTopFlangeMLs.Count / 2
        oStream.WriteLine "Index" & vbTab & "GetMkgLineDist" & vbTab & vbTab & "Markingline Type"
         
        '' Get each object in collection and check for required feature type
        For lCounter = 1 To (oTopFlangeMLs.Count / 2)
            On Error Resume Next
            Set oUnkMark = oTopFlangeMLs.Item(lCounter)
            If Not oUnkMark Is Nothing Then
                strMLName = oTopFlangeMLs.Item(lCounter + (oTopFlangeMLs.Count / 2))
                dMarkingDis = oStrMfgProfileSupport.GetMarkingLineDistFromStartOfUnfoldProfile(oUnkMark)
                strMsg = lCounter & vbTab & Round(1000 * dMarkingDis, ROUND_CONST) & vbTab & vbTab & vbTab & strMLName
'                Set oMarkGeom2d = GetMarkingLineGeom2D(oUnkMark)
'                dDistDist = oStrMfgProfileSupport.DistanceFromStartOfUnfoldedProfile(oMarkGeom2d)
'                Set oMarkGeom2d = Nothing
'                strMsg = strMsg & vbTab & vbTab & Round(1000 * dDistDist, ROUND_CONST) & vbTab & vbTab & strMLName
                oStream.WriteLine strMsg
                Set oUnkMark = Nothing
            End If
        Next
    End If
    
    If oBtmFlangeMLs.Count > 0 Then
        oStream.WriteLine ""
        oStream.WriteLine "Testing GetMarkingLineDistFromStartOfUnfoldProfile for BOTTOMFLANGE"
        oStream.WriteLine "Number of Marking objects: " & oBtmFlangeMLs.Count / 2
        oStream.WriteLine "Index" & vbTab & "GetMkgLineDist" & vbTab & vbTab & "Markingline Type"
        
        'oLogFile.ostream.writeline "UseMarginShrinkage = False"
        'oStrMfgProfileSupport.UseMarginShrinkage = False 'is no longer supported
        '' Get each object in collection and check for required feature type
        For lCounter = 1 To (oBtmFlangeMLs.Count / 2)
            On Error Resume Next
            Set oUnkMark = oBtmFlangeMLs.Item(lCounter)
            If Not oUnkMark Is Nothing Then
                strMLName = oBtmFlangeMLs.Item(lCounter + (oBtmFlangeMLs.Count / 2))
                dMarkingDis = oStrMfgProfileSupport.GetMarkingLineDistFromStartOfUnfoldProfile(oUnkMark)
                strMsg = lCounter & vbTab & Round(1000 * dMarkingDis, ROUND_CONST) & vbTab & vbTab & vbTab & strMLName
'                Set oMarkGeom2d = GetMarkingLineGeom2D(oUnkMark)
'                dDistDist = oStrMfgProfileSupport.DistanceFromStartOfUnfoldedProfile(oMarkGeom2d)
'                Set oMarkGeom2d = Nothing
'                strMsg = strMsg & vbTab & vbTab & Round(1000 * dDistDist, ROUND_CONST) & vbTab & vbTab & strMLName
                oStream.WriteLine strMsg
                Set oUnkMark = Nothing
            End If
        Next
    End If
    
        
    '' *************************************************************************
    ''  TEST : GetSupportValue
    '' *************************************************************************
    oStream.WriteLine ""
    oStream.WriteLine ""
    oStream.WriteLine "Testing GetSupportValue & GetSupportValueRange"
    Dim dProcLength As Double
    Dim dMaxRange As Double
    Dim dMinRange As Double
    Dim dRangeIncr As Double
    Dim dSupportVal As Double
    Dim dDistanceVal As Double
    Dim lMfgGeomTypes(8) As Long
    Dim lIncr As Long
    
    lMfgGeomTypes(0) = STRMFG_BENDING_LINE
    lMfgGeomTypes(1) = STRMFG_FITTINGANGLE
    lMfgGeomTypes(2) = STRMFG_TWIST_INFO
    lMfgGeomTypes(3) = STRMFG_PROFILE_BEND_INFO
    lMfgGeomTypes(4) = STRMFG_PROFILE_ROLL_INFO
    lMfgGeomTypes(5) = STRMFG_PROFILE_DEPTH_AFTERUNTWIST
    lMfgGeomTypes(6) = STRMFG_PROFILE_DEPTH_BEFOREUNTWIST
    lMfgGeomTypes(7) = STRMFG_PROFILE_LENGTH
    lMfgGeomTypes(8) = STRMFG_LANDING_CURVE
    
    
    For lCounter = 0 To 8
        dMinRange = -99
        dMaxRange = -99
        dRangeIncr = -1
        dProcLength = -1
        dSupportVal = -1
        strMsg = ""
        oStrMfgProfileSupport.GetSupportValueRange lMfgGeomTypes(lCounter), JXSEC_WEB_LEFT, dMinRange, dMaxRange
        oStream.WriteLine GetGeometryTypeName(lMfgGeomTypes(lCounter)) & vbTab & "JXSEC_WEB_LEFT" & vbTab & "Min: " & Round(1000 * dMinRange, ROUND_CONST) & vbTab & "Max: " & Round(1000 * dMaxRange, ROUND_CONST)
        If dMaxRange > dMinRange Then
            'dProcLength = oStrMfgProfileSupport.GetProcessingLength(STRMFG_BeforeFeaturesLength, JXSEC_WEB_LEFT)
            oStream.WriteLine "Support values:"
            dRangeIncr = (dMaxRange - dMinRange) / 10
            dDistanceVal = 0
            For lIncr = 0 To 10
                strMsg = strMsg & Round(1000 * dDistanceVal, ROUND_CONST) & vbTab
                dDistanceVal = dDistanceVal + dRangeIncr
            Next lIncr
            oStream.WriteLine strMsg
            strMsg = ""
            dDistanceVal = 0
            For lIncr = 0 To 10
                dSupportVal = oStrMfgProfileSupport.GetSupportValue(lMfgGeomTypes(lCounter), JXSEC_WEB_LEFT, dDistanceVal)
                dDistanceVal = dDistanceVal + dRangeIncr
                If lMfgGeomTypes(lCounter) = STRMFG_FITTINGANGLE Then
                    strMsg = strMsg & Round(RADIANS_TO_DEGREES * dSupportVal, ROUND_CONST) & vbTab
                Else
                    strMsg = strMsg & Round(1000 * dSupportVal, ROUND_CONST) & vbTab
                End If
            Next lIncr
            oStream.WriteLine strMsg
        End If
        oStream.WriteLine ""
        
        dMinRange = -99
        dMaxRange = -99
        dRangeIncr = -1
        dProcLength = -1
        dSupportVal = -1
        strMsg = ""
        oStrMfgProfileSupport.GetSupportValueRange lMfgGeomTypes(lCounter), JXSEC_WEB_RIGHT, dMinRange, dMaxRange
        oStream.WriteLine GetGeometryTypeName(lMfgGeomTypes(lCounter)) & vbTab & "JXSEC_WEB_RIGHT" & vbTab & "Min: " & Round(1000 * dMinRange, ROUND_CONST) & vbTab & "Max: " & Round(1000 * dMaxRange, ROUND_CONST)
        If dMaxRange > dMinRange Then
            'dProcLength = oStrMfgProfileSupport.GetProcessingLength(STRMFG_BeforeFeaturesLength, JXSEC_WEB_RIGHT)
            oStream.WriteLine "Support values:"
            dRangeIncr = (dMaxRange - dMinRange) / 10
            dDistanceVal = 0
            For lIncr = 0 To 10
                strMsg = strMsg & Round(1000 * dDistanceVal, ROUND_CONST) & vbTab
                dDistanceVal = dDistanceVal + dRangeIncr
            Next lIncr
            oStream.WriteLine strMsg
            strMsg = ""
            dDistanceVal = 0
            For lIncr = 0 To 10
                dSupportVal = oStrMfgProfileSupport.GetSupportValue(lMfgGeomTypes(lCounter), JXSEC_WEB_RIGHT, dDistanceVal)
                dDistanceVal = dDistanceVal + dRangeIncr
                If lMfgGeomTypes(lCounter) = STRMFG_FITTINGANGLE Then
                    strMsg = strMsg & Round(RADIANS_TO_DEGREES * dSupportVal, ROUND_CONST) & vbTab
                Else
                    strMsg = strMsg & Round(1000 * dSupportVal, ROUND_CONST) & vbTab
                End If
            Next lIncr
            oStream.WriteLine strMsg
        End If
        oStream.WriteLine ""
        
        dMinRange = -99
        dMaxRange = -99
        dRangeIncr = -1
        dProcLength = -1
        dSupportVal = -1
        strMsg = ""
        oStrMfgProfileSupport.GetSupportValueRange lMfgGeomTypes(lCounter), JXSEC_TOP, dMinRange, dMaxRange
        oStream.WriteLine GetGeometryTypeName(lMfgGeomTypes(lCounter)) & vbTab & "JXSEC_TOP" & vbTab & "Min: " & Round(1000 * dMinRange, ROUND_CONST) & vbTab & "Max: " & Round(1000 * dMaxRange, ROUND_CONST)
        If dMaxRange > dMinRange Then
            'dProcLength = oStrMfgProfileSupport.GetProcessingLength(STRMFG_BeforeFeaturesLength, JXSEC_TOP)
            oStream.WriteLine "Support values:"
            dRangeIncr = (dMaxRange - dMinRange) / 10
            dDistanceVal = 0
            For lIncr = 0 To 10
                strMsg = strMsg & Round(1000 * dDistanceVal, ROUND_CONST) & vbTab
                dDistanceVal = dDistanceVal + dRangeIncr
            Next lIncr
            oStream.WriteLine strMsg
            strMsg = ""
            dDistanceVal = 0
            For lIncr = 0 To 10
                dSupportVal = oStrMfgProfileSupport.GetSupportValue(lMfgGeomTypes(lCounter), JXSEC_TOP, dDistanceVal)
                dDistanceVal = dDistanceVal + dRangeIncr
                If lMfgGeomTypes(lCounter) = STRMFG_FITTINGANGLE Then
                    strMsg = strMsg & Round(RADIANS_TO_DEGREES * dSupportVal, ROUND_CONST) & vbTab
                Else
                    strMsg = strMsg & Round(1000 * dSupportVal, ROUND_CONST) & vbTab
                End If
            Next lIncr
            oStream.WriteLine strMsg
        End If
        oStream.WriteLine ""
        
        dMinRange = -99
        dMaxRange = -99
        dRangeIncr = -1
        dProcLength = -1
        dSupportVal = -1
        strMsg = ""
        oStrMfgProfileSupport.GetSupportValueRange lMfgGeomTypes(lCounter), JXSEC_BOTTOM, dMinRange, dMaxRange
        oStream.WriteLine GetGeometryTypeName(lMfgGeomTypes(lCounter)) & vbTab & "JXSEC_BOTTOM" & vbTab & "Min: " & Round(1000 * dMinRange, ROUND_CONST) & vbTab & "Max: " & Round(1000 * dMaxRange, ROUND_CONST)
        If dMaxRange > dMinRange Then
            'dProcLength = oStrMfgProfileSupport.GetProcessingLength(STRMFG_BeforeFeaturesLength, JXSEC_BOTTOM)
            oStream.WriteLine "Support values:"
            dRangeIncr = (dMaxRange - dMinRange) / 10
            dDistanceVal = 0
            For lIncr = 0 To 10
                strMsg = strMsg & Round(1000 * dDistanceVal, ROUND_CONST) & vbTab
                dDistanceVal = dDistanceVal + dRangeIncr
            Next lIncr
            oStream.WriteLine strMsg
            strMsg = ""
            dDistanceVal = 0
            For lIncr = 0 To 10
                dSupportVal = oStrMfgProfileSupport.GetSupportValue(lMfgGeomTypes(lCounter), JXSEC_BOTTOM, dDistanceVal)
                dDistanceVal = dDistanceVal + dRangeIncr
                If lMfgGeomTypes(lCounter) = STRMFG_FITTINGANGLE Then
                    strMsg = strMsg & Round(RADIANS_TO_DEGREES * dSupportVal, ROUND_CONST) & vbTab
                Else
                    strMsg = strMsg & Round(1000 * dSupportVal, ROUND_CONST) & vbTab
                End If
            Next lIncr
            oStream.WriteLine strMsg
        End If
        oStream.WriteLine ""
        oStream.WriteLine ""
        oStream.WriteLine ""
    Next lCounter


    '' *************************************************************************
    ''  TEST : GetProfileKnuckleInfo
    '' *************************************************************************
    oStream.WriteLine "Testing GetProfileKnuckleInfo"
    Dim oGeomColl As IJDObjectCollection
    Dim oGeom2d As IJMfgGeom2d
    Dim dKnuckleAngle As Double, dKnuckleDist As Double
    Dim bKnuckleThisSide As Boolean
    
    Set oGeomColl = oStrMfgProfileSupport.GetGeometries(STRMFG_AFTER_PROCESS2D, STRMFG_KNUCKLE_MARK)
    If oGeomColl.Count > 0 Then
        lCounter = 1
        For Each oGeom2d In oGeomColl
            bKnuckleThisSide = False
            dKnuckleAngle = -1
            dKnuckleDist = -1
            oStrMfgProfileSupport.GetProfileKnuckleInfo oGeom2d, bKnuckleThisSide, dKnuckleAngle, dKnuckleDist
            oStream.WriteLine "Knuckle Number: " & lCounter
            oStream.WriteLine "Is Knuckle on this side: " & bKnuckleThisSide
            oStream.WriteLine "Knuckle Angle: " & Round(RADIANS_TO_DEGREES * dKnuckleAngle, ROUND_CONST)
            oStream.WriteLine "Knuckle distance from start of unfold profile: " & Round(1000 * dKnuckleDist, ROUND_CONST)
            lCounter = lCounter + 1
        Next oGeom2d
        Set oGeom2d = Nothing
    Else
        oStream.WriteLine "There are no knuckles on this profile to test."
    End If
    Set oGeomColl = Nothing
    oStream.WriteLine ""
    oStream.WriteLine ""
    
    
   '' *************************************************************************
    ''  TEST : GetBendingCurveInfo
    '' *************************************************************************
    oStream.WriteLine "Testing GetBendingCurveInfo"
    Dim oStPoint As IJDPosition, oEndPoint As IJDPosition
    Dim bIsBentUp As Boolean
    
    Set oStPoint = New DPosition
    Set oEndPoint = New DPosition
    bIsBentUp = False
    
    Dim oModelPOM As IJDPOM
    Set oModelPOM = GetPOM()
        
    Set oGeomColl = oStrMfgProfileSupport.GetGeometries(STRMFG_AFTER_PROCESS2D, STRMFG_BENDING_LINE)
    lCounter = 1
    If oGeomColl.Count > 0 Then
        For Each oGeom2d In oGeomColl
            Dim oMoniker As IUnknown
            Set oMoniker = oModelPOM.GetObjectMoniker(oGeom2d)
        
            oStrMfgProfileSupport.GetBendingCurveInfo oMoniker, oStPoint, oEndPoint, bIsBentUp
            If bIsBentUp Then
                oStream.WriteLine "Curve Number: " & lCounter & vbTab & vbTab & vbTab & "Curve IS bent up."
            Else
                oStream.WriteLine "Curve Number: " & lCounter & vbTab & vbTab & vbTab & "Curve is NOT bent up."
            End If
            oStream.WriteLine "Starting Point (x, y, z):" & vbTab & "(" & Round(1000 * oStPoint.x, ROUND_CONST) & ", " & _
                    Round(1000 * oStPoint.y, ROUND_CONST) & ", " & Round(1000 * oStPoint.z, ROUND_CONST) & ")"
            oStream.WriteLine "Ending   Point (x, y, z):" & vbTab & "(" & Round(1000 * oEndPoint.x, ROUND_CONST) & ", " & _
                    Round(1000 * oEndPoint.y, ROUND_CONST) & ", " & Round(1000 * oEndPoint.z, ROUND_CONST) & ")"
            'ostream.writeline "Is Curve Bent up: " & bIsBentUp
            lCounter = lCounter + 1
            oStream.WriteLine ""
            bIsBentUp = False
            Set oStPoint = Nothing
            Set oEndPoint = Nothing
        Next oGeom2d
        Set oGeom2d = Nothing
    Else
        oStream.WriteLine " There are no bending curves to test."
    End If
    Set oGeomColl = Nothing
    oStream.WriteLine ""
    oStream.WriteLine ""
    
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
    Set oModelPOM = Nothing
    
    
    '' *************************************************************************
    ''  TEST : GetCurveHeightByOffset -- STRMFG_BENDING_LINE
    '' *************************************************************************
    Dim Offset As Double
    Offset = 0.5
    oStream.WriteLine ""
    oStream.WriteLine "Testing GetCurveHeightByOffset for STRMFG_BENDING_LINE"
    
    Dim nIBLCurveCount As Long
    nIBLCurveCount = oStrMfgProfileSupport.GetCurveCount(STRMFG_BENDING_LINE, JXSEC_WEB_LEFT)

    If nIBLCurveCount > 0 Then
        lCounter = 0
        Dim nIndex As Long
        ' write header
        strMsg = "Curve" & vbTab
        Do
            strMsg = strMsg & 1000 * (lCounter * Offset) & vbTab
            lCounter = lCounter + 1
        Loop While lCounter * Offset < dProfLength
        oStream.WriteLine strMsg
        
        'write each curve on one line
        For nIndex = 1 To nIBLCurveCount
            strMsg = nIndex & vbTab
            lCounter = 0
            Do
                strMsg = strMsg & Round(1000 * oStrMfgProfileSupport.GetCurveHeightByOffset(STRMFG_BENDING_LINE, JXSEC_WEB_LEFT, nIndex, lCounter * Offset), ROUND_CONST) & vbTab
                lCounter = lCounter + 1
            Loop While lCounter * Offset < dProfLength
            oStream.WriteLine strMsg
        Next
        
        oStream.WriteLine ""
        oStream.WriteLine ""
        oStream.WriteLine "Testing GetCurveHeightAtEnds for STRMFG_BENDING_LINE"
        'write end points and offsets
        Dim StartOffset As Double, EndOffset As Double, StartHeight As Double, EndHeight As Double
        oStream.WriteLine "Curve" & vbTab & "Start" & vbTab & "Height" & vbTab & "End" & vbTab & "Height"
        For nIndex = 1 To nIBLCurveCount
            oStrMfgProfileSupport.GetCurveHeightAtEnds STRMFG_BENDING_LINE, JXSEC_WEB_LEFT, nIndex, StartOffset, StartHeight, EndOffset, EndHeight
            oStream.WriteLine nIndex & vbTab & Round(1000 * StartOffset, ROUND_CONST) _
            & vbTab & Round(1000 * StartHeight, ROUND_CONST) _
            & vbTab & Round(1000 * EndOffset, ROUND_CONST) _
            & vbTab & Round(1000 * EndHeight, ROUND_CONST)
        Next
        
        oStream.WriteLine ""
        oStream.WriteLine ""
        oStream.WriteLine "Testing GetCurveHeightsByOffset for STRMFG_BENDING_LINE"
        'write end points and offsets
        For nIndex = 1 To nIBLCurveCount
            oStream.WriteLine "Curve # " & nIndex
            Dim dHeights() As Double
            oStrMfgProfileSupport.GetCurveHeightAtEnds STRMFG_BENDING_LINE, JXSEC_WEB_LEFT, nIndex, StartOffset, StartHeight, EndOffset, EndHeight
            dHeights = oStrMfgProfileSupport.GetCurveHeightsByOffset(STRMFG_BENDING_LINE, JXSEC_WEB_LEFT, nIndex, StartOffset, Offset)
            Dim strHeader As String
            Dim strHeights As String
            strHeader = " Offset: " & vbTab
            strHeights = " Height: " & vbTab
            For iCount = LBound(dHeights) To UBound(dHeights)
                    strHeader = strHeader & Round(1000 * (StartOffset + (iCount * Offset)), ROUND_CONST) & vbTab
                    strHeights = strHeights & Round(1000 * dHeights(iCount), ROUND_CONST) & vbTab
            Next
            oStream.WriteLine strHeader
            oStream.WriteLine strHeights
        Next
    Else
        oStream.WriteLine " There are no curves of type STRMFG_BENDING_LINE"
    End If
    oStream.WriteLine ""
    oStream.WriteLine ""
    
    
    '' *************************************************************************
    ''  TEST : GetCurveHeightByOffset -- STRMFG_TOPFLANGE_BENDING_LINE
    '' *************************************************************************
    oStream.WriteLine ""
    oStream.WriteLine "Testing GetCurveHeightByOffset for STRMFG_TOPFLANGE_BENDING_LINE"
    
    nIBLCurveCount = oStrMfgProfileSupport.GetCurveCount(STRMFG_BENDING_LINE, JXSEC_TOP)

    If nIBLCurveCount > 0 Then
        lCounter = 0
               
        ' write header
        strMsg = "Curve" & vbTab
        Do
            strMsg = strMsg & 1000 * (lCounter * Offset) & vbTab
            lCounter = lCounter + 1
        Loop While lCounter * Offset < dProfLength
        oStream.WriteLine strMsg
        
        'write each curve on one line
        For nIndex = 1 To nIBLCurveCount
            strMsg = nIndex & vbTab
            lCounter = 0
            Do
                strMsg = strMsg & Round(1000 * oStrMfgProfileSupport.GetCurveHeightByOffset(STRMFG_BENDING_LINE, JXSEC_TOP, nIndex, lCounter * Offset), ROUND_CONST) & vbTab
                lCounter = lCounter + 1
            Loop While lCounter * Offset < dProfLength
            oStream.WriteLine strMsg
        Next
                
        oStream.WriteLine ""
        oStream.WriteLine ""
        oStream.WriteLine "Testing GetCurveHeightAtEnds for STRMFG_TOPFLANGE_BENDING_LINE"
        'write end points and offsets
        oStream.WriteLine "Curve" & vbTab & "Start" & vbTab & "Height" & vbTab & "End" & vbTab & "Height"
        For nIndex = 1 To nIBLCurveCount
            oStrMfgProfileSupport.GetCurveHeightAtEnds STRMFG_BENDING_LINE, JXSEC_TOP, nIndex, StartOffset, StartHeight, EndOffset, EndHeight
            oStream.WriteLine nIndex & vbTab & Round(1000 * StartOffset, ROUND_CONST) _
            & vbTab & Round(1000 * StartHeight, ROUND_CONST) _
            & vbTab & Round(1000 * EndOffset, ROUND_CONST) _
            & vbTab & Round(1000 * EndHeight, ROUND_CONST)
        Next
        
        oStream.WriteLine ""
        oStream.WriteLine ""
        oStream.WriteLine "Testing GetCurveHeightsByOffset for STRMFG_TOPFLANGE_BENDING_LINE"
        'write end points and offsets
        For nIndex = 1 To nIBLCurveCount
            oStream.WriteLine "Curve # " & nIndex
            oStrMfgProfileSupport.GetCurveHeightAtEnds STRMFG_BENDING_LINE, JXSEC_TOP, nIndex, StartOffset, StartHeight, EndOffset, EndHeight
            dHeights = oStrMfgProfileSupport.GetCurveHeightsByOffset(STRMFG_BENDING_LINE, JXSEC_TOP, nIndex, StartOffset, Offset)
            strHeader = " Offset: " & vbTab
            strHeights = " Height: " & vbTab
            For iCount = LBound(dHeights) To UBound(dHeights)
                    strHeader = strHeader & Round(1000 * (StartOffset + (iCount * Offset)), ROUND_CONST) & vbTab
                    strHeights = strHeights & Round(1000 * dHeights(iCount), ROUND_CONST) & vbTab
            Next
            oStream.WriteLine strHeader
            oStream.WriteLine strHeights
        Next
    Else
        oStream.WriteLine " There are no curves of type STRMFG_TOPFLANGE_BENDING_LINE"
    End If
    oStream.WriteLine ""
    oStream.WriteLine ""
    
    '' *************************************************************************
    ''  TEST : GetCurveHeightByOffset -- STRMFG_BOTTOMFLANGE_BENDING_LINE
    '' *************************************************************************
    oStream.WriteLine ""
    oStream.WriteLine "Testing GetCurveHeightByOffset for STRMFG_BOTTOMFLANGE_BENDING_LINE"
    
    nIBLCurveCount = oStrMfgProfileSupport.GetCurveCount(STRMFG_BENDING_LINE, JXSEC_BOTTOM)

    If nIBLCurveCount > 0 Then
        lCounter = 0
               
        ' write header
        strMsg = "Curve" & vbTab
        Do
            strMsg = strMsg & 1000 * (lCounter * Offset) & vbTab
            lCounter = lCounter + 1
        Loop While lCounter * Offset < dProfLength
        oStream.WriteLine strMsg
        
        'write each curve on one line
        For nIndex = 1 To nIBLCurveCount
            strMsg = nIndex & vbTab
            lCounter = 0
            Do
                strMsg = strMsg & Round(1000 * oStrMfgProfileSupport.GetCurveHeightByOffset(STRMFG_BENDING_LINE, JXSEC_BOTTOM, nIndex, lCounter * Offset), ROUND_CONST) & vbTab
                lCounter = lCounter + 1
            Loop While lCounter * Offset < dProfLength
            oStream.WriteLine strMsg
        Next
                
        oStream.WriteLine ""
        oStream.WriteLine "Testing GetCurveHeightAtEnds for STRMFG_BOTTOMFLANGE_BENDING_LINE"
        'write end points and offsets
        oStream.WriteLine "Curve" & vbTab & "Start" & vbTab & "Height" & vbTab & "End" & vbTab & "Height"
        For nIndex = 1 To nIBLCurveCount
            oStrMfgProfileSupport.GetCurveHeightAtEnds STRMFG_BENDING_LINE, JXSEC_BOTTOM, nIndex, StartOffset, StartHeight, EndOffset, EndHeight
            oStream.WriteLine nIndex & vbTab & Round(1000 * StartOffset, ROUND_CONST) _
            & vbTab & Round(1000 * StartHeight, ROUND_CONST) _
            & vbTab & Round(1000 * EndOffset, ROUND_CONST) _
            & vbTab & Round(1000 * EndHeight, ROUND_CONST)
        Next
        
        oStream.WriteLine ""
        oStream.WriteLine "Testing GetCurveHeightsByOffset for STRMFG_BOTTOMFLANGE_BENDING_LINE"
        'write end points and offsets
        For nIndex = 1 To nIBLCurveCount
            oStream.WriteLine "Curve # " & nIndex
            oStrMfgProfileSupport.GetCurveHeightAtEnds STRMFG_BENDING_LINE, JXSEC_BOTTOM, nIndex, StartOffset, StartHeight, EndOffset, EndHeight
            dHeights = oStrMfgProfileSupport.GetCurveHeightsByOffset(STRMFG_BENDING_LINE, JXSEC_BOTTOM, nIndex, StartOffset, Offset)
            strHeader = " Offset: " & vbTab
            strHeights = " Height: " & vbTab
            For iCount = LBound(dHeights) To UBound(dHeights)
                    strHeader = strHeader & Round(1000 * (StartOffset + (iCount * Offset)), ROUND_CONST) & vbTab
                    strHeights = strHeights & Round(1000 * dHeights(iCount), ROUND_CONST) & vbTab
            Next
            oStream.WriteLine strHeader
            oStream.WriteLine strHeights
        Next
    Else
        oStream.WriteLine " There are no curves of type STRMFG_BOTTOMFLANGE_BENDING_LINE"
    End If
    oStream.WriteLine ""
    oStream.WriteLine ""
    
    
   '' *************************************************************************
    ''  TEST : StrokeEdgeNodeIntoPoints
    '' *************************************************************************
    oStream.WriteLine "Testing StrokeEdgeNodeIntoPoints"
    oStream.WriteBlankLines 2
        
    Set oGeomColl = oStrMfgProfileSupport.GetGeometries(STRMFG_AFTER_PROCESS2D, STRMFG_BENDING_LINE)
    lCounter = 1
    If oGeomColl.Count > 0 Then
        For Each oGeom2d In oGeomColl
            Dim sIBLstr As String
            oGeom2d.OutputXML PROFILE_TYPE, sIBLstr, , , , , , , , , , , 0.0001, 0.001, , 1000
            oStream.WriteLine "Regular XML output for Bending line no. " & lCounter
            oStream.WriteLine "**** {{{"
            oStream.WriteLine sIBLstr
            oStream.WriteLine "}}} ****" & vbCrLf
            
            Dim oIBLdoc As New DOMDocument
            If oIBLdoc.loadXML(sIBLstr) Then
                Dim oEdgeNode As IXMLDOMNode
                Set oEdgeNode = oIBLdoc.firstChild
                
                oStream.WriteLine "XML output for Bending line no. " & lCounter & " with Chord Height = 1e-4m"
                oStream.WriteLine "**** {{{"
                oStream.WriteLine ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oEdgeNode, MfgStroke_ChordHeight, 0.0001).xml
                oStream.WriteLine "}}} ****" & vbCrLf
                
                oStream.WriteLine "XML output for Bending line no. " & lCounter & " with Arc Length = 30mm"
                oStream.WriteLine "**** {{{"
                oStream.WriteLine ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oEdgeNode, MfgStroke_ArcLength, 0.03).xml
                oStream.WriteLine "}}} ****" & vbCrLf
                
                oStream.WriteLine "XML output for Bending line no. " & lCounter & " with X-distance = 30mm"
                oStream.WriteLine "**** {{{"
                oStream.WriteLine ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oEdgeNode, MfgStroke_Xdistance, 0.03).xml
                oStream.WriteLine "}}} ****" & vbCrLf
            Else
                oStream.WriteLine "Failed to generate valid XML document from above XML!"
            End If
            oStream.WriteBlankLines 1
            lCounter = lCounter + 1
            Set oIBLdoc = Nothing
        Next oGeom2d
        Set oGeom2d = Nothing
        
        ' Generate Full XMLs too.
        
        ' First set up a good prefix for the XML filenames
        Set oNamedItem = oMfgProfilePart
        
        Dim oMfgEntityHelper As IJMfgEntityHelper
        Set oMfgEntityHelper = New MfgEntityHelper
        
        Dim sCleanMfgPartName As String
        oMfgEntityHelper.ReplaceIllegalChars oNamedItem.Name, sCleanMfgPartName
        
        Dim sPrefixForXMLfiles As String
        sPrefixForXMLfiles = Left(strLogFile, Len(strLogFile) - 4) & "--" & sCleanMfgPartName & "--"
        
        Dim sDefaultXML As String
        Set oIBLdoc = oMfgProfilePart.GetXMLData2(sDefaultXML)
        
        ' Now create the XMLs - one for each input type
        Dim oXMLstream As TextStream
        Set oFSO = New FileSystemObject
        
        Set oXMLstream = oFSO.OpenTextFile(sPrefixForXMLfiles & "Default.xml", ForWriting, True)
        oXMLstream.Write sDefaultXML
        oXMLstream.Close
        
        Dim oNodeList As IXMLDOMNodeList
        Dim oNode As IXMLDOMNode
        
        Set oNodeList = oIBLdoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_bending_line']/SMS_EDGE")
        For Each oNode In oNodeList
            oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode, MfgStroke_ArcLength, 0.03)
            oNode.parentNode.removeChild oNode
        Next
        Set oNodeList = oIBLdoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_topflange_bending_line']/SMS_EDGE")
        For Each oNode In oNodeList
            oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode, MfgStroke_ArcLength, 0.03)
            oNode.parentNode.removeChild oNode
        Next
        Set oNodeList = oIBLdoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_bottomflange_bending_line']/SMS_EDGE")
        For Each oNode In oNodeList
            oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode, MfgStroke_ArcLength, 0.03)
            oNode.parentNode.removeChild oNode
        Next
        
        Set oXMLstream = oFSO.OpenTextFile(sPrefixForXMLfiles & "30mmArcLength.xml", ForWriting, True)
        oXMLstream.Write oIBLdoc.xml
        oXMLstream.Close
        
        Set oIBLdoc = New DOMDocument
        oIBLdoc.loadXML sDefaultXML
        
        Set oNodeList = oIBLdoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_bending_line']/SMS_EDGE")
        For Each oNode In oNodeList
            oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode, MfgStroke_ChordHeight, 0.0001)
            oNode.parentNode.removeChild oNode
        Next
        Set oNodeList = oIBLdoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_topflange_bending_line']/SMS_EDGE")
        For Each oNode In oNodeList
            oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode, MfgStroke_ChordHeight, 0.0001)
            oNode.parentNode.removeChild oNode
        Next
        Set oNodeList = oIBLdoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_bottomflange_bending_line']/SMS_EDGE")
        For Each oNode In oNodeList
            oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode, MfgStroke_ChordHeight, 0.0001)
            oNode.parentNode.removeChild oNode
        Next
        
        Set oXMLstream = oFSO.OpenTextFile(sPrefixForXMLfiles & "1e-4ChordHeight.xml", ForWriting, True)
        oXMLstream.Write oIBLdoc.xml
        oXMLstream.Close
        
        Set oIBLdoc = New DOMDocument
        oIBLdoc.loadXML sDefaultXML
        
        Set oNodeList = oIBLdoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_bending_line']/SMS_EDGE")
        For Each oNode In oNodeList
            oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode, MfgStroke_Xdistance, 0.03)
            oNode.parentNode.removeChild oNode
        Next
        Set oNodeList = oIBLdoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_topflange_bending_line']/SMS_EDGE")
        For Each oNode In oNodeList
            oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode, MfgStroke_Xdistance, 0.03)
            oNode.parentNode.removeChild oNode
        Next
        Set oNodeList = oIBLdoc.selectNodes("//SMS_MARKING[@TYPE='strmfg_bottomflange_bending_line']/SMS_EDGE")
        For Each oNode In oNodeList
            oNode.parentNode.appendChild ConvertEdgeNodeWithArcsIntoEdgeNodeWithPoints(oNode, MfgStroke_Xdistance, 0.03)
            oNode.parentNode.removeChild oNode
        Next
        
        Set oXMLstream = oFSO.OpenTextFile(sPrefixForXMLfiles & "30mmXdistance.xml", ForWriting, True)
        oXMLstream.Write oIBLdoc.xml
        oXMLstream.Close
        
        Set oFSO = Nothing
        
    Else
        oStream.WriteLine " There are no bending curves to test."
    End If
    Set oGeomColl = Nothing
    oStream.WriteLine ""
    oStream.WriteLine ""
    
    oStream.WriteLine ""
    oStream.WriteLine "============================End of report ==========="
    

    'Close Log File
    oStream.Close

CleanUp:
  
    Set oMfgProfilePart = Nothing
    Set oNamedItem = Nothing
    Set oSDPartWrapper = Nothing
 

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

