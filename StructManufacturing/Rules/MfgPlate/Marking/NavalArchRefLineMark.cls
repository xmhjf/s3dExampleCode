VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NavalArchRefLineMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Naval Architecture Reference Line Markings
'               in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH        feb. 6. 2002    created
'      ssreedha   april 8 2004    Fix for TR-CP·46915  Markings that should not be sent to XML are being sent for shell plates
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.NavalArchRefLineMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub


Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "NavalArchRefLineMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

    CreateAPSMarkings STRMFG_NAVALARCHLINE, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

    If bSelectiveRecompute Then
        Exit Function
    End If
    
    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr

    Dim oMfgPart As IJMfgPlatePart
    Set oPlateWrapper.object = Part
    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part
    
    If Not oSDPlateWrapper.plateType = Hull Then Exit Function

    'Get the Naval Arch reference curves that intersect the part
    Dim oReferenceCurvesCol  As Collection
    Dim nIndex As Long
    Dim nIndex1 As Long

    ' Always use base side to fix TR 35918
    oSDPlateWrapper.ReferenceCurves PlateBaseSide, oReferenceCurvesCol

    If oReferenceCurvesCol Is Nothing Then
        'Since there is nothing to be marked we can leave the function
        GoTo CleanUp
    End If

    'Can 't use "For Each" when user defined types in collection
    For nIndex = 1 To oReferenceCurvesCol.Count
        Dim oWB As IJWireBody
        Dim oCSColl As IJElements
        Dim oSystemMark As IJMfgSystemMark
        Dim oMarkingInfo As MarkingInfo
        Dim oGeom3dLine As IJMfgGeom3D
        Dim oCS As IJComplexString
        Dim oMarkLines As Collection
        Dim oRefCurveData As IJRefCurveData

        Set oRefCurveData = oReferenceCurvesCol.Item(nIndex)
        If Not oRefCurveData Is Nothing Then
            Dim RefCurveType As JSRCOS_TYPES
            RefCurveType = oRefCurveData.Type
            Dim eStrMfgGeomType As StrMfgGeometryType
            Select Case RefCurveType
                Case JSRCOS_KNUCKLE
                    eStrMfgGeomType = STRMFG_KNUCKLE_MARK
                Case JSRCOS_REFERENCE
                    eStrMfgGeomType = STRMFG_GENERAL_MARK
                Case JSRCOS_TANGENT
                    'As the Curve type for Both Flat_of_Bottom  and Flat_of_Side Ref curves will
                    'be same ie JSRCOS_TANGENT, so inorder to determine the reference curve marking type,
                    'Use Reference Curve Name for checking the types
                    'Fix for TR-CP·46915  Markings that should not be sent to XML are being sent for shell plates
                    'S.V.Suresh - 08 April 2004
                    
                    'Search for String "Flat_of_bottom" in Curve name from start in text compare mode
                    Dim iPos As Integer
                    iPos = InStr(1, oRefCurveData.Name, "Flat_of_Bottom", vbTextCompare)
                    If iPos = 1 Then 'ie at the start of string
                        eStrMfgGeomType = STRMFG_FLAT_OF_BOTTOM_MARK
                    Else
                        iPos = InStr(1, oRefCurveData.Name, "Flat_of_Side", vbTextCompare)
                        If iPos = 1 Then
                            eStrMfgGeomType = STRMFG_FLAT_OF_SIDE_MARK
                        Else
                            eStrMfgGeomType = STRMFG_GENERAL_MARK
                        End If
                    End If
                Case JSRCOS_UNKNOWN 'Unknown case put the Geometry Type to "General"
                    eStrMfgGeomType = STRMFG_GENERAL_MARK
                Case Else
                    GoTo NextRefCurve
            End Select
            
            Set oMarkLines = oRefCurveData.GetMarkingLineCollection
            For nIndex1 = 1 To oMarkLines.Count
                Set oWB = oMarkLines.Item(nIndex1)
                If Not oWB Is Nothing Then
                    'Convert the IJWireBody to a IJComplexString
                    Set oCSColl = m_oMfgRuleHelper.WireBodyToComplexStrings(oWB)
                    If Not oCSColl Is Nothing Then
                        If oCSColl.Count = 0 Then
                            Set oCSColl = Nothing
                        End If
                    End If
                        
                    If (oCSColl Is Nothing) Then
                        GoTo NextMarking
                    End If
                    
                    For Each oCS In oCSColl
                        
                        'Create a SystemMark object to store additional information
                        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                        
                        'Set the marking side
                        oSystemMark.SetMarkingSide UpSide

                        'QI for the MarkingInfo object on the SystemMark
                        Set oMarkingInfo = oSystemMark
                        oMarkingInfo.Name = oRefCurveData.Name
                                                                        
                        Set oGeom3dLine = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                        oGeom3dLine.PutGeometry oCS
                        oGeom3dLine.PutGeometrytype eStrMfgGeomType
                                                
                        'ApplyDirection oSystemMark, oConnectionData.ConnectingPort
                        oSystemMark.Set3dGeometry oGeom3dLine
                        oGeomCol3d.AddGeometry 1, oGeom3dLine
                    Next
                End If
NextMarking:
                Set oWB = Nothing
                Set oCS = Nothing
                If Not oCSColl Is Nothing Then
                    oCSColl.Clear
                    Set oCSColl = Nothing
                End If
                Set oSystemMark = Nothing
                Set oMarkingInfo = Nothing
            Next nIndex1

            Set oMarkLines = Nothing
            Set oRefCurveData = Nothing
        End If
NextRefCurve:
        ' We shouldn't remove the current object from the collection.
        'oReferenceCurvesCol.Remove nIndex

    Next nIndex

    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d

CleanUp:
    Set oMfgPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oSDPlateWrapper = Nothing
    Set oGeomCol3d = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1015, , "RULES")
    GoTo CleanUp
End Function
 
