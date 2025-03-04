VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RefMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project: MfgTemplateMarking
'
'  Abstract:    Creates Reference marks on Template
'
'  History:
'      Siva        2nd September 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingPlate.RefMarks"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler

    Dim oMfgGeomChild   As IJMfgGeomChild
    Set oMfgGeomChild = oMfgTemplate
    
    Dim oMfgTemplateSetRep  As IJMfgTemplateSetReport
    Set oMfgTemplateSetRep = oMfgGeomChild.GetParent
    
    Dim oPlatePart  As Object
    Set oPlatePart = oMfgTemplateSetRep.GetParentPlatePart
    
    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = oPlatePart
    
    ' If the part is of hull type, reference marks rule will do
    If oSDPlateWrapper.plateType = Hull Then
        Exit Function
    End If
    
    Dim oMfgRuleHelper  As MfgRuleHelpers.Helper
    Set oMfgRuleHelper = New Helper
    
    Dim oMfgTemplateSet  As IJDMfgTemplateSet
    Set oMfgTemplateSet = oMfgGeomChild.GetParent
    
    Dim oProcSettings As IJMfgTemplateProcessSettings
    Set oProcSettings = oMfgTemplateSet.GetProcessSettings
        
    Dim strTemplateSide As String
    strTemplateSide = oProcSettings.TemplateSide

    Dim oReferenceCurvesCol  As Collection
    
    If strTemplateSide = "BaseSide" Then
        oSDPlateWrapper.ReferenceCurves PlateBaseSide, oReferenceCurvesCol
    Else
        oSDPlateWrapper.ReferenceCurves PlateOffsetSide, oReferenceCurvesCol
    End If
    
    If oReferenceCurvesCol Is Nothing Then
        'Since there is nothing to be marked we can leave the function
        GoTo CleanUp
    End If
    
    Dim oTemplateBottomCurve   As IJCurve
    Set oTemplateBottomCurve = oMfgTemplate.GetTemplateLocationMarkLine
    
    If oTemplateBottomCurve Is Nothing Then
        Exit Function
    End If
    
    Dim oGeomElems As IJElements
    Set oGeomElems = New JObjectCollection
    
    Dim nRefCurveIndex As Long, nMarkIndex  As Long
    
    For nRefCurveIndex = 1 To oReferenceCurvesCol.Count
        
        Dim oRefCurveData As IJRefCurveData
        Set oRefCurveData = oReferenceCurvesCol.Item(nRefCurveIndex)
        
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
            
            Dim oMarkLines As Collection
            Set oMarkLines = oRefCurveData.GetMarkingLineCollection
            For nMarkIndex = 1 To oMarkLines.Count
                
                Dim oWB As IJWireBody
                Set oWB = oMarkLines.Item(nMarkIndex)
                
                If Not oWB Is Nothing Then
                    'Convert the IJWireBody to a IJComplexString
                    Dim oCSColl As IJElements
                    Set oCSColl = oMfgRuleHelper.WireBodyToComplexStrings(oWB)
                    If Not oCSColl Is Nothing Then
                        If oCSColl.Count = 0 Then
                            GoTo NextMarking
                        End If
                    Else
                        GoTo NextMarking
                    End If
                    
                    Dim oRefCS  As IJComplexString
                    For Each oRefCS In oCSColl
                    
                        Dim dMinDist As Double, dDummy As Double, dX As Double, dY As Double, dZ As Double
                        oTemplateBottomCurve.DistanceBetween oRefCS, dMinDist, dX, dY, dZ, dDummy, dDummy, dDummy
                        
                        ' Check the distance between reference curve and the template bottom line
                        If dMinDist < 0.001 Then
                        
                            Dim oMarkPosition   As IJDPosition
                            Set oMarkPosition = New DPosition
                            
                            oMarkPosition.Set dX, dY, dZ
                            
                            Dim oMarkVec    As IJDVector
                            Set oMarkVec = GetMarkVector(oMfgTemplate, oTemplateBottomCurve, oMarkPosition)
                            
                            Dim oMarkCS As IJComplexString
                            Set oMarkCS = CreateMarkAtPosition(oMarkPosition, oMarkVec, TEMPLATE_KNUCKLE_MARK_LENGTH, True)
                            
                            Dim oGeom3D As IJMfgGeom3d
                            Set oGeom3D = CreateGeom3D(oMarkCS, eStrMfgGeomType, , oRefCurveData.Name)
                            
                            oGeomElems.Add oGeom3D
                        End If
                    Next
                End If
NextMarking:
            Next nMarkIndex

            Set oMarkLines = Nothing
            Set oRefCurveData = Nothing
        End If
NextRefCurve:
    Next nRefCurveIndex
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oGeomElems
    
CleanUp:
    Exit Function
    
ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3020, , "RULES")
End Function
