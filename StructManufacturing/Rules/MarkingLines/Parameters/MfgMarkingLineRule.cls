VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MfgMarkingLineRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2009, Intergraph Corp.  All rights reserved.
'
' Project:  StrMfgMarkingParRule.vbp
' Module:   MfgMarkingLineRule.cls
' Description:  Rule for creating Marking Lines by Marking Line command.
'
' Comments:
'   History :
'   Suma Mallena        28-MAR-2009     Creation

'*******************************************************************************
Option Explicit
Const MODULE = "MfgMarkingLineRule:"

''''''''''''''   Definition of Interfaces used for Filter Criteria.''''''''''''''
'- Plate ------------------------------------------------------------'
Private Const IJPlateSystem = "{E0B23CD4-7CEB-11d3-B351-0050040EFC17}"
'- Profile ----------------------------------------------------------'
Private Const IJStiffenerSystem = "{E0B23CD5-7CEB-11d3-B351-0050040EFC17}"
'- PlatePart --------------------------------------------------------'
Private Const IJPlatePart = "{780F26C2-82E9-11D2-B339-080036024603}"
'- ProfilePart ------------------------------------------------------'
Private Const IJProfilePart = "{69F3E7BF-40A0-11D2-B324-080036024603}"
'- IJRefCurveOnSurface --------------------------------------------'
Private Const IJRefCurveOnSurface = "{EBE54C96-77B1-11D5-8A1D-00C04F79B54E}"
'- Seam -------------------------------------------------------------'
Private Const IJSeam = "{02C1327F-2C31-11D2-8329-0800367F3D03}"

Implements IJDMarkingFilter


'****************************************************************************************************************************
'Method:
'    GetCriteria
'       Inputs: Input parts, Command Step, Marking Type.
'       Output: Filter criteria as a string.
'Description:
'   Depending upon the command step (Reference Curve step or Related Part step), the filter criteria can be controlled for marking lines of various types.

'****************************************************************************************************************************

Private Function IJDMarkingFilter_GetCriteria(ByVal pInputParts As IJElements, ByVal eFilterStep As GSCADMfgRulesDefinitions.EnumFilterSteps, ByVal lMarkingType As Long) As String
Const METHOD = "IJDMarkingFilter_GetCriteria "
    On Error GoTo ErrorHandler
    
    Dim strFilterCriteria As String
    
    Select Case eFilterStep
    
        Case ReferenceStep
        
            Select Case lMarkingType
            
                Case STRMFG_PLATELOCATION_MARK, _
                     STRMFG_PROFILELOCATION_MARK
                
                    strFilterCriteria = IJRefCurveOnSurface & _
                                        " OR " & "[StrMfgMarkingLinesCmd.CMarkingFilter,IsLogicalConnection]" & _
                                        " OR " & "[StrMfgMarkingLinesCmd.CMarkingFilter,IsEdgePort]"
                Case STRMFG_MARGIN_MARK
                
                    strFilterCriteria = IJRefCurveOnSurface & _
                                        " OR " & "[StrMfgMarkingLinesCmd.CMarkingFilter,IsLogicalConnection]" & _
                                        " OR " & "[StrMfgMarkingLinesCmd.CMarkingFilter,IsEdgePort]"
                Case STRMFG_PINJIG_MARK
                    
                    strFilterCriteria = "[StrMfgMarkingLinesCmd.CMarkingFilter,IsLogicalConnection]"
                    
                Case Else
                
                    strFilterCriteria = IJRefCurveOnSurface & _
                                        " OR " & "[StrMfgMarkingLinesCmd.CMarkingFilter,IsLogicalConnection]" & _
                                        " OR " & "[StrMfgMarkingLinesCmd.CMarkingFilter,IsEdgePort]"
                    
            End Select
        
        Case RelatedPartStep
            
            Select Case lMarkingType
            
                Case STRMFG_PLATELOCATION_MARK
                
                    strFilterCriteria = IJPlateSystem & " OR " & IJPlatePart
  
                                        
                Case STRMFG_PROFILELOCATION_MARK
                
                    strFilterCriteria = IJStiffenerSystem & " OR " & IJProfilePart
                
                Case STRMFG_PINJIG_MARK
                    
                    strFilterCriteria = IJPlateSystem & _
                                        " OR " & IJPlatePart & _
                                        " OR " & IJStiffenerSystem & _
                                        " OR " & IJProfilePart & _
                                        " OR " & IJRefCurveOnSurface & _
                                        " OR " & IJSeam

                Case Else
                
                    strFilterCriteria = IJPlateSystem & _
                                        " OR " & IJPlatePart & _
                                        " OR " & IJStiffenerSystem & _
                                        " OR " & IJProfilePart & _
                                        " OR " & IJRefCurveOnSurface & _
                                        " OR " & "[StrMfgMarkingLinesCmd.CMarkingFilter,IsLogicalConnection]" & _
                                        " OR " & "[StrMfgMarkingLinesCmd.CMarkingFilter,IsEdgePort]" & _
                                        " OR " & IJSeam
            
            End Select

    
    End Select
    
    IJDMarkingFilter_GetCriteria = strFilterCriteria

Exit Function

ErrorHandler:
 
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4034, , "RULES")

End Function

'****************************************************************************************************************************
'Method:
'    GetCustomAttributes
'       Inputs: Marking Lines, Reference Curve, Related Part, Marking Type.
'       Output: Collection of Attributes. Each attribute is a pair of "Attribute Name" and it's "Value"
'Description:
'   The values for the Custom Attributes on IJMfgSketchLocation can be controlled here.
'   Depending on the Type of Marking Line, user can fill the collection with the "Name of the attribute" and it's "Value".
'   If the "Value" is filled as "Nothing", then an empty value is stored with the attribute.
'   If no attribute is added to the collection, then the system evaluates the missing attributes and populates with system generated default value.
'****************************************************************************************************************************
Private Function IJDMarkingFilter_GetCustomAttributes(ByVal pMarkingLines As IJElements, ByVal pDispReferencePart As Object, ByVal pDispRelatedPart As Object, ByVal lMarkingType As Long) As Object
Const METHOD = "IJDMarkingFilter_GetCustomAttributes"
    On Error GoTo ErrorHandler

    ' Instead of using a collection we have to use a dictionary which can give a name value pair.
    ' The normal VB Collection doesn't have the capability to return the name value pair as needed for this rule

    Dim oCollection As Dictionary
    Set oCollection = New Dictionary

    Select Case lMarkingType

        Case STRMFG_PLATELOCATION_MARK

            oCollection.Add "Direction", "unknown"

        Case STRMFG_PROFILELOCATION_MARK

            oCollection.Add "Direction", "unknown"

        Case STRMFG_MARGIN_MARK

            oCollection.Add "MaxAssyMarginValue", 0.02
            oCollection.Add "MaxFabMarginValue", 0.03
            oCollection.Add "MaxCustomMarginValue", 0.04

            oCollection.Add "Direction", Nothing 'This attribute value is set to empty


    End Select

    Set IJDMarkingFilter_GetCustomAttributes = oCollection

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4034, , "RULES")
End Function

'****************************************************************************************************************************
'Method:
'    GetMarkingTypes
'       Inputs: Marking Lines,Input parts
'       Output: Colection of Marking Types that will be shown in the MarkingType drop-down in the ribbonbar of Marking Line command
'Description:
'   Marking Type as LONG value is added to the output collection. this number should match with the code-list number defined in
'   \StructManufacturing\Data\bulkload\AllStructMfgCodeLists.xls

'****************************************************************************************************************************

Private Function IJDMarkingFilter_GetMarkingTypes(ByVal pMarkingLines As IJElements, ByVal pInputParts As IJElements) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
Const METHOD = "IJDMarkingFilter_GetMarkingTypes "
    On Error GoTo ErrorHandler

    Dim oCollection  As New Collection
    Dim strMarkingLinesInputType As String
    
    If TypeOf pInputParts.Item(1) Is IJPlatePart Then
        strMarkingLinesInputType = "Plate"
    ElseIf TypeOf pInputParts.Item(1) Is IJProfilePart Then
        strMarkingLinesInputType = "Profile"
    ElseIf TypeOf pInputParts.Item(1) Is ISPSMemberPartPrismatic Then
        strMarkingLinesInputType = "Member"
    Else
        strMarkingLinesInputType = "PinJig"
    End If
             
    Select Case (strMarkingLinesInputType)
        Case "Plate"
            
            oCollection.Add STRMFG_BLOCK_MARK                '"Block Mark"
            oCollection.Add STRMFG_PAINTING_MARK             '"Painting Mark"
            oCollection.Add STRMFG_NAME_MARK                 '"Name Mark"
            oCollection.Add STRMFG_ROBOT_MARK                '"Robot Mark"
            oCollection.Add STRMFG_USERDEFINED_MARK          '"User Defined Mark"
            
            oCollection.Add STRMFG_PLATELOCATION_MARK        '"Plate Location Mark"
            oCollection.Add STRMFG_PROFILELOCATION_MARK      '"Profile Location Mark"
            
            oCollection.Add STRMFG_PLATE_TO_PLATE_TJOINT_MARK '"Plate to Plate Tjoint Mark"
            oCollection.Add STRMFG_PROFILE_TO_PLATE_MARK      '"Profile To Plate Mark"
            
            oCollection.Add STRMFG_SEAM_MARK                  '"SeamControlMark"
            oCollection.Add STRMFG_ROLL_LINES_MARK            '"RollLineMark"
            oCollection.Add STRMFG_ROLL_BOUNDARIES_MARK       '"RollBoundaryMark"
            oCollection.Add STRMFG_KNUCKLE_MARK               '"Knuckle Line Mark"
            oCollection.Add STRMFG_NAVALARCHLINE              '"Naval Arch Ref Line"
            oCollection.Add STRMFG_DIRECTION                  '"ShipDirectionMark"
            
            oCollection.Add STRMFG_PINJIG_MARK                '"Pinjig Marking"
            
            oCollection.Add STRMFG_BASELINE_MARK             '"Base Control Line Mark"
            oCollection.Add STRMFG_BENDING_CONTROLLINES_MARK '"Bending Control Line Mark"
            oCollection.Add STRMFG_BRACKETLOCATION_MARK      '"Bracket Location Mark"
'            oCollection.Add STRMFG_BRIDGE_MARK               '"Bridge Mark"
            oCollection.Add STRMFG_BUTTOCKLINE_MARK          '"Buttock Line Mark"

            oCollection.Add STRMFG_COLLARPLATELOCATION_MARK  '"Collar Plate Loc Mark"
            oCollection.Add STRMFG_DIAGONALS_MARK            '"Diagonal Mark"
            oCollection.Add STRMFG_END_MARK                  '"End Connection Mark"
            oCollection.Add STRMFG_FRAMELINE_MARK            '"Frameline Mark"
'
            oCollection.Add STRMFG_LAP_MARK                  '"Lap Connection Mark"
'            oCollection.Add STRMFG_MARGIN_MARK               '"Margin Mark"
'            oCollection.Add STRMFG_NAME_MARK                 '"Mark Ship Name"
            oCollection.Add STRMFG_EDGE_CHECKLINES_MARK      '"Plate Edge Check Mark"
'
'            oCollection.Add STRMFG_FEATURE_MARK               '"Plate Feature Tab"
'            oCollection.Add STRMFG_WELDTAB_MARK               '"Plate Weld Tab"

            oCollection.Add STRMFG_TEMPLATE_MARK              '"TemplatePosMark"
            oCollection.Add STRMFG_WATERLINE_MARK             '"WaterLineMark"
            oCollection.Add STRMFG_FRAMES_CHECKLINES_MARK     '"Frame Check Lines Mark"
            
            oCollection.Add STRMFG_MOUNT_ANGLE_MARK             '"Mount Angle Mark"
            oCollection.Add STRMFG_OPENING_MARK                 '"Opening Mark"
            oCollection.Add STRMFG_CUTTING_REF_MARK             '"Cutting Ref Mark"
            oCollection.Add STRMFG_GEOMETRY_CUT_MARK            '"Geometry Cut Mark"
            oCollection.Add STRMFG_ENDCUT_MARK                  '"End Cut Mark"
            oCollection.Add STRMFG_HOLE_TRACE_MARK              '"Hole Trace Mark"
            oCollection.Add STRMFG_HOLE_REF_MARK                '"Hole Ref Mark"
            
            oCollection.Add STRMFG_SEAM_BEND_MARK               '"Seam Bend Mark"
            oCollection.Add STRMFG_COMMON_SEAM_MARK             '"Common Seam Mark"
            oCollection.Add STRMFG_EDGE_LOCATION_MARK           '"Edge Location Mark"
            
            oCollection.Add STRMFG_CONN_PART_MARK               '"Conn Part Mark"
            oCollection.Add STRMFG_KNUCKLE_TAB_MARK             '"Knuckle Tab Mark"
            oCollection.Add STRMFG_TAB_REFERENCE_MARK          '"Tab Reference Mark"
            oCollection.Add STRMFG_REFERENCE_X_MARK             '"Reference X Mark"
            oCollection.Add STRMFG_REFERENCE_Y_MARK             '"Reference Y Mark"
            oCollection.Add STRMFG_REFERENCE_Z_MARK             '"Reference Z Mark"
            oCollection.Add STRMFG_REFERENCE_CENTER_MARK        '"Reference Center Mark"
            
            oCollection.Add STRMFG_LAP_TRACE_MARK               '"Lap Trace Mark"
            oCollection.Add STRMFG_LAP_REF_MARK                 '"Lap Ref Mark"
            oCollection.Add STRMFG_LAP_EDGE_MARK                '"Lap Edge Mark"
            oCollection.Add STRMFG_ROLL_SECTION_REF_MARK        '"Roll Section Ref Mark"
            oCollection.Add STRMFG_ELONGATION_MARK              '"Elongation Mark"
    
        Case "Profile"
         
            oCollection.Add STRMFG_BLOCK_MARK                '"Block Mark"
            oCollection.Add STRMFG_PAINTING_MARK             '"Painting Mark"
            oCollection.Add STRMFG_NAME_MARK                 '"Name Mark"
            oCollection.Add STRMFG_ROBOT_MARK                '"Robot Mark"
            oCollection.Add STRMFG_USERDEFINED_MARK          '"User Defined Mark"
            
            oCollection.Add STRMFG_END_MARK                  '"End Connection Mark"
            oCollection.Add STRMFG_SEAM_MARK                 '"SeamControlMark"
            oCollection.Add STRMFG_FITTING_MARK              '"FittingMark"
            oCollection.Add STRMFG_PLATELOCATION_MARK        '"Plate Location Mark"
            oCollection.Add STRMFG_KNUCKLE_MARK              '"Knuckle Line Mark"
            
            oCollection.Add STRMFG_MOUNT_ANGLE_MARK             '"Mount Angle Mark"
            oCollection.Add STRMFG_OPENING_MARK                 '"Opening Mark"
            oCollection.Add STRMFG_CUTTING_REF_MARK             '"Cutting Ref Mark"
            oCollection.Add STRMFG_GEOMETRY_CUT_MARK            '"Geometry Cut Mark"
            oCollection.Add STRMFG_ENDCUT_MARK                  '"End Cut Mark"
            oCollection.Add STRMFG_HOLE_TRACE_MARK              '"Hole Trace Mark"
            oCollection.Add STRMFG_HOLE_REF_MARK                '"Hole Ref Mark"
            
            oCollection.Add STRMFG_PROFILE_TO_PROFILE_PENETRATION_MARK   '"Profile To Profile Penetration Mark"
            oCollection.Add STRMFG_EDGE_LOCATION_MARK           '"Edge Location Mark"
            oCollection.Add STRMFG_CONN_PART_MARK               '"Conn Part Mark"
            
            oCollection.Add STRMFG_DIRECTION                  '"ShipDirectionMark"
            oCollection.Add STRMFG_LAP_MARK                   '"Lap Connection Mark"
            oCollection.Add STRMFG_WATERLINE_MARK             '"WaterLineMark"
            oCollection.Add STRMFG_FRAMELINE_MARK             '"Frameline Mark"
            oCollection.Add STRMFG_BUTTOCKLINE_MARK           '"Buttock Line Mark"
            oCollection.Add STRMFG_BUILTUP_CONNECTION_MARK    ' "BuiltUp Connection Mark"
        
        Case "Member"
         
            oCollection.Add STRMFG_BLOCK_MARK                '"Block Mark"
            oCollection.Add STRMFG_PAINTING_MARK             '"Painting Mark"
            oCollection.Add STRMFG_NAME_MARK                 '"Name Mark"
            oCollection.Add STRMFG_ROBOT_MARK                '"Robot Mark"
            oCollection.Add STRMFG_USERDEFINED_MARK          '"User Defined Mark"
            
'            oCollection.Add STRMFG_END_MARK                  '"End Connection Mark"
'            oCollection.Add STRMFG_SEAM_MARK                 '"SeamControlMark"
'            oCollection.Add STRMFG_FITTING_MARK              '"FittingMark"
'            oCollection.Add STRMFG_PLATELOCATION_MARK        '"Plate Location Mark"
'            oCollection.Add STRMFG_KNUCKLE_MARK              '"Knuckle Line Mark"
            
            
        Case "PinJig"
            
            oCollection.Add STRMFG_EXTEND_PINJIG_INTERSECTION    '"Extend PinJig Intersection"
            oCollection.Add STRMFG_PINJIG_MARK                   '"Pinjig Marking"
            oCollection.Add STRMFG_PINJIG_DIAGONAL               '"Pinjig Diagonal"

    End Select


    Set IJDMarkingFilter_GetMarkingTypes = oCollection


 Exit Function


ErrorHandler:
   
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 4034, , "RULES")

End Function

