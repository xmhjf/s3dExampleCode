Attribute VB_Name = "Helper"
Option Explicit

Private Const Module = "MfgHelper"
Public m_bLogError As Boolean

'Descrption: To Display Parent objects like ParentPart, Assembly and Block.
'Logic:
' The Input is the selected object. Based on the type of input and according to
' Specification, the Parent String has been Selected.
Public Sub DisplayParentObjectsInformation(oSelectedObject As Object, oStream As TextStream)
    Const METHODNAME = "DisplayParentObjectsInformation"
    On Error GoTo ErrorHandler
    
    Dim oMfgChild           As IJMfgChild
    Dim oParent             As IJMfgParent
    Dim strParent           As String
    Dim bParentRequired     As Boolean
    Dim oNamedItem          As IJNamedItem

    Set oMfgChild = oSelectedObject
    Set oParent = oMfgChild.getParent
    
    If TypeOf oParent Is IJMfgMarkingFolder Then
        Set oMfgChild = oParent
        Set oParent = oMfgChild.getParent
    End If
    
    bParentRequired = True

'    Conditional loop to Set the Parent String in the Relationship Tab
    If TypeOf oSelectedObject Is IJMfgPlatePart Then
        strParent = "PLATE_NAME"
    ElseIf TypeOf oSelectedObject Is IJMfgProfilePart Then
        strParent = "PROFILE_NAME"
    ElseIf TypeOf oParent Is IJPlanningAssembly Then
        strParent = "PARENT_ASSEMBLY_NAME"
    Else
        If TypeOf oSelectedObject Is IJPinJig Then
            bParentRequired = False
        Else
            strParent = "PARENT"
        End If
    End If
    
    If bParentRequired Then
        Set oNamedItem = oParent
        oStream.WriteLine " " & strParent & "  : " & oNamedItem.Name
    End If
    
    Set oMfgChild = oParent
    Set oParent = oMfgChild.getParent
    
    'To Check if the Parent of the Detailed Plate/Profile is not Nothing
    If (Not (oParent Is Nothing)) Then
        ''  Get the Assembly
        '   if condition to skip a case where plate part is directly under Block.
        If TypeOf oParent Is IJAssembly And Not (TypeOf oParent Is IJBlock) Then
            Set oNamedItem = oParent
            oStream.WriteLine " ASSEMBLY_NAME" & "  : " & oNamedItem.Name
            '   if condition for a case where the parent is the root project,
            '   which does not support IJMfgChild interface
            If TypeOf oParent Is IJMfgChild Then
                Set oMfgChild = oParent
                Set oParent = oMfgChild.getParent
            End If
        End If
        
        ''  Get the Block
        '   if condition to check if the parent at this point is config root
        If (Not (oParent Is Nothing)) And TypeOf oParent Is IJMfgChild Then
            '   Loop to navigate to the first encountered block
            While Not TypeOf oParent Is IJBlock
                Set oMfgChild = oParent
                Set oParent = oMfgChild.getParent
            Wend
            Set oNamedItem = oParent
            oStream.WriteLine " BLOCK_NAME" & "  : " & oNamedItem.Name
        End If
    End If

Cleanup:
    Set oMfgChild = Nothing
    Set oParent = Nothing
    Set oNamedItem = Nothing
    Exit Sub
ErrorHandler:
    oStream.WriteLine METHODNAME & "Error number: " & Err.Number & vbCrLf & _
                    "Error description: " & Err.Description
    GoTo Cleanup
End Sub
'***********************************************************************
' METHOD: GetSystemMarkingLines
'
' DESCRIPTION: Returns three collections of marking lines based on
'               location (web, top flange, or bottom flange)
'***********************************************************************
Public Sub GetMarkingLines(ByVal oPart As Object, _
                                oWebMLs As Collection, _
                                oTopFlangeMLs As Collection, _
                                oBtmFlangeMLs As Collection)
    Const METHOD = "GetMarkingLines"
    On Error GoTo ErrorHandler
    
    Dim oMfgProfilePart As IJMfgProfilePart
    Dim oMfgGeomCol2d  As IJMfgGeomCol2d
    Dim lCount As Long
    Dim lFaceID As Long
    Dim oGeom2d As IJMfgGeom2d
    Dim oMoniker As IMoniker
    Dim eStrMfgGeomType As StrMfgGeometryType
    Dim oObjectUnk As IUnknown
    Dim lIndex As Long
    
    
    Set oMfgProfilePart = oPart
    Set oMfgGeomCol2d = oMfgProfilePart.FinalGeometriesAfterProcess2D
    Set oMfgProfilePart = Nothing
        
    If oMfgGeomCol2d.Getcount > 0 Then
        For lCount = 1 To oMfgGeomCol2d.Getcount
            Set oGeom2d = oMfgGeomCol2d.GetGeometry(lCount)
            
            eStrMfgGeomType = oGeom2d.GetGeometryType
            lFaceID = oGeom2d.FaceId
            
            If eStrMfgGeomType = STRMFG_BLOCK_MARK _
               Or eStrMfgGeomType = STRMFG_ROBOT_MARK _
               Or eStrMfgGeomType = STRMFG_USERDEFINED_MARK _
               Or eStrMfgGeomType = STRMFG_BRACKETLOCATION_MARK _
               Or eStrMfgGeomType = STRMFG_FRAMELINE_MARK Then
                        
               On Error Resume Next
               Set oMoniker = oGeom2d.GetMoniker
               Set oObjectUnk = BindToObject(oMoniker)
                
               'We're adding all the marking objects in order to the first half of the
               '    collection, and a string representing their corresponding GeomTypes
               '    to the second half of the same collection.
               Select Case lFaceID
                    Case JXSEC_TOP
                        If 0 = oTopFlangeMLs.Count Then
                            oTopFlangeMLs.Add oObjectUnk
                        Else
                            lIndex = oTopFlangeMLs.Count / 2
                            oTopFlangeMLs.Add oObjectUnk, , , after:=lIndex
                        End If
                        lIndex = oTopFlangeMLs.Count
                        oTopFlangeMLs.Add GetGeometryTypeName(eStrMfgGeomType), , , after:=lIndex
                    Case JXSEC_BOTTOM
                        If 0 = oBtmFlangeMLs.Count Then
                            oBtmFlangeMLs.Add oObjectUnk
                        Else
                            lIndex = oBtmFlangeMLs.Count / 2
                            oBtmFlangeMLs.Add oObjectUnk, , , after:=lIndex
                        End If
                        lIndex = oBtmFlangeMLs.Count
                        oBtmFlangeMLs.Add GetGeometryTypeName(eStrMfgGeomType), , , after:=lIndex
                    Case Else
                        If 0 = oWebMLs.Count Then
                            oWebMLs.Add oObjectUnk
                        Else
                            lIndex = oWebMLs.Count / 2
                            oWebMLs.Add oObjectUnk, , , after:=lIndex
                        End If
                        lIndex = oWebMLs.Count
                        oWebMLs.Add GetGeometryTypeName(eStrMfgGeomType), , , after:=lIndex
               End Select
                
               Set oObjectUnk = Nothing
               Set oMoniker = Nothing
            End If
            eStrMfgGeomType = -1
            lFaceID = -1
            Set oGeom2d = Nothing
        Next lCount
    End If
    Set oMfgGeomCol2d = Nothing
     
    Exit Sub
ErrorHandler:
    Set oMfgProfilePart = Nothing
    Set oMfgGeomCol2d = Nothing
    Set oGeom2d = Nothing
    Set oMoniker = Nothing
    Set oObjectUnk = Nothing
End Sub

''**************************************************************************************
'' Routine      : GetMarkingLineGeom2D
'' Description  : Returns the corresponding Geometry 2D object for the Marking Line
''
''**************************************************************************************
Private Function GetMarkingLineGeom2D(oSystemMarkingLine As Object) As IJMfgGeom2d
    Const METHOD = "GetMarkingLineGeom2D"

    Const IID_IJMfgSystemMark As String = "{FB6CADE6-782A-4E5D-8C05-BF8D7D33806C}"
    
    Dim oAssocRel As IJDAssocRelation
    Dim oUnkColl As Object
    Dim oTargetObjCol As IJDTargetObjectCol

    Set oAssocRel = oSystemMarkingLine
    Set oUnkColl = oAssocRel.CollectionRelations(IID_IJMfgSystemMark, "SystemMark2dChild")
    Set oTargetObjCol = oUnkColl

    Set oAssocRel = Nothing
    Set oUnkColl = Nothing
    
    If oTargetObjCol.Count > 0 Then
        Set GetMarkingLineGeom2D = oTargetObjCol.Item(1)
        
    End If
    Set oTargetObjCol = Nothing
    Exit Function
End Function

Public Function GetGeometryTypeName(oGeomType As StrMfgGeometryType) As String
    Select Case oGeomType
        Case STRMFG_OUTER_CONTOUR
            GetGeometryTypeName = "STRMFG_OUTER_CONTOUR"
        Case STRMFG_INNER_CONTOUR
            GetGeometryTypeName = "STRMFG_INNER_CONTOUR"
        Case STRMFG_INNER_CONTOUR_MARK
            GetGeometryTypeName = "STRMFG_INNER_CONTOUR_MARK"
        Case STRMFG_FEATURE
            GetGeometryTypeName = "STRMFG_FEATURE"
        Case STRMFG_DIRECTION
            GetGeometryTypeName = "STRMFG_DIRECTION"
        Case STRMFG_FEATURE_MARK
            GetGeometryTypeName = "STRMFG_FEATURE_MARK"
        Case STRMFG_UV_MARK
            GetGeometryTypeName = "STRMFG_UV_MARK"
        Case STRMFG_SYSTEM_MARK
            GetGeometryTypeName = "STRMFG_SYSTEM_MARK"
        Case STRMFG_LOCATION_MARK
            GetGeometryTypeName = "STRMFG_LOCATION_MARK"
        Case STRMFG_END_MARK
            GetGeometryTypeName = "STRMFG_END_MARK"
        Case STRMFG_LAP_MARK
            GetGeometryTypeName = "STRMFG_LAP_MARK"
        Case STRMFG_FITTING_MARK
            GetGeometryTypeName = "STRMFG_FITTING_MARK"
        Case STRMFG_REF_MARK
            GetGeometryTypeName = "STRMFG_REF_MARK"
        Case STRMFG_BLOCK_MARK
            GetGeometryTypeName = "STRMFG_BLOCK_MARK"
        Case STRMFG_SEAM_MARK
            GetGeometryTypeName = "STRMFG_SEAM_MARK"
        Case STRMFG_TEMPLATE_MARK
            GetGeometryTypeName = "STRMFG_TEMPLATE_MARK"
        Case STRMFG_BASELINE_MARK
            GetGeometryTypeName = "STRMFG_BASELINE_MARK"
        Case STRMFG_ROLL_LINES_MARK
            GetGeometryTypeName = "STRMFG_ROLL_LINES_MARK"
        Case STRMFG_KNUCKLE_MARK
            GetGeometryTypeName = "STRMFG_KNUCKLE_MARK"
        Case STRMFG_USERDEFINED_MARK
            GetGeometryTypeName = "STRMFG_USERDEFINED_MARK"
        Case STRMFG_FLAT_OF_BOTTOM_MARK
            GetGeometryTypeName = "STRMFG_FLAT_OF_BOTTOM_MARK"
        Case STRMFG_FLAT_OF_SIDE_MARK
            GetGeometryTypeName = "STRMFG_FLAT_OF_SIDE_MARK"
        Case STRMFG_ROLL_BOUNDARIES_MARK
            GetGeometryTypeName = "STRMFG_ROLL_BOUNDARIES_MARK"
        Case STRMFG_LABELS_MARK
            GetGeometryTypeName = "STRMFG_LABELS_MARK"
        Case STRMFG_DIAGONALS_MARK
            GetGeometryTypeName = "STRMFG_DIAGONALS_MARK"
        Case STRMFG_ROBOT_MARK
            GetGeometryTypeName = "STRMFG_ROBOT_MARK"
        Case STRMFG_PAINTING_MARK
            GetGeometryTypeName = "STRMFG_PAINTING_MARK"
        Case STRMFG_NAME_MARK
            GetGeometryTypeName = "STRMFG_NAME_MARK"
        Case STRMFG_EDGE_CHECKLINES_MARK
            GetGeometryTypeName = "STRMFG_EDGE_CHECKLINES_MARK"
        Case STRMFG_BENDING_LINE
            GetGeometryTypeName = "STRMFG_BENDING_LINE"
        Case STRMFG_BENDING_CONTROLLINES_MARK
            GetGeometryTypeName = "STRMFG_BENDING_CONTROLLINES_MARK"
        Case STRMFG_SIGHTLINE_MARK
            GetGeometryTypeName = "STRMFG_SIGHTLINE_MARK"
        Case STRMFG_FRAMES_CHECKLINES_MARK
            GetGeometryTypeName = "STRMFG_FRAMES_CHECKLINES_MARK"
        Case STRMFG_MARGIN_MARK
            GetGeometryTypeName = "STRMFG_MARGIN_MARK"
        Case STRMFG_BEVEL_MARK
            GetGeometryTypeName = "STRMFG_BEVEL_MARK"
        Case STRMFG_GENERAL_MARK
            GetGeometryTypeName = "STRMFG_GENERAL_MARK"
        Case STRMFG_PLATELOCATION_MARK
            GetGeometryTypeName = "STRMFG_PLATELOCATION_MARK"
        Case STRMFG_PROFILELOCATION_MARK
            GetGeometryTypeName = "STRMFG_PROFILELOCATION_MARK"
        Case STRMFG_COLLARPLATELOCATION_MARK
            GetGeometryTypeName = "STRMFG_COLLARPLATELOCATION_MARK"
        Case STRMFG_BRACKETLOCATION_MARK
            GetGeometryTypeName = "STRMFG_BRACKETLOCATION_MARK"
        Case STRMFG_PLATE_TO_PLATE_BUTTJOINT_MARK
            GetGeometryTypeName = "STRMFG_PLATE_TO_PLATE_BUTTJOINT_MARK"
        Case STRMFG_PLATE_TO_PLATE_TJOINT_MARK
            GetGeometryTypeName = "STRMFG_PLATE_TO_PLATE_TJOINT_MARK"
        Case STRMFG_PROFILE_TO_PLATE_PENETRATION_MARK
            GetGeometryTypeName = "STRMFG_PROFILE_TO_PLATE_PENETRATION_MARK"
        Case STRMFG_PROFILE_TO_PLATE_MARK
            GetGeometryTypeName = "STRMFG_PROFILE_TO_PLATE_MARK"
        Case STRMFG_FRAMELINE_MARK
            GetGeometryTypeName = "STRMFG_FRAMELINE_MARK"
        Case STRMFG_WATERLINE_MARK
            GetGeometryTypeName = "STRMFG_WATERLINE_MARK"
        Case STRMFG_BUTTOCKLINE_MARK
            GetGeometryTypeName = "STRMFG_BUTTOCKLINE_MARK"
        Case STRMFG_WELDTAB_MARK
            GetGeometryTypeName = "STRMFG_WELDTAB_MARK"
        Case STRMFG_FEATURETAB_MARK
            GetGeometryTypeName = "STRMFG_FEATURETAB_MARK"
        Case STRMFG_BRIDGE_MARK
            GetGeometryTypeName = "STRMFG_BRIDGE_MARK"
        Case STRMFG_TemplateSet
            GetGeometryTypeName = "STRMFG_TemplateSet"
        Case STRMFG_Template
            GetGeometryTypeName = "STRMFG_Template"
        Case STRMFG_TemplateCtlLine
            GetGeometryTypeName = "STRMFG_TemplateCtlLine"
        Case STRMFG_TemplateBasePlane
            GetGeometryTypeName = "STRMFG_TemplateBasePlane"
        Case STRMFG_TemplateLocationMarkLine
            GetGeometryTypeName = "STRMFG_TemplateLocationMarkLine"
        Case STRMFG_TopLine
            GetGeometryTypeName = "STRMFG_TopLine"
        Case STRMFG_BoundaryLine3D
            GetGeometryTypeName = "STRMFG_BoundaryLine3D"
        Case STRMFG_TemplateSightLine
            GetGeometryTypeName = "STRMFG_TemplateSightLine"
        Case STRMFG_PinJigOutput
            GetGeometryTypeName = "STRMFG_PinJigOutput"
        Case STRMFG_PinJigPart3D
            GetGeometryTypeName = "STRMFG_PinJigPart3D"
        Case STRMFG_PinJigProcessData
            GetGeometryTypeName = "STRMFG_PinJigProcessData"
        Case STRMFG_MfgPinSets
            GetGeometryTypeName = "STRMFG_MfgPinSets"
        Case STRMFG_MfgPin
            GetGeometryTypeName = "STRMFG_MfgPin"
        Case STRMFG_PinJigProjectedData
            GetGeometryTypeName = "STRMFG_PinJigProjectedData"
        Case STRMFG_PinJigRemarkingLine3D
            GetGeometryTypeName = "STRMFG_PinJigRemarkingLine3D"
        Case STRMFG_PinJigContourLine3D
            GetGeometryTypeName = "STRMFG_PinJigContourLine3D"
        Case STRMFG_PinJigRemarkingLine2D
            GetGeometryTypeName = "STRMFG_PinJigRemarkingLine2D"
        Case STRMFG_PinJigContourLine2D
            GetGeometryTypeName = "STRMFG_PinJigContourLine2D"
        Case STRMFG_PinJigCornerPoint
            GetGeometryTypeName = "STRMFG_PinJigCornerPoint"
        Case STRMFG_PinJigRemarkingPoint
            GetGeometryTypeName = "STRMFG_PinJigRemarkingPoint"
        Case STRMFG_PinJigCrossPoint
            GetGeometryTypeName = "STRMFG_PinJigCrossPoint"
        Case STRMFG_PinJigOriginPoint
            GetGeometryTypeName = "STRMFG_PinJigOriginPoint"
        Case STRMFG_PinJigSurface
            GetGeometryTypeName = "STRMFG_PinJigSurface"
        Case STRMFG_NEUTRALAXIS_CURVE
            GetGeometryTypeName = "STRMFG_NEUTRALAXIS_CURVE"
        Case STRMFG_FITTINGANGLE
            GetGeometryTypeName = "STRMFG_FITTINGANGLE"
        Case STRMFG_TWIST_INFO
            GetGeometryTypeName = "STRMFG_TWIST_INFO"
        Case STRMFG_PROFILE_BEND_INFO
            GetGeometryTypeName = "STRMFG_PROFILE_BEND_INFO"
        Case STRMFG_PROFILE_ROLL_INFO
            GetGeometryTypeName = "STRMFG_PROFILE_ROLL_INFO"
        Case STRMFG_PROFILE_DEPTH_AFTERUNTWIST
            GetGeometryTypeName = "STRMFG_PROFILE_DEPTH_AFTERUNTWIST"
        Case STRMFG_PROFILE_DEPTH_BEFOREUNTWIST
            GetGeometryTypeName = "STRMFG_PROFILE_DEPTH_BEFOREUNTWIST"
        Case STRMFG_PROFILE_LENGTH
            GetGeometryTypeName = "STRMFG_PROFILE_LENGTH"
        Case STRMFG_AllGeometryType
            GetGeometryTypeName = "STRMFG_AllGeometryType"
        Case STRMFG_AllGeometryCol3dType
            GetGeometryTypeName = "STRMFG_AllGeometryCol3dType"
        Case STRMFG_BOUNDARY_DIR
            GetGeometryTypeName = "STRMFG_BOUNDARY_DIR"
        Case STRMFG_BOUNDARY_REVERSE_DIR
            GetGeometryTypeName = "STRMFG_BOUNDARY_REVERSE_DIR"
        Case STRMFG_LANDING_CURVE
            GetGeometryTypeName = "STRMFG_LANDING_CURVE"
        Case STRMFG_OPENING_EDGE
            GetGeometryTypeName = "STRMFG_OPENING_EDGE"
        Case STRMFG_PRIMARY_SHRINKAGE_MARK
            GetGeometryTypeName = "STRMFG_PRIMARY_SHRINKAGE_MARK"
        Case STRMFG_SECONDARY_SHRINKAGE_MARK
            GetGeometryTypeName = "STRMFG_SECONDARY_SHRINKAGE_MARK"
        Case STRMFG_NAVALARCHLINE
            GetGeometryTypeName = "STRMFG_NAVALARCHLINE"
        Case STRMFG_CROSSSECTION_MARK
            GetGeometryTypeName = "STRMFG_CROSSSECTION_MARK"
        Case STRMFG_WATERLINE_REND_MARK
            GetGeometryTypeName = "STRMFG_WATERLINE_REND_MARK"
        Case STRMFG_WELD_COMPENSATION_MARK
            GetGeometryTypeName = "STRMFG_WELD_COMPENSATION_MARK"
        Case STRMFG_WELD_COMPENSATION_WC1_MARK
            GetGeometryTypeName = "STRMFG_WELD_COMPENSATION_WC1_MARK"
        Case STRMFG_WELD_COMPENSATION_WC2_MARK
            GetGeometryTypeName = "STRMFG_WELD_COMPENSATION_WC2_MARK"
        Case STRMFG_OUTER_CONTOUR_PREWELD
            GetGeometryTypeName = "STRMFG_OUTER_CONTOUR_PREWELD"
        Case STRMFG_SketchedTemplateLocationMarkLine
            GetGeometryTypeName = "STRMFG_SketchedTemplateLocationMarkLine"
        Case STRMFG_SketchedTemplate
            GetGeometryTypeName = "STRMFG_SketchedTemplate"
        Case STRMFG_EXTEND_PINJIG_INTERSECTION
            GetGeometryTypeName = "STRMFG_EXTEND_PINJIG_INTERSECTION"
        Case STRMFG_PINJIG_MARK
            GetGeometryTypeName = "STRMFG_PINJIG_MARK"
        Case STRMFG_PINJIG_DIAGONAL
            GetGeometryTypeName = "STRMFG_PINJIG_DIAGONAL"
        Case STRMFG_MARKINGLINES_CUSTOM
            GetGeometryTypeName = "STRMFG_MARKINGLINES_CUSTOM"
        Case STRMFG_PROFILE_FINALOUTPUT
            GetGeometryTypeName = "STRMFG_PROFILE_FINALOUTPUT"
        Case STRMFG_PROFILE_SUPPORTOUTPUT
            GetGeometryTypeName = "STRMFG_PROFILE_SUPPORTOUTPUT"
        Case STRMFG_PROFILE_BEFOREUNFOLD
            GetGeometryTypeName = "STRMFG_PROFILE_BEFOREUNFOLD"
        Case STRMFG_PROFILE_AFTERUNFOLD
            GetGeometryTypeName = "STRMFG_PROFILE_AFTERUNFOLD"
        Case STRMFG_PINJIG_ARCHETYPE
            GetGeometryTypeName = "STRMFG_PINJIG_ARCHETYPE"
        Case STRMFG_FET_MARK
            GetGeometryTypeName = "STRMFG_FET_MARK"
        Case STRMFG_PROFILE_ORIGIN
            GetGeometryTypeName = "STRMFG_PROFILE_ORIGIN"
        Case STRMFG_BUILTUP_CONNECTION_MARK
            GetGeometryTypeName = "STRMFG_BUILTUP_CONNECTION_MARK"
        Case STRMFG_MEMBER_OPENING_MARK
            GetGeometryTypeName = "STRMFG_MEMBER_OPENING_MARK"
        Case STRMFG_MEMBER_LOCATION_MARK
            GetGeometryTypeName = "STRMFG_MEMBER_LOCATION_MARK"
        Case STRMFG_MEMBER_REFERENCE_MARK
            GetGeometryTypeName = "STRMFG_MEMBER_REFERENCE_MARK"
        Case Default
            GetGeometryTypeName = ""
    End Select

End Function

Public Function GetPOM() As IJDPOM
    On Error GoTo Cleanup
    
    Const sMODELDATABASE = "Model"
    Const sCONNECTMIDDLE = "ConnectMiddle"
    Const sDBTYPECONFIG = "DBTypeConfiguration"
    
    Dim objJContext   As IJContext
    Set objJContext = GetJContext()
    
    Dim oDBTypeConfig     As IJDBTypeConfiguration
    Set oDBTypeConfig = objJContext.GetService(sDBTYPECONFIG)
    
    Dim sModelDBID          As Variant
    sModelDBID = oDBTypeConfig.get_DataBaseFromDBType(sMODELDATABASE)
    
    Dim oConnectMiddle    As IJDAccessMiddle    'IJDConnectMiddle
    Set oConnectMiddle = objJContext.GetService(sCONNECTMIDDLE)
    
    Set GetPOM = oConnectMiddle.GetResourceManager(sModelDBID)
    
Cleanup:
    Set objJContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Exit Function
ErrorHandler:
    GetPOM = Nothing
    GoTo Cleanup
End Function

'******************************************************************************
' Routine: GetPOMEx
'
' Abstract:
' This method returns the Persistent Object Manager (POM) for the
' specified database.
'
' Description:
'
' Inputs:
'
' Outputs:
'
'******************************************************************************
Public Function GetPOMEx(databaseType As String) As IJDPOM
    Const METHOD As String = "GetPOMEx"
    On Error GoTo ErrorHandler

    Dim sProgressMessage As String
    sProgressMessage = ""
        
    ' Get the Server Context.
    Dim oPOM As IJDPOM
    Dim oContext As IJContext
    Dim oAccessMiddle As IJDAccessMiddle
    
    sProgressMessage = "Get the Server Context"
    Set oContext = GetJContext()

    ' Get the AccessMiddle object.
    sProgressMessage = "Get the IJDAccessMiddle"
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")

    ' Get the Persistent Object Managers (POM).
    sProgressMessage = "Get the Persistent Object Manager (POM)"
    Set oPOM = oAccessMiddle.GetResourceManagerFromType(databaseType)
   
    ' Return the POM.
    sProgressMessage = "Return the POM"
    Set GetPOMEx = oPOM
   
    Set oContext = Nothing
    Set oAccessMiddle = Nothing
    Set oPOM = Nothing

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, Module, METHOD, sProgressMessage).Number
End Function

Private Function BindToObject(oID As IMoniker) As IUnknown
    On Error GoTo Cleanup
    
    Dim oModelPOM As IJDPOM
    Set oModelPOM = GetPOM
    
    Set BindToObject = oModelPOM.GetObject(oID)
    
Cleanup:
    Set oModelPOM = Nothing
    
    Exit Function
ErrorHandler:
    Set BindToObject = Nothing
    GoTo Cleanup
End Function

'********************************************************************
' ' Routine: LogError
'
' Description:  default Error Reporter
'********************************************************************
Public Function LogError(oErrObject As ErrObject, _
                            Optional strSourceFile As String = "", _
                            Optional strMethod As String = "", _
                            Optional strExtraInfo As String = "") As IJError
     
    Dim strErrSource As String
    Dim strErrDesc As String
    Dim lErrNumber As Long
    Dim oEditErrors As IJEditErrors
     
    m_bLogError = True
    lErrNumber = oErrObject.Number
    strErrSource = oErrObject.Source
    strErrDesc = oErrObject.Description
     
     ' retrieve the error service
    Set oEditErrors = GetJContext().GetService("Errors")
       
    ' add the error to the service : the error is also logged to the file specified by
    ' "HKEY_LOCAL_MACHINE/SOFTWARE/Intergraph/Sp3D/Core/OperationParameter/ReportErrors_Log"
    Set LogError = oEditErrors.Add(lErrNumber, _
                                      strErrSource, _
                                      strErrDesc, _
                                      , _
                                      , _
                                      , _
                                      strMethod & ": " & strExtraInfo, _
                                      , _
                                      strSourceFile)
    Set oEditErrors = Nothing
End Function

