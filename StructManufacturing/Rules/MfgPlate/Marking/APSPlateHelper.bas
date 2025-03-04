Attribute VB_Name = "APSPlateHelper"
'----------------------------------------------------------------------------------------------------------------------------------------------------
' Copyright (C) 2009 Intergraph Corporation. All rights reserved.
'
' File Info:
'     Project: MfgPlateMarking
'     Module:  APSPlateHelper.cls
'
' Abstract:
'      Helper for the MfgPlateMarking rules.
'
' Notes:
'      Because of invalid data from Structural Detailing, some of the marks may not be shown in the manufacturing output.
'      For example, if physical connection between the plate and profile is missing then the part monitor will not show the profile location mark.
'
'      To address these kinds of problems with the connections and inability to project reference curves on the plate geometry,
'      as a workaround solution, we use the marking line to be constructing this marking geometry for generating the Manufacturing output.
'
'
' History:
'     Suma Mallena      April 20. 2009      Created.
'
'----------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Private Const MODULE = " MfgPlateMarking.APSPlateHelper "

' Enumerator for the mode by which the user defined marking is created in the marking line command.
Private Enum enumUserMarkingLineMode
    eIntersectionMode = 0   ' Intersection
    eSketchMode = 1         ' 2D Projection
    eReferenceMode = 2      ' Reference Curve
End Enum

' ************************************************************************************************************************************
' Public Function CreateAPSMarkings
'
' Description:  Helper function which fills the Geom3DCollection with Geom3d objects created from the user defined marking lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'
' ************************************************************************************************************************************
Public Sub CreateAPSMarkings(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oOutputGeomCol3d As MfgGeomCol3d)
         
    If m_oMfgRuleHelper Is Nothing Then
        Set m_oMfgRuleHelper = New MfgRuleHelpers.Helper
    End If
    
	Dim oGeomCol3d As MfgGeomCol3d
	
    Select Case eMarkingType
    
        ' Plate/Profile Location Marks,Bracket marks,Collar marks
        Case STRMFG_PLATELOCATION_MARK, STRMFG_PROFILELOCATION_MARK, STRMFG_BRACKETLOCATION_MARK, STRMFG_COLLARPLATELOCATION_MARK, STRMFG_MOUNT_ANGLE_MARK:
    
            CreateLocationMarks eMarkingType, ReferenceObjColl, oGeomCol3d
        
        ' Plate/Profile Fitting Marks
        Case STRMFG_PLATE_TO_PLATE_TJOINT_MARK, STRMFG_PROFILE_TO_PLATE_MARK:
        
            CreateFittingMarks eMarkingType, ReferenceObjColl, oGeomCol3d
        
        ' Seam Control MArks
        Case STRMFG_SEAM_MARK:
        
            CreateSeamControlMarks eMarkingType, ReferenceObjColl, oGeomCol3d
        
        ' Roll Marks
        Case STRMFG_ROLL_LINES_MARK, STRMFG_ROLL_BOUNDARIES_MARK:
        
            CreateRollMarks eMarkingType, ReferenceObjColl, oGeomCol3d
            
        'Knuckle Marks
        Case STRMFG_KNUCKLE_MARK
        
            CreateKnuckleMarks eMarkingType, ReferenceObjColl, oGeomCol3d
            
        'Naval Arch Lines
        Case STRMFG_NAVALARCHLINE
        
            CreateNavalArchMarks eMarkingType, ReferenceObjColl, oGeomCol3d
        
        'Ship Direction Mark
        Case STRMFG_DIRECTION
            CreateShipDirectionMarks eMarkingType, ReferenceObjColl, oGeomCol3d
            
        'Lap connection marks,End connection marks
        Case STRMFG_LAP_MARK, STRMFG_END_MARK:
            CreateConnectionMarks eMarkingType, ReferenceObjColl, oGeomCol3d
            
        'Hole Trace marks, Hole Ref marks
        Case STRMFG_HOLE_TRACE_MARK, STRMFG_HOLE_REF_MARK:
            CreateHoleMarks eMarkingType, ReferenceObjColl, oGeomCol3d
            
        'Conn Part Mark
        Case STRMFG_CONN_PART_MARK
            CreateConnectionPartMarks eMarkingType, ReferenceObjColl, oGeomCol3d
            
        Case Else
            CreateMarks eMarkingType, ReferenceObjColl, oGeomCol3d
        
    End Select

    If oOutputGeomCol3d Is Nothing Then
        Set oOutputGeomCol3d = oGeomCol3d
    Else
        If Not oGeomCol3d Is Nothing Then
        If (oGeomCol3d.GetCount > 0) Then
            Dim ind As Long
            For ind = oGeomCol3d.GetCount To 1 Step -1
                Dim oMfgGeom3d As IJMfgGeom3D
                Set oMfgGeom3d = oGeomCol3d.GetGeometry(ind)
                oGeomCol3d.RemoveGeometry oMfgGeom3d
                oOutputGeomCol3d.AddGeometry (oOutputGeomCol3d.GetCount + 1), oMfgGeom3d
                Set oMfgGeom3d = Nothing
            Next
            Dim oIJDObject As IJDObject
            Set oIJDObject = oGeomCol3d
            oIJDObject.Remove
            Set oIJDObject = Nothing
            End If
        End If
    End If

End Sub

'*************************************************************************************************************************************
'*************************              STRMFG_PLATELOCATION_MARK            *********************************************************
'*************************              STRMFG_PROFILELOCATION_MARK          *********************************************************
' ************************************************************************************************************************************
' Private Function CreateLocationMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below type:
'       STRMFG_PLATELOCATION_MARK
'       STRMFG_PROFILELOCATION_MARK
'       STRMFG_BRACKETLOCATION_MARK
'       STRMFG_COLLARPLATELOCATION_MARK
'   2. For each Geom3d object get the corresponding Marking Line.
'   3. The marking line has custom attributes stored on the custom interface 'IJMfgSketchLocation'. Get the values of these attributes.
'   4. Create a system mark with the information on these attributes.
' ************************************************************************************************************************************
Private Sub CreateLocationMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
                            
    Const METHOD = " CreateLocationMarks "
    On Error GoTo ErrorHandler
    
    Dim lGeomCount          As Long
    Dim lLoopCounter        As Long
    Dim oGeom3d             As IJMfgGeom3D
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lMarkingSide        As Long
    Dim oThisPart           As Object
    Dim oRelatedPart        As Object
    Dim oCS                 As IJComplexString
    Dim oThisPort           As IJPort
    
    lGeomCount = 0

    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    'If (eMarkingType = STRMFG_PLATELOCATION_MARK) Or (eMarkingType = STRMFG_PROFILELOCATION_MARK) Then
        Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, eMarkingType)
    'End If

    lGeomCount = oGeomCol3d.GetCount
    
    'Get the Coordinate System object
    Dim oMfgFrameSystem As IJDCoordinateSystem
    Dim nIndex As Long
    For nIndex = 1 To ReferenceObjColl.Count
        If TypeOf ReferenceObjColl.Item(nIndex) Is IJDCoordinateSystem Then
            Set oMfgFrameSystem = ReferenceObjColl.Item(nIndex)
            Exit For
        End If
    Next nIndex

    For lLoopCounter = 1 To lGeomCount
        'Get the Geom3d object from Geom3dCollection
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
        
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
       
        'Get Marking Side.
        GetMarkingSideAndPort oGeom3d, lMarkingSide, oThisPort

        'Create a system mark
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oSystemMark.SetMarkingSide lMarkingSide

        'Set the Marking Info
        Set oMarkingInfo = oSystemMark
        
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Get Marking Line Custom Attributes on the interface "IJMfgSketchLocation" and set them on MarkingInfo
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim oAttribute As IJDAttribute
        Dim oAttributes As IJDAttributes
        Dim oAttributesCol As IJDAttributesCol
        Dim INTERFACE As Variant
        Dim oMetaDataHelp As IJDAttributeMetaData
        
        Set oMetaDataHelp = oMarkingLineAE
        INTERFACE = oMetaDataHelp.IID("IJMfgSketchLocation")

        Set oAttributes = oMarkingLineAE
        Set oAttributesCol = oAttributes.CollectionOfAttributes(INTERFACE)

        'Get Custom Attribute value for RELATED PART NAME
        Set oAttribute = oAttributesCol.Item("RelatedPartName")
        oMarkingInfo.Name = oAttribute.Value

        'Get Custom Attribute value for DIRECTION
        Set oAttribute = oAttributesCol.Item("Direction")
        oMarkingInfo.Direction = GetCodeListStringValue("SMMarkingThicknessDirection", oAttribute.Value)

        ' Get THICKNESS DIRECTION VECTOR from DIRECTION
        Dim strDir As String
        strDir = oMarkingInfo.Direction

        If strDir = "in" Or strDir = "out" Then
            Dim oCurve As IJCurve
            Dim Xs As Double
            Dim Xe As Double
            Dim Ys As Double
            Dim Ye As Double
            Dim Zs As Double
            Dim Ze As Double
            oGeom3d.GetGeometry.GetCurve 1, oCurve
            oCurve.EndPoints Xs, Ys, Zs, Xe, Ye, Ze
            oMarkingInfo.ThicknessDirection = GetDirVector(strDir, Ys / Abs(Ys))
        Else
            oMarkingInfo.ThicknessDirection = GetDirVector(strDir, 0)
        End If

        'Custom Attribute: FITTING ANGLE
        Set oAttribute = oAttributesCol.Item("FittingAngle")
        oMarkingInfo.FittingAngle = CDbl(oAttribute.Value)
        
        'Custom Attribute: FLANGE DIRECTION
        Set oAttribute = oAttributesCol.Item("FlangeDirection")
        If oAttribute.Value <> vbNullString Then
            oMarkingInfo.FlangeDirection = GetCodeListStringValue("SMMarkingFlangeDirection", oAttribute.Value)
        End If

        'Finally set the Geometry of the System Mark
        oSystemMark.Set3dGeometry oGeom3d
        
        If (eMarkingType = STRMFG_PLATELOCATION_MARK) Or (eMarkingType = STRMFG_PROFILELOCATION_MARK) Then
        
             'Get the part(on which marking line is existing)
              Set oThisPart = oMarkingLineAE.GetMfgMarkingPart
              
              Set oRelatedPart = oMarkingLineAE.GetMfgMarkingRelatedObject
        
             Set oCS = oGeom3d
             lGeomCount = lGeomCount + 1
                                      
             Dim sMoldedSide As String
             Dim oConnectionData As ConnectionData
             CreateDeclivityMarks oThisPart, oConnectionData, oCS, oMfgFrameSystem, oGeom3d, oGeomCol3d, lGeomCount, lMarkingSide, sMoldedSide, oRelatedPart
        End If

        
NextLocationMark:
        Set oGeom3d = Nothing
        Set oMarkingLineAE = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        
    Next lLoopCounter

    Exit Sub
    
ErrorHandler:
   
    If (eMarkingType = STRMFG_PLATELOCATION_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1019, , "RULES"
    ElseIf (eMarkingType = STRMFG_PROFILELOCATION_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1022, , "RULES"
            ElseIf (eMarkingType = STRMFG_BRACKETLOCATION_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1004, , "RULES"
    ElseIf (eMarkingType = STRMFG_COLLARPLATELOCATION_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1007, , "RULES"

    End If
    
    GoTo NextLocationMark

End Sub

'*************************************************************************************************************************************
'*************************              STRMFG_PLATE_TO_PLATE_TJOINT_MARK            *************************************************
'*************************              STRMFG_PROFILE_TO_PLATE_MARK                 *************************************************
' ************************************************************************************************************************************
' Private Function CreateFittingMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   Case 1 :
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below types:
'       STRMFG_PLATE_TO_PLATE_TJOINT_MARK,
'       STRMFG_PROFILE_TO_PLATE_MARK
'   2. For each Geom3d object get back the Marking Line.
'   3. Get the information of Marking side from MArking Line.
'   4. Create a system mark with this information.
'
'   Case 2 :
'   (If there is a User Defined Location Mark, then the System Fitting Marks can be generated even if there are no User defined Fitting Marks.)
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below types:
'       STRMFG_PLATELOCATION_MARK,
'       STRMFG_PROFILELOCATION_MARK
'   2. For each Geom3d object get back the User Defined Location Marks.
'   3. Get the marking side, wire bodies of Location Marks.
'   4. Calculate the positions on the plate and create curves for Fitting marks.
'   5. Create system marks with this information.

' ************************************************************************************************************************************
Private Sub CreateFittingMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
    
    Const METHOD = " CreateFittingMarks "
    On Error GoTo ErrorHandler
    
    Dim lGeomCount          As Long
    Dim lLoopCounter        As Long
    Dim strMarkingName      As String
    Dim oGeom3d             As IJMfgGeom3D
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lMarkingSide        As Long
    
    lGeomCount = 0
    '*************************************************************************************************************************************
    '*********************************************          Case 1      '*****************************************************************
    '*************************************************************************************************************************************
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    If eMarkingType = STRMFG_PLATE_TO_PLATE_TJOINT_MARK Then
        strMarkingName = "PLATE TJOINT FITTING MARK"
        Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_PLATE_TO_PLATE_TJOINT_MARK)
    ElseIf eMarkingType = STRMFG_PROFILE_TO_PLATE_MARK Then
        strMarkingName = "PROFILE FITTING MARK"
        Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_PROFILE_TO_PLATE_MARK)
    End If
    
    lGeomCount = oGeomCol3d.GetCount
            
    For lLoopCounter = 1 To lGeomCount
        
        On Error GoTo ErrorHandler1
        'Get the Geom3d object from Geom3dCollection
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
        
        'Get Marking Side
        GetMarkingSideAndPort oGeom3d, lMarkingSide, Nothing
                
        'Create a system mark
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oSystemMark.SetMarkingSide lMarkingSide

        'Set the Marking Info
        Set oMarkingInfo = oSystemMark
        oMarkingInfo.Name = strMarkingName

        'Set the geometry of thesystem mark
        oSystemMark.Set3dGeometry oGeom3d
        
NextFittingMark:

        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
            
    Next lLoopCounter
    
        
    '*************************************************************************************************************************************
    '*********************************************          Case 2      '*****************************************************************
    '*************************************************************************************************************************************
       
    'If there is a User Defined Location Mark, then the System Fitting Marks can be generated even if there are no User defined Fitting Marks.
    Dim oMarkingPort        As IJPort
    Dim oComplexString      As IJComplexString
    Dim oWB                 As IJWireBody
    Dim MfgMGHelper         As IJMfgMGHelper
    
    Dim eMarkingLineMode    As enumUserMarkingLineMode
    Dim oGeom3dCustom       As IJMfgGeom3D
    Dim oMoniker            As IMoniker
    Dim oCS                 As IJComplexString
    Dim oThisPart           As Object
    Dim oThisPartSurface    As IJSurfaceBody
    Dim oConnectedPart      As Object
    Dim oConnPartSurface    As IJSurfaceBody
    
    Dim oPartSupport        As IJPartSupport
    Dim oProfilePartSupport As IJProfilePartSupport
    Dim oPlateSystem        As IJPlateSystem
    Dim oWebPort            As IJPort
    Dim oMarkPosColl        As Collection
    Dim oMarkPosObj         As IJDPosition 'iteration object
    Dim oGeomCol3DCustom    As IJMfgGeomCol3d
    Dim lGeomCountCustom    As Long
         
    lGeomCountCustom = 0
    On Error GoTo ErrorHandler
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    If eMarkingType = STRMFG_PLATE_TO_PLATE_TJOINT_MARK Then
        Set oGeomCol3DCustom = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_PLATELOCATION_MARK)
    ElseIf eMarkingType = STRMFG_PROFILE_TO_PLATE_MARK Then
        Set oGeomCol3DCustom = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_PROFILELOCATION_MARK)
    End If
    
    lGeomCountCustom = oGeomCol3DCustom.GetCount
    
    For lLoopCounter = 1 To lGeomCountCustom
    
         On Error GoTo ErrorHandler2
         'Get the Geom3d object from Geom3dCollection
         Set oGeom3dCustom = oGeomCol3DCustom.GetGeometry(lLoopCounter)
         
         'Get the Marking Line AE from Geom3d object
         Set oMarkingLineAE = GetMarkingLineAE(oGeom3dCustom)
         
         'Get Marking Line Side and the Port of the plate on which the Marking Line is existing.
         GetMarkingSideAndPort oGeom3dCustom, lMarkingSide, oMarkingPort
          
         'Create a system mark
         Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
         oSystemMark.SetMarkingSide lMarkingSide

         'Set the Marking Info
         Set oMarkingInfo = oSystemMark
         oMarkingInfo.Name = strMarkingName
         
         'Get the Marking Line geometry as complex string
         Set oComplexString = oGeom3dCustom
         
         'Get the Wire Body from Complex String
         Set MfgMGHelper = New GSCADMathGeom.MfgMGHelper
         MfgMGHelper.ComplexStringToWireBody oComplexString, oWB
                
         'Get the part(on which marking line is existing)
         Set oThisPart = oMarkingLineAE.GetMfgMarkingPart
          
         'Get the part surface (on which marking line is existing)
         Set oThisPartSurface = oMarkingPort.Geometry
          
         'Get connected part (on which marking line is existing)
         Set oConnectedPart = oMarkingLineAE.GetMfgMarkingRelatedObject
          
         'Get connected part's parent system
         If TypeOf oConnectedPart Is IJPlatePart Then
              Set oPlateSystem = m_oMfgRuleHelper.GetTopMostParentSystem(oConnectedPart)
              Set oConnPartSurface = oPlateSystem
         ElseIf TypeOf oConnectedPart Is IJStiffener Then
              Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
              Set oPartSupport.Part = oConnectedPart
              Set oProfilePartSupport = oPartSupport
              Set oWebPort = oProfilePartSupport.GetFacePortAdjacentToMountingFace(lMarkingSide)
              Set oConnPartSurface = oWebPort.Geometry
         Else
              Set oConnPartSurface = oConnectedPart
         End If
       
         'Get collection of mark positions
         Set oMarkPosColl = GetAPSFittingMarkPositions(oThisPart, oWB, lMarkingSide)
             
         'Iterate through the collection of mark points in oMarkPosColl
         For Each oMarkPosObj In oMarkPosColl

             'Create mark line through oMarkPosObj
             Set oCS = CreateLocationFittingMark(oMarkPosObj, oThisPartSurface, oConnPartSurface)
             
             Set oGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
             oGeom3d.PutGeometry oCS
             oGeom3d.PutGeometrytype eMarkingType
             Set oMoniker = oGeom3dCustom.GetMoniker
             oGeom3d.PutMoniker oMoniker
        
             oGeomCol3d.AddGeometry lLoopCounter, oGeom3d
         
             oSystemMark.Set3dGeometry oGeom3d

         Next oMarkPosObj
NextMark:
         Set oMarkPosColl = Nothing
         Set oMarkingLineAE = Nothing
         Set oSystemMark = Nothing
         Set oMarkingInfo = Nothing
         Set oGeom3d = Nothing
         Set oGeom3dCustom = Nothing

    Next lLoopCounter
    
    On Error GoTo ErrorHandler
    ' Get rid of oGeomCol3DCustom and oGeom3dCustom
    Dim oGeomObject As IJDObject
    While oGeomCol3DCustom.GetCount > 0
        Set oGeom3dCustom = oGeomCol3DCustom.GetGeometry(1)
        oGeomCol3DCustom.RemoveGeometry oGeom3dCustom
    Wend

    Exit Sub
    
ErrorHandler1:
    If (eMarkingType = STRMFG_PLATE_TO_PLATE_TJOINT_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1020, , "RULES"
    ElseIf (eMarkingType = STRMFG_PROFILE_TO_PLATE_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1023, , "RULES"
    End If

    ' Continue with the next fitting MArk of CASE 1
    GoTo NextFittingMark

ErrorHandler2:
    If (eMarkingType = STRMFG_PLATE_TO_PLATE_TJOINT_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1020, , "RULES"
    ElseIf (eMarkingType = STRMFG_PROFILE_TO_PLATE_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1023, , "RULES"
    End If
    ' Continue with the next fitting MArk of CASE 2
    GoTo NextMark
    
ErrorHandler:
    If (eMarkingType = STRMFG_PLATE_TO_PLATE_TJOINT_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1020, , "RULES"
    ElseIf (eMarkingType = STRMFG_PROFILE_TO_PLATE_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1023, , "RULES"
    End If
    
End Sub

'*************************************************************************************************************************************
'*************************              STRMFG_SEAM_MARK            *****************************************************************
'*************************************************************************************************************************************
' Private Function CreateSeamControlMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below type:
'       STRMFG_SEAM_MARK'
'   2. For each Geom3d object get back the Marking Line.
'   3. Get the information of Marking side from MArking Line.
'   4. Create a system mark with this information.
'
' Note:
'   1. When the user defined marking Line is created by using a reference curve from Marking Line command,
'       there is a need to offset the system mark which is generated from this user defined marking line.
'
'   2. When the user defined marking Line is created either in Sketch2d mode or Intersection mode,
'       there is no need to offset the system mark created from the user defined marking line and the system mark will be located at the same location.
'
' ************************************************************************************************************************************
Private Sub CreateSeamControlMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
    Const METHOD = " CreateSeamControlMarks "
    On Error GoTo ErrorHandler
    
    Dim lGeomCount          As Long
    Dim lLoopCounter        As Long
    Dim strMarkingName      As String
    Dim oGeom3d             As IJMfgGeom3D
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    
    Dim oMarkingPort        As IJPort
    Dim oComplexString      As IJComplexString
    Dim oWB                 As IJWireBody
    Dim MfgMGHelper         As IJMfgMGHelper
    Dim lMarkingSide        As Long
    Dim eMarkingLineMode    As enumUserMarkingLineMode
    Dim oCS                 As IJComplexString
     
    If Not (eMarkingType = STRMFG_SEAM_MARK) Then
        Exit Sub
    End If
    
    lGeomCount = 0
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, eMarkingType)
    strMarkingName = "SEAM CONTROL"
    
    lGeomCount = oGeomCol3d.GetCount
     
    For lLoopCounter = 1 To lGeomCount
         
         'Get the Geom3d object from Geom3dCollection
         Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
         
         'Get the Marking Line AE from Geom3d object
         Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
         
         'Get Marking Line Mode
         eMarkingLineMode = GetUserDefinedMarkingLineMode(oMarkingLineAE)
         
         'Get Marking Line Side and the Port of the plate on which the Marking Line is existing.
         GetMarkingSideAndPort oGeom3d, lMarkingSide, oMarkingPort
                  
         'Create a system mark
         Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
         oSystemMark.SetMarkingSide lMarkingSide

         'Set the Marking Info
         Set oMarkingInfo = oSystemMark
         oMarkingInfo.Name = strMarkingName

         ' 1. When the user defined marking Line is created by using a reference curve in marking line command,
         ' there is a need to offset the system mark line generated from this user defined marking line.
         
         ' 2. When the user defined marking Line is created either in Sketch2d mode or Intersection mode,
         ' there is no need to offset the system mark created from the user defined marking line and the system mark will be located at the same location.
                     
         If eMarkingLineMode = eReferenceMode Then
             ' The user defined marking Line is created by using a reference curve.
             ' Need to offset the system mark line generated from this user defined marking line.
             
             'Get the Marking Line geometry as complex string
             Set oComplexString = oGeom3d
             
             'Get the Wire Body from Complex String
             Set MfgMGHelper = New GSCADMathGeom.MfgMGHelper
             MfgMGHelper.ComplexStringToWireBody oComplexString, oWB
             
             'Get the middle point of the Wire Body
             Dim oWBMidPoint As IJDPosition
             Dim oGeomOpsToolBox As New DGeomOpsToolBox
             oGeomOpsToolBox.GetMiddlePointOfCompositeCurve oWB, oWBMidPoint
          
             'Get the center of gravity of the plate upside surface.
             Dim oUpsideSurface As IJSurfaceBody
             Dim oPlateCG As IJDPosition
             
             Set oUpsideSurface = oMarkingPort.Geometry
             oUpsideSurface.GetCenterOfGravity oPlateCG
             
             'Create a vector from the mid-point of Wire Body to CG of plate.
             Dim oVector As IJDVector
             Set oVector = New DVector
             oVector.Set (oWBMidPoint.x - oPlateCG.x), (oWBMidPoint.y - oPlateCG.y), (oWBMidPoint.z - oPlateCG.z)
             m_oMfgRuleHelper.ScaleVector oVector, -1
             
             'Get the offset distance by which the Wire Body has to be moved.
             Dim dOffSetDist As Double
             dOffSetDist = GetSeamDistance
             
             'Offset the Wire Body in the direction of the vector.
             Dim oUnkOffsetCurve As IUnknown
             Set oUnkOffsetCurve = m_oMfgRuleHelper.OffsetCurve(oUpsideSurface, oWB, oVector, dOffSetDist, False)
             
             'Get the Complex string from the Wire Body.
             Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oUnkOffsetCurve)
     
             oGeom3d.PutGeometry oCS
             oSystemMark.Set3dGeometry oGeom3d
             Set oUnkOffsetCurve = Nothing
         
         Else
             ' The user defined marking Line is created either in Sketch2d mode or Intersection mode.
             ' There is no need to offset the system mark created from the user defined marking line.
             oSystemMark.Set3dGeometry oGeom3d
         
         End If
         
NextSeamMark:
         Set oGeom3d = Nothing
         Set oMarkingLineAE = Nothing
         Set oSystemMark = Nothing
         Set oMarkingInfo = Nothing

     Next lLoopCounter

    Exit Sub
    
ErrorHandler:
    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1028, , "RULES"
    GoTo NextSeamMark
End Sub

'**************************************************************************************************************************************
'*************************              STRMFG_ROLL_LINES_MARK            *************************************************************
'*************************              STRMFG_ROLL_BOUNDARIES_MARK            ********************************************************
' *************************************************************************************************************************************
' Private Function CreateRollMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below types:
'       STRMFG_ROLL_LINES_MARK,
'       STRMFG_ROLL_BOUNDARIES_MARK
'   2. For each Geom3d object get back the Marking Line.
'   3. Get the information of Marking side from MArking Line.
'   4. Create a system mark with this information.
' ************************************************************************************************************************************

Private Sub CreateRollMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
    Const METHOD = " CreateRollMarks "
    On Error GoTo ErrorHandler

    Dim lGeomCount          As Long
    Dim lLoopCounter        As Long
    Dim strMarkingName      As String
    Dim oGeom3d             As IJMfgGeom3D
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lMarkingSide        As Long
    
    lGeomCount = 0
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    If eMarkingType = STRMFG_ROLL_LINES_MARK Then
        strMarkingName = "ROLL LINE"
        Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_ROLL_LINES_MARK)
    ElseIf eMarkingType = STRMFG_ROLL_BOUNDARIES_MARK Then
        strMarkingName = "ROLL BOUNDARY"
        Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_ROLL_BOUNDARIES_MARK)
    End If
    
    lGeomCount = oGeomCol3d.GetCount
    
    For lLoopCounter = 1 To lGeomCount
         
         'Get the Geom3d object from Geom3dCollection
         Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
         
         'Get the Marking Line AE from Geom3d object
         Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
         
         'Get Marking Line Side
         GetMarkingSideAndPort oGeom3d, lMarkingSide
                  
         'Create a system mark
         Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
         oSystemMark.SetMarkingSide lMarkingSide

         'Set the Marking Info
         Set oMarkingInfo = oSystemMark
         oMarkingInfo.Name = strMarkingName
             
         'Finally set the Geometry of the System Mark
         oSystemMark.Set3dGeometry oGeom3d
         
NextRollMark:
         Set oGeom3d = Nothing
         Set oMarkingLineAE = Nothing
         Set oSystemMark = Nothing
         Set oMarkingInfo = Nothing
         
     Next lLoopCounter

    Exit Sub
    
ErrorHandler:

    If (eMarkingType = STRMFG_ROLL_LINES_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1027, , "RULES"
    ElseIf (eMarkingType = STRMFG_ROLL_BOUNDARIES_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1026, , "RULES"
    End If
    
    GoTo NextRollMark
    
End Sub

' ***********************************************************************************
' Private Function GetAPSFittingMarkPositions
'
' Description:  Helper function to find fitting mark positions for the input part.
'               Input arguments: this part, Wire Body, marking side.
'               Output argument: collection of found mark point positions.'
'
' ***********************************************************************************
Private Function GetAPSFittingMarkPositions(oThisPart As Object, _
                                            oCurve As IJWireBody, _
                                            UpSide As Long) As Collection
    
    Const sMETHOD = "GetAPSFittingMarkPositions"

    On Error GoTo ErrorHandler

    '- Take two ends positions of the wirebody and create mark points positions at FittingMark Distance
    '- Save final mark points positions in the oMarkPosColl and return

    'Prepare different helper objects
    Dim oMarkPosColl As New Collection
    Dim oEndPos1 As IJDPosition
    Dim oEndPos2 As IJDPosition
    Dim oMarkPointPos1 As IJDPosition
    Dim oMarkPointPos2 As IJDPosition

     'get end points of WB
     oCurve.GetEndPoints oEndPos1, oEndPos2
     
     'find two mark points at 'd' from both ends
     Dim oCompStr As IJComplexString
     Dim o3dCurve As IJCurve
     Dim dLength As Double
     Dim oCurveElems As IJElements
     
     Set oCurveElems = m_oMfgRuleHelper.WireBodyToComplexStrings(oCurve)
     
     For Each oCompStr In oCurveElems
         Set o3dCurve = oCompStr
         dLength = dLength + o3dCurve.Length
     Next
     If dLength < (2 * FittingMarkDistance) Then
         Exit Function
     End If
     
     Set oMarkPointPos1 = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oCurve, oEndPos1, FittingMarkDistance, oEndPos2)

     Set oMarkPointPos2 = m_oMfgRuleHelper.GetPointAlongCurveAtDistance(oCurve, oEndPos2, FittingMarkDistance, oEndPos1)

     'Add mark points position to the collection
     oMarkPosColl.Add oMarkPointPos1
     oMarkPosColl.Add oMarkPointPos2

    Set GetAPSFittingMarkPositions = oMarkPosColl
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'**************************************************************************************************************************************
'*************************              STRMFG_KNUCKLE_MARK            *************************************************************
' *************************************************************************************************************************************
' Private Function CreateKnuckleMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below types:
'       STRMFG_KNUCKLE_MARK
'   2. For each Geom3d object get back the Marking Line.
'   3. Get the information of Marking side from MArking Line.
'   4. Create a system mark with this information.
' ************************************************************************************************************************************

Private Sub CreateKnuckleMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
    Const METHOD = " CreateKnuckleMarks "
    On Error GoTo ErrorHandler

    Dim lGeomCount          As Long
    Dim lLoopCounter        As Long
    Dim strMarkingName      As String
    Dim oGeom3d             As IJMfgGeom3D
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lMarkingSide        As Long
   
    lGeomCount = 0
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    If eMarkingType = STRMFG_KNUCKLE_MARK Then
        strMarkingName = "KNUCKLE MARK"
        Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_KNUCKLE_MARK)
    End If
    
    lGeomCount = oGeomCol3d.GetCount
    
    For lLoopCounter = 1 To lGeomCount

        'Get the Geom3d object from Geom3dCollection
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
        
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
       
        'Get Marking Side.
        GetMarkingSideAndPort oGeom3d, lMarkingSide
      
        'Create a system mark
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oSystemMark.SetMarkingSide lMarkingSide

        'Set the Marking Info
        Set oMarkingInfo = oSystemMark
    
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Get Marking Line Custom Attributes on the interface "IJMfgSketchLocation" and set them on MarkingInfo
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim oAttribute As IJDAttribute
        Dim oAttributes As IJDAttributes
        Dim oAttributesCol As IJDAttributesCol
        Dim INTERFACE As Variant
        Dim oMetaDataHelp As IJDAttributeMetaData
        
        Set oMetaDataHelp = oMarkingLineAE
        INTERFACE = oMetaDataHelp.IID("IJMfgSketchLocation")

        Set oAttributes = oMarkingLineAE
        Set oAttributesCol = oAttributes.CollectionOfAttributes(INTERFACE)

        oMarkingInfo.Name = strMarkingName

        'Get Custom Attribute value for DIRECTION
        Set oAttribute = oAttributesCol.Item("Direction")
        oMarkingInfo.Direction = GetCodeListStringValue("SMMarkingThicknessDirection", oAttribute.Value)

        ' Get THICKNESS DIRECTION VECTOR from DIRECTION
        Dim strDir As String
        strDir = oMarkingInfo.Direction

        If strDir = "in" Or strDir = "out" Then
            Dim oCurve As IJCurve
            Dim Xs As Double
            Dim Xe As Double
            Dim Ys As Double
            Dim Ye As Double
            Dim Zs As Double
            Dim Ze As Double
            oGeom3d.GetGeometry.GetCurve 1, oCurve
            oCurve.EndPoints Xs, Ys, Zs, Xe, Ye, Ze
            oMarkingInfo.ThicknessDirection = GetDirVector(strDir, Ys / Abs(Ys))
        Else
            oMarkingInfo.ThicknessDirection = GetDirVector(strDir, 0)
        End If

        'Custom Attribute: FITTING ANGLE
        Set oAttribute = oAttributesCol.Item("FittingAngle")
        oMarkingInfo.FittingAngle = CDbl(oAttribute.Value)
        
                    
        'Finally set the Geometry of the System Mark
        oSystemMark.Set3dGeometry oGeom3d
        
NextKnuckleMark:
        Set oGeom3d = Nothing
        Set oMarkingLineAE = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        
    Next lLoopCounter


    Exit Sub
    
ErrorHandler:
    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1011, , "RULES"
    GoTo NextKnuckleMark
End Sub

'**************************************************************************************************************************************
'*************************              STRMFG_NAVALARCHLINE            *************************************************************
' *************************************************************************************************************************************
' Private Function CreateNavalArchMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below types:
'       STRMFG_NAVALARCHLINE
'   2. For each Geom3d object get back the Marking Line.
'   3. Get the information of Marking side from MArking Line.
'   4. Create a system mark with this information.
' ************************************************************************************************************************************

Private Sub CreateNavalArchMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
    Const METHOD = " CreateNavalArchMarks "
    On Error GoTo ErrorHandler

    Dim lGeomCount          As Long
    Dim lLoopCounter        As Long
    Dim strMarkingName      As String
    Dim oGeom3d             As IJMfgGeom3D
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lMarkingSide        As Long
    
    lGeomCount = 0
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    If eMarkingType = STRMFG_NAVALARCHLINE Then
        strMarkingName = "NAVAL ARCH MARK"
        Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_NAVALARCHLINE)
    End If
    
    lGeomCount = oGeomCol3d.GetCount
    
    For lLoopCounter = 1 To lGeomCount

        'Get the Geom3d object from Geom3dCollection
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
        
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
       
        'Get Marking Side.
        GetMarkingSideAndPort oGeom3d, lMarkingSide
      
        'Create a system mark
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oSystemMark.SetMarkingSide lMarkingSide

        'Set the Marking Info
        Set oMarkingInfo = oSystemMark

        Dim oRefInput As Object
        Dim oNamedItem As IJNamedItem
        
        On Error Resume Next
        Set oRefInput = oMarkingLineAE.GetMfgMarkingRefInput
        Set oNamedItem = oRefInput
        
        strMarkingName = oNamedItem.Name
        
        oMarkingInfo.Name = strMarkingName
                    
        'As the Curve type for Both Flat_of_Bottom  and Flat_of_Side Ref curves will
        'be same ie STRMFG_NAVALARCHLINE, so inorder to determine the reference curve marking type,
        'Use Reference Curve Name for checking the types
        
        'Search for String "Flat_of_bottom" in Curve name from start in text compare mode
        Dim eStrMfgGeomType As StrMfgGeometryType
        Dim iPos As Integer
        iPos = InStr(1, strMarkingName, "Flat_of_Bottom", vbTextCompare)
        If iPos = 1 Then 'ie at the start of string
            eStrMfgGeomType = STRMFG_FLAT_OF_BOTTOM_MARK
        Else
            iPos = InStr(1, strMarkingName, "Flat_of_Side", vbTextCompare)
            If iPos = 1 Then
                eStrMfgGeomType = STRMFG_FLAT_OF_SIDE_MARK
            Else
                eStrMfgGeomType = STRMFG_GENERAL_MARK
            End If
        End If
                    
        oGeom3d.PutGeometrytype eStrMfgGeomType
        
        On Error GoTo ErrorHandler
        
        'Finally set the Geometry of the System Mark
        oSystemMark.Set3dGeometry oGeom3d
        
        
NextNavalMark:
        Set oGeom3d = Nothing
        Set oMarkingLineAE = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        
    Next lLoopCounter


    Exit Sub
    
ErrorHandler:
    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1015, , "RULES"
    GoTo NextNavalMark
End Sub



'**************************************************************************************************************************************
'*************************              STRMFG_DIRECTION            *************************************************************
' *************************************************************************************************************************************
' Private Function CreateShipDirectionMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below types:
'       STRMFG_DIRECTION
'   2. For each Geom3d object get back the Marking Line.
'   3. Get the information of Marking side from MArking Line.
'   4. Create a system mark with this information.
' ************************************************************************************************************************************

Private Sub CreateShipDirectionMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
    Const METHOD = " CreateShipDirectionMarks "
    On Error GoTo ErrorHandler

    Dim lGeomCount          As Long
    Dim lLoopCounter        As Long
    Dim strMarkingName      As String
    Dim oGeom3d             As IJMfgGeom3D
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lMarkingSide        As Long
    
    lGeomCount = 0
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    If eMarkingType = STRMFG_DIRECTION Then
        strMarkingName = "SHIP DIRECTION MARK"
        Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_DIRECTION)
    End If
    
    lGeomCount = oGeomCol3d.GetCount
   
    For lLoopCounter = 1 To lGeomCount

        'Get the Geom3d object from Geom3dCollection
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
        
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
       
        'Get Marking Side.
        GetMarkingSideAndPort oGeom3d, lMarkingSide
      
        'Create a system mark
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oSystemMark.SetMarkingSide lMarkingSide

        'Set the Marking Info
        Set oMarkingInfo = oSystemMark
        
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Get Marking Line Custom Attributes on the interface "IJMfgSketchLocation" and set them on MarkingInfo
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim oAttribute As IJDAttribute
        Dim oAttributes As IJDAttributes
        Dim oAttributesCol As IJDAttributesCol
        Dim INTERFACE As Variant
        Dim oMetaDataHelp As IJDAttributeMetaData
        
        Set oMetaDataHelp = oMarkingLineAE
        INTERFACE = oMetaDataHelp.IID("IJMfgSketchLocation")

        Set oAttributes = oMarkingLineAE
        Set oAttributesCol = oAttributes.CollectionOfAttributes(INTERFACE)

        'Get Custom Attribute value for DIRECTION
        Set oAttribute = oAttributesCol.Item("Direction")
        oMarkingInfo.Name = GetCodeListStringValue("SMMarkingThicknessDirection", oAttribute.Value)
        
        If oMarkingInfo.Name = vbNullString Then
            oMarkingInfo.Name = "unknown"
        End If
                    
        'Finally set the Geometry of the System Mark
        oSystemMark.Set3dGeometry oGeom3d
        
NextShipDirectionMark:
        Set oGeom3d = Nothing
        Set oMarkingLineAE = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        
    Next lLoopCounter


    Exit Sub
    
ErrorHandler:
    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1029, , "RULES"
    GoTo NextShipDirectionMark
End Sub



'*************************************************************************************************************************************
'*************************              STRMFG_LAP_MARK            *********************************************************
'*************************              STRMFG_END_MARK            *********************************************************
' ************************************************************************************************************************************
' Private Function CreateConnectionMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below type:
'       STRMFG_END_MARK
'       STRMFG_LAP_MARK
'   2. For each Geom3d object get the corresponding Marking Line.
'   3. The marking line has custom attributes stored on the custom interface 'IJMfgSketchLocation'. Get the values of these attributes.
'   4. Create a system mark with the information on these attributes.
' ************************************************************************************************************************************
Private Sub CreateConnectionMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
                            
    Const METHOD = " CreateConnectionMarks "
    On Error GoTo ErrorHandler
    
    Dim lGeomCount          As Long
    Dim lLoopCounter        As Long
    Dim strMarkingName      As String
    Dim oGeom3d             As IJMfgGeom3D
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lMarkingSide        As Long
    
    lGeomCount = 0
    
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    If (eMarkingType = STRMFG_END_MARK) Or (eMarkingType = STRMFG_LAP_MARK) Then
        Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, eMarkingType)
    End If
    
    lGeomCount = oGeomCol3d.GetCount
            
    For lLoopCounter = 1 To lGeomCount
        
        'Get the Geom3d object from Geom3dCollection
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
              
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
               
        'Get Marking Side
        GetMarkingSideAndPort oGeom3d, lMarkingSide, Nothing
                
        'Create a system mark
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oSystemMark.SetMarkingSide lMarkingSide

        'Set the Marking Info
        Set oMarkingInfo = oSystemMark
                
        Dim oAttribute As IJDAttribute
        Dim oAttributes As IJDAttributes
        Dim oAttributesCol As IJDAttributesCol
        Dim INTERFACE As Variant
        Dim oMetaDataHelp As IJDAttributeMetaData
        
        Set oMetaDataHelp = oMarkingLineAE
        INTERFACE = oMetaDataHelp.IID("IJMfgSketchLocation")

        Set oAttributes = oMarkingLineAE
        Set oAttributesCol = oAttributes.CollectionOfAttributes(INTERFACE)

        'Get Custom Attribute value for RELATED PART NAME
        Set oAttribute = oAttributesCol.Item("RelatedPartName")
        oMarkingInfo.Name = oAttribute.Value

        'Set the geometry of the system mark
        oSystemMark.Set3dGeometry oGeom3d
        
NextConnectionMark:

        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
            
    Next lLoopCounter
    
Exit Sub
    
ErrorHandler:
    If (eMarkingType = STRMFG_END_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1009, , "RULES"
    ElseIf (eMarkingType = STRMFG_LAP_MARK) Then
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1012, , "RULES"
    End If
    GoTo NextConnectionMark
End Sub



'**************************************************************************************************************************************
'*************************              STRMFG_HOLE_TRACE_MARK             *************************************************************
'*************************              STRMFG_HOLE_REF_MARK            ********************************************************
' *************************************************************************************************************************************
' Private Function CreateRollMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below types:
'       STRMFG_HOLE_TRACE_MARK,
'       STRMFG_HOLE_REF_MARK
'   2. For each Geom3d object get back the Marking Line.
'   3. Get the information of Marking side from MArking Line.
'   4. Create cross marks system mark with this information.
' ************************************************************************************************************************************

Private Sub CreateHoleMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
    Const METHOD = " CreateHoleMarks "
    On Error GoTo ErrorHandler

    Dim lLoopCounter        As Long
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oThisPart           As Object
    Dim lMarkingSide        As Long
    Dim oMarkingPort        As IJPort
    Dim oComplexString      As IJComplexString
    Dim oWB                 As IJWireBody
    Dim MfgMGHelper         As IJMfgMGHelper
    
    Dim oGeom3dCustom       As IJMfgGeom3D
    Dim oGeomCol3DCustom    As IJMfgGeomCol3d
    Dim lGeomCountCustom    As Long
         
    lGeomCountCustom = 0
    On Error GoTo ErrorHandler
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    If eMarkingType = STRMFG_HOLE_TRACE_MARK Then
        Set oGeomCol3DCustom = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_HOLE_TRACE_MARK)
    ElseIf eMarkingType = STRMFG_HOLE_REF_MARK Then
        Set oGeomCol3DCustom = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_HOLE_REF_MARK)
    End If
    
    lGeomCountCustom = oGeomCol3DCustom.GetCount
    
    For lLoopCounter = 1 To lGeomCountCustom
    
         On Error GoTo ErrorHandler
         'Get the Geom3d object from Geom3dCollection
         Set oGeom3dCustom = oGeomCol3DCustom.GetGeometry(lLoopCounter)
         
         'Get the Marking Line AE from Geom3d object
         Set oMarkingLineAE = GetMarkingLineAE(oGeom3dCustom)
        
        'Get the part(on which marking line is existing)
         Set oThisPart = oMarkingLineAE.GetMfgMarkingPart
         
         'Get Marking Line Side and the Port of the plate on which the Marking Line is existing.
         GetMarkingSideAndPort oGeom3dCustom, lMarkingSide, oMarkingPort

         'Get the Marking Line geometry as complex string
         Set oComplexString = oGeom3dCustom
         
         'Get the Wire Body from Complex String
         Set MfgMGHelper = New GSCADMathGeom.MfgMGHelper
         MfgMGHelper.ComplexStringToWireBody oComplexString, oWB
                       
         'Create Cross Marks and add to oGeomCol3d
         CreateHoleCrossMarks oThisPart, oWB, lMarkingSide, Nothing, oGeom3dCustom, oGeomCol3d
             

NextMark:
      
         Set oMarkingLineAE = Nothing
         Set oGeom3dCustom = Nothing

    Next lLoopCounter
    
    On Error GoTo ErrorHandler
    ' Get rid of oGeomCol3DCustom and oGeom3dCustom
    Dim oGeomObject As IJDObject
    While oGeomCol3DCustom.GetCount > 0
        Set oGeom3dCustom = oGeomCol3DCustom.GetGeometry(1)
        oGeomCol3DCustom.RemoveGeometry oGeom3dCustom
    Wend

    Exit Sub
    
ErrorHandler:

    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1067, , "RULES"
    
    GoTo NextMark
    
End Sub


'*************************************************************************************************************************************
'*************************              STRMFG_CONN_PART_MARK            *********************************************************
' ************************************************************************************************************************************
' Private Function CreateConnectionMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below type:
'       STRMFG_CONN_PART_MARK
'   2. For each Geom3d object get the corresponding Marking Line.
'   3. The marking line has custom attributes stored on the custom interface 'IJMfgSketchLocation'. Get the values of these attributes.
'   4. Create a system mark with the information on these attributes.
' ************************************************************************************************************************************
Private Sub CreateConnectionPartMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
                            
    Const METHOD = " CreateConnectionPartMarks "
    On Error GoTo ErrorHandler
    
    Dim lGeomCount          As Long
    Dim lLoopCounter        As Long
    Dim strMarkingName      As String
    Dim oGeom3d             As IJMfgGeom3D
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lMarkingSide        As Long
    Dim dBaseBaseDist As Double, dOffsetOffsetDist As Double
    Dim dOffsetBaseDist As Double, dBaseOffsetDist As Double, dDistDifference As Double
    Dim oRelatedPart        As Object
    Dim Part                As Object

    lGeomCount = 0
    
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, eMarkingType)
    
    lGeomCount = oGeomCol3d.GetCount
            
    For lLoopCounter = 1 To lGeomCount
        
        'Get the Geom3d object from Geom3dCollection
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
              
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
               
        'Get Marking Side
        GetMarkingSideAndPort oGeom3d, lMarkingSide, Nothing
                
        'Create a system mark
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oSystemMark.SetMarkingSide lMarkingSide

        'Set the Marking Info
        Set oMarkingInfo = oSystemMark
                
        Dim oAttribute As IJDAttribute
        Dim oAttributes As IJDAttributes
        Dim oAttributesCol As IJDAttributesCol
        Dim INTERFACE As Variant
        Dim oMetaDataHelp As IJDAttributeMetaData
        
        Set oMetaDataHelp = oMarkingLineAE
        INTERFACE = oMetaDataHelp.IID("IJMfgSketchLocation")

        Set oAttributes = oMarkingLineAE
        Set oAttributesCol = oAttributes.CollectionOfAttributes(INTERFACE)

        'Get Custom Attribute value for RELATED PART NAME
        Set oAttribute = oAttributesCol.Item("RelatedPartName")
        oMarkingInfo.Name = oAttribute.Value
        
        Set oRelatedPart = oMarkingLineAE.GetMfgMarkingRelatedObject
        
        '*** Check if the connected object is a plate ***'
        If Not oRelatedPart Is Nothing Then

            If TypeOf oRelatedPart Is IJPlatePart Then
            
                Dim oSDConnPlateWrapper As New StructDetailObjects.PlatePart
                Dim oSupp               As IJPartSupport
                Set oSupp = New PartSupport
                Dim oPos1 As New DPosition, oPos2 As New DPosition
            
                '*** Get difference between Upside ***'
                If lMarkingSide = BaseSide Then
                    dDistDifference = IIf(dBaseBaseDist < dBaseOffsetDist, dBaseBaseDist, dBaseOffsetDist)
                ElseIf lMarkingSide = OffsetSide Then
                    dDistDifference = IIf(dOffsetOffsetDist < dOffsetBaseDist, dOffsetOffsetDist, dOffsetBaseDist)
                End If
                
                Set Part = oMarkingLineAE.GetMfgMarkingPart
                
                'Create the SD plate Wrapper and initialize it
                Dim oSDPlateWrapper As StructDetailObjects.PlatePart
                Set oSDPlateWrapper = New StructDetailObjects.PlatePart
                Set oSDPlateWrapper.object = Part
            
                Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
                Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
                Set oPlateWrapper.object = Part
                
                '*** Get base surface geometry of this plate ***'
                Dim oPlateBasePort As IJPort
                Dim oPartBaseSurface As IUnknown
                Set oPlateBasePort = oSDPlateWrapper.BasePort(BPT_Base)
                Set oPartBaseSurface = oPlateBasePort.Geometry
                
                '*** Get ModelBody from geometry ***'
                Dim oPlateBaseModelBody As IJDModelBody
                Set oPlateBaseModelBody = oPartBaseSurface
                
                'Dim sPBFileName As String
                'sPBFileName = Environ("TEMP")
                'If sPBFileName = "" Or sPBFileName = vbNullString Then
                '    sPBFileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
                'End If
                'sPBFileName = sPBFileName & "\oPlateBase.sat"
                'oPlateBaseModelBody.DebugToSATFile sPBFileName
                '***********************************************'
                
                '*** Get base surface geometry of this plate ***'
                Dim oPlateOffsetPort As IJPort
                Dim oPartOffsetSurface As IUnknown
                Set oPlateOffsetPort = oSDPlateWrapper.BasePort(BPT_Offset)
                Set oPartOffsetSurface = oPlateOffsetPort.Geometry
                
                '*** Get ModelBody from geometry ***'
                Dim oPlateOffsetModelBody As IJDModelBody
                Set oPlateOffsetModelBody = oPartOffsetSurface
                '***********************************************'
                
                '*** Initialize the plate wrapper and the Physical Connection wrapper ***'
                Set oSDConnPlateWrapper.object = oRelatedPart
                        
                '*** Check if base surfaces are aligned ***'
                'Get base surface geometry of connected plate
                Dim oConnPlateBasePort As IJPort
                Dim oConnPartSurface As IUnknown
                Dim oConnPlateModelBody As IJDModelBody
                Dim oConnPlateOffsetPort As IJPort
                '***********************************************'
                
                '*** Get BaseBaseDist ***'
                Set oConnPlateBasePort = oSDConnPlateWrapper.BasePort(BPT_Base)
                Set oConnPartSurface = oConnPlateBasePort.Geometry
                Set oConnPlateModelBody = oConnPartSurface
                
                oPlateBaseModelBody.GetMinimumDistance oConnPlateModelBody, oPos1, oPos2, dBaseBaseDist
        
                '*** Get OffsetBaseDist ***'
                oPlateOffsetModelBody.GetMinimumDistance oConnPlateModelBody, oPos1, oPos2, dOffsetBaseDist
                
                '*** Get OffsetOffsetDist ***'
                Set oConnPlateOffsetPort = oSDConnPlateWrapper.BasePort(BPT_Offset)
                Set oConnPartSurface = oConnPlateOffsetPort.Geometry
                Set oConnPlateModelBody = oConnPartSurface
                
                oPlateOffsetModelBody.GetMinimumDistance oConnPlateModelBody, oPos1, oPos2, dOffsetOffsetDist
                
                '*** Get BaseOffsetDist ***'
                oPlateBaseModelBody.GetMinimumDistance oConnPlateModelBody, oPos1, oPos2, dBaseOffsetDist
        
                oMarkingInfo.SetAttributeNameAndValue "BaseBaseDifference", dBaseBaseDist
                oMarkingInfo.SetAttributeNameAndValue "BaseOffsetDifference", dBaseOffsetDist
                oMarkingInfo.SetAttributeNameAndValue "UpSideDifference", dDistDifference
                oMarkingInfo.SetAttributeNameAndValue "UpSide", lMarkingSide
                oMarkingInfo.SetAttributeNameAndValue "OffsetBaseDifference", dOffsetBaseDist
                oMarkingInfo.SetAttributeNameAndValue "OffsetOffsetDifference", dOffsetOffsetDist

            End If

        End If
        
        oGeom3d.PutGeometrytype STRMFG_CONN_PART_MARK

        'Set the geometry of the system mark
        oSystemMark.Set3dGeometry oGeom3d
        
NextConnectionMark:

        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
            
    Next lLoopCounter
    
Exit Sub
    
ErrorHandler:

    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1072, , "RULES"

    GoTo NextConnectionMark
End Sub

'**************************************************************************************************************************************
' Private Function CreateMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects created from User Defined MArking Lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines
'   2. For each Geom3d object get back the Marking Line.
'   3. Get the information of Marking side from MArking Line.
'   4. Create a system mark with this information.
' ************************************************************************************************************************************

Private Sub CreateMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
    Const METHOD = " CreateMarks "
    On Error GoTo ErrorHandler

    Dim lGeomCount          As Long
    Dim lLoopCounter        As Long
    Dim oGeom3d             As IJMfgGeom3D
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lMarkingSide        As Long
    
    lGeomCount = 0
    ' Prepares Geom3dCollection having Geom3D objects with the geometries of user created marking lines of given type.
    Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, eMarkingType)
    
    lGeomCount = oGeomCol3d.GetCount
    
    For lLoopCounter = 1 To lGeomCount
         
         'Get the Geom3d object from Geom3dCollection
         Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
         
         'Get the Marking Line AE from Geom3d object
         Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
         
         'Get Marking Line Side
         GetMarkingSideAndPort oGeom3d, lMarkingSide
                  
         'Create a system mark
         Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
         oSystemMark.SetMarkingSide lMarkingSide

         'Set the Marking Info
'         Set oMarkingInfo = oSystemMark
             
         'Finally set the Geometry of the System Mark
         oSystemMark.Set3dGeometry oGeom3d
         
         oGeom3d.PutGeometrytype eMarkingType
         
NextMark:
        
         Set oGeom3d = Nothing
         Set oMarkingLineAE = Nothing
         Set oSystemMark = Nothing
         Set oMarkingInfo = Nothing
         
     Next lLoopCounter

    Exit Sub
    
ErrorHandler:

    StrMfgLogError Err, MODULE, METHOD
    GoTo NextMark
    
End Sub

