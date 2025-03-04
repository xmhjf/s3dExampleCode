Attribute VB_Name = "APSProfileHelper"
'----------------------------------------------------------------------------------------------------------------------------------------------------
' Copyright (C) 2009 Intergraph Corporation. All rights reserved.
'
' File Info:
'     Project: MfgProfileMarking
'     Module:  APSProfileHelper.bas
'
' Abstract:
'      Helper for the MfgProfileMarking rules.
'
' Notes:
'      Because of invalid data from Structural Detailing, some of the marks may not be shown in the manufacturing output.
'      For example, if physical connection between the plate and profile is missing then the part monitor will not show the plate location mark in MfgProfile Output.
'
'      To address these kinds of problems with the connections and inability to project reference curves on the profile geometry,
'      as a workaround solution, we use the marking line to be constructing this marking geometry for generating the Manufacturing output.
'
'
' History:
'     B K Teja      April 24. 2009      Created.
'
'----------------------------------------------------------------------------------------------------------------------------------------------------

Option Explicit

Const MODULE = "APSProfileHelper.bas"

'*************************************************************************************************************************************
' CreateAPSMarkingsPlateLocationMark
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects of Plate location Marks created by the User.
'               Input arguments: reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below type:
'       STRMFG_PLATELOCATION_MARK
'   2. For each Geom3d object get the corresponding Marking Line.
'   3. The marking line has custom attributes stored on the custom interface 'IJMfgSketchLocation'. Get the values of these attributes.
'   4. Create a system mark with the information on these attributes.
' ************************************************************************************************************************************

Private Sub CreateAPSMarkingsPlateLocationMark(eMarkingType As Long, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, oGeomCol3d As MfgGeomCol3d)
Const METHOD = "CreateAPSMarkingsPlateLocationMark"
On Error GoTo ErrorHandler
        
    Dim oGeom3d             As IJMfgGeom3d
        
    Dim lLoopCounter        As Long
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lGeomCount          As Long
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim lThicknessSide      As Long
    Dim strDir              As String
    Dim dThickness          As Double
    Dim oThisPart           As Object
    Dim oRelatedPart        As Object
    Dim oCS                 As IJComplexString
    Dim oThisPort           As IJPort
    Dim lMarkingSide        As Long
    
    'Get user created Plate Location marks
    Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, eMarkingType)

    lGeomCount = oGeomCol3d.Getcount
        
    For lLoopCounter = 1 To lGeomCount
  
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
        
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
        
        'Get Marking Line Side.
        lThicknessSide = GetProfileMarkingSide(oMarkingLineAE)

        'Get the Related part's  thickness
        dThickness = GetRelatedPartThickness(oMarkingLineAE)
        
        'create system mark object
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        
        'set marking side in SystemMark object
        oSystemMark.SetMarkingSide lThicknessSide
        
        Set oMarkingInfo = oSystemMark

        'Get Custom Attributes
        
        Dim oAttribute As IJDAttribute
        Dim oAttributes As IJDAttributes
        Dim oAttributesCol As IJDAttributesCol
        Dim INTERFACE As Variant
        Dim oMetaDataHelp As IJDAttributeMetaData
        
        Set oAttributes = oMarkingLineAE
        Set oMetaDataHelp = oMarkingLineAE
        INTERFACE = oMetaDataHelp.IID("IJMfgSketchLocation")
        
        Set oAttributesCol = oAttributes.CollectionOfAttributes(INTERFACE)
        
        'Get Related part name
        Set oAttribute = oAttributesCol.Item("RelatedPartName")
        oMarkingInfo.name = oAttribute.Value
        
        'get direction of the marking
        Set oAttribute = oAttributesCol.Item("Direction")
        strDir = GetCodeListStringValue("SMMarkingThicknessDirection", oAttribute.Value)
        oMarkingInfo.direction = strDir
        
        oMarkingInfo.thickness = dThickness
        
        'set ThicknessDirection on MarkingInfo object
        If strDir = "In" Or strDir = "Out" Then
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
        
        'set FittingAngle
        Set oAttribute = oAttributesCol.Item("FittingAngle")
        oMarkingInfo.FittingAngle = CDbl(oAttribute.Value)
        
        'set Geometry on the SystemMark
        oSystemMark.Set3dGeometry oGeom3d
        
        If eMarkingType = STRMFG_PLATELOCATION_MARK Then ' Create Mount angle Marks from the Location MArk
            'Get the part(on which marking line is existing)
              Set oThisPart = oMarkingLineAE.GetMfgMarkingPart
              
              Set oRelatedPart = oMarkingLineAE.GetMfgMarkingRelatedObject
        
             Set oCS = oGeom3d
             Dim sMoldedSide As String
             Dim oConnectionData As ConnectionData
             
            'Get Marking Line Side.
            lMarkingSide = GetProfileMarkingSide(oMarkingLineAE)
            
            lGeomCount = lGeomCount + 1
            
            CreateDeclivityMarksOnProfile oThisPart, oConnectionData, oCS, lMarkingSide, _
                                  oGeom3d, oGeomCol3d, lGeomCount, sMoldedSide, oRelatedPart
        End If
                              
NextItem:
        Set oMarkingLineAE = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing

    Next lLoopCounter

Exit Sub

ErrorHandler:
    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 2008, , "RULES"
    GoTo NextItem
    
End Sub

'*************************************************************************************************************************************
' CreateAPSMarkingsEndConnectionMark
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects of End Connection Marks created by the User.
'               Input arguments: reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below type:
'       STRMFG_END_MARK
'   2. For each Geom3d object get the corresponding Marking Line.
'   3. The marking line has custom attributes stored on the custom interface 'IJMfgSketchLocation'. Get the values of these attributes.
'   4. Get the Face port of the end of the connected profile where marking line is placed
'   4. Create a system mark with the information on these attributes.
' ************************************************************************************************************************************
Private Sub CreateAPSMarkingsEndConnectionMark(ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, oGeomCol3d As MfgGeomCol3d)
Const METHOD = "CreateAPSMarkingsEndConnectionMark"
On Error GoTo ErrorHandler
        
    Dim oGeom3d             As IJMfgGeom3d
    
    Dim lLoopCounter        As Long
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lGeomCount          As Long
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oMetaDataHelp       As IJDAttributeMetaData
    Dim oAttribute          As IJDAttribute
    Dim oAttributes         As IJDAttributes
    Dim oAttributesCol      As IJDAttributesCol
    Dim lThicknessSide      As Long
    Dim INTERFACE           As Variant
    Dim oPort               As IJPort
    Dim dThickness          As Double
    
    
    Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_END_MARK)

    lGeomCount = oGeomCol3d.Getcount
    
    For lLoopCounter = 1 To lGeomCount
       
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
        
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
        
        'Get Marking Line Side.
        lThicknessSide = GetProfileMarkingSide(oMarkingLineAE)
        
        'Get the Related part's  thickness
        dThickness = GetRelatedPartThickness(oMarkingLineAE)
        
        'Create Systemmark Object
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        oSystemMark.SetMarkingSide lThicknessSide
        
        Set oMarkingInfo = oSystemMark
        
        '****************************
        'Get Custom Attributes
        '****************************
        Set oAttributes = oMarkingLineAE

        Set oMetaDataHelp = oMarkingLineAE
        INTERFACE = oMetaDataHelp.IID("IJMfgSketchLocation")

        Set oAttributesCol = oAttributes.CollectionOfAttributes(INTERFACE)
        
        'get Related Part's name
        Set oAttribute = oAttributesCol.Item("RelatedPartName")
        oMarkingInfo.name = oAttribute.Value
        oMarkingInfo.thickness = dThickness
        
        'get FittingAngle
        Set oAttribute = oAttributesCol.Item("FittingAngle")
        oMarkingInfo.FittingAngle = CDbl(oAttribute.Value)
        
        Dim oObject As Object
        Set oObject = oMarkingLineAE.GetMfgMarkingRefInput
        
        If TypeOf oObject Is IJPort Then
            Set oPort = oObject
            'Get the FacePort of the edge port used for creating marking line
            Dim oFacePort As IJPort
            Get_FacePortFromEdgePort oPort, oFacePort
             
            'get the port on which marking line is placed
            Dim oProfilePort As IJPort
            Set oProfilePort = oMarkingLineAE.GetMfgMarkingPortForProfilePart(oMarkingLineAE.GetMfgMarkingPart)
    
            'get common geometry of the ports
            Dim oWireBody As IJWireBody
            Set oWireBody = m_oMfgRuleHelper.GetCommonGeometry(oFacePort.Geometry, oProfilePort.Geometry, False)
             'Converting wire body into complex string
            Dim oCS As IJComplexString
            Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oWireBody)
            
            'Crete a geom3d object with the new geometry
            oGeom3d.PutGeometry oCS
    
            oSystemMark.Set3dGeometry oGeom3d
            
        Else
            ' if RefInput of the marking line is not a edge port, add the geom3d object to the system mark object
            oSystemMark.Set3dGeometry oGeom3d
        End If
NextItem:
        Set oMarkingLineAE = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oMetaDataHelp = Nothing
        Set oAttribute = Nothing
        Set oAttributes = Nothing
        Set oAttributesCol = Nothing
        Set INTERFACE = Nothing
        Set oPort = Nothing

    Next lLoopCounter
Exit Sub

ErrorHandler:
    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 2003, , "RULES"
    GoTo NextItem
End Sub

'*************************************************************************************************************************************
' CreateAPSMarkingsSeamControlMark
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects of Seam Control Marks created by the User.
'               Input arguments: reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below type:
'       STRMFG_SEAM_MARK
'   2. For each Geom3d object get the corresponding Marking Line.
'   3. Offset the seam control mark if it is placed by reference curve method
'   4. Create a system mark with the information.
' ************************************************************************************************************************************

Private Sub CreateAPSMarkingsSeamControlMark(ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, oGeomCol3d As MfgGeomCol3d)
Const METHOD = "CreateAPSMarkingsSeamControlMark"
On Error GoTo ErrorHandler
        
    Dim oGeom3d             As IJMfgGeom3d
          
    Dim lLoopCounter        As Long
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lGeomCount          As Long
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim lThicknessSide      As Long
    Dim eMarkingLineMode    As enumUserMarkingLineMode
    Dim oPort               As IJPort
    
    Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_SEAM_MARK)

    lGeomCount = oGeomCol3d.Getcount
    
    For lLoopCounter = 1 To lGeomCount
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
            
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
        
        'Get Marking Line Side.
        lThicknessSide = GetProfileMarkingSide(oMarkingLineAE)
        
        'Get Marking Line Mode
         eMarkingLineMode = GetUserDefinedMarkingLineMode(oMarkingLineAE)
        
        'Create a system mark object
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oSystemMark.SetMarkingSide lThicknessSide

        Set oMarkingInfo = oSystemMark
        
        'set marking name attribute
        oMarkingInfo.name = "SEAM CONTROL"
        
        If eMarkingLineMode = eReferenceMode Then
            ' The user defined marking Line is created either in reference curve mode.
            ' offset the system mark created from the user defined marking line by GetSeamDistance.
            
            'get the port of the profile on which marking is placed
            Set oPort = oMarkingLineAE.GetMfgMarkingPortForProfilePart(oMarkingLineAE.GetMfgMarkingPart)

            'get the marking line placed by the user as a complex string
            Dim oSurface As IJSurfaceBody
            Set oSurface = oPort.Geometry
            Dim oComplexString      As IJComplexString
            Set oComplexString = oGeom3d 'oMarkingLineAE.GeometryAsComplexString
            
            'get the wirebody from the complex string
            Dim oWB As IJWireBody
            Dim MfgMGHelper         As IJMfgMGHelper
            Set MfgMGHelper = New GSCADMathGeom.MfgMGHelper
            MfgMGHelper.ComplexStringToWireBody oComplexString, oWB
            Set MfgMGHelper = Nothing
            
            'get the midpoint of the marking line
            Dim oMidPoint As IJDPosition
            Dim oGeomOps As IJDTopologyToolBox
            Set oGeomOps = New DGeomOpsToolBox
            oGeomOps.GetMiddlePointOfCompositeCurve oWB, oMidPoint

            'get the center of gravity of the marking surface to find the vector direction
            Dim oCOG As IJDPosition
            oSurface.GetCenterOfGravity oCOG
            
            ' get a vector to indicate the side where the marking line needs to be offseted
            Dim oVector As IJDVector
            Set oVector = New DVector
            oVector.Set (oMidPoint.x - oCOG.x), (oMidPoint.y - oCOG.y), (oMidPoint.z - oCOG.z)
            
            m_oMfgRuleHelper.ScaleVector oVector, -1
            
            Dim dOffSet As Double
            Dim oUnkTransientSheet As IUnknown
            
            'offset the curve and add the geometry to the Geom3d object
            dOffSet = GetSeamDistance
            Set oUnkTransientSheet = m_oMfgRuleHelper.OffsetCurve(oSurface, oWB, oVector, dOffSet, False)
            
            Set oComplexString = Nothing
            
            Set oComplexString = m_oMfgRuleHelper.WireBodyToComplexString(oUnkTransientSheet)
            Set oUnkTransientSheet = Nothing

            oGeom3d.PutGeometry oComplexString
            oSystemMark.Set3dGeometry oGeom3d
        
        Else
            ' The user defined marking Line is created either in Sketch2d mode or Intersection mode.
            ' There is no need to offset the system mark created from the user defined marking line.
            oSystemMark.Set3dGeometry oGeom3d
        
        End If
NextItem:
        Set oMarkingLineAE = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oPort = Nothing
    Next lLoopCounter
Exit Sub

ErrorHandler:
    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 2015, , "RULES"
    GoTo NextItem
End Sub

'*************************************************************************************************************************************
' CreateAPSMarkingsFittingMark
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects of Fitting Marks created by the User.
'               Input arguments: reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below type:
'       STRMFG_FITTING_MARK
'   2. For each Geom3d object get the corresponding Marking Line.
'   4. Create a system mark with the information of the user created marking line.
' ************************************************************************************************************************************

Private Sub CreateAPSMarkingsFittingMark(ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, oGeomCol3d As MfgGeomCol3d)
Const METHOD = "CreateAPSMarkingsFittingMark"
On Error GoTo ErrorHandler
    
    Dim oGeom3d As IJMfgGeom3d
    Dim lLoopCounter        As Long
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lGeomCount          As Long
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim lThicknessSide      As Long

    Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_FITTING_MARK)

    lGeomCount = oGeomCol3d.Getcount
    
    For lLoopCounter = 1 To lGeomCount
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)

        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
        
        'Get Marking Line Side.
        lThicknessSide = GetProfileMarkingSide(oMarkingLineAE)
        
        'create System mark object
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oSystemMark.SetMarkingSide lThicknessSide
        
        Set oMarkingInfo = oSystemMark
        
        'Set marking name
        oMarkingInfo.name = "FITTING"
    
        oSystemMark.Set3dGeometry oGeom3d
NextItem:
        Set oMarkingLineAE = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
    Next lLoopCounter
    
Exit Sub

ErrorHandler:
    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 2004, , "RULES"
    GoTo NextItem
End Sub

'*************************************************************************************************************************************
' CreateAPSMarkingsKnuckleMark
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects of Fitting Marks created by the User.
'               Input arguments: reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of below type:
'       STRMFG_KNUCKLE_MARK
'   2. For each Geom3d object get the corresponding Marking Line.
'   4. Create a system mark with the information of the user created marking line.
' ************************************************************************************************************************************

Private Sub CreateAPSMarkingsKnuckleMark(ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, oGeomCol3d As MfgGeomCol3d)
Const METHOD = "CreateAPSMarkingsKnuckleMark"
On Error GoTo ErrorHandler
        
    Dim oGeom3d As IJMfgGeom3d
    Dim lLoopCounter        As Long
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lGeomCount          As Long
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oMetaDataHelp       As IJDAttributeMetaData
    Dim oAttribute          As IJDAttribute
    Dim oAttributes         As IJDAttributes
    Dim oAttributesCol      As IJDAttributesCol
    Dim oAttributeInfo      As IJDAttributeInfo
    Dim lThicknessSide      As Long
    Dim INTERFACE           As Variant

    
    Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, STRMFG_KNUCKLE_MARK)

    lGeomCount = oGeomCol3d.Getcount
    
    For lLoopCounter = 1 To lGeomCount
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
        
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
        
        'Get Marking Line Side.
        lThicknessSide = GetProfileMarkingSide(oMarkingLineAE)
        
        'create System mark object
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        oSystemMark.SetMarkingSide lThicknessSide
        Set oMarkingInfo = oSystemMark

        'set marking info name
        oMarkingInfo.name = "KNUCKLE-U-" & lLoopCounter
        On Error GoTo EndGetAttributes:

        'get attribute collection
        Set oAttributes = oMarkingLineAE
        
        Set oMetaDataHelp = oMarkingLineAE
        INTERFACE = oMetaDataHelp.IID("IJMfgSketchLocation")

        Set oAttributesCol = oAttributes.CollectionOfAttributes(INTERFACE)
        
        'get direction attribute
        Set oAttribute = oAttributesCol.Item("Direction")
        oMarkingInfo.direction = GetCodeListStringValue("SMMarkingThicknessDirection", oAttribute.Value)

        'get FittingAngle attribute
        Set oAttribute = oAttributesCol.Item("FittingAngle")
        oMarkingInfo.FittingAngle = CDbl(oAttribute.Value)
EndGetAttributes:
        On Error GoTo ErrorHandler
        
        'set oGeom3d ob system mark
        oSystemMark.Set3dGeometry oGeom3d
NextItem:
        Set oMarkingLineAE = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing
        Set oMetaDataHelp = Nothing
        Set oAttribute = Nothing
        Set oAttributes = Nothing
        Set oAttributesCol = Nothing
        Set oAttributeInfo = Nothing
        Set INTERFACE = Nothing

    Next lLoopCounter
    
Exit Sub

ErrorHandler:
    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 2006, , "RULES"
    GoTo NextItem
End Sub

' ************************************************************************************************************************************
' Public Function CreateAPSMarkings
'
' Description:  Helper function which fills the Geom3DCollection with Geom3d objects created from the user defined marking lines.
'               Input arguments: marking type, reference object collection.
'               Input/ Output argument: Geom3dCollection.
'
' ************************************************************************************************************************************

Public Sub CreateAPSProfileMarkings(eMarkingType As Long, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByRef oOutputGeomCol3d As MfgGeomCol3d)
Const METHOD = "CreateAPSProfileMarkings: "
    
    'if reference object collection is nothing then exit the method
    If ReferenceObjColl Is Nothing Then
        Exit Sub
    ElseIf ReferenceObjColl.Count = 0 Then
        Exit Sub
    End If
    
    If m_oMfgRuleHelper Is Nothing Then
        Set m_oMfgRuleHelper = New MfgRuleHelpers.Helper
    End If
    
        Dim oGeomCol3d As MfgGeomCol3d
        
    Select Case eMarkingType
    
    Case STRMFG_PLATELOCATION_MARK, STRMFG_MOUNT_ANGLE_MARK:  'Plate Location Mark and Mount Angle Mark
        CreateAPSMarkingsPlateLocationMark eMarkingType, ReferenceObjColl, oGeomCol3d
        
    Case STRMFG_END_MARK:   'End Connection Mark
        CreateAPSMarkingsEndConnectionMark ReferenceObjColl, oGeomCol3d
        
    Case STRMFG_SEAM_MARK:  'Seam Control Mark
        CreateAPSMarkingsSeamControlMark ReferenceObjColl, oGeomCol3d
      
    Case STRMFG_FITTING_MARK:   'Fitting Profile Mark
        CreateAPSMarkingsFittingMark ReferenceObjColl, oGeomCol3d
        
    Case STRMFG_KNUCKLE_MARK:  'Knuckle mark
        CreateAPSMarkingsKnuckleMark ReferenceObjColl, oGeomCol3d
        
    Case STRMFG_HOLE_TRACE_MARK, STRMFG_HOLE_REF_MARK: 'Hole Trace marks, Hole Ref marks
        CreateAPSMarkingsHoleMarks eMarkingType, ReferenceObjColl, oGeomCol3d
        
    Case Else
        CreateProfileMarks eMarkingType, ReferenceObjColl, oGeomCol3d
      
    End Select

    If oOutputGeomCol3d Is Nothing Then
        Set oOutputGeomCol3d = oGeomCol3d
    Else
        If Not oGeomCol3d Is Nothing Then
            If (oGeomCol3d.Getcount > 0) Then
                Dim ind As Long
                For ind = oGeomCol3d.Getcount To 1 Step -1
                    Dim oMfgGeom3d As IJMfgGeom3d
                    Set oMfgGeom3d = oGeomCol3d.GetGeometry(ind)
                    oGeomCol3d.RemoveGeometry oMfgGeom3d
                    oOutputGeomCol3d.AddGeometry (oOutputGeomCol3d.Getcount + 1), oMfgGeom3d
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

Private Sub CreateAPSMarkingsHoleMarks(eMarkingType As Long, _
                            ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, _
                            ByRef oGeomCol3d As MfgGeomCol3d)
    Const METHOD = " CreateAPSMarkingsHoleMarks "
    On Error GoTo ErrorHandler

    Dim lLoopCounter        As Long
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim oThisPart           As Object
    Dim lMarkingSide        As Long
    Dim oMarkingPort        As IJPort
    Dim oComplexString      As IJComplexString
    Dim oWB                 As IJWireBody
    Dim MfgMGHelper         As IJMfgMGHelper
    
    Dim oGeom3dCustom       As IJMfgGeom3d
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
    
    lGeomCountCustom = oGeomCol3DCustom.Getcount
    
    For lLoopCounter = 1 To lGeomCountCustom
    
         On Error GoTo ErrorHandler
         'Get the Geom3d object from Geom3dCollection
         Set oGeom3dCustom = oGeomCol3DCustom.GetGeometry(lLoopCounter)
         
         'Get the Marking Line AE from Geom3d object
         Set oMarkingLineAE = GetMarkingLineAE(oGeom3dCustom)
        
        'Get the part(on which marking line is existing)
         Set oThisPart = oMarkingLineAE.GetMfgMarkingPart
         
        'Get Marking Line Side.
        lMarkingSide = GetProfileMarkingSide(oMarkingLineAE)

         'Get the Marking Line geometry as complex string
         Set oComplexString = oGeom3dCustom
         
         'Get the Wire Body from Complex String
         Set MfgMGHelper = New GSCADMathGeom.MfgMGHelper
         MfgMGHelper.ComplexStringToWireBody oComplexString, oWB
                       
         'Create Cross Marks and add to oGeomCol3d
         CreateProfileHoleMarks oThisPart, oWB, lMarkingSide, Nothing, oGeom3dCustom, oGeomCol3d
             

NextMark:
      
         Set oMarkingLineAE = Nothing
         Set oGeom3dCustom = Nothing

    Next lLoopCounter
    
    On Error GoTo ErrorHandler
    ' Get rid of oGeomCol3DCustom and oGeom3dCustom
    Dim oGeomObject As IJDObject
    While oGeomCol3DCustom.Getcount > 0
        Set oGeom3dCustom = oGeomCol3DCustom.GetGeometry(1)
        oGeomCol3DCustom.RemoveGeometry oGeom3dCustom
    Wend

    Exit Sub
    
ErrorHandler:

    StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", 1020, , "RULES"
    
    GoTo NextMark
    
End Sub


'*************************************************************************************************************************************
' CreateProfileMarks
'
' Description:  Function which fills the Geom3DCollection with Geom3d objects of Marks created by the User.
'               Input arguments: reference object collection.
'               Input/ Output argument: Geom3dCollection.
'   Algorithm:
'   1. Prepare Geom3dCollection of Geom3D objects with the geometries of user created marking lines of given type
'   2. For each Geom3d object get the corresponding Marking Line.
'   3. Create a system mark with the information of the user created marking line.
' ************************************************************************************************************************************

Private Sub CreateProfileMarks(eMarkingType As Long, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, oGeomCol3d As MfgGeomCol3d)
Const METHOD = "CreateProfileMarks"
On Error GoTo ErrorHandler
        
    Dim oGeom3d             As IJMfgGeom3d
    Dim lLoopCounter        As Long
    Dim oSystemMark         As IJMfgSystemMark
    Dim oMarkingInfo        As MarkingInfo
    Dim lGeomCount          As Long
    Dim oMarkingLineAE      As IJMfgMarkingLines_AE
    Dim lThicknessSide      As Long
    
    Set oGeomCol3d = m_oMfgRuleHelper.GetUserDefinedMarkGeometriesFromColl(ReferenceObjColl, eMarkingType)

    lGeomCount = oGeomCol3d.Getcount
    
    For lLoopCounter = 1 To lGeomCount
        Set oGeom3d = oGeomCol3d.GetGeometry(lLoopCounter)
        
        'Get the Marking Line AE from Geom3d object
        Set oMarkingLineAE = GetMarkingLineAE(oGeom3d)
        
        'Get Marking Line Side.
        lThicknessSide = GetProfileMarkingSide(oMarkingLineAE)
        
        'create System mark object
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        oSystemMark.SetMarkingSide lThicknessSide
        Set oMarkingInfo = oSystemMark

        'set marking info name
        'oMarkingInfo.name = ""
               
        'set oGeom3d ob system mark
        oSystemMark.Set3dGeometry oGeom3d
        
NextItem:

        Set oMarkingLineAE = Nothing
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oGeom3d = Nothing

    Next lLoopCounter
    
Exit Sub

ErrorHandler:
    StrMfgLogError Err, MODULE, METHOD
    GoTo NextItem
End Sub

