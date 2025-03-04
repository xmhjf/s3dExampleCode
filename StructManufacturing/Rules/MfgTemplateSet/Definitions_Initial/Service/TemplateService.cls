VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'    Template service for Creation of BasePlane, ControlLine, TemplateOutContour etc.
'    which can be customized by shipyard.
'
'Description:
'History :
'   ByungSam Lee            07/09/2001      Creation.
'   ByungSam Lee            02/23/2002      Updated totally CreateTemplateOutContour Method
'   ByungSam Lee            06/24/2002      Updated this procedure according to the modification of Process Settings
'   Marcel Veldhuizen       2004/04/22      Included correct error handling
'   Thakur                  2005/05/02      Major modification for CR 67910
'   Ram Kandimalla          2005/09/23      Major code revision for stability
'   Venkat                  2006/10/12      Reviewed MsgBox usage -TR:103745
'   Sivaprasad              2007/21/02      DI-CP·114374  Template command should gracefully handle "no frame system"
'   Suma Mallena            2009/08/06      DI-CP-169008  Expose new balance types
'******************************************************************************************************************
Option Explicit
Implements IJDMfgTemplateService

Private Const EPSILON = 0.0001  '0.1mm
Private Const FRAME_EDGE_DIST = 0.009 ' consider frame is on edge if dist < FRAME_EDGE_DIST
Private Const PI = 3.14286
Private Const PROC_PARAM_FRAME = "Frame"
Private Const PROC_PARAM_TRANSVERSAL = "Transversal"
Private Const PROC_PARAM_ALONGFRAME = "AlongFrame"
Private Const PROC_PARAM_NOTUSED = "NotUsed"
Private Const PROC_PARAM_POSITIONFRAME = "PositionFrame"
Private Const PROC_PARAM_PERPENDICULAR = "Perpendicular"
Private Const PROC_PARAM_POSITIONEVEN = "PositionEven"
Private Const BASEPLANECODELISTNAME = "SMTemplateProcessBasePlane"
Private Const USERDEFVALUESRULENAME = "TemplateUserDefinedValues"
Private Const TOL_MM = 0.005

Private Const MODULE = "GSCADStrMfgTemplate.TemplateService"
Private Const PART_SURFACE_BASE = 2

Private m_oMfgGeomHelper As MfgGeomHelper
Private m_oMfgMGHelper   As GSCADMathGeom.MfgMGHelper
Private m_oPlatePart     As IJPlatePart
Private m_dOffset As Double
Private m_dExt As Double

Private Sub Class_Initialize()
 Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
            
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'********************************************************************************************************************************
' Function Name:    CreateBasePlane
' Interface:        IJDMfgTemplateService
' Inputs:           PlatePart -- selected plate
'                   pProcessSettings -- Template settings, type, orientation, direction, side etc.
' Output:           Object - Baseplane IJPlane
' Assumption:
' Algorithm:        1. Check the Template Side which is BaseSide or OffsetSide
'                   2. Check the Template Type and If it is a special plate or not
'                       2.1. Template Type is PerPendicualrXY(Near to Stem/Stern) or
'                           Pentagonal plate or Triangular plate which edge is not parallel to Frame line
'                           => Perpendicular XY Algorithm
'                       2.2. Template Type is Stem/Stern
'                           => Stem/Stern Type, There is no triangular and pentagonal plate(only quadrangle plate).
'                           => refer to Analysis- DefineCtlLineAndBPlane.doc
'                           => Exactly symmetrical to Center line
'                       2.3. Other Case
'                           Check the BasePlane Setting Value
'                           - BySystem - Check the area of Plate Part if more than one third(33%)  =>  Flat part is Base Plane
'                                                                                    otherwise => Make the BasePlane using the diagonal
'                           - NormalToPlate - Make the BasePlane using the diagonal
'                           - ParallelToPlate - Flat part is BasePlane
'
'********************************************************************************************************************************
Private Function IJDMfgTemplateService_CreateBasePlane(ByVal pPlatePart As Object, ByVal pProcessSettings As Object, ByVal oTemplateSet As Object, Optional bUserDefined As Boolean = False) As Object
 Const METHOD = "IJDMfgTemplateService_CreateBasePlane"
    On Error GoTo ErrorHandler
    
    bUserDefined = False
    
    Dim oTemplateSetData As TemplateSetData
    Set oTemplateSetData = New TemplateSetData
    oTemplateSetData.InitSettings pPlatePart, pProcessSettings, oTemplateSet
    
    Dim oBasePlane As IJPlane
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper

    Dim bBaseSide As Boolean
    bBaseSide = oTemplateSetData.TemplateSide
    
    
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
 
    Dim strBasePlane        As String
    Dim lBasePlaneType      As Long
    Dim oPOM                As IUnknown
    Dim oCodeListMetaData   As IJDCodeListMetaData
    
    strBasePlane = pProcessSettings.BasePlane
    
    Set oPOM = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    Set oCodeListMetaData = oPOM
    
    If Not oCodeListMetaData Is Nothing Then
        lBasePlaneType = oCodeListMetaData.ValueIDByShortString(BASEPLANECODELISTNAME, strBasePlane)
    End If
    Set oCodeListMetaData = Nothing
    Set oPOM = Nothing
    
    Dim ratio           As Double
    
    If lBasePlaneType = 5190 Then
        '''''''''''''''''''''''''''''''''''''''''''''
        '"AverageOfCornersPlane" / "NormalToPlate"'''
        '''''''''''''''''''''''''''''''''''''''''''''
        Set oBasePlane = CreateAvgOfCornersPlane(pPlatePart, oTemplateSetData.Direction, oSurfaceBody, bBaseSide)

    ElseIf lBasePlaneType = 5191 Then
        '''''''''''''''''''''''''''''''''''''''''''''
        '"MostPlanarNatural" / "ParallelTo plate"''''
        '''''''''''''''''''''''''''''''''''''''''''''
         On Error Resume Next 'based on surface complexity it might fail
         oGeomHelper.GetMostPlanarFaceOfSheetBody2 oSurfaceBody, ratio, oBasePlane
         If Err.Number < 0 Then
             LogMessage Err, MODULE, METHOD
         End If
         On Error GoTo ErrorHandler
                         
    ElseIf (lBasePlaneType = 5192) Then
        '''''''''''''''''''''''''''''''''''''''''''''
        ' By System''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''
        Dim ePlateType As enumPlateType
        ePlateType = CheckAndGetPlateType(pProcessSettings, oTemplateSetData.SurfaceBody)
       
        Select Case ePlateType
            Case SymToCenterLine
                ' exactly symmetrical to Center line and quadrandgle plate.' Using the Same algorithm as Normal Case
                Set oBasePlane = CreateAvgOfCornersPlane(pPlatePart, oTemplateSetData.Direction, oSurfaceBody, bBaseSide)
            Case PerpendicularXY 'Next To Stem/Stern
                On Error Resume Next
                Set oBasePlane = CreateBPlaneOfPerpendicularXY(pProcessSettings, pPlatePart, oSurfaceBody, bBaseSide)
                On Error GoTo ErrorHandler
            Case MostFlatPlate
                
                On Error Resume Next 'based on surface complexity it might fail
                oGeomHelper.GetMostPlanarFaceOfSheetBody2 oSurfaceBody, ratio, oBasePlane
                If Err.Number < 0 Then
                    LogMessage Err, MODULE, METHOD
                End If
                On Error GoTo ErrorHandler
                    
            Case NormalPlate 'quadrangle
                Set oBasePlane = CreateAvgOfCornersPlane(pPlatePart, oTemplateSetData.Direction, oSurfaceBody, bBaseSide)
            Case Box
                Set oBasePlane = CreateParallelAxisPlane(pPlatePart, oSurfaceBody)
        End Select
        
    ElseIf lBasePlaneType = 5193 Then
        '''''''''''''''''''''''''''''''''''''''''''''
        '"ParallelAxis"'''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''
        Set oBasePlane = CreateParallelAxisPlane(pPlatePart, oSurfaceBody)
        
    ElseIf lBasePlaneType = 5194 Then
        '''''''''''''''''''''''''''''''''''''''''''''
        ' "TrueNatural"'''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''
        Set oBasePlane = GetTrueNaturalBalance(pPlatePart)

    ElseIf (lBasePlaneType = 5195) Or (lBasePlaneType = 5196) Or (lBasePlaneType = 5197) Or (lBasePlaneType = 5198) Then
        '''''''''''''''''''''''''''''''''''''''''''''
        '''"UpperAftCorners"'''''''''''''''''''''''''
        '''"LowerAftCorners"'''''''''''''''''''''''''
        '''"LowerForeCorners"''''''''''''''''''''''''
        '''"UpperForeCorners"''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''
        
        Set oBasePlane = GetBasePlaneFromCorners(pPlatePart, strBasePlane, oSurfaceBody, bBaseSide)
        
    ElseIf lBasePlaneType = 5199 Then 'userdeinfed
        bUserDefined = True
        Set IJDMfgTemplateService_CreateBasePlane = Nothing
        Exit Function
    End If
    
    If Not oBasePlane Is Nothing Then
        Dim oTemplateHelper As IJMfgTemplateHelper
        Set oTemplateHelper = New MfgTemplateHelper
        
        Dim dRootX As Double, dRootY As Double, dRootZ As Double
        Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
        
        oBasePlane.GetRootPoint dRootX, dRootY, dRootZ
        oBasePlane.GetNormal dNormalX, dNormalY, dNormalZ
        
        Dim oGeomElements As IJElements
        Set oGeomElements = New JObjectCollection
        oGeomElements.Add pPlatePart
        
        ' The base plane(oBasePlane) that we got before is infinite one, it needs to be bounded by range of the plate with some tolerance
        Set IJDMfgTemplateService_CreateBasePlane = oTemplateHelper.MakeTemplateBasePlane(oGeomElements, _
                                                                                          dRootX, dRootY, dRootZ, _
                                                                                          dNormalX, dNormalY, dNormalZ)
    Else
        Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBP_FaliedToCreateBasePlane, , "RULES")
        Exit Function
    End If
       
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBP_FaliedToCreateBasePlane, , "RULES")
End Function


'********************************************************************************************************************************
' Function Name:    CreateControlLine
' Interface:        IJDMfgTemplateService
' Inputs:           PlatePart -- selected plate
'                   pProcessSettings -- Template settings, type, orientation, direction, side etc.
' Output:           Object - Baseplane IJPlane
' Assumption:
' Algorithm:        1. Check the Template Side which is BaseSide or OffsetSide
'                   2. Check the Template Type and If it is a special plate or not
'                       2.1. Template Type is PerPendicualrXY(Near to Stem/Stern) or
'                           Pentagonal plate or Triangular plate which edge is not parallel to Frame line,
'                           => Perpendicular XY Algorithm
'                       2.2. Template Type is Stem/Stern
'                           => Stem/Stern Type, There is no triangular and pentagonal plate(only quadrangle plate).
'                           => refer to Analysis- DefineCtlLineAndBPlane.doc
'                           => Exactly symmetrical to Center line
'                       2.3. Other Case
'
'********************************************************************************************************************************
Private Function IJDMfgTemplateService_CreateControlLine(ByVal pPlatePart As Object, ByVal pProcessSettings As Object, ByVal pBasePlane As Object, ByVal oTemplateSet As Object) As Object
Const METHOD = "IJDMfgControlLineService_CreateControlLine"
    On Error GoTo ErrorHandler

    Dim oTemplateSetData As TemplateSetData
    Set oTemplateSetData = New TemplateSetData
    oTemplateSetData.InitSettings pPlatePart, pProcessSettings, oTemplateSet
          
    Dim oControlLine As Object
    
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
    
    Dim bBaseSide As Boolean
    bBaseSide = oTemplateSetData.TemplateSide

    Dim strDirection As String
    strDirection = oTemplateSetData.Direction
    
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
        
    Dim ePlateType As enumPlateType
        
    '2. Check the Template Type and If it is a special plate or not
    ePlateType = CheckAndGetPlateType(pProcessSettings, oSurfaceBody)

    Select Case ePlateType
        Case SymToCenterLine 'Center line itself is the ControlLine
            
            On Error Resume Next
            Set oControlLine = CreateBCtlLineStemStern(oSurfaceBody)
            On Error GoTo ErrorHandler
            
        Case PerpendicularXY 'Next To Stem/Stern
            
            On Error Resume Next
            Set oControlLine = CreateBCtlLineOfPerpendicularXY(pPlatePart, strDirection, oSurfaceBody, pBasePlane, bBaseSide)
            If oControlLine Is Nothing Then 'try normal plate
                LogMessage Err, MODULE, METHOD, "Failed to create control line with CreateBCtlLineOfPerpendicularXY. Using alternate method 'CreateBCtlLineOfNormalPlate' now."
                Set oControlLine = CreateBCtlLineOfNormalPlate(pPlatePart, strDirection, oSurfaceBody, pBasePlane, bBaseSide)
            End If
            On Error GoTo ErrorHandler
        
        Case Box
             Set IJDMfgTemplateService_CreateControlLine = Nothing
             Exit Function
             
        Case Else 'normal plate
            On Error Resume Next
            Set oControlLine = CreateBCtlLineOfNormalPlate(pPlatePart, strDirection, oSurfaceBody, pBasePlane, bBaseSide)
            On Error GoTo ErrorHandler
    End Select
     
    If Not oControlLine Is Nothing Then
        Set IJDMfgTemplateService_CreateControlLine = oControlLine
    Else
        StrMfgLogError Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FaliedToCreateBaseControlLine, , "RULES"
        Err.Raise vbObjectError + TPL_CBCL_FaliedToCreateBaseControlLine
        Exit Function
    End If

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_CBCL_FaliedToCreateBaseControlLine, , "RULES")
End Function

Private Function IJDMfgTemplateService_CreateMarkingLine(ByVal pPlatePart As Object, ByVal pProcessSettings As Object, ByVal pBasePlane As Object) As Object

End Function

Private Function IJDMfgTemplateService_CreateSightLines(ByVal pTemplateSet As Object, ByVal numOfSightLine As Long, ByVal pMarkingSetting As Object) As Object
Const METHOD = "IJDMfgSightLineService_CreateSightLines"
On Error GoTo ErrorHandler

'    Dim oTemplateSet As IJDMfgTemplateSet
'    Set oTemplateSet = pTemplateSet
'
'    Dim oTemplateSetRpt As IJMfgTemplateSetReport
'    Set oTemplateSetRpt = oTemplateSet
'
'    Dim oPlatePart As IJPlatePart
'    Set oPlatePart = oTemplateSetRpt.GetParentPlatePart
'
'    Dim oSettingsHelper As MfgSettingsHelper
'    Dim strProgId As String
'    Dim oSightLineOffset As IJMfgTemplateSightLineRule
'    Dim dOffset As Double
'
'    Set oSettingsHelper = pMarkingSetting
'    strProgId = oSettingsHelper.GetProgIDFromAttr("TmSightLine")
'    Set oSightLineOffset = CreateObject(strProgId)
'    dOffset = oSightLineOffset.GetSightLineOffset(oPlatePart)
'
'    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
'    oTemplateSet.GetBasePlaneNormalVector dNormalX, dNormalY, dNormalZ
'
'    Dim oTemplates As IJElements
'    Dim nNumOfTemplates As Long
'    Set oTemplates = oTemplateSet.GetTemplates
'    nNumOfTemplates = oTemplates.count
'
'    Dim oOutputElems As IJElements
'    Set oOutputElems = New JObjectCollection
'
'    Dim dLength As Double
'    dLength = 0.01 'later, It's default for marking
'
'    Dim nCount As Long
'    For nCount = 1 To nNumOfTemplates
'
'        Dim oTemplate As MfgTemplate
'        Set oTemplate = oTemplates.Item(nCount)
'
'        Dim oTopLine As Line3d
'        Set oTopLine = oTemplate.TopLine
'
'        Dim dStartX As Double, dStartY As Double, dStartZ As Double
'        Dim dEndX As Double, dEndY As Double, dEndZ As Double
'
'        oTopLine.GetStartPoint dStartX, dStartY, dStartZ
'        oTopLine.GetEndPoint dEndX, dEndY, dEndZ
'
'        Dim oOffsetVector As DVector
'        Set oOffsetVector = New DVector
'        oOffsetVector.Set dEndX - dStartX, dEndY - dStartY, dEndZ - dStartZ
'        oOffsetVector.Length = 1#
'
'        Dim dStartSightLineX As Double, dStartSightLineY As Double, dStartSightLineZ As Double
'        dStartSightLineX = (dStartX + dEndX) / 2 + oOffsetVector.x * dOffset
'        dStartSightLineY = (dStartY + dEndY) / 2 + oOffsetVector.y * dOffset
'        dStartSightLineZ = (dStartZ + dEndZ) / 2 + oOffsetVector.z * dOffset
'
'        Dim dEndSightLineX As Double, dEndSightLineY As Double, dEndSightLineZ As Double
'        dEndSightLineX = dStartSightLineX + dNormalX * dLength
'        dEndSightLineY = dStartSightLineY + dNormalY * dLength
'        dEndSightLineZ = dStartSightLineZ + dNormalZ * dLength
'
'        Dim oSightLine As Line3d
'        Set oSightLine = New Line3d
'        oSightLine.DefineBy2Points dStartSightLineX, dStartSightLineY, dStartSightLineZ, _
'                                                dEndSightLineX, dEndSightLineY, dEndSightLineZ
'
'        oOutputElems.Add oSightLine
'
'        Set oTemplate = Nothing
'        Set oTopLine = Nothing
'        Set oOffsetVector = Nothing
'
'    Next nCount
'
'    Set IJDMfgTemplateService_CreateSightLines = oOutputElems
'
'    'Release Objects
'    Set oTemplateSet = Nothing
'    Set oTemplates = Nothing
'    Set oSightLine = Nothing
'    Set oOutputElems = Nothing

   Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'********************************************************************************************************************************
' Function Name:    CreateTemplateOutContour
'
' Interface:        IJDMfgTemplateService
'
' Inputs:           PlatePart -- selected plate
'                   TemplateSet -- TemplateSet object created by semantic
'                   ProcessSetting -- contains the settings for TemplateSet ex: Type, Side etc
'                   SketchingLines -- Collection of lines sketched by user
'
' Output:           Collection of Geom3d Objects.
'
' Assumption:       Within the output collection, a set of 4 objects make one template. each object represent
'                   the bottom curve, side, top , side of the template
'
' Algorithm:        1. GetSettingValues From ProcessRules
'                   2. Create Template planes based on TemplateType
'                   3. Intersect the Plane with Surface to get Bottom curve of template
'                   4. Apply Minimum Extension, Margin, Shrinkage, and Offset
'                   5. Adjust the baseplane to the min template height
'                   6. Project end points of Bottom curves to baseplane and connect points to form template contour
'********************************************************************************************************************************

Private Function IJDMfgTemplateService_CreateTemplateOutContour(ByVal pPlatePart As Object, ByVal pUnkTemplateSet As Object, ByVal pProcessSettings As Object, ByVal pSketchingLines As Object) As Object
    Const METHOD = "IJDMfgTemplateService_CreateTemplateOutContour"
    On Error Resume Next
        
    Dim oTemplateSetData As TemplateSetData
    Set oTemplateSetData = New TemplateSetData
    oTemplateSetData.InitSettings pPlatePart, pProcessSettings, pUnkTemplateSet
    
    Dim oTemplateContours As IJElements
    Dim oBasePlane As IJPlane
    Dim oTemplateSet As IJDMfgTemplateSet
    
    Set oTemplateSet = pUnkTemplateSet
    Set oBasePlane = oTemplateSet.GetBasePlane
    Set m_oPlatePart = pPlatePart
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   1.      Create Template planes based on TemplateType
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oTemplateDataElems As IJElements

    If oTemplateSetData.TemplateType = "UserDefined" Then
        CreateTemplateDataSketched pUnkTemplateSet, oTemplateSetData, pSketchingLines, oTemplateDataElems
        CreateTemplateDataPlanes pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        
    ElseIf oTemplateSetData.TemplateType = "Aft/Forward" Then
        CreateTemplateDataAftFore pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        CreateTemplateDataSketched pUnkTemplateSet, oTemplateSetData, pSketchingLines, oTemplateDataElems
        CreateTemplateDataPlanes pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        
    ElseIf oTemplateSetData.TemplateType = "Frame" Then
        CreateTemplateDataAftFore pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        CreateTemplateDataFrames pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        CreateTemplateDataSketched pUnkTemplateSet, oTemplateSetData, pSketchingLines, oTemplateDataElems
        CreateTemplateDataPlanes pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        
    ElseIf oTemplateSetData.TemplateType = "Perpendicular" Then
        'this same as Frame with TemplateDirection = Perpendicular
        CreateTemplateDataAftFore pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        CreateTemplateDataFrames pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        CreateTemplateDataSketched pUnkTemplateSet, oTemplateSetData, pSketchingLines, oTemplateDataElems
        CreateTemplateDataPlanes pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
    
    ' The next two are special do not create sketched or call CreateTemplateDataPlanes since they use
    ' different alogrithms to define the planes
    ElseIf oTemplateSetData.TemplateType = "PerpendicularXY" Then
        CreatePerpendicularXYTemplatePlanes pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
    ElseIf oTemplateSetData.TemplateType = "Stem/Stern" Then
        CreateStemSternTemplatePlanes pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
    ElseIf oTemplateSetData.TemplateType = "CenterLine" Then
        CreateTemplateDataAftFore pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        CreateTemplateDataFrames pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        CreateTemplateDataSketched pUnkTemplateSet, oTemplateSetData, pSketchingLines, oTemplateDataElems
        CreateTemplateDataPlanes pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
    ElseIf oTemplateSetData.TemplateType = "Box" Then
        CreateTemplateDataBox pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        CreateTemplateDataBoxSketched pUnkTemplateSet, oTemplateSetData, pSketchingLines, oTemplateDataElems
        CreateTemplateDataBoxPlanes pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
    ElseIf oTemplateSetData.TemplateType = "UserDefined Box" Then
        CreateTemplateDataBoxSketched pUnkTemplateSet, oTemplateSetData, pSketchingLines, oTemplateDataElems
        CreateTemplateDataBoxPlanes pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
    ElseIf oTemplateSetData.TemplateType = "UserDefined Box With Edges" Then
        CreateTemplateDataBoxWithEdges pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
        CreateTemplateDataBoxSketched pUnkTemplateSet, oTemplateSetData, pSketchingLines, oTemplateDataElems
        CreateTemplateDataBoxPlanes pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
    Else
         GoTo CleanUp
    End If
    

    '   2.      get Bottom curve of template
    CreateTemplateBottomLines pPlatePart, oTemplateDataElems, oTemplateSetData

    '   3.      Apply Minimum Extension, Margin, Shrinkage, and Offset
        '       Extension is controlled by the catalog value
    ApplyExtension oBasePlane, oTemplateDataElems, oTemplateSetData
   
        '       ApplyMargin is optional.
    Dim bApplymargin As Boolean
    bApplymargin = GetLongTypeUserChoice("ApplyMarginInProcess", pUnkTemplateSet)
    
    If bApplymargin = True Then
        '  ApplyMargin is optional.
        ApplyMargin pPlatePart, oTemplateDataElems
    End If
    
    ApplyModifiedBottomlines pPlatePart, pUnkTemplateSet, oTemplateSetData, oTemplateDataElems
    
        '       ApplyOffset is optional, ApplyOffset implements SHI's MinimumExtension Offset rule
        'm_dOffset = oTemplateSetData.Offset
        'm_dExt = oTemplateSetData.TemplateExtension
        'ApplyOffset oBasePlane, oTemplateDataElems
        
   '   4.      Adjust the baseplane to the min template height
    AdjustBasePlane pPlatePart, oTemplateDataElems, oTemplateSetData, oBasePlane, pUnkTemplateSet, pProcessSettings
         
    '5. Project end points of Bottom curves to baseplane and connect points to form template contour
    Set oTemplateContours = CreateTemplateContoursFromBottomCurves(oBasePlane, oTemplateDataElems)

    If Not oTemplateContours Is Nothing Then
        Set IJDMfgTemplateService_CreateTemplateOutContour = oTemplateContours
    Else
        Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToCreateTemplateContours, , "RULES")
        GoTo CleanUp
    End If

CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToCreateTemplateContours, , "RULES")
End Function


'********************************************************************************************************************************
' Function Name:    CreatePerpendicularXYTemplatePlanes'
' Interface:        <private>'
' Inputs:           oTemplateSet --  TemplateSet'
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the TemplateDataElems by setting the controlpoint, plane etc
'
' Assumption:       SkecthingLines are not required
'
' Algorithm:        1. Create XY plane
'                   2. Create Template Planes At Corner Points of plate on XY Plane
'                   3. Sort the created templates
'                   4. Interpolate templates based on offset and interval on XY plane
'********************************************************************************************************************************
Private Sub CreatePerpendicularXYTemplatePlanes(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
    Const METHOD = "CreatePerpendicularXYTemplatePlanes"
    On Error GoTo ErrorHandler
    
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        Set oTemplateDataElems = New JObjectCollection
    End If
        
    '1. Create XY plane
    Dim oPlane As Plane3d
    Set oPlane = New Plane3d
    oPlane.SetNormal 0, 0, 1
    
    CreateTemplatePlanesAtCornerPoints oTemplateSet, oPlane, oTemplateSetData, oTemplateDataElems
    
    SortTemplateDataElems oTemplateSet, oTemplateDataElems
    
    CreateTemplatePlanesInterPolated oTemplateSet, oPlane, oTemplateSetData, oTemplateDataElems
    
CleanUp:
    Set oPlane = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


'********************************************************************************************************************************
' Function Name:    CreatePerpendicularXYTemplatePlanes'
' Interface:        <private>'
' Inputs:           oTemplateSet --  TemplateSet'
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the TemplateDataElems by setting the controlpoint, plane etc
'
' Assumption:       SkecthingLines are not required
'
' Algorithm:        1. Create XZ plane
'                   2. Create Template Planes At Corner Points of plate on XZ Plane
'                   3. Create Template Planes At Inflection Points of plate on XZ Plane
'                   4. Sort the created templates
'                   5. Interpolate templates based on offset and interval on XZ plane
'********************************************************************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Stem/Stern shell plates are exactly symmetrical to the center line of ship.
'   There is no triangle, no pentagonal plate on the stem/stern part
'       1. Project Symmetric Part of Plate onto XZ Plane
'       2. Get Inflection Points on ControlLine
'       3. Get Interpolation Points on ControlLine
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateStemSternTemplatePlanes(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
    Const METHOD = "CreateStemSternTemplatePlanes"
    On Error GoTo ErrorHandler
    
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        Set oTemplateDataElems = New JObjectCollection
    End If
    
    Dim oPlane As Plane3d
    Set oPlane = New Plane3d
    oPlane.SetNormal 1, 0, 0
    
    CreateTemplatePlanesAtCornerPoints oTemplateSet, oPlane, oTemplateSetData, oTemplateDataElems
    oPlane.SetNormal 0, 1, 0
    CreateTemplatePlanesAtInflectionPt oTemplateSet, oPlane, oTemplateSetData, oTemplateDataElems
    SortTemplateDataElems oTemplateSet, oTemplateDataElems
    CreateTemplatePlanesInterPolated oTemplateSet, oPlane, oTemplateSetData, oTemplateDataElems

CleanUp:
    Set oPlane = Nothing
       
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub



'********************************************************************************************************************************
' Function Name:    CreateTemplateBottomLines
'
' Interface:        <private>
'
' Inputs:           PlatePart -- selected plate
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'
' Output:           <indirect> modifies the TemplateDataElems by setting the BottomLine property
'
' Assumption:       TemplateDataElems is already created by CreateTemplatePlanes
'
' Algorithm:        1. Get Surface from Plate Part for given side
'                   2. Intersect surface with TemplatePlane OR if template is on planar edge, use the edge itself
'                   3. Save intersection curve as BottomLine in TemplateData object
'********************************************************************************************************************************
Private Sub CreateTemplateBottomLines(oPlatePart As IJPlatePart, oTemplateDataElems As IJElements, oTemplateSetData As TemplateSetData)
    Const METHOD = "CreateTemplateBottomLines"
    On Error GoTo ErrorHandler
    
    Dim oSurfaceBody As IUnknown
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
    
    Dim oValidTemplateDataElems As IJElements
    Set oValidTemplateDataElems = New JObjectCollection
    
    Dim nIndex As Long
    Dim oIntersectionCurve As IUnknown
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    Dim oTemplateData As TemplateData
    
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    
    For nIndex = 1 To oTemplateDataElems.count
        Set oTemplateData = oTemplateDataElems.Item(nIndex)
        If oTemplateData.TemplateType = Edge Then
            Set oIntersectionCurve = oTemplateData.DefiningObj
        Else 'Adjust the skecthed curve botomline based on the plane, otherwise template will not be planar
            ' for all template types just intersect the
            ' template plane with the surface body to get a bottom line
            m_oMfgGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oTemplateData.Plane, oIntersectionCurve, oStartPos, oEndPos
        End If
        If Not oIntersectionCurve Is Nothing Then
            Set oTemplateData.BottomLine = oIntersectionCurve
        End If
        
        If Not oTemplateData.BottomLine Is Nothing Then
            oValidTemplateDataElems.Add oTemplateData
        End If
        
        Set oIntersectionCurve = Nothing
        Set oStartPos = Nothing
        Set oEndPos = Nothing
        Set oTemplateData = Nothing
    Next
 
    oTemplateDataElems.Clear
    
    oTemplateDataElems.AddElements oValidTemplateDataElems
 
    Set oValidTemplateDataElems = Nothing
    Set oSurfaceBody = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToCreateBottomLines, , "RULES")
End Sub

'********************************************************************************************************************************
' Function Name:    CreateTemplateContoursFromBottomCurves
' Interface:        <private>'
' Inputs:           oBasePlane
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           IJElements -- lines representing the 'Bottom-Side1-Top-Side2 of the template
'
' Assumption:       Bottom-Side1-Top-Side2 order is important
'
' Algorithm:        1. Project end points of Bottom curves to baseplane and connect points to form template contour
'********************************************************************************************************************************
Public Function CreateTemplateContoursFromBottomCurves(oBasePlane As IJPlane, oTemplateDataElems As IJElements) As IJElements
    Const METHOD = "CreateTemplateContoursFromBottomCurves"
    On Error GoTo ErrorHandler
    
    Dim nIndex As Long
    Dim oTemplateColl As IJElements
    Set oTemplateColl = New JObjectCollection
     
    Dim dBPNormalX As Double, dBPNormalY As Double, dBPNormalZ As Double
    Dim dBPRootX As Double, dBPRootY As Double, dBPRootZ As Double
    
    oBasePlane.GetRootPoint dBPRootX, dBPRootY, dBPRootZ
    oBasePlane.GetNormal dBPNormalX, dBPNormalY, dBPNormalZ

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New GeometryFactory

    Dim oLine As Line3d
    Set oLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, dBPRootX, dBPRootY, dBPRootZ, dBPNormalX, dBPNormalY, dBPNormalZ, 1)
        
    Dim oGeomMisc As DGeomOpsMisc
    Set oGeomMisc = New DGeomOpsMisc
    
    Dim oWireBody As Object
    oGeomMisc.CreateModelGeometryFromGType Nothing, oLine, Nothing, oWireBody
    
    Dim oTemplateData As TemplateData
    Dim oSBottomLinePos As IJDPosition, oEBottomLinePos As IJDPosition
    Dim oTemplatePlane As IJPlane
    Dim oInfBasePlane As IJPlane
    Set oInfBasePlane = New Plane3d
 
    oInfBasePlane.DefineByPointNormal dBPRootX, dBPRootY, dBPRootZ, dBPNormalX, dBPNormalY, dBPNormalZ
    
    Dim lNoDirTempCount As Long
    For nIndex = 1 To oTemplateDataElems.count
        Dim oTemplateInfo As IJElements
        Set oTemplateInfo = New JObjectCollection
        Set oTemplateData = oTemplateDataElems.Item(nIndex)
        'Get BottomLine EndPoints
        GetEndPoints oTemplateData.BottomLine, oSBottomLinePos, oEBottomLinePos
        Set oTemplatePlane = oTemplateData.Plane
 
        Dim oProjectedLine As Object
        Set oProjectedLine = GetProjectedCurveOnPlane(oWireBody, oTemplatePlane)
        
        Dim oStartPos As IJDPosition, oEndPos As IJDPosition
        GetEndPoints oProjectedLine, oStartPos, oEndPos

        Dim oProjectionVec As IJDVector
        Set oProjectionVec = oEndPos.Subtract(oStartPos)
        oProjectionVec.Length = 1

        Dim oSTopLinePos As IJDPosition, oETopLinePos As IJDPosition
        Set oSTopLinePos = GetInterSectionPointBtwLineAndPlane(oSBottomLinePos, oProjectionVec, oInfBasePlane)
        Set oETopLinePos = GetInterSectionPointBtwLineAndPlane(oEBottomLinePos, oProjectionVec, oInfBasePlane)

                
        'sometimes Template may have zero minimum height and side lines may not be valid
        Dim oSide1 As Line3d, oSide2 As Line3d, oTopLine As Line3d
        If oSTopLinePos.DistPt(oSBottomLinePos) > 0.0001 Then
            Set oSide1 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oSBottomLinePos.x, oSBottomLinePos.y, oSBottomLinePos.z, _
                                                                    oSTopLinePos.x, oSTopLinePos.y, oSTopLinePos.z)
        Else
            'there is no side line, Make Start Top Line pos same as start bottom line pos
            oSTopLinePos.Set oSBottomLinePos.x, oSBottomLinePos.y, oSBottomLinePos.z
            
        End If
        
        If oETopLinePos.DistPt(oEBottomLinePos) > 0.0001 Then
            Set oSide2 = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oETopLinePos.x, oETopLinePos.y, oETopLinePos.z, _
                                                                    oEBottomLinePos.x, oEBottomLinePos.y, oEBottomLinePos.z)
        Else
            'there is no side line, Make end Top Line pos same as end bottom line pos
            oETopLinePos.Set oEBottomLinePos.x, oEBottomLinePos.y, oEBottomLinePos.z
        End If
        
        Set oTopLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oSTopLinePos.x, oSTopLinePos.y, oSTopLinePos.z, _
                                                                    oETopLinePos.x, oETopLinePos.y, oETopLinePos.z)
        
        'check if topline and bottom line are
        Dim oGeomHelper As MfgGeomHelper
        Set oGeomHelper = New MfgGeomHelper
        Dim bOverlap As Double
        bOverlap = oGeomHelper.CheckOverlapBetweenTwoCurvesWithinTol(oTemplateData.BottomLine, oTopLine, TOL_MM)  '5mm tolerance
        
        If bOverlap = True Then GoTo EndOfFor 'discard this template, it is just a line
        
        oTemplateInfo.Add oTemplateData.BottomLine
        oTemplateInfo.Add oTopLine
        If Not oSide1 Is Nothing Then
            oTemplateInfo.Add oSide1
        End If
        If Not oSide2 Is Nothing Then
            oTemplateInfo.Add oSide2
        End If
        
        'Add GroupIndex and Sketched information
        Dim oTempPropColl As Collection
        Set oTempPropColl = New Collection
        
        Dim lGroupIndex As Long
        If oTemplateData.GroupType = Primary Then
            lGroupIndex = 1
        ElseIf oTemplateData.GroupType = Secondary Then
            lGroupIndex = 2
        Else
            lNoDirTempCount = lNoDirTempCount + 1
            lGroupIndex = 2 + lNoDirTempCount
        End If
        
        Dim strKey As String
        If oTemplateData.TemplateType = Sketched Then
            strKey = "Sketched"
        ElseIf oTemplateData.TemplateType = SystemTemplateWithModifiedBL Then
            strKey = "SystemModifiedBL"
        ElseIf oTemplateData.TemplateType = SketchedTemplateWithModifiedBL Then
            strKey = "SketchedModifiedBL"
        ElseIf oTemplateData.TemplateType = TemplateWithAddedBL Then
            strKey = "AddedBL"
        Else 'system or edge
            strKey = ""
        End If
        
        oTempPropColl.Add lGroupIndex
        oTempPropColl.Add strKey
        
        oTemplateInfo.Add oTempPropColl
        
        oTemplateColl.Add oTemplateInfo
        
        Set oSide1 = Nothing
        Set oSide2 = Nothing
        Set oTopLine = Nothing
        Set oTemplatePlane = Nothing
        Set oProjectedLine = Nothing
        Set oTemplateData = Nothing
        Set oSTopLinePos = Nothing
        Set oETopLinePos = Nothing
        Set oProjectionVec = Nothing
EndOfFor:
    Next nIndex

    Set CreateTemplateContoursFromBottomCurves = oTemplateColl

CleanUp:
    Set oGeomFactory = Nothing
    Set oLine = Nothing
    Set oGeomMisc = Nothing
    Set oTemplateData = Nothing
    Set oTemplatePlane = Nothing
    Set oProjectedLine = Nothing
    Set oProjectionVec = Nothing
    Set oSTopLinePos = Nothing
    Set oETopLinePos = Nothing
    Set oSide1 = Nothing
    Set oSide2 = Nothing
    Set oTopLine = Nothing
    Set oInfBasePlane = Nothing
   
Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function


'********************************************************************************************************************************
' Function Name:    AdjustBasePlane
'
' Interface:        <private>
'
' Inputs:           PlatePart -- selected plate
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   BasePlane -- Baseplane defined for the templateset
'
' Output:           <indirect> modifies the baseplane based on the min height of template
'
' Assumption:       BottomLines are already created
'
' Algorithm:        1. Get Baseplane Normal
'                   2. Contstruct min box for TemplateBottomLines using baseplane normal and MfgGeomHelper routine
'                   3. Find the center point of min box's plane
'                   4. adjust root point of plane to new location m_dMinHeight distance from min box plane
'********************************************************************************************************************************'

Private Sub AdjustBasePlane(pPlatePart As Object, oTemplateDataElems As IJElements, oTemplateSetData As TemplateSetData, oBasePlane As IJPlane, oTemplateSet As Object, oProcessSettings As Object)
    Const METHOD = "AdjustBasePlane"
    On Error GoTo ErrorHandler
    
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    
    Dim dBasePlaneRootX As Double
    Dim dBasePlaneRootY As Double
    Dim dBasePlaneRootZ As Double
    Dim dBasePlaneNormalX As Double
    Dim dBasePlaneNormalY As Double
    Dim dBasePlaneNormalZ As Double
    Dim dPlateRootX As Double
    Dim dPlateRootY As Double
    Dim dPlateRootZ As Double
    Dim dPlateNormalX As Double
    Dim dPlateNormalY As Double
    Dim dPlateNormalZ As Double
    Dim oPlateNormalVec As IJDVector
    Dim oBasePlaneNormalVec As IJDVector
    Dim oMinBoxPoints As IJElements
    Dim oMinBoxRootPoint As IJDPosition
    Dim Points(1 To 4) As IJDPosition
    Dim dBoxHeight As Double
    
'   1. Get Baseplane Normal
    oBasePlane.GetNormal dBasePlaneNormalX, dBasePlaneNormalY, dBasePlaneNormalZ
    Set oBasePlaneNormalVec = New DVector
    oBasePlaneNormalVec.Set dBasePlaneNormalX, dBasePlaneNormalY, dBasePlaneNormalZ
    
    Dim oTemplateBottomLines As IJElements
    Set oTemplateBottomLines = New JObjectCollection
    Dim oTemplateData As TemplateData
    Dim nIndex As Long
    
    For nIndex = 1 To oTemplateDataElems.count
        Set oTemplateData = oTemplateDataElems.Item(nIndex)
        oTemplateBottomLines.Add oTemplateData.BottomLine
        Set oTemplateData = Nothing
    Next
    
    If oTemplateBottomLines.count = 0 Then
        'No bottom lines -> cannot adjust base plane
        Exit Sub
    End If
    
'   2. Contstruct min box for TemplateBottomLines using baseplane normal and MfgGeomHelper routine
    Set oMinBoxPoints = m_oMfgGeomHelper.GetGeometryMinBoxByVector(oTemplateBottomLines, oBasePlaneNormalVec)
    
    Set Points(1) = oMinBoxPoints.Item(1)
    Set Points(2) = oMinBoxPoints.Item(2)
    Set Points(3) = oMinBoxPoints.Item(3)
    Set Points(4) = oMinBoxPoints.Item(4)
    
    Dim oBoxNormalVec As IJDVector
    Set oBoxNormalVec = New DVector
    
    oBoxNormalVec.Set Points(4).x - Points(1).x, Points(4).y - Points(1).y, Points(4).z - Points(1).z
    oBoxNormalVec.Length = 1
        
'   3. Find the center point of min box's plane
    ' set the new rootpoint for baseplane at center of min box plane (center of diaganol from point 1 to 3)
    Set oMinBoxRootPoint = New DPosition
    oMinBoxRootPoint.x = (Points(1).x + Points(3).x) / 2
    oMinBoxRootPoint.y = (Points(1).y + Points(3).y) / 2
    oMinBoxRootPoint.z = (Points(1).z + Points(3).z) / 2
    
    dBoxHeight = Points(1).DistPt(Points(4))
    
'   4. adjust root point of plane to new location m_dMinHeight distance from min box plane
    ' negative sign since we need to move in opposite direction as the normal
    ' the base plane normal is always pointed towards the platepart
    m_oMfgGeomHelper.GetPlatePartAvgPointAvgNormal oTemplateSetData.SurfaceBody, oTemplateSetData.IsOnBaseSide, dPlateRootX, dPlateRootY, dPlateRootZ, dPlateNormalX, dPlateNormalY, dPlateNormalZ
    Set oPlateNormalVec = New DVector
    
    ' If the SurfaceType of template is PART_SURFACE_BASE then calculate the normal of plate surface again
    If oTemplateSet.SurfaceType = PART_SURFACE_BASE Then
        Dim oPortSurfBody As IJSurfaceBody
        Set oPortSurfBody = GetPortSurface(pPlatePart, oProcessSettings.TemplateSide)
        
        If Not oPortSurfBody Is Nothing Then
            m_oMfgGeomHelper.GetPlatePartAvgPointAvgNormal oPortSurfBody, oTemplateSetData.IsOnBaseSide, dPlateRootX, dPlateRootY, dPlateRootZ, dPlateNormalX, dPlateNormalY, dPlateNormalZ
        End If
    End If
    
    oPlateNormalVec.Set dPlateNormalX, dPlateNormalY, dPlateNormalZ
    oPlateNormalVec.Length = 1
    
    ' DI-CP·82454  Template does not change side (Baseside/Offsetside)
    ' TR-CP·83461  Template sets not created correctly for parameter side value OffsetSide
    'check which side to extend.
        
    If oBoxNormalVec.Dot(oPlateNormalVec) < 0 Then
        dBasePlaneRootX = oMinBoxRootPoint.x - oBoxNormalVec.x * oTemplateSetData.MinHeight
        dBasePlaneRootY = oMinBoxRootPoint.y - oBoxNormalVec.y * oTemplateSetData.MinHeight
        dBasePlaneRootZ = oMinBoxRootPoint.z - oBoxNormalVec.z * oTemplateSetData.MinHeight
    Else
        dBasePlaneRootX = oMinBoxRootPoint.x + oBoxNormalVec.x * (oTemplateSetData.MinHeight + dBoxHeight)
        dBasePlaneRootY = oMinBoxRootPoint.y + oBoxNormalVec.y * (oTemplateSetData.MinHeight + dBoxHeight)
        dBasePlaneRootZ = oMinBoxRootPoint.z + oBoxNormalVec.z * (oTemplateSetData.MinHeight + dBoxHeight)
    End If
    
    If (oBasePlaneNormalVec.Dot(oPlateNormalVec) > 0) Then
        'reverse the normal to point towards the plate for consistency
        oBasePlane.SetNormal (-1 * oBasePlaneNormalVec.x), (-1 * oBasePlaneNormalVec.y), (-1 * oBasePlaneNormalVec.z)
    End If
        
    oBasePlane.SetRootPoint dBasePlaneRootX, dBasePlaneRootY, dBasePlaneRootZ
    
    '    MsgBox "IsOnBaseSide : " & oTemplateSetData.IsOnBaseSide & vbNewLine & vbTab _
    '         & "Plate Root   : " & dPlateRootX & " - " & dPlateRootY & " - " & dPlateRootZ & " - " & vbNewLine & vbTab _
    '         & "Plate Normal : " & dPlateNormalX & " - " & dPlateNormalY & " - " & dPlateNormalZ & " - " & vbNewLine & vbTab _
    '         & "Plane Root   : " & dBasePlaneRootX & " - " & dBasePlaneRootY & " - " & dBasePlaneRootZ & " - " & vbNewLine & vbTab _
    '         & "Plane Normal : " & oBasePlaneNormalVec.x & " - " & oBasePlaneNormalVec.y & " - " & oBasePlaneNormalVec.z
       
CleanUp:
    Set Points(1) = Nothing
    Set Points(2) = Nothing
    Set Points(3) = Nothing
    Set Points(4) = Nothing
    Set oMinBoxPoints = Nothing
    Set oMinBoxRootPoint = Nothing
    Set oBasePlaneNormalVec = Nothing
    Set oPlateNormalVec = Nothing
    Set oTemplateBottomLines = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


'********************************************************************************************************************************
' Function Name:    ApplyExtension
'
' Interface:        <private>
'
' Inputs:           TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'
' Output:           <indirect> modifies the bottomline of each object in TemplateDataElems
'
' Assumption:       BottomLines are already created
'
' Algorithm:        Apply the extension on both ends of the bottom lines
'                   If the m_dTemplateExtesion is less than or equal to zero just copy the curve with no extension
'                   ExtraPolateCurve will throw errors if the extension value is zero
'********************************************************************************************************************************'

Private Sub ApplyExtension(ByRef oBasePlane As IJPlane, ByRef oTemplateDataElems As IJElements, oTemplateSetData As TemplateSetData)
    Const METHOD = "ApplyExtension"
    On Error GoTo ErrorHandler

    Dim nIndex As Long, i As Long, j As Long
    Dim oTempExtendedBtmCurve As IUnknown
    Dim oTemplateData As TemplateData
    Dim dExt As Double, dMinExt As Double, dOffset As Double
    
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    dOffset = oTemplateSetData.Offset
    
    If dOffset > 0 Then 'Variable extn
        dMinExt = oTemplateSetData.TemplateMinimumExtension
        
        ApplyExtnByOffset oBasePlane, oTemplateDataElems, dOffset, dMinExt
    Else 'const extn
        dExt = oTemplateSetData.TemplateExtension
        
        If dExt > 0 Then
            For nIndex = 1 To oTemplateDataElems.count
                'Extend the bottom curve by m_TemplateExtension
                Set oTemplateData = oTemplateDataElems.Item(nIndex)
                Set oTempExtendedBtmCurve = m_oMfgGeomHelper.ExtraPolateCurve(oTemplateData.BottomLine, 1, 0, 2, Nothing, Nothing, dExt)
                Set oTemplateData.BottomLine = Nothing
                Set oTemplateData.BottomLine = oTempExtendedBtmCurve
                Set oTempExtendedBtmCurve = Nothing
                Set oTemplateData = Nothing
            Next
        End If
    End If
    
    Exit Sub
ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToApplyExtToLocationMarkLines, , "RULES")
End Sub


'********************************************************************************************************************************
' Function Name:    ApplyMargin
'
' Interface:        <private>
'
' Inputs:           PlatePart -- Plate to get margin and shrinkage values
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the bottomline of each object in TemplateDataElems
'
' Assumption:       BottomLines are already created
'
' Algorithm:        1. Get Margin and Shrinkage from the plate
'                   2. for each bottomline check if it intersects the edge and get corresponding margin
'********************************************************************************************************************************'

Private Sub ApplyMargin(ByVal oPlatePart As Object, ByRef oTemplateDataElems As IJElements)
    Const METHOD = "ApplyMargin"
    On Error GoTo ErrorHandler
        
    Dim i As Long

    Dim oTemplateData As TemplateData
        
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    
    Dim oExtendedCurve1 As IUnknown
    On Error Resume Next
    
    '2. for each bottomline check if it intersects the edge and get corresponding margin
    For i = 1 To oTemplateDataElems.count
        Set oTemplateData = oTemplateDataElems.Item(i)

        Set oExtendedCurve1 = m_oMfgGeomHelper.ExtendCurveWithMargins(oTemplateData.BottomLine, oPlatePart)

        Set oTemplateData.BottomLine = oExtendedCurve1

        Set oTemplateData = Nothing
        Set oExtendedCurve1 = Nothing
    Next i

CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToApplyMargin, , "RULES")
End Sub

'********************************************************************************************************************************
' SHI's offset requirement. Extend bottomcurve by one offset if (offset - overhang) < MINIMUMEXTENSION
'********************************************************************************************************************************'
Public Sub ApplyOffset(oBasePlane As IJPlane, oTemplateDataElems As IJElements)
    Const METHOD = "ApplyOffset"
    On Error Resume Next
    
    If m_dExt > 0 Then
        
        Dim nIndex As Long
    
        Dim oTemplateContours As IJElements
        Set oTemplateContours = New JObjectCollection
         
        Dim dBPNormalX As Double, dBPNormalY As Double, dBPNormalZ As Double
        Dim dBPRootX As Double, dBPRootY As Double, dBPRootZ As Double
        
        oBasePlane.GetRootPoint dBPRootX, dBPRootY, dBPRootZ
        oBasePlane.GetNormal dBPNormalX, dBPNormalY, dBPNormalZ
    
    
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New GeometryFactory
    
        Dim oLine As Line3d
        Set oLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, dBPRootX, dBPRootY, dBPRootZ, dBPNormalX, dBPNormalY, dBPNormalZ, 1)
        
        Dim oGeomMisc As DGeomOpsMisc
        Set oGeomMisc = New DGeomOpsMisc
        
        Dim oWireBody As Object
        oGeomMisc.CreateModelGeometryFromGType Nothing, oLine, Nothing, oWireBody
        
        Dim oTemplateData As TemplateData
        Dim oSBottomLinePos As IJDPosition, oEBottomLinePos As IJDPosition
        Dim oTemplatePlane As IJPlane
        Dim oInfBasePlane As IJPlane
        Set oInfBasePlane = New Plane3d
     
        oInfBasePlane.DefineByPointNormal dBPRootX, dBPRootY, dBPRootZ, dBPNormalX, dBPNormalY, dBPNormalZ
        
        For nIndex = 1 To oTemplateDataElems.count
            Set oTemplateData = oTemplateDataElems.Item(nIndex)
            If oTemplateData Is Nothing Then GoTo SkipThisTemplate
            
            'Get BottomLine EndPoints
            GetEndPoints oTemplateData.BottomLine, oSBottomLinePos, oEBottomLinePos
            Set oTemplatePlane = oTemplateData.Plane
            If oTemplatePlane Is Nothing Then GoTo SkipThisTemplate
     
            Dim oProjectedLine As Object
            Set oProjectedLine = GetProjectedCurveOnPlane(oWireBody, oTemplatePlane)
            If oProjectedLine Is Nothing Then GoTo SkipThisTemplate
            
            Dim oStartPos As IJDPosition, oEndPos As IJDPosition
            GetEndPoints oProjectedLine, oStartPos, oEndPos
            
            'Hoonsik modified ************************************* 2006.02.17
            Dim oSideLineVec As IJDVector
            Set oSideLineVec = oEndPos.Subtract(oStartPos)
            If oSideLineVec Is Nothing Then GoTo SkipThisTemplate
            oSideLineVec.Length = 1
          
            Dim oControlPoint As IJDPosition
            Dim oCtlPtOnBPlane As IJDPosition
            
            Set oControlPoint = oTemplateData.ControlPoint
            If oControlPoint Is Nothing Then GoTo SkipThisTemplate
            
            Dim oTopLineTmp As IJLine
            Dim dX As Double, dY As Double, dZ As Double
            Dim dirX As Double, dirY As Double, dirZ As Double
            Dim oPt1 As Object, oPt2 As Object
            
            m_oMfgGeomHelper.IntersectSurfaceWithPlane oTemplatePlane, oInfBasePlane, oTopLineTmp, oPt1, oPt2
            If oTopLineTmp Is Nothing Then GoTo SkipThisTemplate
            
            oControlPoint.Get dX, dY, dZ
    
            oTopLineTmp.GetDirection dirX, dirY, dirZ
            
            Dim oBCLPlane As IJPlane
            'make a transient plane to get the control point on the top line
            m_oMfgGeomHelper.MakeTransientPlane dX, dY, dZ, dirX, dirY, dirZ, oBCLPlane
            If oBCLPlane Is Nothing Then GoTo SkipThisTemplate
    
            'get the projected control point on the base plane
            Set oCtlPtOnBPlane = m_oMfgGeomHelper.IntersectCurveWithPlane(oTopLineTmp, oBCLPlane)
            If oCtlPtOnBPlane Is Nothing Then GoTo SkipThisTemplate
                   
            Dim oTopLineVec As IJDVector
            Dim oFirstPtOfTopLine As IJDPosition
            Dim oSecondPtOfTopLine As IJDPosition
            Dim oTemplateSideLineVec As IJDVector
            
            Set oTemplateSideLineVec = oCtlPtOnBPlane.Subtract(oControlPoint)
            
            Dim oFirstSideLine As Line3d, oSecondSideLine As Line3d
            
            Set oFirstSideLine = CreateInfiniteLine(oSBottomLinePos, oTemplateSideLineVec)
            Set oFirstPtOfTopLine = m_oMfgGeomHelper.IntersectCurveWithPlane(oFirstSideLine, oInfBasePlane)
            If oFirstPtOfTopLine Is Nothing Then GoTo SkipThisTemplate
    
            Set oSecondSideLine = CreateInfiniteLine(oEBottomLinePos, oTemplateSideLineVec)
            Set oSecondPtOfTopLine = m_oMfgGeomHelper.IntersectCurveWithPlane(oSecondSideLine, oInfBasePlane)
            If oSecondPtOfTopLine Is Nothing Then GoTo SkipThisTemplate
    
            'Get the TopLine direction
            Set oTopLineVec = oSecondPtOfTopLine.Subtract(oFirstPtOfTopLine)
            oTopLineVec.Length = 1
                            
            Dim dExtAtFirstPtOfTLine As Double, dExtAtSecondPtOfTLine As Double
            Dim dDistBtwFirstAndCtl As Double, dDistBtwSecondAndCtl As Double
    
            Dim nNumber As Integer
            Dim dExtension As Double
            
            dDistBtwFirstAndCtl = oFirstPtOfTopLine.DistPt(oCtlPtOnBPlane)
    '        MsgBox "DistA: " & Round(dDistBtwFirstAndCtl, 3) * 1000
      
            nNumber = Fix((Round(dDistBtwFirstAndCtl, 3) * 1000) / (Round(m_dOffset, 3) * 1000))
    '        MsgBox "L: " & Round(dDistBtwFirstAndCtl - m_dOffset * nNumber, 3) * 1000
            dExtension = m_dOffset - (dDistBtwFirstAndCtl - m_dOffset * nNumber)
            
            'Check if the First Ext. is less than m_dExt, then add m_dOffset value.
            If Round(dExtension, 3) * 1000 > m_dExt Then
                dExtAtFirstPtOfTLine = dExtension
            Else
                dExtAtFirstPtOfTLine = m_dOffset + dExtension
            End If
    '        MsgBox "dExtAtFirstPtOfTLine: " & Round(dExtAtFirstPtOfTLine, 3) * 1000
    
            dDistBtwSecondAndCtl = oSecondPtOfTopLine.DistPt(oCtlPtOnBPlane)
    '        MsgBox "DistB: " & Round(dDistBtwSecondAndCtl, 3) * 1000
            
            nNumber = Fix((Round(dDistBtwSecondAndCtl, 3) * 1000) / (Round(m_dOffset, 3) * 1000))
            dExtension = m_dOffset - (dDistBtwSecondAndCtl - m_dOffset * nNumber)
    '        MsgBox "U: " & Round(dDistBtwSecondAndCtl - m_dOffset * nNumber, 3) * 1000
            
            'Check if the Second Ext. is less than m_dExt, then add m_dOffset value.
            If Round(dExtension, 3) * 1000 > m_dExt Then
                dExtAtSecondPtOfTLine = dExtension
            Else
                dExtAtSecondPtOfTLine = m_dOffset + dExtension
            End If
            
    '/////////
    
    
            Dim oFirstSidePlane As IJPlane, oSecondSidePlane As IJPlane
            Dim dFirstExtDistOfBtmCrv As Double, dSecondExtDistOfBtmCrv As Double
            Dim oStPointOfBtmCrv As IJDPosition, oEdPointOfBtmCrv As IJDPosition
            
            'Calculate the first point of the top line
            oFirstPtOfTopLine.x = oFirstPtOfTopLine.x - oTopLineVec.x * dExtAtFirstPtOfTLine
            oFirstPtOfTopLine.y = oFirstPtOfTopLine.y - oTopLineVec.y * dExtAtFirstPtOfTLine
            oFirstPtOfTopLine.z = oFirstPtOfTopLine.z - oTopLineVec.z * dExtAtFirstPtOfTLine
            
            m_oMfgGeomHelper.MakeTransientPlane oFirstPtOfTopLine.x, oFirstPtOfTopLine.y, oFirstPtOfTopLine.z, oTopLineVec.x, oTopLineVec.y, oTopLineVec.z, oFirstSidePlane
            
            Dim oFirstExtPosOfBtmCrv As IJDPosition, oSecondExtPosOfBtmCrv As IJDPosition
            Dim oTempExtendedBtmCurve1 As IUnknown, oTempExtendedBtmCurve2 As IUnknown
            
            Set oTempExtendedBtmCurve1 = m_oMfgGeomHelper.ExtraPolateCurve(oTemplateData.BottomLine, 1, 0, 2, Nothing, Nothing, 0.5)
            If oTempExtendedBtmCurve1 Is Nothing Then GoTo SkipThisTemplate
            
            Set oFirstExtPosOfBtmCrv = m_oMfgGeomHelper.IntersectCurveWithPlane(oTempExtendedBtmCurve1, oFirstSidePlane)
            If oFirstExtPosOfBtmCrv Is Nothing Then GoTo SkipThisTemplate
            
            m_oMfgGeomHelper.GetLengthBet2Points oTempExtendedBtmCurve1, oSBottomLinePos, oFirstExtPosOfBtmCrv, dFirstExtDistOfBtmCrv
            
            'Calculate the second point of the top line
            oSecondPtOfTopLine.x = oSecondPtOfTopLine.x + oTopLineVec.x * dExtAtSecondPtOfTLine
            oSecondPtOfTopLine.y = oSecondPtOfTopLine.y + oTopLineVec.y * dExtAtSecondPtOfTLine
            oSecondPtOfTopLine.z = oSecondPtOfTopLine.z + oTopLineVec.z * dExtAtSecondPtOfTLine
               
            m_oMfgGeomHelper.MakeTransientPlane oSecondPtOfTopLine.x, oSecondPtOfTopLine.y, oSecondPtOfTopLine.z, oTopLineVec.x, oTopLineVec.y, oTopLineVec.z, oSecondSidePlane
            Set oSecondExtPosOfBtmCrv = m_oMfgGeomHelper.IntersectCurveWithPlane(oTempExtendedBtmCurve1, oSecondSidePlane)
            If oSecondExtPosOfBtmCrv Is Nothing Then GoTo SkipThisTemplate
            
            m_oMfgGeomHelper.GetLengthBet2Points oTempExtendedBtmCurve1, oEBottomLinePos, oSecondExtPosOfBtmCrv, dSecondExtDistOfBtmCrv
       
            If Abs(dFirstExtDistOfBtmCrv) > EPSILON Then
                Set oTempExtendedBtmCurve2 = m_oMfgGeomHelper.ExtraPolateCurve(oTemplateData.BottomLine, 1, 0, 0, Nothing, Nothing, dFirstExtDistOfBtmCrv)  'Extend curve from Start point
                If oTempExtendedBtmCurve2 Is Nothing Then GoTo SkipThisTemplate
            Else
                Set oTempExtendedBtmCurve2 = oTemplateData.BottomLine
            End If
            
            Dim oExtendedCurve As IUnknown
            
            If Abs(dSecondExtDistOfBtmCrv) > EPSILON Then
                Set oExtendedCurve = m_oMfgGeomHelper.ExtraPolateCurve(oTempExtendedBtmCurve2, 1, 0, 1, Nothing, Nothing, dSecondExtDistOfBtmCrv) 'Extend curve from End point
                If oExtendedCurve Is Nothing Then GoTo SkipThisTemplate
            Else
                Set oExtendedCurve = oTempExtendedBtmCurve2
            End If
    
            Set oTemplateData.BottomLine = oExtendedCurve
       
SkipThisTemplate:
        Next nIndex
        
CleanUp:
        Set oTemplateContours = Nothing
        Set oGeomFactory = Nothing
        Set oLine = Nothing
        Set oGeomMisc = Nothing
        Set oTemplateData = Nothing
        Set oTemplatePlane = Nothing
        Set oProjectedLine = Nothing
        Set oSideLineVec = Nothing
        Set oInfBasePlane = Nothing
        
    End If

Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

Private Sub ApplyExtnByOffset(oBasePlane As IJPlane, ByRef oTemplateDataElems As IJElements, dOffset As Double, dMinExtn As Double)
    Const METHOD = "ApplyExtnByOffset"
    On Error Resume Next
    
    If dMinExtn > 0 Then
        
        Dim nIndex As Long
    
        Dim oTemplateContours As IJElements
        Set oTemplateContours = New JObjectCollection
         
        Dim dBPNormalX As Double, dBPNormalY As Double, dBPNormalZ As Double
        Dim dBPRootX As Double, dBPRootY As Double, dBPRootZ As Double
        
        oBasePlane.GetRootPoint dBPRootX, dBPRootY, dBPRootZ
        oBasePlane.GetNormal dBPNormalX, dBPNormalY, dBPNormalZ
    
        Dim oGeomFactory As IngrGeom3D.GeometryFactory
        Set oGeomFactory = New GeometryFactory
    
        Dim oLine As Line3d
        Set oLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, dBPRootX, dBPRootY, dBPRootZ, dBPNormalX, dBPNormalY, dBPNormalZ, 1)
        
        Dim oGeomMisc As DGeomOpsMisc
        Set oGeomMisc = New DGeomOpsMisc
        
        Dim oWireBody As Object
        oGeomMisc.CreateModelGeometryFromGType Nothing, oLine, Nothing, oWireBody
        
        Dim oTemplateData As TemplateData
        Dim oSBottomLinePos As IJDPosition, oEBottomLinePos As IJDPosition
        Dim oTemplatePlane As IJPlane
        Dim oInfBasePlane As IJPlane
        Set oInfBasePlane = New Plane3d
     
        oInfBasePlane.DefineByPointNormal dBPRootX, dBPRootY, dBPRootZ, dBPNormalX, dBPNormalY, dBPNormalZ
        
        For nIndex = 1 To oTemplateDataElems.count
            Set oTemplateData = oTemplateDataElems.Item(nIndex)
            If oTemplateData Is Nothing Then GoTo SkipThisTemplate
            
            'Get BottomLine EndPoints
            GetEndPoints oTemplateData.BottomLine, oSBottomLinePos, oEBottomLinePos
            Set oTemplatePlane = oTemplateData.Plane
            If oTemplatePlane Is Nothing Then GoTo SkipThisTemplate
     
            Dim oProjectedLine As Object
            Set oProjectedLine = GetProjectedCurveOnPlane(oWireBody, oTemplatePlane)
            If oProjectedLine Is Nothing Then GoTo SkipThisTemplate
            
            Dim oStartPos As IJDPosition, oEndPos As IJDPosition
            GetEndPoints oProjectedLine, oStartPos, oEndPos
            
            Dim oSideLineVec As IJDVector
            Set oSideLineVec = oEndPos.Subtract(oStartPos)
            If oSideLineVec Is Nothing Then GoTo SkipThisTemplate
            oSideLineVec.Length = 1
          
            Dim oControlPoint As IJDPosition
            Dim oCtlPtOnBPlane As IJDPosition
            
            Set oControlPoint = oTemplateData.ControlPoint
            If oControlPoint Is Nothing Then GoTo SkipThisTemplate
                               
            Dim oTopLineVec As IJDVector
            Dim oFirstPtOfTopLine As IJDPosition
            Dim oSecondPtOfTopLine As IJDPosition
                        
            Dim oFirstSideLine As Line3d, oSecondSideLine As Line3d, oCtlPtSideLine As Line3d
            
            Set oCtlPtSideLine = CreateInfiniteLine(oControlPoint, oSideLineVec)
            Set oCtlPtOnBPlane = m_oMfgGeomHelper.IntersectCurveWithPlane(oCtlPtSideLine, oInfBasePlane)
            If oCtlPtOnBPlane Is Nothing Then GoTo SkipThisTemplate
           
            Set oFirstSideLine = CreateInfiniteLine(oSBottomLinePos, oSideLineVec)
            Set oFirstPtOfTopLine = m_oMfgGeomHelper.IntersectCurveWithPlane(oFirstSideLine, oInfBasePlane)
            If oFirstPtOfTopLine Is Nothing Then GoTo SkipThisTemplate
    
            Set oSecondSideLine = CreateInfiniteLine(oEBottomLinePos, oSideLineVec)
            Set oSecondPtOfTopLine = m_oMfgGeomHelper.IntersectCurveWithPlane(oSecondSideLine, oInfBasePlane)
            If oSecondPtOfTopLine Is Nothing Then GoTo SkipThisTemplate
    
            'Get the TopLine direction
            Set oTopLineVec = oSecondPtOfTopLine.Subtract(oFirstPtOfTopLine)
            oTopLineVec.Length = 1
                            
            Dim dExtAtFirstPtOfTLine As Double, dExtAtSecondPtOfTLine As Double
            Dim dDistBtwFirstAndCtl As Double, dDistBtwSecondAndCtl As Double
    
            Dim nNumber As Integer
            Dim dExtension As Double
            
            dDistBtwFirstAndCtl = oFirstPtOfTopLine.DistPt(oCtlPtOnBPlane)
      
            nNumber = Fix((Round(dDistBtwFirstAndCtl, 3)) / (Round(dOffset, 3)))
            dExtension = dOffset - (dDistBtwFirstAndCtl - dOffset * nNumber)
            
            'Check if the First Ext. is less than m_dExt, then add m_dOffset value.
            If Round(dExtension, 3) > dMinExtn Then
                dExtAtFirstPtOfTLine = dExtension
            Else
                dExtAtFirstPtOfTLine = dOffset + dExtension
            End If
    
            dDistBtwSecondAndCtl = oSecondPtOfTopLine.DistPt(oCtlPtOnBPlane)
            
            nNumber = Fix((Round(dDistBtwSecondAndCtl, 3)) / (Round(dOffset, 3)))
            dExtension = dOffset - (dDistBtwSecondAndCtl - dOffset * nNumber)
            
            'Check if the Second Ext. is less than m_dExt, then add m_dOffset value.
            If Round(dExtension, 3) > dMinExtn Then
                dExtAtSecondPtOfTLine = dExtension
            Else
                dExtAtSecondPtOfTLine = dOffset + dExtension
            End If
            
    '/////////
    
    
            Dim oFirstSidePlane As IJPlane, oSecondSidePlane As IJPlane
            Dim dFirstExtDistOfBtmCrv As Double, dSecondExtDistOfBtmCrv As Double
            Dim oStPointOfBtmCrv As IJDPosition, oEdPointOfBtmCrv As IJDPosition
            
            'Calculate the first point of the top line
            oFirstPtOfTopLine.x = oFirstPtOfTopLine.x - oTopLineVec.x * dExtAtFirstPtOfTLine
            oFirstPtOfTopLine.y = oFirstPtOfTopLine.y - oTopLineVec.y * dExtAtFirstPtOfTLine
            oFirstPtOfTopLine.z = oFirstPtOfTopLine.z - oTopLineVec.z * dExtAtFirstPtOfTLine
            
            m_oMfgGeomHelper.MakeTransientPlane oFirstPtOfTopLine.x, oFirstPtOfTopLine.y, oFirstPtOfTopLine.z, oTopLineVec.x, oTopLineVec.y, oTopLineVec.z, oFirstSidePlane
            
            Dim oFirstExtPosOfBtmCrv As IJDPosition, oSecondExtPosOfBtmCrv As IJDPosition
            Dim oTempExtendedBtmCurve1 As IUnknown, oTempExtendedBtmCurve2 As IUnknown
            
            Set oTempExtendedBtmCurve1 = m_oMfgGeomHelper.ExtraPolateCurve(oTemplateData.BottomLine, 1, 0, 2, Nothing, Nothing, 0.5)
            If oTempExtendedBtmCurve1 Is Nothing Then GoTo SkipThisTemplate
            
            Set oFirstExtPosOfBtmCrv = m_oMfgGeomHelper.IntersectCurveWithPlane(oTempExtendedBtmCurve1, oFirstSidePlane)
            If oFirstExtPosOfBtmCrv Is Nothing Then GoTo SkipThisTemplate
            
            m_oMfgGeomHelper.GetLengthBet2Points oTempExtendedBtmCurve1, oSBottomLinePos, oFirstExtPosOfBtmCrv, dFirstExtDistOfBtmCrv
            
            'Calculate the second point of the top line
            oSecondPtOfTopLine.x = oSecondPtOfTopLine.x + oTopLineVec.x * dExtAtSecondPtOfTLine
            oSecondPtOfTopLine.y = oSecondPtOfTopLine.y + oTopLineVec.y * dExtAtSecondPtOfTLine
            oSecondPtOfTopLine.z = oSecondPtOfTopLine.z + oTopLineVec.z * dExtAtSecondPtOfTLine
               
            m_oMfgGeomHelper.MakeTransientPlane oSecondPtOfTopLine.x, oSecondPtOfTopLine.y, oSecondPtOfTopLine.z, oTopLineVec.x, oTopLineVec.y, oTopLineVec.z, oSecondSidePlane
            Set oSecondExtPosOfBtmCrv = m_oMfgGeomHelper.IntersectCurveWithPlane(oTempExtendedBtmCurve1, oSecondSidePlane)
            If oSecondExtPosOfBtmCrv Is Nothing Then GoTo SkipThisTemplate
            
            m_oMfgGeomHelper.GetLengthBet2Points oTempExtendedBtmCurve1, oEBottomLinePos, oSecondExtPosOfBtmCrv, dSecondExtDistOfBtmCrv
       
            If Abs(dFirstExtDistOfBtmCrv) > EPSILON Then
                Set oTempExtendedBtmCurve2 = m_oMfgGeomHelper.ExtraPolateCurve(oTemplateData.BottomLine, 1, 0, 0, Nothing, Nothing, dFirstExtDistOfBtmCrv)  'Extend curve from Start point
                If oTempExtendedBtmCurve2 Is Nothing Then GoTo SkipThisTemplate
            Else
                Set oTempExtendedBtmCurve2 = oTemplateData.BottomLine
            End If
            
            Dim oExtendedCurve As IUnknown
            
            If Abs(dSecondExtDistOfBtmCrv) > EPSILON Then
                Set oExtendedCurve = m_oMfgGeomHelper.ExtraPolateCurve(oTempExtendedBtmCurve2, 1, 0, 1, Nothing, Nothing, dSecondExtDistOfBtmCrv) 'Extend curve from End point
                If oExtendedCurve Is Nothing Then GoTo SkipThisTemplate
            Else
                Set oExtendedCurve = oTempExtendedBtmCurve2
            End If
    
            Set oTemplateData.BottomLine = oExtendedCurve
       
SkipThisTemplate:
        Next nIndex
        
CleanUp:
        Set oTemplateContours = Nothing
        Set oGeomFactory = Nothing
        Set oLine = Nothing
        Set oGeomMisc = Nothing
        Set oTemplateData = Nothing
        Set oTemplatePlane = Nothing
        Set oProjectedLine = Nothing
        Set oSideLineVec = Nothing
        Set oInfBasePlane = Nothing
        
    End If

Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp

End Sub


'********************************************************************************************************************************
' Function Name:    ApplyModifiedBottomlines
'
' Interface:        <private>
'
' Inputs:           oTemplateSet --
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the TemplateDataElems
'
' Assumption:
'
' Algorithm:        1.Get Modified Bottom Lines from TemplateSet
'                   2.Check what existing Bottom Lines(TemplateData) they coincide with and replace
'                   3.Get Added BottomLines from TemplateSet and add to the Bottom Lines collection
'********************************************************************************************************************************'
Private Sub ApplyModifiedBottomlines(oPlatePart As IJPlatePart, oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
    Const METHOD = "ApplyModifiedBottomlines"
    On Error GoTo ErrorHandler
    
    Dim oModifiedBLColl As IJElements
    Set oModifiedBLColl = oTemplateSet.GetInputModifiedBottomLines
    
    Dim oTemplateData   As TemplateData
    Dim oBLCurve        As Object
    Dim oCopiedCurve    As IJComplexString
    Dim i As Long, j As Long
    
    For i = 1 To oModifiedBLColl.count
    
        Set oBLCurve = oModifiedBLColl.Item(i)
        
        For j = 1 To oTemplateDataElems.count
            
            Set oTemplateData = oTemplateDataElems.Item(j)
            Dim oDataCurve As IJCurve
            Set oDataCurve = oTemplateData.BottomLine
            
            Dim dMinDist As Double, dSrcX As Double, dSrcY As Double, dSrcZ As Double
            Dim dDestX As Double, dDestY As Double, dDestZ As Double
            
            'first, check if they intersect
            oDataCurve.DistanceBetween oBLCurve, dMinDist, dSrcX, dSrcY, dSrcZ, dDestX, dDestY, dDestZ
            
            If dMinDist < 0.001 Then
                'Do they overlap?
                Dim oTopologyIntersect As IJDTopologyIntersect
                Set oTopologyIntersect = New DGeomOpsIntersect
                                        
                Dim bHasOverLap As Boolean
                bHasOverLap = False
        
                Dim oHelper                     As New MfgRuleHelpers.Helper

                Dim oCS1 As IJComplexString
                Set oCS1 = New ComplexString3d
                oCS1.AddCurve oDataCurve, True
                
                Dim oWireBody1 As IJWireBody
                Set oWireBody1 = oHelper.ComplexStringToWireBody(oCS1)
                
                Dim oCS2 As IJComplexString
                Set oCS2 = New ComplexString3d
                oCS2.AddCurve oBLCurve, True
                
                Dim oWireBody2 As IJWireBody
                Set oWireBody2 = oHelper.ComplexStringToWireBody(oCS2)

                ''oTopologyIntersect.HasOverlappingGeometry oWireBody1, oWireBody2, bHasOverLap
                
                Dim oOverlapUnknown         As IUnknown
                Dim oOverlapWireBody        As IJWireBody
     
                oTopologyIntersect.PlaceOverlappingObject Nothing, oWireBody1, oWireBody2, Nothing, oOverlapUnknown
                
                On Error Resume Next
                Set oOverlapWireBody = oOverlapUnknown
                On Error GoTo ErrorHandler
                
                If Not oOverlapWireBody Is Nothing Then

                    Set oCopiedCurve = GetCopyOfSymbolOutput(oBLCurve)
                    
                    Set oTemplateData.BottomLine = oCopiedCurve
                    If oTemplateData.TemplateType = System Or oTemplateData.TemplateType = Edge Then
                        oTemplateData.TemplateType = SystemTemplateWithModifiedBL
                    ElseIf oTemplateData.TemplateType = Sketched Then
                        oTemplateData.TemplateType = SketchedTemplateWithModifiedBL
                    End If
                    Set oCopiedCurve = Nothing
                    
                    Exit For
               
                End If
            End If
            
        Next
        
        Set oTemplateData = Nothing
        Set oBLCurve = Nothing
    Next
      
    Dim oAddedBLColl As IJElements
    Set oAddedBLColl = oTemplateSet.GetInputAddedBottomLines
        
    For i = 1 To oAddedBLColl.count
        Set oBLCurve = oAddedBLColl.Item(i)
        
        Dim oControlPoint As IJDPosition
        
        Dim lGroupType As Long
        lGroupType = Primary 'initalize
        
        'in case of box overwrite Control line and group
        If oTemplateSetData.TemplateType = "Box" Or _
            oTemplateSetData.TemplateType = "UserDefined Box" Or _
            oTemplateSetData.TemplateType = "UserDefined Box With Edges" Then
        
            Dim oSurfaceBody As IJSurfaceBody
            Set oSurfaceBody = oTemplateSetData.SurfaceBody
            
            If IsCurveAlongDirection(oBLCurve, oSurfaceBody, oTemplateSetData.Direction) Then
                lGroupType = Primary
            Else
                If IsCurveAlongDirection(oBLCurve, oSurfaceBody, oTemplateSetData.SecondaryDirection) Then
                    lGroupType = Secondary
                Else 'neither along primary nor along secondary
                    lGroupType = NotInPrimaryOrSecondary
            End If
        End If
    
            'Add mid point as control point
            Dim oStartPos As IJDPosition, oEndPos As IJDPosition

            GetEndPoints oBLCurve, oStartPos, oEndPos
            Set oControlPoint = GetEdgeMidPoint(oStartPos, oEndPos)
        Else
            Dim oControlLine As Object
            Set oControlLine = oTemplateSet.GetControlLine 'Get the primary group BCL
        Set oControlPoint = GetIntersectionBetweenCurves(oBLCurve, oControlLine)
        End If
        
        If Not oControlPoint Is Nothing Then
            Set oCopiedCurve = GetCopyOfSymbolOutput(oBLCurve)
            
            Set oTemplateData = New TemplateData
            
            oTemplateData.TemplateType = TemplateWithAddedBL
            Set oTemplateData.DefiningObj = oCopiedCurve
            Set oTemplateData.ControlPoint = oControlPoint
            Set oTemplateData.BottomLine = oCopiedCurve
            oTemplateData.GroupType = lGroupType
            Set oTemplateData.Plane = CreateTemplatePlaneForAddedCurve(oPlatePart, oTemplateData, oTemplateSet)
            oTemplateDataElems.Add oTemplateData
        End If
        Set oTemplateData = Nothing
        Set oBLCurve = Nothing
    Next
        
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub




'********************************************************************************************************************************
' Function Name:    CreateTemplateDataFrames
'
' Interface:        <private>
'
' Inputs:           oTemplateSet --
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the TemplateDataElems by setting the controlpoint, plane etc
'
' Assumption:
'
' Algorithm:        1. Get Frames
'                   2. if PositionEven = NotUsed go for positonframes
'                   2a. at each frame location compute intersection, check if on edge and if no template already at edge
'                   2b. For positioneven, compute the controlpoint by offset along controlline, define plane at this point
'********************************************************************************************************************************'
Private Sub CreateTemplateDataFrames(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
    Const METHOD = "CreateTemplateDataFrames"
    On Error GoTo ErrorHandler

    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        Set oTemplateDataElems = New JObjectCollection
    End If
    
    Dim oTemplateData       As TemplateData
        
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
 
    Dim nIndex As Long
    Dim oControlLine As Object
    Dim oControlStartPos As IJDPosition, oControlEndPos As IJDPosition
    Dim oExtendedControlLine As IUnknown
    Dim oPlateEdge              As Object
    Dim oAftForeEdges As IJElements
    Dim oFrame As Object
    Dim oControlPoint As IJDPosition
    Dim oFrameNormal As IJDVector
    Dim oIHFrame As IHFrame
    Dim oFramePos As IJDPosition
    Dim oFrameElems As IJElements
 
    
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
    
    Set oControlLine = oTemplateSet.GetControlLine
    Set oExtendedControlLine = m_oMfgGeomHelper.ExtraPolateCurve(oControlLine, 1, 1, 2, Nothing, Nothing, 0.3) 'Extend 0.3meter

    Set oAftForeEdges = GetTemplateAftForeEdges(oTemplateSet, oTemplateSetData)
           
    'compute for frames

''''    Dim oIJDModelBody As IJDModelBody
''''    Set oIJDModelBody = oSurfaceBody
''''    If Not oIJDModelBody Is Nothing Then
''''        Dim FileName As String
''''        FileName = Environ("TEMP")
''''        If FileName = "" Or FileName = vbNullString Then
''''            FileName = "C:\Temp" 'Only use C:\Temp if there is a %TEMP% failure
''''        End If
''''        If oTemplateSetData.IsOnBaseSide Then
''''            FileName = FileName & "\SurfaceWithoutFeatures_BaseSide.sat"
''''        Else
''''            FileName = FileName & "\SurfaceWithoutFeatures_OffsetSide.sat"
''''        End If
''''        oIJDModelBody.DebugToSATFile FileName
''''        Set oIJDModelBody = Nothing
''''    End If
     
    ' make PositionFrames higher priority than PositionEven.
    If Not oTemplateSetData.PositionFrames = "NotUsed" Then
        Set oFrameElems = oTemplateSetData.FrameElems
        For nIndex = 1 To oFrameElems.count
            Set oFrame = oFrameElems.Item(nIndex)
            Set oIHFrame = oFrame
            oIHFrame.GetPositionVector oFramePos, oFrameNormal
            
            On Error GoTo ICWP_ErrorHandler 'ICWP-->IntersectCurveWithPlane(this does not retunr if there is no intersection)
            Set oControlPoint = m_oMfgGeomHelper.IntersectCurveWithPlane(oControlLine, oFrame)
            On Error GoTo ErrorHandler
            
            If Not oControlPoint Is Nothing Then
                'check if this frame is on the edge
                Set oPlateEdge = GetEdgeOnFrame(oFrame, oAftForeEdges, oControlLine)
                If Not oPlateEdge Is Nothing Then
                    'Edge  templates are already defined. Do nothing
                    'this frame is on the edge
                    If Not IsEdgeTemplateDefined(oTemplateDataElems, oControlPoint) Then
                        Set oTemplateData = New TemplateData
                        oTemplateData.TemplateType = Edge
                        Set oTemplateData.DefiningObj = oPlateEdge
                        Set oTemplateData.ControlPoint = oControlPoint
                        oTemplateDataElems.Add oTemplateData
                    End If
                Else
                    ' this is not a edge frame
                    Set oTemplateData = New TemplateData
                    oTemplateData.TemplateType = System
                    Set oTemplateData.DefiningObj = oFrame
                    Set oTemplateData.ControlPoint = oControlPoint
                    oTemplateDataElems.Add oTemplateData
                End If
            End If
        Next nIndex
    Else ' PositionEven Note: we do not have edge templates in positon even
        Dim dControlLineLength As Double
        If oFrameNormal Is Nothing Then Set oFrameNormal = New DVector
        'get normal so that we can create plane in same direction
        If oTemplateSetData.Direction = "Longitudinal" Then  'X - Direction(Buttock)
            oFrameNormal.Set 0, 1, 0
        ElseIf oTemplateSetData.Direction = "Transversal" Then  'Y - Direction(Frame)
            oFrameNormal.Set 1, 0, 0
        Else 'Z - Direction(WaterLine)
            oFrameNormal.Set 0, 0, 1
        End If
            
        GetEndPoints oControlLine, oControlStartPos, oControlEndPos
        m_oMfgGeomHelper.GetLengthBet2Points oControlLine, oControlStartPos, oControlEndPos, dControlLineLength

        Dim dInterval As Double
        dInterval = (dControlLineLength - (2 * oTemplateSetData.EdgeOffset)) / (oTemplateSetData.TemplatesCount - 1)
        Dim oTempStartPos As IJDPosition
        Dim oPlane As Plane3d
        'loop unrolled for first control point
        Set oControlPoint = m_oMfgGeomHelper.GetPointAtDistAlongCurve(oControlLine, oControlStartPos, oTemplateSetData.EdgeOffset)
        Set oTempStartPos = oControlPoint
        If Not oControlPoint Is Nothing Then
            Set oTemplateData = New TemplateData
            oTemplateData.TemplateType = System
            ' there is no frame at this offset so create one
            Set oPlane = New Plane3d
            oPlane.DefineByPointNormal oControlPoint.x, oControlPoint.y, oControlPoint.z, _
                                        oFrameNormal.x, oFrameNormal.y, oFrameNormal.z
            Set oTemplateData.DefiningObj = oPlane
            Set oTemplateData.ControlPoint = oControlPoint
            oTemplateDataElems.Add oTemplateData
            Set oPlane = Nothing
        End If
        For nIndex = 2 To oTemplateSetData.TemplatesCount
            Set oControlPoint = m_oMfgGeomHelper.GetPointAtDistAlongCurve(oControlLine, oTempStartPos, dInterval)
            Set oTempStartPos = oControlPoint
            If Not oControlPoint Is Nothing Then
                Set oTemplateData = New TemplateData
                oTemplateData.TemplateType = System
                ' there is no frame at this offset so create one
                Set oPlane = New Plane3d
                oPlane.DefineByPointNormal oControlPoint.x, oControlPoint.y, oControlPoint.z, _
                                             oFrameNormal.x, oFrameNormal.y, oFrameNormal.z
                Set oTemplateData.DefiningObj = oPlane
                Set oTemplateData.ControlPoint = oControlPoint
                oTemplateDataElems.Add oTemplateData
                Set oPlane = Nothing
            End If
        Next
    End If
    
CleanUp:
    Set oAftForeEdges = Nothing
    Set oControlLine = Nothing
    Set oExtendedControlLine = Nothing
    Set oPlateEdge = Nothing
    Set oFrame = Nothing
    Set oFrameElems = Nothing
    Set oTemplateData = Nothing
    Set oControlPoint = Nothing
    Set oTempStartPos = Nothing
    Set oFrameNormal = Nothing
    Set oIHFrame = Nothing
    Set oFramePos = Nothing
    Exit Sub
ICWP_ErrorHandler: 'ICWP-->IntersectCurveWithPlane
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToIntersectCurveWithPlane, , "RULES")
    GoTo CleanUp
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub


'********************************************************************************************************************************
' Function Name:    CreateTemplateDataSketched
'
' Interface:        <private>
'
' Inputs:           oTemplateSet --
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   oSketchingLineElems -- collection of lines skecthed by user
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the TemplateDataElems by setting the controlpoint, plane etc
'
' Assumption:
'
' Algorithm:        1. Get Surface
'                   2. Project sketching line to surface
'                   2. if projected sketching line intersects control line create template
'********************************************************************************************************************************'

Private Sub CreateTemplateDataSketched(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oSketchingLineElems As IJElements, oTemplateDataElems As IJElements)
    Const METHOD = "CreateTemplateDataSketched"
    On Error GoTo ErrorHandler
        
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oSketchingLineElems Is Nothing Then
        strErrorMsg = strErrorMsg & "oSketchingLineElems is Nothing "
        GoTo ErrorHandler
    End If
    If oSketchingLineElems.count = 0 Then Exit Sub
    If oTemplateDataElems Is Nothing Then
        Set oTemplateDataElems = New JObjectCollection
    End If
    
    Dim oTemplateData       As TemplateData
        
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Dim nIndex As Long
    Dim oPlateNormalVector As IJDVector, oBasePlaneNormalVector  As IJDVector
    
    Dim oControlLine As Object
    Dim oProjectionDir As IJDVector
    Dim oSketchedLine As Object
    Dim oControlPoint As IJDPosition
    Dim oProjectedCS  As IJComplexString
    Dim oSurfaceBody As IJSurfaceBody
    
    Set oControlLine = oTemplateSet.GetControlLine
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
    
    'get sketching plane of TemplateSet
    Dim oSketchingPlane As IJPlane
    
    Set oSketchingPlane = oTemplateSet.GetSketchingPlane
    
    If Not oSketchingPlane Is Nothing Then
        Set oProjectionDir = New DVector
        Dim xPlaneNormal As Double, yPlaneNormal As Double, zPlaneNormal As Double
        Call oSketchingPlane.GetNormal(xPlaneNormal, yPlaneNormal, zPlaneNormal)
        Call oProjectionDir.Set(xPlaneNormal, yPlaneNormal, zPlaneNormal)
    End If
    
    ' compute for Sketched lines
    For nIndex = 1 To oSketchingLineElems.count
        Set oSketchedLine = oSketchingLineElems.Item(nIndex)
        
        On Error GoTo PCSTS_ErrorHandler 'PCSTS-->ProjectComplexStringToSurface
        m_oMfgMGHelper.ProjectComplexStringToSurface oSketchedLine, oSurfaceBody, oProjectionDir, oProjectedCS
        
        On Error Resume Next 'It is possible that a curve or two may not have intersection
        Set oControlPoint = GetIntersectionBetweenCurves(oProjectedCS, oControlLine)
        On Error GoTo ErrorHandler
        
        If Not oControlPoint Is Nothing Then
            Set oTemplateData = New TemplateData
            oTemplateData.TemplateType = Sketched
            Set oTemplateData.DefiningObj = oProjectedCS
            Set oTemplateData.ControlPoint = oControlPoint
            oTemplateDataElems.Add oTemplateData
        End If
        Set oProjectedCS = Nothing
    Next nIndex
       
CleanUp:
    Set oProjectionDir = Nothing
    Set oPlateNormalVector = Nothing
    Set oControlLine = Nothing
    Set oBasePlaneNormalVector = Nothing
    Set oSurfaceBody = Nothing
    Set oTemplateData = Nothing
    Set oControlPoint = Nothing
    Set oSketchedLine = Nothing
    Exit Sub
PCSTS_ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FaliedToProjectComplexStringToSurf, , "RULES")
    GoTo CleanUp
GIBC_ErrorHandler: 'GIBC-->GetIntersectionBetweenCurves
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FaliedToGetIntersectionBetweenCurves, , "RULES")
    GoTo CleanUp
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub


'********************************************************************************************************************************
' Function Name:    CreateTemplateDataAftFore
'
' Interface:        <private>
'
' Inputs:           oTemplateSet --
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the TemplateDataElems by setting the controlpoint, plane etc
'
' Assumption:
'
' Algorithm:        1. Get AftFore Edges
'                   2. check orientation required, define template normal
'                   3. check if Edge is planar, edge plane normal same as template normal
'                   4. check if edge intersects controlline
'********************************************************************************************************************************'
Private Sub CreateTemplateDataAftFore(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
    Const METHOD = "CreateTemplateDataAftFore"
    On Error GoTo ErrorHandler
        
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        Set oTemplateDataElems = New JObjectCollection
    End If
    
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
 
    Dim oTemplateData       As TemplateData
    
    Dim nIndex As Long
    Dim oBasePlaneNormalVector    As IJDVector
    Dim oControlLine As Object
    Dim oExtendedControlLine As IUnknown
    Dim oPlateEdge              As Object
    Dim oBasePlane              As IJPlane
    Dim dBasePlNormalX          As Double, dBasePlNormalY As Double, dBasePlNormalZ As Double
    Dim oAftForeEdges As IJElements
    Dim oTemplatePlaneVector As IJDVector
    Dim oControlPoint As IJDPosition
    Dim oSurfaceBody As IJSurfaceBody
   
    Set oControlLine = oTemplateSet.GetControlLine
    Set oExtendedControlLine = m_oMfgGeomHelper.ExtraPolateCurve(oControlLine, 1, 1, 2, Nothing, Nothing, 0.3) 'Extend 0.3meter

    ' find the edges that intersect with BaseControlLine which we will call as AftForeEdges
    ' Theoretically there should only be two edges, but to be on safe side let us create IJElements
    Set oAftForeEdges = GetTemplateAftForeEdges(oTemplateSet, oTemplateSetData)
    
    'get the Base plane object from Template object
    Set oBasePlane = oTemplateSet.GetBasePlane
    oBasePlane.GetNormal dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ

    Set oBasePlaneNormalVector = New DVector
    oBasePlaneNormalVector.Set dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ
    oBasePlaneNormalVector.Length = 1
    
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
   
    Dim bIsEdgePlanar As Boolean
    Dim oEdgePlaneDir As IJDVector

    For nIndex = 1 To oAftForeEdges.count
        Set oPlateEdge = oAftForeEdges.Item(nIndex)
        
        On Error GoTo GIBC_ErrorHandler 'GetIntersectionBetweenCurves
        Set oControlPoint = GetIntersectionBetweenCurves(oPlateEdge, oControlLine)
        On Error GoTo ErrorHandler
    
        If Not oControlPoint Is Nothing Then
            'we need to check if "plane" of edge is same as templateplane
            ' in case of "Aft/Forward" do not consider Orientation
'            TR 121340 Aft Fore templates should be created for all Types
'            If oTemplateSetData.TemplateType = "Aft/Forward" Then
                    Set oTemplatePlaneVector = New DVector
                    If oTemplateSetData.Direction = "Longitudinal" Then 'X - Direction(Buttock)
                        oTemplatePlaneVector.Set 0, 1, 0
                    ElseIf oTemplateSetData.Direction = "Transversal" Then 'Y - Direction(Frame)
                        oTemplatePlaneVector.Set 1, 0, 0
                    Else 'Z - Direction(WaterLine)
                        oTemplatePlaneVector.Set 0, 0, 1
                    End If
'            Else
'
'                If oTemplateSetData.Orientation = "NormalToBasePlane" Then
'                    Set oTemplatePlaneVector = oBasePlaneNormalVector.Clone
'                ElseIf oTemplateSetData.Orientation = "AlongFrame" Then
'                    Set oTemplatePlaneVector = New DVector
'                    If oTemplateSetData.Direction = "Longitudinal" Then 'X - Direction(Buttock)
'                        oTemplatePlaneVector.Set 0, 1, 0
'                    ElseIf oTemplateSetData.Direction = "Transversal" Then 'Y - Direction(Frame)
'                        oTemplatePlaneVector.Set 1, 0, 0
'                    Else 'Z - Direction(WaterLine)
'                        oTemplatePlaneVector.Set 0, 0, 1
'                    End If
'                Else 'perpendicular 'use plate normal at the control point
'                    oSurfaceBody.GetNormalFromPosition oControlPoint, oTemplatePlaneVector
'                End If
'            End If
            
            oTemplatePlaneVector.Length = 1
            
            If IsCurveOnGivenPlane(oPlateEdge, oTemplatePlaneVector) Then ' they are almost normal to each other
                    Set oTemplateData = New TemplateData
                    oTemplateData.TemplateType = Edge
                    Set oTemplateData.DefiningObj = oPlateEdge
                    Set oTemplateData.ControlPoint = oControlPoint
                    oTemplateDataElems.Add oTemplateData
                    Set oTemplateData = Nothing
            End If 'else we do not place a template if the edge is planar in some other direction
        End If
        Set oTemplatePlaneVector = Nothing
    Next nIndex
    
CleanUp:
     
    Set oAftForeEdges = Nothing
    Set oControlLine = Nothing
    Set oExtendedControlLine = Nothing
    Set oBasePlane = Nothing
    Set oBasePlaneNormalVector = Nothing
    Set oPlateEdge = Nothing
    Set oSurfaceBody = Nothing
    Set oTemplateData = Nothing
    Set oControlPoint = Nothing
    Set oSurfaceBody = Nothing
    Set oEdgePlaneDir = Nothing
                
    Exit Sub
GIBC_ErrorHandler: 'GetIntersectionBetweenCurves
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FaliedToGetIntersectionBetweenCurves, , "RULES")
    GoTo CleanUp
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub


'********************************************************************************************************************************
' Function Name:    CreateTemplatePlanesAtCornerPoints
'
' Interface:        <private>
'
' Inputs:           oTemplateSet --
'                   oPlane -- plane for projection XZ for stemstern XY for PerpendicularXY
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the TemplateDataElems by setting the controlpoint, plane etc
'
' Assumption:
'
' Algorithm:         1. Project control line onto given plane
'                    2.  Compute the templateplane at the plate edge
'                    3.  for each of the edge get the end points
'                           There are three cases for edge and controlline intersection
'       case1: The edge and the controlline intersection are on a straight line, (the edge is planar)
'       case2: both ends are on same side of edge w.r.t intersection point
'           case2a: both ends will form same template plane
'           case2a: both ends will form different template planes
'       case3: ends are on opposite side of edge w.r.t intersection point
'  case 3 and case2 are handled same, we just check if a plane placed at the edge will intersect the control line
'
'********************************************************************************************************************************'

Private Sub CreateTemplatePlanesAtCornerPoints(oTemplateSet As IJDMfgTemplateSet, oPlane As IJPlane, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
    Const METHOD = "CreateTemplatePlanesAtCornerPoints"
    On Error GoTo ErrorHandler

    Dim strErrorMsg As String
 
     'check if inputs are okay
    If oPlane Is Nothing Then
        strErrorMsg = strErrorMsg & "oPlane is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        ' create a new collection
        Set oTemplateDataElems = New JObjectCollection
    End If

    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''   1. Project control line onto given plane
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oControlLine As Object
    Dim oPrjctdCntrlLn As IJCurve
    Dim oCntrlLnSPos As IJDPosition, oCntrlLnEPos As IJDPosition
    Dim oPrjctdCntrlLnSPos As IJDPosition, oPrjctdCntrlLnEPos As IJDPosition

    Set oControlLine = oTemplateSet.GetControlLine
    m_oMfgGeomHelper.ProjectCurveOntoPlane oControlLine, oPlane, oPrjctdCntrlLn
    GetEndPoints oControlLine, oCntrlLnSPos, oCntrlLnEPos
    GetEndPoints oPrjctdCntrlLn, oPrjctdCntrlLnSPos, oPrjctdCntrlLnEPos


    Dim oTemplateData As TemplateData
    Dim oTemplatePlane As IJPlane
    Dim dRootX As Double, dRootY As Double, dRootZ As Double
    Dim dSrcX As Double, dSrcY As Double, dSrcZ As Double
    Dim dInX As Double, dInY As Double, dinZ As Double
    Dim dMinDist As Double
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    Dim nIndex As Long, nCount As Integer
    Dim dLength As Double
    Dim dInterval As Double

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''  2.  Compute the templateplane at the plate edge
''''      Construct a plane at the edge end point
''''      check if there is an intersection between plane and control line
''''             create template set data and store the dist from the control point
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oAftForeEdgeElems As IJElements
    Set oAftForeEdgeElems = GetTemplateAftForeEdges(oTemplateSet, oTemplateSetData)

    Dim oEdgePos1 As IJDPosition, oEdgePos2 As IJDPosition
    Dim oPrjctEdgePos1 As IJDPosition, oPrjctEdgePos2 As IJDPosition
    Dim oEdge As Object
    Dim oProjectedEdge As IUnknown
    Dim oIntersectionPos As IJDPosition
    Dim oPrjctIntersectionPos As IJDPosition
    Dim oIntersectionPosEdge As IJDPosition
    Dim oTemplatePlaneNormal  As IJDVector
    Dim oPrjctControlPos1 As IJDPosition
    Dim oPrjctControlPos2 As IJDPosition
    Dim oTemplatePlaneEdge1 As IJPlane
    Dim oTemplatePlaneEdge2 As IJPlane

    '3. for each of the edge get the end points
    ' There are three cases for edge and controlline intersection
    ' case1: The edge and the controlline intersection are on a straight line, (the edge is planar)
    ' case2: both ends are on same side of edge w.r.t intersection point
    '   case2a: both ends will form same template plane
    '   case2a: both ends will form different template planes
    ' case3: ends are on opposite side of edge w.r.t intersection point
    ' case 3 and case2 are handled same, we just check if a plane placed at the edge will intersect the control line
    
    For nIndex = 1 To oAftForeEdgeElems.count
        Set oEdge = oAftForeEdgeElems.Item(nIndex)
        Set oIntersectionPos = GetIntersectionBetweenCurves(oEdge, oControlLine)
        'create template plane at intersection between control line and plateedge
        If Not oIntersectionPos Is Nothing Then
            Set oPrjctIntersectionPos = GetProjectedPointOnPlane(oIntersectionPos, oPlane)
            Set oTemplatePlaneNormal = m_oMfgGeomHelper.GetTangentByPointOnCurve(oPrjctdCntrlLn, oPrjctIntersectionPos)
            oIntersectionPos.Get dRootX, dRootY, dRootZ
            oTemplatePlaneNormal.Get dNormalX, dNormalY, dNormalZ
            m_oMfgGeomHelper.MakeTransientPlane dRootX, dRootY, dRootZ, dNormalX, dNormalY, dNormalZ, oTemplatePlane

            Set oTemplateData = New TemplateData
            Set oTemplateData.ControlPoint = oIntersectionPos
            Set oTemplateData.Plane = oTemplatePlane
            ' case1: The edge and the controlline intersection are on a straight line, (the edge is planar)
            If IsEdgeOnPlane(oEdge, oTemplatePlane, oControlLine) Then
                Set oTemplateData.DefiningObj = oEdge
                oTemplateData.TemplateType = Edge
                oTemplateDataElems.Add oTemplateData
                Set oTemplateData = Nothing
                Set oTemplatePlane = Nothing
                Set oTemplatePlaneNormal = Nothing
            Else
                Set oTemplateData.DefiningObj = oTemplatePlane
                oTemplateData.TemplateType = System
                oTemplateDataElems.Add oTemplateData
                Set oTemplateData = Nothing
                Set oTemplatePlane = Nothing
                Set oTemplatePlaneNormal = Nothing
                
                ' compute planes at the end of the edge
                GetEndPoints oEdge, oEdgePos1, oEdgePos2
            
                Set oPrjctEdgePos1 = GetProjectedPointOnPlane(oEdgePos1, oPlane)
                oPrjctdCntrlLn.DistanceBetween oPrjctEdgePos1, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dinZ
                Set oPrjctControlPos1 = New Point3d
                oPrjctControlPos1.Set dSrcX, dSrcY, dSrcZ
                Set oTemplatePlaneNormal = m_oMfgGeomHelper.GetTangentByPointOnCurve(oPrjctdCntrlLn, oPrjctControlPos1)
                
                oEdgePos1.Get dRootX, dRootY, dRootZ
                oTemplatePlaneNormal.Get dNormalX, dNormalY, dNormalZ
                m_oMfgGeomHelper.MakeTransientPlane dRootX, dRootY, dRootZ, dNormalX, dNormalY, dNormalZ, oTemplatePlaneEdge1
                
                Set oPrjctEdgePos2 = GetProjectedPointOnPlane(oEdgePos2, oPlane)
                oPrjctdCntrlLn.DistanceBetween oPrjctEdgePos2, dMinDist, dSrcX, dSrcY, dSrcZ, dInX, dInY, dinZ
                Set oPrjctControlPos2 = New Point3d
                oPrjctControlPos2.Set dSrcX, dSrcY, dSrcZ
                Set oTemplatePlaneNormal = m_oMfgGeomHelper.GetTangentByPointOnCurve(oPrjctdCntrlLn, oPrjctControlPos2)
                
                oEdgePos2.Get dRootX, dRootY, dRootZ
                oTemplatePlaneNormal.Get dNormalX, dNormalY, dNormalZ
                m_oMfgGeomHelper.MakeTransientPlane dRootX, dRootY, dRootZ, dNormalX, dNormalY, dNormalZ, oTemplatePlaneEdge2
                      
                ' case2: both ends are on same side of edge w.r.t intersection point
                ' case3: ends are on opposite side of edge w.r.t intersection point
                ' case 3 and case2 are handled same, we just check if a plane placed at the edge will intersect the control line
                
                If IsEqualPoint(oPrjctEdgePos1, oPrjctEdgePos2) Then
                    '   case2a: both ends will form same template plane
                    ' check if control line and the templateplane will intersect
                    Set oIntersectionPos = GetIntersectionBetweenCurveAndPlane(oControlLine, oTemplatePlaneEdge1)
                    If Not oIntersectionPos Is Nothing Then
                            Set oTemplateData = New TemplateData
                            oTemplateData.TemplateType = System ' this is NOT a frame, but a plane
                            Set oTemplateData.ControlPoint = oIntersectionPos
                            Set oTemplateData.Plane = oTemplatePlaneEdge1
                            Set oTemplateData.DefiningObj = oTemplatePlaneEdge1
                            oTemplateDataElems.Add oTemplateData
                            Set oTemplateData = Nothing
                    End If
                Else
                    '   case2b: both ends will form different template planes or case 3
                    Set oIntersectionPos = GetIntersectionBetweenCurveAndPlane(oControlLine, oTemplatePlaneEdge1)
                    If Not oIntersectionPos Is Nothing Then
                            Set oTemplateData = New TemplateData
                            oTemplateData.TemplateType = System ' this is NOT a frame, but a plane
                            Set oTemplateData.ControlPoint = oIntersectionPos
                            Set oTemplateData.Plane = oTemplatePlaneEdge1
                            Set oTemplateData.DefiningObj = oTemplatePlaneEdge1
                            oTemplateDataElems.Add oTemplateData
                            Set oTemplateData = Nothing
                    End If
                    Set oIntersectionPos = GetIntersectionBetweenCurveAndPlane(oControlLine, oTemplatePlaneEdge2)
                    If Not oIntersectionPos Is Nothing Then
                            Set oTemplateData = New TemplateData
                            oTemplateData.TemplateType = System ' this is NOT a frame, but a plane
                            Set oTemplateData.ControlPoint = oIntersectionPos
                            Set oTemplateData.Plane = oTemplatePlaneEdge2
                            Set oTemplateData.DefiningObj = oTemplatePlaneEdge2
                            oTemplateDataElems.Add oTemplateData
                            Set oTemplateData = Nothing
                    End If
                End If
            End If
        End If
    Next

    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToDefTemplatePlaneFromCornerPts, , "RULES")
End Sub


Private Sub SortTemplateDataElems(oTemplateSet As IJDMfgTemplateSet, oTemplateDataElems As IJElements)
    Const METHOD = "SortTemplateDataElems"
    On Error GoTo ErrorHandler
    Dim strErrorMsg As String
        
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateDataElems is Nothing "
        GoTo ErrorHandler
    End If
    
    Dim oSortedTemplateDataElems As IJElements
    Set oSortedTemplateDataElems = New JObjectCollection
    Dim oTemplateData       As TemplateData
    
    Dim oControlLine As Object
    Set oControlLine = oTemplateSet.GetControlLine
        
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
        
    Dim nCount As Long
    Dim nIndex As Long, nMinIndex As Long
    Dim dMinLength As Double, dLength As Double
    
    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
    GetEndPoints oControlLine, oStartPos, oEndPos
 
    Do While oTemplateDataElems.count > 0
        nMinIndex = 0
        dLength = 1000
        dMinLength = 10000
        For nIndex = 1 To oTemplateDataElems.count
            Set oTemplateData = oTemplateDataElems.Item(nIndex)
            m_oMfgGeomHelper.GetLengthBet2Points oControlLine, oStartPos, oTemplateData.ControlPoint, dLength
            If dLength < dMinLength Then
                dMinLength = dLength
                nMinIndex = nIndex
            End If
        Next
        Set oTemplateData = oTemplateDataElems.Item(nMinIndex)
        oSortedTemplateDataElems.Add oTemplateData
        oTemplateDataElems.Remove (nMinIndex)
    Loop
    oTemplateDataElems.Clear
    Set oTemplateDataElems = oSortedTemplateDataElems
    Set oSortedTemplateDataElems = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function GetTemplateAftForeEdges(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData) As IJElements
    Const METHOD = "GetTemplateAftForeEdges"
    On Error GoTo ErrorHandler
        
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
        
    Dim nIndex As Long
    Dim oPlateNormalVector As IJDVector, oBaseControlLineVector As IJDVector, oBasePlaneNormalVector  As IJDVector
    Dim oEdgeElems As IJElements
    Dim oAftForeEdgeElems As IJElements
    Dim oControlLine As Object
    Dim oControlStartPos As IJDPosition, oControlEndPos As IJDPosition
    Dim oExtendedControlLine As IUnknown
    Dim oPlateEdge              As Object
    Dim oBasePlane              As IJPlane
    Dim dBasePlNormalX          As Double, dBasePlNormalY          As Double, dBasePlNormalZ          As Double
    Dim oBaseControlLinePlane   As IJPlane
    Dim oCrossVector            As IJDVector
    Dim oIntersectionPtObj     As Object
   
    Set oAftForeEdgeElems = New JObjectCollection
    
    Set oPlateNormalVector = oTemplateSetData.PlateNormal.Clone
        
    On Error Resume Next
    Set oEdgeElems = m_oMfgGeomHelper.GetPlatePartEdgesInIJElements(oTemplateSetData.SurfaceBody, oTemplateSetData.IsOnBaseSide)

    On Error GoTo ErrorHandler
    
    Set oControlLine = oTemplateSet.GetControlLine
    Set oExtendedControlLine = m_oMfgGeomHelper.ExtraPolateCurve(oControlLine, 1, 1, 2, Nothing, Nothing, 0.3) 'Extend 0.3meter

    GetEndPoints oControlLine, oControlStartPos, oControlEndPos
    Set oBaseControlLineVector = New DVector
    oBaseControlLineVector.Set oControlStartPos.x - oControlEndPos.x, oControlStartPos.y - oControlEndPos.y, oControlStartPos.z - oControlEndPos.z
    oBaseControlLineVector.Length = 1
    
    'get the Base plane object from Template object
    Set oBasePlane = oTemplateSet.GetBasePlane
    oBasePlane.GetNormal dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ
    
    Set oBasePlaneNormalVector = New DVector
    oBasePlaneNormalVector.Set dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ
    oBasePlaneNormalVector.Length = 1
    
    ' We may not always get the intersection point if we try intersecting two curves.
    ' Hence, create the plane along base control line and intersect with edges.
    If oTemplateSetData.Orientation = "NormalToBasePlane" Then
        Set oCrossVector = oBaseControlLineVector.Cross(oBasePlaneNormalVector)
    Else  'plate normal
        Set oCrossVector = oBaseControlLineVector.Cross(oPlateNormalVector)
    End If
    m_oMfgGeomHelper.MakeTransientPlane oControlStartPos.x, oControlStartPos.y, oControlStartPos.z, _
            oCrossVector.x, oCrossVector.y, oCrossVector.z, oBaseControlLinePlane
            
    For nIndex = 1 To oEdgeElems.count
        Set oPlateEdge = oEdgeElems.Item(nIndex)
        'Intersect plane with curve
        
        On Error Resume Next 'Not all edges expected to intersect
        Set oIntersectionPtObj = m_oMfgGeomHelper.IntersectCurveWithPlane(oPlateEdge, oBaseControlLinePlane)
        On Error GoTo ErrorHandler
        
        'if edge intersects BaseControlLine
        If Not oIntersectionPtObj Is Nothing Then
            If oEdgeElems.count = 3 Then 'triangle plate
                If Not IsPointAtCurveEnd(oIntersectionPtObj, oPlateEdge) Then
                    oAftForeEdgeElems.Add oPlateEdge    'add it to collection of edge
                End If
            Else
                oAftForeEdgeElems.Add oPlateEdge    'add it to collection of edges
            End If
            Set oIntersectionPtObj = Nothing
        End If
        Set oPlateEdge = Nothing
    Next ' for each edge
    
    Set GetTemplateAftForeEdges = oAftForeEdgeElems
    
CleanUp:
    Set oAftForeEdgeElems = Nothing
    Set oPlateNormalVector = Nothing
    Set oEdgeElems = Nothing
    Set oCrossVector = Nothing
    Set oBasePlaneNormalVector = Nothing
    Set oBasePlane = Nothing
    Set oBaseControlLineVector = Nothing
    Set oControlLine = Nothing
    Set oExtendedControlLine = Nothing
    Set oAftForeEdgeElems = Nothing
    Exit Function
GPE_ErrorHandler: 'GPE-->GetPlatepartEdges
    Err.Raise TPL_TC_FailedToGetEdgesFromPlatePart + vbObjectError, MODULE + METHOD, "Failed to get edges from PlatePart"
    GoTo CleanUp
ICWP_ErrorHandler: 'ICWP-->IntersectCurveWithPlane
    Err.Raise TPL_TC_FailedToIntersectCurveWithPlane + vbObjectError, MODULE + METHOD, "Failed to intersect Curve with Plane"
    GoTo CleanUp
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function


Private Sub CreateTemplatePlanesInterPolated(oTemplateSet As IJDMfgTemplateSet, oPlane As IJPlane, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
    Const METHOD = "CreateTemplatePlanesInterPolated"
    On Error GoTo ErrorHandler
    
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oPlane Is Nothing Then
        strErrorMsg = strErrorMsg & "oPlane is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateDataElems is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems.count < 2 Then
        strErrorMsg = strErrorMsg & "Cannot interpolate if num of Templates is less than two"
        GoTo ErrorHandler
    End If

    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Dim nCount As Long, nIndex As Long, nTemplate As Long
    Dim dLength As Double
    
    Dim oStartPos As IJDPosition, oPoint As IJDPosition, oIntersectionPoint As IJDPosition, oPrjctPoint As IJDPosition
    Dim oTangent As IJDVector, oTemplatePlane As IJPlane, oProjectVec As IJDVector
    Dim oTemplateData As TemplateData, oTemplateData1 As TemplateData, oTemplateDataNew As TemplateData
    Dim dInterval As Double
    Dim oControlLine As Object
    Dim oPrjctdCtrlLn As Object
    dInterval = oTemplateSetData.Interval
    
    Set oControlLine = oTemplateSet.GetControlLine
    m_oMfgGeomHelper.ProjectCurveOntoPlane oControlLine, oPlane, oPrjctdCtrlLn
    
    If dInterval < 0.001 Then
        GoTo CleanUp
    End If
    
    On Error Resume Next
     For nTemplate = 1 To oTemplateDataElems.count - 1
        Set oTemplateData = oTemplateDataElems.Item(nTemplate)
        Set oTemplateData1 = oTemplateDataElems.Item(nTemplate + 1)
        
        dLength = 0
        m_oMfgGeomHelper.GetLengthBet2Points oControlLine, oTemplateData.ControlPoint, oTemplateData1.ControlPoint, dLength
        
        If dLength > dInterval Then
            
            nCount = Int(dLength / dInterval)
            Set oStartPos = oTemplateData.ControlPoint
            
            For nIndex = 1 To nCount
                Set oPoint = m_oMfgGeomHelper.GetPointAtDistAlongCurve(oControlLine, oStartPos, dInterval)
                Set oPrjctPoint = GetProjectedPointOnPlane(oPoint, oPlane)
                            
                Set oTangent = m_oMfgGeomHelper.GetTangentByPointOnCurve(oPrjctdCtrlLn, oPrjctPoint)
                
                m_oMfgGeomHelper.MakeTransientPlane oPoint.x, oPoint.y, oPoint.z, _
                                        oTangent.x, oTangent.y, oTangent.z, oTemplatePlane
                
                Set oTemplateDataNew = New TemplateData
                oTemplateDataNew.TemplateType = System ' this is NOT a frame, but a plane
                Set oTemplateDataNew.ControlPoint = oPoint
                Set oTemplateDataNew.Plane = oTemplatePlane
                Set oTemplateDataNew.DefiningObj = oTemplatePlane
                oTemplateDataElems.Add oTemplateDataNew
                Set oTemplateDataNew = Nothing
                Set oStartPos = Nothing
                Set oStartPos = oPoint
                Set oTemplatePlane = Nothing
                Set oTangent = Nothing
                Set oPoint = Nothing
            Next
        End If
      Next
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub CreateTemplatePlanesAtInflectionPt(oTemplateSet As IJDMfgTemplateSet, oPlane As IJPlane, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
    Const METHOD = "CreateTemplatePlanesAtInflectionPt"
    On Error GoTo ErrorHandler
    
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oPlane Is Nothing Then
        strErrorMsg = strErrorMsg & "oPlane is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        ' create a new collection
        Set oTemplateDataElems = New JObjectCollection
    End If

    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    

    Dim nIndex As Long


    Dim oControlLine As Object
    Dim oPrjctdCtrlLn As Object
    Dim oIntersectionPt As IJDPosition
    
    Set oControlLine = oTemplateSet.GetControlLine
    m_oMfgGeomHelper.ProjectCurveOntoPlane oControlLine, oPlane, oPrjctdCtrlLn
    
    
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    1. Get Inflection Points on Projected ControlLine
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oInflectionPosElems As IJElements
    Set oInflectionPosElems = m_oMfgGeomHelper.GetInflectionPointsOnCurve(oPrjctdCtrlLn)

'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    2. Create Template Plane with tangent at inflection point as normal
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim oTemplatePlane As IJPlane
    Dim oInflectionPos As IJDPosition
    Dim oTangentVec As IJDVector
    Dim oTemplateData As TemplateData

    If Not oInflectionPosElems Is Nothing Then
        For nIndex = 1 To oInflectionPosElems.count
            Set oInflectionPos = oInflectionPosElems.Item(nIndex)
            Set oTangentVec = m_oMfgGeomHelper.GetTangentByPointOnCurve(oPrjctdCtrlLn, oInflectionPos)
             
           '  MsgBox "Point:" & oInflectionPos.x & " , " & oInflectionPos.y & " , " & oInflectionPos.z
             m_oMfgGeomHelper.MakeTransientPlane oInflectionPos.x, oInflectionPos.y, oInflectionPos.z, _
                                    oTangentVec.x, oTangentVec.y, oTangentVec.z, oTemplatePlane
                                    
            Set oIntersectionPt = GetIntersectionBetweenCurveAndPlane(oControlLine, oTemplatePlane)
            
            If Not oIntersectionPt Is Nothing Then
                Set oTemplateData = New TemplateData
                oTemplateData.TemplateType = System ' this is NOT a frame, but a plane
                Set oTemplateData.ControlPoint = oIntersectionPt
                Set oTemplateData.Plane = oTemplatePlane
                Set oTemplateData.DefiningObj = oTemplatePlane
                oTemplateDataElems.Add oTemplateData
            End If
                
            Set oTemplateData = Nothing
            Set oTemplatePlane = Nothing
            Set oTangentVec = Nothing
            Set oInflectionPos = Nothing
            Set oIntersectionPt = Nothing
        Next nIndex
    End If
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub CreateTemplateDataPlanes(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
Const METHOD = "CreateTemplateDataPlanes"
On Error GoTo ErrorHandler
        
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        Set oTemplateDataElems = New JObjectCollection
    End If
    
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper

        
        
    ' A)calculate Udirection vector for the templates
    ' B)calculate Vdirection vector for the templates
    ' Note:
    '   TemplateType = Frame
    '       Direction = Transversal only
    '       PositionEven = NotUsed
    '       PositionFrames = PositionFrame
    '   other TemplatesTypes
    '       any direction is valid and should be used
    ' C) calculate Normal vector for the templates
    
    Dim nIndex As Long
    Dim oBasePlaneNormalVector    As IJDVector
    Dim oBasePlane              As IJPlane
    Dim dBasePlNormalX          As Double, dBasePlNormalY As Double, dBasePlNormalZ As Double

    Dim oSurfaceBody As IJSurfaceBody
    Dim oTemplateData As TemplateData
   
    Dim oStartPos           As IJDPosition
    Dim oEndPos             As IJDPosition
    Dim oUVecElem           As IJElements
    Dim oIntersectCurve     As Object
    Dim oPlane As IJPlane
    Dim dRootX As Double, dRootY As Double, dRootZ As Double
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    Dim oUVector As IJDVector, oVVector As IJDVector, oTemplatePlaneNormal As IJDVector
    On Error Resume Next
      
    'get the Base plane object from Template object
    Set oBasePlane = oTemplateSet.GetBasePlane
    oBasePlane.GetNormal dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ

    Set oBasePlaneNormalVector = New DVector
    oBasePlaneNormalVector.Set dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ
    oBasePlaneNormalVector.Length = 1
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
    For nIndex = 1 To oTemplateDataElems.count
        
        Set oTemplateData = oTemplateDataElems.Item(nIndex)
        
        ' A)calculate Udirection vector for the templates
        If oTemplateData.TemplateType = System Then
            On Error GoTo ISWP_ErrorHandler 'ISWP-->IntersectSurfaceWithPlane
            m_oMfgGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oTemplateData.DefiningObj, oIntersectCurve, oStartPos, oEndPos
            On Error GoTo ErrorHandler
        Else    ' both edge and sketched templates have curves as defining object
            GetEndPoints oTemplateData.DefiningObj, oStartPos, oEndPos
        End If
        Set oUVector = oEndPos.Subtract(oStartPos)
        oUVector.Length = 1
    
        ' B)calculate Vdirection vector for the templates
       
        Set oTemplatePlaneNormal = New DVector
        ' in case of "Aft/Forward" do not consider Orientation
'       TR 121340 Aft Fore templates should be created for all Types
'       If oTemplateSetData.TemplateType = "Aft/Forward" And oTemplateData.TemplateType = Edge Then
        If oTemplateData.TemplateType = Edge Then
                If oTemplateSetData.Direction = "Longitudinal" Then  'X - Direction(Buttock)
                    oTemplatePlaneNormal.Set 0, 1, 0
                ElseIf oTemplateSetData.Direction = "Transversal" Then  'Y - Direction(Frame)
                    oTemplatePlaneNormal.Set 1, 0, 0
                Else 'Z - Direction(WaterLine)
                    oTemplatePlaneNormal.Set 0, 0, 1
                End If
        Else
            If oTemplateSetData.Orientation = "AlongFrame" Then
                If oTemplateSetData.Direction = "Longitudinal" Then  'X - Direction(Buttock)
                    oTemplatePlaneNormal.Set 0, 1, 0
                ElseIf oTemplateSetData.Direction = "Transversal" Then  'Y - Direction(Frame)
                    oTemplatePlaneNormal.Set 1, 0, 0
                Else 'Z - Direction(WaterLine)
                    oTemplatePlaneNormal.Set 0, 0, 1
                End If
            Else
                If oTemplateSetData.Orientation = "NormalToBasePlane" Then
                    Set oVVector = oBasePlaneNormalVector.Clone
                Else 'perpendicular 'use plate normal at the control point
                    oSurfaceBody.GetNormalFromPosition oTemplateData.ControlPoint, oVVector
                End If
                oVVector.Length = 1
                Set oTemplatePlaneNormal = oUVector.Cross(oVVector)
                oTemplatePlaneNormal.Length = 1
            End If
        End If
        oTemplateData.ControlPoint.Get dRootX, dRootY, dRootZ
        oTemplatePlaneNormal.Get dNormalX, dNormalY, dNormalZ
        m_oMfgGeomHelper.MakeTransientPlane dRootX, dRootY, dRootZ, dNormalX, dNormalY, dNormalZ, oPlane
        Set oTemplateData.Plane = oPlane
        Set oPlane = Nothing
    Next
CleanUp:
    Set oBasePlane = Nothing
    Set oBasePlaneNormalVector = Nothing
    Set oSurfaceBody = Nothing
    Set oTemplateData = Nothing
    Set oPlane = Nothing
    Exit Sub
ISWP_ErrorHandler: 'ISWP-->IntersectSurfaceWithPlane
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToIntersectSurfaceWithPlane, , "RULES")
    GoTo CleanUp
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub
Private Sub CreateTemplateDataBoxPlanes(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
Const METHOD = "CreateTemplateDataBoxPlanes"
On Error GoTo ErrorHandler
        
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        Set oTemplateDataElems = New JObjectCollection
    End If
    
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper

        
        
    ' A)calculate Udirection vector for the templates
    ' B)calculate Vdirection vector for the templates
    ' Note:
    '   TemplateType = Frame
    '       Direction = Transversal only
    '       PositionEven = NotUsed
    '       PositionFrames = PositionFrame
    '   other TemplatesTypes
    '       any direction is valid and should be used
    ' C) calculate Normal vector for the templates
    
    Dim nIndex As Long
    Dim oBasePlaneNormalVector    As IJDVector
    Dim oBasePlane              As IJPlane
    Dim dBasePlNormalX          As Double, dBasePlNormalY As Double, dBasePlNormalZ As Double

    Dim oSurfaceBody As IJSurfaceBody
    Dim oTemplateData As TemplateData
   
    Dim oStartPos           As IJDPosition
    Dim oEndPos             As IJDPosition
    Dim oUVecElem           As IJElements
    Dim oIntersectCurve     As Object
    Dim oPlane As IJPlane
    Dim dRootX As Double, dRootY As Double, dRootZ As Double
    Dim dNormalX As Double, dNormalY As Double, dNormalZ As Double
    Dim oUVector As IJDVector, oVVector As IJDVector, oTemplatePlaneNormal As IJDVector
    On Error Resume Next
      
    'get the Base plane object from Template object
    Set oBasePlane = oTemplateSet.GetBasePlane
    oBasePlane.GetNormal dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ

    Set oBasePlaneNormalVector = New DVector
    oBasePlaneNormalVector.Set dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ
    oBasePlaneNormalVector.Length = 1
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
    For nIndex = 1 To oTemplateDataElems.count
        
        Set oTemplateData = oTemplateDataElems.Item(nIndex)
        
        ' A)calculate Udirection vector for the templates
        If oTemplateData.TemplateType = System Then
            On Error GoTo ISWP_ErrorHandler 'ISWP-->IntersectSurfaceWithPlane
            m_oMfgGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oTemplateData.DefiningObj, oIntersectCurve, oStartPos, oEndPos
            On Error GoTo ErrorHandler
        Else    ' both edge and sketched templates have curves as defining object
            GetEndPoints oTemplateData.DefiningObj, oStartPos, oEndPos
        End If
        Set oUVector = oEndPos.Subtract(oStartPos)
        
    
        ' B)calculate Vdirection vector for the templates
       
        Set oTemplatePlaneNormal = New DVector
        If oTemplateSetData.Orientation = "AlongFrame" Then
            If oTemplateData.GroupType = Primary Then
                If oTemplateSetData.Direction = "Longitudinal" Then  'X - Direction(Buttock)
                    oTemplatePlaneNormal.Set 0, 1, 0
                ElseIf oTemplateSetData.Direction = "Transversal" Then  'Y - Direction(Frame)
                    oTemplatePlaneNormal.Set 1, 0, 0
                Else 'Z - Direction(WaterLine)
                    oTemplatePlaneNormal.Set 0, 0, 1
                End If
            Else 'Frame tempaltes can have only two directions/groups
                If oTemplateSetData.SecondaryDirection = "Longitudinal" Then  'X - Direction(Buttock)
                    oTemplatePlaneNormal.Set 0, 1, 0
                ElseIf oTemplateSetData.SecondaryDirection = "Transversal" Then  'Y - Direction(Frame)
                    oTemplatePlaneNormal.Set 1, 0, 0
                Else 'Z - Direction(WaterLine)
                    oTemplatePlaneNormal.Set 0, 0, 1
                End If
            End If
        Else
            If oTemplateSetData.Orientation = "NormalToBasePlane" Then
                Set oVVector = oBasePlaneNormalVector.Clone
            Else 'perpendicular 'use plate normal at the control point (Does not apply for Box)
                oSurfaceBody.GetNormalFromPosition oTemplateData.ControlPoint, oVVector
            End If
            oVVector.Length = 1
            Set oTemplatePlaneNormal = oUVector.Cross(oVVector)
            oTemplatePlaneNormal.Length = 1
        End If
        
        Dim oControlPoint As IJDPosition
        Set oControlPoint = oTemplateData.ControlPoint
        oControlPoint.Get dRootX, dRootY, dRootZ 'Take a point on the curve
        oTemplatePlaneNormal.Get dNormalX, dNormalY, dNormalZ
        m_oMfgGeomHelper.MakeTransientPlane dRootX, dRootY, dRootZ, dNormalX, dNormalY, dNormalZ, oPlane
        Set oTemplateData.Plane = oPlane
        Set oPlane = Nothing
    Next
CleanUp:
    Set oBasePlane = Nothing
    Set oBasePlaneNormalVector = Nothing
    Set oSurfaceBody = Nothing
    Set oTemplateData = Nothing
    Set oPlane = Nothing
    Exit Sub
ISWP_ErrorHandler: 'ISWP-->IntersectSurfaceWithPlane
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FailedToIntersectSurfaceWithPlane, , "RULES")
    GoTo CleanUp
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

Private Function IJDMfgTemplateService_CreateTemplatePlane(ByVal pPlatePart As Object, ByVal pUnkPosition As Object, ByVal pProcessSettings As Object, ByVal pDispTemplateSet As Object, Optional ByVal lGroupNumber As Long = 1&) As Object
Const METHOD = "IJDMfgTemplateService_CreateTemplatePlane"
On Error GoTo ErrorHandler

    Dim strErrorMsg As String

    If pProcessSettings Is Nothing Then
        strErrorMsg = strErrorMsg & "pProcessSettings is Nothing "
        GoTo ErrorHandler
    End If

    Dim oTemplateSetData As TemplateSetData
    Set oTemplateSetData = New TemplateSetData
    oTemplateSetData.InitSettings pPlatePart, pProcessSettings, pDispTemplateSet
    
    Dim strOrientation As String

    'TemplateOrientation
    strOrientation = pProcessSettings.TemplateOrientation


    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    
    Dim oRootPoint As IJDPosition
    Set oRootPoint = pUnkPosition
        
    Dim oTemplateSet As IJDMfgTemplateSet
    Set oTemplateSet = pDispTemplateSet
    
    If strOrientation = PROC_PARAM_ALONGFRAME Then
        'If 'Along frame', construct template plane at the point along Frame
        Dim oTemplatePlaneVector As IJDVector
        Set oTemplatePlaneVector = New DVector
        If oTemplateSetData.Direction = "Longitudinal" Then 'X - Direction(Buttock)
            oTemplatePlaneVector.Set 0, 1, 0
        ElseIf oTemplateSetData.Direction = "Transversal" Then 'Y - Direction(Frame)
            oTemplatePlaneVector.Set 1, 0, 0
        Else 'Z - Direction(WaterLine)
            oTemplatePlaneVector.Set 0, 0, 1
        End If
        
        'create a plane at pUnkPosition in oTemplatePlaneVector direction
        Dim oPlaneFrame As IJPlane
        Set oPlaneFrame = New Plane3d
        oPlaneFrame.SetRootPoint oRootPoint.x, oRootPoint.y, oRootPoint.z
        oPlaneFrame.SetNormal oTemplatePlaneVector.x, oTemplatePlaneVector.y, oTemplatePlaneVector.z
        
        Set IJDMfgTemplateService_CreateTemplatePlane = oPlaneFrame
        
    Else
        '   Create a Plane using BCL direction as normal
        '   Get the intersection of plane and plate surface
        '   Get the direction from intersection result. This is U Vector
        
        Dim oSurfaceBody As IJSurfaceBody
        Dim oIntersectCurve     As Object

        Set oSurfaceBody = oTemplateSetData.SurfaceBody

        Dim oUVector As IJDVector, oVVector As IJDVector, oTempVector As IJDVector
        Dim oStartPos As IJDPosition
        Dim oEndPos As IJDPosition

        Set oTempVector = New DVector
        Dim strDirection As String
        
        If lGroupNumber = Primary Then
            strDirection = oTemplateSetData.Direction
        Else
            strDirection = oTemplateSetData.SecondaryDirection
        End If
               
         Select Case strDirection
            Case "Transversal"
                oTempVector.Set 1, 0, 0
            Case "Longitudinal"
                oTempVector.Set 0, 1, 0
            Case "Waterline"
                oTempVector.Set 0, 0, 1
            End Select
                
        Dim oPlaneUVector As IJPlane
        Set oPlaneUVector = New Plane3d
        oPlaneUVector.SetRootPoint oRootPoint.x, oRootPoint.y, oRootPoint.z
        oPlaneUVector.SetNormal oTempVector.x, oTempVector.y, oTempVector.z
        
        m_oMfgGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oPlaneUVector, oIntersectCurve, oStartPos, oEndPos
        
        Set oUVector = oStartPos.Subtract(oEndPos)
        oUVector.Length = 1
        
        If strOrientation = PROC_PARAM_PERPENDICULAR Then
            '   Plate Normal at the point is V vector
            '   U cross V is desired template plane normal
            oSurfaceBody.GetNormalFromPosition oRootPoint, oVVector

'            Dim oPlanePerpendicularToPlate As IJPlane
'            Set oPlanePerpendicularToPlate = New Plane3d
'            oPlanePerpendicularToPlate.SetRootPoint oRootPoint.x, oRootPoint.y, oRootPoint.z
'            oPlanePerpendicularToPlate.GetNormal oFinalVector.x, oFinalVector.y, oFinalVector.z
'
'            Set IJDMfgTemplateService_CreateTemplatePlane = oPlanePerpendicularToPlate

        Else
            '   Normal To Base Plane is V vector
            '   U cross V is desired template plane normal
                
            Dim oBasePlane              As IJPlane
            Dim dBasePlNormalX          As Double, dBasePlNormalY As Double, dBasePlNormalZ As Double
            
            'get the Base plane object from Template object
            Set oBasePlane = oTemplateSet.GetBasePlane
            oBasePlane.GetNormal dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ
        
            Set oVVector = New DVector
            oVVector.Set dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ
            oVVector.Length = 1
            
        End If
        
        Dim oFinalVector As IJDVector
        Set oFinalVector = oUVector.Cross(oVVector)
        oFinalVector.Length = 1
        
        Dim oOutputPlane As IJPlane
        Set oOutputPlane = New Plane3d
        oOutputPlane.SetRootPoint oRootPoint.x, oRootPoint.y, oRootPoint.z
        oOutputPlane.SetNormal oFinalVector.x, oFinalVector.y, oFinalVector.z
        
        Set IJDMfgTemplateService_CreateTemplatePlane = oOutputPlane

    End If

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, strErrorMsg, "SMCustomWarningMessages", TPL_TC_FailedToCreateBottomLines, , "RULES")
End Function



Private Function IJDMfgTemplateService_ValidateProcessSettings(ByVal oProcessSettings As Object) As Boolean
Const METHOD = "IJDMfgTemplateService_ValidateProcessSettings"
On Error GoTo ErrorHandler

'1. If Type is Frame, Direction should always be Transversal (i.e. Y-direction).
'2. If Orientation is AlongFrame, PositionEven should be NotUsed and PositionFrames should be PositionFrame
'3. If Type is Perpendicular, Orientation is Perpendicular. (redundant type when you can choose Frame with orientation=Perpendicular)
'4. PositionEven and PositionFrame are mutually exclusive ( cannot occur together

    Dim strErrorMsg As String

    If oProcessSettings Is Nothing Then
        strErrorMsg = strErrorMsg & "oProcessSettings is Nothing "
        GoTo ErrorHandler
    End If

    'get values from template process settings object
    Dim strTempType As String
    Dim strOrientation As String
    Dim strDirection As String
    Dim strPositionEven As String
    Dim strPositionFrames As String

    'TemplateType
    strTempType = oProcessSettings.TemplateType

    'TemplateOrientation
    strOrientation = oProcessSettings.TemplateOrientation

    'TemplateDirection
    strDirection = oProcessSettings.TemplateDirection

    'TemplatePositionEven
    strPositionEven = oProcessSettings.TemplatePositionEven

    'TemplatePositionFrames
    strPositionFrames = oProcessSettings.TemplatePositionFrames

    '1. If Type is Frame, Direction should always be Transversal (i.e. Y-direction).
    If strTempType = PROC_PARAM_FRAME Then
        If strDirection <> PROC_PARAM_TRANSVERSAL Then
            Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_VPS_InvalidDirectionForFrameType, , "RULES")
            Exit Function
        End If
    End If

    '2. If Orientation is AlongFrame, PositionEven should be NotUsed and PositionFrames should be PositionFrame
        If strOrientation = PROC_PARAM_ALONGFRAME Then
        If strPositionEven <> PROC_PARAM_NOTUSED Then
            Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_VPS_InvalidPositionEvenForAlongFrameOrientation, , "RULES")
            Exit Function
        End If
        If strPositionFrames <> PROC_PARAM_POSITIONFRAME Then
            Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_VPS_InvalidPositionFramesForAlongFrameOrientation, , "RULES")
            Exit Function
        End If
    End If

    '3. If Type is Perpendicular, Orientation is Perpendicular. (redundant type when you can choose Frame with orientation=Perpendicular)
    If strTempType = PROC_PARAM_PERPENDICULAR Then
        If strOrientation <> PROC_PARAM_PERPENDICULAR Then
            Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_VPS_InvalidOrientationForPerpendicularType, , "RULES")
            Exit Function
       End If
    End If

    '4. PositionEven and PositionFrame are mutually exclusive ( cannot occur together)
    If strPositionEven = PROC_PARAM_POSITIONEVEN And strPositionFrames = PROC_PARAM_POSITIONFRAME Then
        Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_VPS_InvalidPositionEvenAndPositionFrame, , "RULES")
        Exit Function
    End If
    IJDMfgTemplateService_ValidateProcessSettings = True
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_FailedToValidateProcessSettings, , "RULES")
End Function


'********************************************************************************************************************************
' Function Name:    CreateTemplateDataBox
'
' Interface:        <private>
'
' Inputs:           oTemplateSet --
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the TemplateDataElems by setting the controlpoint, plane etc
'
' Assumption:
'
' Algorithm:        1. Get Frames in the range of the surface
'                   2. Make Template Data Collection
'                        a. For each Frame, do the following
'                             Get the diection of frame (If not Primary or Secondary, ignore)
'                             Check if it intersects the surface
'                             If it intersects along edge, ignore it
'                             Create Template Data
'                             Store the Frame
'                             Store the GroupType based on direction
'                   3. Get Edges
'                        a. For each Edge, do following
'                            Check if it is in Primary (if so, create Template Data)
'                            Check if it is in Secondary (if so, create Template Data)
'********************************************************************************************************************************'
Private Sub CreateTemplateDataBox(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
    Const METHOD = "CreateTemplateDataBox"
    On Error GoTo ErrorHandler

      
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        Set oTemplateDataElems = New JObjectCollection
    End If
    
    Dim oTemplateData       As TemplateData
        
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
 
    Dim nIndex As Long
    Dim oControlLine As Object
    Dim oControlStartPos As IJDPosition, oControlEndPos As IJDPosition
    Dim oExtendedControlLine As IUnknown
    Dim oPlateEdge              As Object
    Dim oAftForeEdges As IJElements
    Dim oFrame As Object
    Dim oControlPoint As IJDPosition
    Dim oFrameNormal As IJDVector
    Dim oIHFrame As IHFrame
    Dim oFramePos As IJDPosition
    Dim oFrameElems As IJElements
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    Dim oMidPos As IJDPosition
    
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
    
    'Create template data at Edges
    Dim oEdgeElems As IJElements
    Set oEdgeElems = m_oMfgGeomHelper.GetPlatePartEdgesInIJElements(oTemplateSetData.SurfaceBody, oTemplateSetData.IsOnBaseSide)
    Dim oEdge As Object
    
    'get the Base plane object from Template object
    Dim oBasePlane As IJPlane
    Dim dBasePlNormalX As Double, dBasePlNormalY As Double, dBasePlNormalZ As Double
    
    Set oBasePlane = oTemplateSet.GetBasePlane
    oBasePlane.GetNormal dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ

    Dim oBasePlaneNormalVector As IJDVector
    Set oBasePlaneNormalVector = New DVector
    oBasePlaneNormalVector.Set dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ
    oBasePlaneNormalVector.Length = 1
    
    For nIndex = 1 To oEdgeElems.count
        Set oEdge = oEdgeElems.Item(nIndex)
        Dim oEdgePlaneNormal As IJDVector
        Dim bIsLine As Boolean
        Dim oUVector As IJDVector
        Dim oVVector As IJDVector
        Dim oTemplatePlaneNormal As IJDVector
            
        GetEndPoints oEdge, oStartPos, oEndPos
        Set oMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
        Set oUVector = oEndPos.Subtract(oStartPos)
        
        If oTemplateSetData.Orientation = "NormalToBasePlane" Then
            Set oVVector = oBasePlaneNormalVector.Clone
            Set oTemplatePlaneNormal = oUVector.Cross(oVVector)
        ElseIf oTemplateSetData.Orientation = "AlongFrame" Then
            If IsCurveAlongDirection(oEdge, oSurfaceBody, oTemplateSetData.Direction) Then
                Set oTemplatePlaneNormal = New DVector
                If oTemplateSetData.Direction = "Longitudinal" Then 'X - Direction(Buttock)
                    oTemplatePlaneNormal.Set 0, 1, 0
                ElseIf oTemplateSetData.Direction = "Transversal" Then 'Y - Direction(Frame)
                    oTemplatePlaneNormal.Set 1, 0, 0
                Else 'Z - Direction(WaterLine)
                    oTemplatePlaneNormal.Set 0, 0, 1
                End If
            Else
                If IsCurveAlongDirection(oEdge, oSurfaceBody, oTemplateSetData.SecondaryDirection) Then
                    Set oTemplatePlaneNormal = New DVector
                    If oTemplateSetData.SecondaryDirection = "Longitudinal" Then 'X - Direction(Buttock)
                        oTemplatePlaneNormal.Set 0, 1, 0
                    ElseIf oTemplateSetData.SecondaryDirection = "Transversal" Then 'Y - Direction(Frame)
                        oTemplatePlaneNormal.Set 1, 0, 0
                    Else 'Z - Direction(WaterLine)
                        oTemplatePlaneNormal.Set 0, 0, 1
                    End If
                Else
                    'oTemplatePlaneVector will be nothing
                End If
            End If
            
        Else 'perpendicular 'use plate normal at the control point
            oSurfaceBody.GetNormalFromPosition oMidPos, oVVector
            Set oTemplatePlaneNormal = oUVector.Cross(oVVector)
        End If
        
        If Not oTemplatePlaneNormal Is Nothing Then
            If IsCurveOnGivenPlane(oEdge, oTemplatePlaneNormal) Then
                'We can place templates as per orientation
                Set oTemplateData = New TemplateData
                oTemplateData.TemplateType = Edge
                Set oTemplateData.DefiningObj = oEdge
                If IsCurveAlongDirection(oEdge, oSurfaceBody, oTemplateSetData.Direction) Then
                    oTemplateData.GroupType = Primary
                Else
                    If IsCurveAlongDirection(oEdge, oSurfaceBody, oTemplateSetData.SecondaryDirection) Then
                        oTemplateData.GroupType = Secondary
                    Else 'neither along primary nor along secondary
                        oTemplateData.GroupType = NotInPrimaryOrSecondary
                    End If
                End If
                
                'Add mid point as control point
                Set oTemplateData.ControlPoint = oMidPos
    
                oTemplateDataElems.Add oTemplateData
            End If
        End If
    Next nIndex
    
    'create template data at Frames
    Set oFrameElems = oTemplateSetData.FrameElems
    For nIndex = 1 To oFrameElems.count
        Set oFrame = oFrameElems.Item(nIndex)
        Dim oIntersectionCurve As Object
        m_oMfgGeomHelper.IntersectSurfaceWithPlane oSurfaceBody, oFrame, oIntersectionCurve, oStartPos, oEndPos
        
        If Not oIntersectionCurve Is Nothing Then
            If Not IsCurveOneOfEdges(oIntersectionCurve, oEdgeElems) Then
                Set oTemplateData = New TemplateData
                oTemplateData.TemplateType = System
                Set oTemplateData.DefiningObj = oFrame
                If IsFrameAlongDirection(oFrame, oTemplateSetData.Direction) Then
                    oTemplateData.GroupType = Primary
                Else
                    oTemplateData.GroupType = Secondary
                End If
                'Add mid point as control point
                Set oMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
                Set oTemplateData.ControlPoint = oMidPos
                
                oTemplateDataElems.Add oTemplateData
            End If
        End If
    Next nIndex
        
    
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

'********************************************************************************************************************************
' Function Name:    CreateTemplateDataBoxSketched
'
' Interface:        <private>
'
' Inputs:           oTemplateSet --
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the TemplateDataElems by setting the controlpoint, plane etc
'
' Assumption:
'
' Algorithm:     a. For each SketchingLine, do the following
'                        Project the sketching Line onto surface
'                        Check if projected line is along Primary or Secondary
'                        Create Template Data
'                        Store the Projected Sketched Line
'                        Store the GroupType
'********************************************************************************************************************************'
Private Sub CreateTemplateDataBoxSketched(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oSketchingLineElems As IJElements, oTemplateDataElems As IJElements)
    Const METHOD = "CreateTemplateDataBoxSketched"
    On Error GoTo ErrorHandler

      
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    
    If oSketchingLineElems Is Nothing Then
        Exit Sub
    End If
        
    If oTemplateDataElems Is Nothing Then
        Set oTemplateDataElems = New JObjectCollection
    End If
    
    Dim oTemplateData       As TemplateData
        
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
 
    Dim nIndex As Long
    Dim oControlLine As Object
    Dim oControlStartPos As IJDPosition, oControlEndPos As IJDPosition
    Dim oExtendedControlLine As IUnknown
    Dim oPlateEdge              As Object
    Dim oAftForeEdges As IJElements
    Dim oFrame As Object
    Dim oControlPoint As IJDPosition
    Dim oFrameNormal As IJDVector
    Dim oIHFrame As IHFrame
    Dim oFramePos As IJDPosition
    Dim oFrameElems As IJElements
 
    
    Dim oSurfaceBody As IJSurfaceBody
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
        
    'get sketching plane of TemplateSet
    Dim oSketchingPlane As IJPlane
    
    Set oSketchingPlane = oTemplateSet.GetSketchingPlane
    
    Dim oProjectionDir As IJDVector
    If Not oSketchingPlane Is Nothing Then
        Set oProjectionDir = New DVector
        Dim xPlaneNormal As Double, yPlaneNormal As Double, zPlaneNormal As Double
        Call oSketchingPlane.GetNormal(xPlaneNormal, yPlaneNormal, zPlaneNormal)
        Call oProjectionDir.Set(xPlaneNormal, yPlaneNormal, zPlaneNormal)
    End If
    
    Dim oProjectedCS  As IJComplexString
 
    'create template data at oSketchingLineElems
    For nIndex = 1 To oSketchingLineElems.count
        Dim oSketchedLine As Object
        Set oSketchedLine = oSketchingLineElems.Item(nIndex)
        
        On Error GoTo PCSTS_ErrorHandler 'PCSTS-->ProjectComplexStringToSurface
        m_oMfgMGHelper.ProjectComplexStringToSurface oSketchedLine, oSurfaceBody, oProjectionDir, oProjectedCS
        On Error Resume Next 'It is possible that a curve or two may not have intersection
                
        If Not oProjectedCS Is Nothing Then
            Set oTemplateData = New TemplateData
            oTemplateData.TemplateType = Sketched
            Set oTemplateData.DefiningObj = oProjectedCS
            If IsCurveAlongDirection(oProjectedCS, oSurfaceBody, oTemplateSetData.Direction) Then
                oTemplateData.GroupType = Primary
            Else
                If IsCurveAlongDirection(oProjectedCS, oSurfaceBody, oTemplateSetData.SecondaryDirection) Then
                    oTemplateData.GroupType = Secondary
                Else
                    oTemplateData.GroupType = NotInPrimaryOrSecondary
                End If
            End If
            'Add mid point as control point
            Dim oStartPos As IJDPosition
            Dim oEndPos As IJDPosition
            Dim oMidPos As IJDPosition
            GetEndPoints oProjectedCS, oStartPos, oEndPos
            Set oMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
            Set oTemplateData.ControlPoint = oMidPos
            
            oTemplateDataElems.Add oTemplateData
        End If
    Next nIndex
        
CleanUp:
    Exit Sub
PCSTS_ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", TPL_TC_FaliedToProjectComplexStringToSurf, , "RULES")
    GoTo CleanUp
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub
'********************************************************************************************************************************
' Function Name:    CreateTemplateDataBoxWithEdges
'
' Interface:        <private>
'
' Inputs:           oTemplateSet --
'                   TemplateSetData -- contains the settings for TemplateSet ex: Type, Side etc
'                   TemplateDataElems -- Collection of objects which specify control point, plane etc. for each template
'
' Output:           <indirect> modifies the TemplateDataElems by setting the controlpoint, plane etc
'
' Assumption:
'
' Algorithm:
'                   1. Make Template Data Collection
'                   2. Get Edges
'                        a. For each Edge, do following
'                            Check if it is in Primary (if so, create Template Data)
'                            Check if it is in Secondary (if so, create Template Data)
'********************************************************************************************************************************'
Private Sub CreateTemplateDataBoxWithEdges(oTemplateSet As IJDMfgTemplateSet, oTemplateSetData As TemplateSetData, oTemplateDataElems As IJElements)
    Const METHOD = "CreateTemplateDataBoxWithEdges"
    On Error GoTo ErrorHandler

      
    Dim strErrorMsg As String
    
    'check if inputs are okay
    If oTemplateSet Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSet is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateSetData Is Nothing Then
        strErrorMsg = strErrorMsg & "oTemplateSetData is Nothing "
        GoTo ErrorHandler
    End If
    If oTemplateDataElems Is Nothing Then
        Set oTemplateDataElems = New JObjectCollection
    End If
    
    Dim oTemplateData       As TemplateData
        
    If m_oMfgGeomHelper Is Nothing Then Set m_oMfgGeomHelper = New MfgGeomHelper
    If m_oMfgMGHelper Is Nothing Then Set m_oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
 
    Dim nIndex As Long
    Dim oControlLine As Object
    Dim oControlStartPos As IJDPosition, oControlEndPos As IJDPosition
    Dim oExtendedControlLine As IUnknown
    Dim oPlateEdge              As Object
    Dim oAftForeEdges As IJElements
    Dim oFrame As Object
    Dim oControlPoint As IJDPosition
    Dim oFrameNormal As IJDVector
    Dim oIHFrame As IHFrame
    Dim oFramePos As IJDPosition
    Dim oFrameElems As IJElements
    Dim oSurfaceBody As IJSurfaceBody
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    Dim oMidPos As IJDPosition
    Set oSurfaceBody = oTemplateSetData.SurfaceBody
    
    'Create template data at Edges
    Dim oEdgeElems As IJElements
    Set oEdgeElems = m_oMfgGeomHelper.GetPlatePartEdgesInIJElements(oTemplateSetData.SurfaceBody, oTemplateSetData.IsOnBaseSide)
    Dim oEdge As Object
    'get the Base plane object from Template object
    Dim oBasePlane As IJPlane
    Dim dBasePlNormalX As Double, dBasePlNormalY As Double, dBasePlNormalZ As Double
    
    Set oBasePlane = oTemplateSet.GetBasePlane
    oBasePlane.GetNormal dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ

    Dim oBasePlaneNormalVector As IJDVector
    Set oBasePlaneNormalVector = New DVector
    oBasePlaneNormalVector.Set dBasePlNormalX, dBasePlNormalY, dBasePlNormalZ
    oBasePlaneNormalVector.Length = 1
    
    For nIndex = 1 To oEdgeElems.count
        Set oEdge = oEdgeElems.Item(nIndex)
        Dim oEdgePlaneNormal As IJDVector
        Dim bIsLine As Boolean
        Dim oUVector As IJDVector
        Dim oVVector As IJDVector
        Dim oTemplatePlaneNormal As IJDVector
            
        GetEndPoints oEdge, oStartPos, oEndPos
        Set oMidPos = GetEdgeMidPoint(oStartPos, oEndPos)
        Set oUVector = oEndPos.Subtract(oStartPos)
        
        If oTemplateSetData.Orientation = "NormalToBasePlane" Then
            Set oVVector = oBasePlaneNormalVector.Clone
            Set oTemplatePlaneNormal = oUVector.Cross(oVVector)
        ElseIf oTemplateSetData.Orientation = "AlongFrame" Then
            If IsCurveAlongDirection(oEdge, oSurfaceBody, oTemplateSetData.Direction) Then
                Set oTemplatePlaneNormal = New DVector
                If oTemplateSetData.Direction = "Longitudinal" Then 'X - Direction(Buttock)
                    oTemplatePlaneNormal.Set 0, 1, 0
                ElseIf oTemplateSetData.Direction = "Transversal" Then 'Y - Direction(Frame)
                    oTemplatePlaneNormal.Set 1, 0, 0
                Else 'Z - Direction(WaterLine)
                    oTemplatePlaneNormal.Set 0, 0, 1
                End If
            Else
                If IsCurveAlongDirection(oEdge, oSurfaceBody, oTemplateSetData.SecondaryDirection) Then
                    Set oTemplatePlaneNormal = New DVector
                    If oTemplateSetData.SecondaryDirection = "Longitudinal" Then 'X - Direction(Buttock)
                        oTemplatePlaneNormal.Set 0, 1, 0
                    ElseIf oTemplateSetData.SecondaryDirection = "Transversal" Then 'Y - Direction(Frame)
                        oTemplatePlaneNormal.Set 1, 0, 0
                    Else 'Z - Direction(WaterLine)
                        oTemplatePlaneNormal.Set 0, 0, 1
                    End If
                Else
                    'oTemplatePlaneVector will be nothing
                End If
            End If
            
        Else 'perpendicular 'use plate normal at the control point
            oSurfaceBody.GetNormalFromPosition oMidPos, oVVector
            Set oTemplatePlaneNormal = oUVector.Cross(oVVector)
        End If
        
        If Not oTemplatePlaneNormal Is Nothing Then
            If IsCurveOnGivenPlane(oEdge, oTemplatePlaneNormal) Then
                'We can place templates as per orientation
                Set oTemplateData = New TemplateData
                oTemplateData.TemplateType = Edge
                Set oTemplateData.DefiningObj = oEdge
                If IsCurveAlongDirection(oEdge, oSurfaceBody, oTemplateSetData.Direction) Then
                    oTemplateData.GroupType = Primary
                Else
                    If IsCurveAlongDirection(oEdge, oSurfaceBody, oTemplateSetData.SecondaryDirection) Then
                        oTemplateData.GroupType = Secondary
                    Else 'neither along primary nor along secondary
                        oTemplateData.GroupType = NotInPrimaryOrSecondary
                    End If
                End If
                
                'Add mid point as control point
                Set oTemplateData.ControlPoint = oMidPos
    
                oTemplateDataElems.Add oTemplateData
            End If
        End If
    Next nIndex
    
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

'Description: Function checks to see if the Frame is along givne direction(Ex:F1 ia along Transversal)
'Algorithm:
'    Based on Direction, check if Frame normal is along Axis(If Transversal, it should be along X axis)
Private Function IsFrameAlongDirection(oFrameObj As Object, strDirection As String) As Boolean
    Const METHOD = "IsFrameAlongDirection"
    On Error GoTo ErrorHandler
    
        Dim oFrame As IHFrame
        Dim oFrameAxis As IHFrameAxis
        Dim oFrameNormal As IJDVector
        Set oFrame = oFrameObj
        Set oFrameAxis = oFrame.FrameAxis
        
        Select Case strDirection
        Case "Transversal"
            If oFrameAxis.AxisType = fsXAxis Then
                IsFrameAlongDirection = True
            End If
        Case "Longitudinal"
            If oFrameAxis.AxisType = fsYAxis Then
                IsFrameAlongDirection = True
            End If
        Case "Waterline"
            If oFrameAxis.AxisType = fsZAxis Then
                IsFrameAlongDirection = True
            End If
        End Select
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
'Description: Function checks to see if the Curve is along given direction
'Algorithm:
'Get the vector along curve
'Get the Axis close to the normal of the plate
'Get the cross product of the above two vectors
'Get the  Axis based on the template direction(Ex:X for Transversal)
'They should be same for curve to be along given direction
Private Function IsCurveAlongDirection(oCurveObj As Object, oSurfaceBody As IJSurfaceBody, strDirection As String) As Boolean
    Const METHOD = "IsCurveAlongDirection"
    On Error GoTo ErrorHandler
    
    'Get Vector along Curve
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    GetEndPoints oCurveObj, oStartPos, oEndPos
    
    Dim oCurveVec As IJDVector
    Set oCurveVec = oStartPos.Subtract(oEndPos)
        
    'Get the Axis close to normal of plate(parallle Axis)
    Dim oParallelAxis As IJDVector
    Set oParallelAxis = GetParallelAxis(oSurfaceBody)
    
    'Get the cross product
    Dim oCurvePlaneNormal As IJDVector
    Set oCurvePlaneNormal = oParallelAxis.Cross(oCurveVec)
    
    'Get template plane normal
    Dim oTemplateNormal As IJDVector
    Set oTemplateNormal = New DVector
    
    Select Case strDirection
    Case "Transversal"
        oTemplateNormal.Set 1, 0, 0
    Case "Longitudinal"
        oTemplateNormal.Set 0, 1, 0
    Case "Waterline"
        oTemplateNormal.Set 0, 0, 1
    End Select
        
    'How close are these normals?
    oCurvePlaneNormal.Length = 1
    If ArcCos(Abs(oTemplateNormal.Dot(oCurvePlaneNormal))) < (15 * PI / 180) Then '15 deg tolerance
        IsCurveAlongDirection = True
    Else
        IsCurveAlongDirection = False
    End If
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Private Function IsCurveOneOfEdges(oCurve As Object, oEdges As IJElements) As Boolean
    Const METHOD = "IsCurveOneOfEdges"
    On Error GoTo ErrorHandler
    
    Dim oTopologyIntersect As IJDTopologyIntersect
    Set oTopologyIntersect = New DGeomOpsIntersect
    
    Dim i As Long
    For i = 1 To oEdges.count
        Dim bOverlap As Boolean
        oTopologyIntersect.HasOverlappingGeometry oCurve, oEdges.Item(i), bOverlap
        If bOverlap Then
            IsCurveOneOfEdges = True
            Exit Function
        End If
    Next

CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
Private Function ArcCos(x As Double) As Double
   If Abs(x) < 1 Then
      ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
   Else
      ArcCos = 0
   End If
End Function

Private Function GetCopyOfSymbolOutput(oSO As Object) As IJComplexString
    Const METHOD = "GetCopyOfSymbolOutput"
    On Error GoTo ErrorHandler
    
    Dim oComplexString3D As IComplexStrings3d
    Dim oCurves As IJElements
    
    Dim oComplexString As IJComplexString
    Set oComplexString = oSO
    oComplexString.GetCurves oCurves
    
    Set oComplexString3D = New GeometryFactory
    Set GetCopyOfSymbolOutput = oComplexString3D.CreateByCurves(Nothing, oCurves)

CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

Private Function CreateTemplatePlaneForAddedCurve(oPlatePart As IJPlatePart, oTemplateData As TemplateData, oTemplateSet As IJDMfgTemplateSet) As IJPlane
    Const METHOD = "CreateTemplatePlaneForAddedCurve"
    On Error GoTo ErrorHandler
    
    Dim oBCP As IJDPosition
    Dim oProcessSettings As IJMfgTemplateProcessSettings
    Dim lGroupNumber As Long
        
    'from tempalte data, get control point
    Set oBCP = oTemplateData.ControlPoint
    'from template, get processsettings
    Set oProcessSettings = oTemplateSet.GetProcessSettings
    'from template data, get grouptype
    If oTemplateData.GroupType = Primary Then
        lGroupNumber = Primary
    Else
        lGroupNumber = Secondary
    End If

    Set CreateTemplatePlaneForAddedCurve = IJDMfgTemplateService_CreateTemplatePlane(oPlatePart, oBCP, oProcessSettings, oTemplateSet, lGroupNumber)
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function


Private Function GetLongTypeUserChoice(strParamterName As String, oTemplateSetObj As Object) As Boolean
    Const METHOD = "GetLongTypeUserChoice"
    On Error GoTo ErrorHandler
    
    GetLongTypeUserChoice = False
    
    Dim strUserDefinedValuesProgID  As String
    Dim oSRDQuery As IJSRDQuery
    Dim oRuleQuery As IJSRDRuleQuery
    Dim oRule As IJSRDRule
    Dim oRuleObj As Object

    Set oSRDQuery = New SRDQuery
    Set oRuleQuery = oSRDQuery.GetRulesQuery
    
    Set oRuleObj = oRuleQuery.GetRule(USERDEFVALUESRULENAME)
    Set oRule = oRuleObj

    strUserDefinedValuesProgID = oRule.ProgId
    
    Dim oMfgTemplateParameters As IJDMfgTemplateParameters
    Set oMfgTemplateParameters = CreateObject(strUserDefinedValuesProgID)
    
    oMfgTemplateParameters.object = oTemplateSetObj
    
    Dim lApplymargin As Long
    lApplymargin = oMfgTemplateParameters.GetLongValue(strParamterName)
    
    If lApplymargin = 1 Then
        GetLongTypeUserChoice = True
    End If
    
    Set oSRDQuery = Nothing
    Set oRuleQuery = Nothing
    Set oRuleObj = Nothing
    Set oRule = Nothing
    Set oMfgTemplateParameters = Nothing
    
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function


