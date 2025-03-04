VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'******************************************************************
' Copyright (C) 2006, Intergraph Corporation. All rights reserved.
'
'File
'    hrTypeA.cls
'
'Author
'       SS
'
'Description
'
'Notes
'
'History:
'  10 Sept 2004 MD
' Skiniing option added which can be used to create handrail surface based on requirement
' Skinning option is used to set "brkcrv " argument of CreateBySingleSweep() function and also
' through CreateSurface wrapper provided in CrossSectionService.bas. So proper argument need to be sent
' to CreateSurface() through CreateProjectionFromCSProfile(). For more details of
' appropriate skiining  option see definition of CreateProjectionFromCSProfile() & documentation of CreateBySingleSweep
' Skinning option will not work for cases like toprail section radius is high and handrail path bend radius is
' too small for given radius. So validation code added to check if CreateBySingleSweep works
' for given path & cross section. If not then skinning option will be changed. This check is done for
' Physical  representations. Can also be added for operational representation if based on requirement
' 6 Jun 05  JMS Went to correct a problem where the Error handlers never had m_Errors
'                       declared (TR#79298) & discovered other issues -- the
'                       judiscious use of "On Error Resume Next" was masking many problems. To
'                       use this function correctly one should immediately reset back to the error
'                       handler after the offending statement; otherwise, other coding problems later
'                       in the code will be masked (which was found in several instances).
'                       Also, resume next is not necessary to determine if an object supports
'                       an interface "Typeof" should be used instead.
'   09-May-06   JMS DI#97751 - All lines are constructed through the
'                   geometry factory instead of being "new"ed
'
'   05-oct-06   SS      TR#107161 - commented out msgbox as we should not be raising any msgboxes from
'                       symbol could. If such error condition happens during sync, ifc it could
'                       stop the process.
'
'   11-Oct-08   SS      CR#38818 - add handrail path (physical and not sketched path which can be called logical)
'                       as its output,locatable but not displayable
'   05-Aug-08   WR      TR-CP-131048 - Added new method to ISPSHandrailSymbolServices to get the top rail radius
'
'   02-Oct-08   SS      CR#148021 - Add Centerline representation to support Drawings for CR112776.
'                       DLL version is bumped and will require a synchronize.
'
'   02-Oct-08   SS      CR#141607 - Removed Detailed Physical from handrail representations as it is same
'                       as SimplePhysical and redundant.
'   12-Aug-09   GG      DM-169707  - Fixed the weight issues and some COG issues
'   04-Sept-09  GG      DM-171219 Changed the way to calculate weight and COG for handrails. Now the WCG calculation
'                       is independent to the Physical evaluation. The main changes in this module are:
'                       1. Moved CalcWCG from ComputeStuff module with modifications
'                       2. Modified CreatePosts to handle the new option for calculate WCG
'                       3. Created subroutine CalcPostsVolumeCG
'
'   09-Sep-09   WR      TR-144290 - Added new method CreateProjection which nows directly calls the new method added by
'                       Core Math CreateBySingleSweepWCapsOpts to sweep the ToePlate always in the Z direction.
'   20-Nov-09 GG        DM-175208  Circular end treatment adapts to the post section size
'                       1. Always use toprail cross section for circular treatment;
'                       2. Changed path direction to be consistent with toprail;
'                       3. Fixed the mirror direction for circular treatment;
'                       4. If the assumption for circular radius is not a good one, a formular is used to get the radius.
'   14-Dec-09   MJ      TR-169906 Horizontal Offsets are not applied when figuring COG of handrail
'                       The fix for this TR is to ensure that the symbol's Calculate Weight and CG
'                       follows the same logic that is used by it's physical representation calculation.
'                       This involved significant code refactoring so that both custom methods are able
'                       to call the same code
'   10-Feb-10   GG      Fixed the following issues:
'                       1.  Use top-rail cross section for the rectangular treatment (TR-173639);
'                       2.  Corrected the orientation of the rectangular end treatment or the last post if there is no end treatment(DM -176290);
'                       3.  Fixed the centerline graphics for treatments (TR-169532, TR-169001);
'                       4.  Made the pad length and pad width consistent to the legend (TR-161195);
'                       5.  The Spacing between posts is always greater than the minimum clearance if "With Post at Turn" is false (CR-58237).
'   17-Aug-10   GG      DM 186125. Circular end treatment of handrail is not correct if the path is sloped
'                       Removed un-used function CreateCirEndTreatment
'   29-Dec-10   GG      TR 191191  The handrail computing fails for some path patterns. Also removed some obsolete subs
''*******************************************************************

Const definitionProgId As String = "SPSHandrailMacros.TypeA"

Implements IJDUserSymbolServices
Implements ISPSHandrailSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck
Implements ISPSHandrailConvertServices

Private m_complex As IngrGeom3D.ComplexString3d
Private m_transform As IngrGeom3D.IJDT4x4
Private m_GeomFactory As IngrGeom3D.GeometryFactory
Dim bOnPreLoad As Boolean

Private TopRailSection As String
Private TopRailSecStandard As String
Private TopRailSectionDepth As Double
Private TopRailSectionWidth As Double
Private TopRailSectionCP As Integer
Private ToprailSectionAngle As Double
Private TreatmentRadius As Double

'see CreateProjectionFromCSProfile() in CrossSectionService.bas for more details of skin option
Private SkinOption As Long

Private Const MODULE = "SPSHandrailMacros.PlaceTypeA"

Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259
Private m_oLocalizer As IJLocalizer

Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"

Private Sub Class_Initialize()
    Set m_GeomFactory = New IngrGeom3D.GeometryFactory
    Set m_complex = Nothing
    Set m_transform = Nothing
    Set m_oErrors = New IMSErrorLog.JServerErrors
    Set m_oLocalizer = New IMSLocalizer.Localizer
    m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
    Set m_complex = Nothing
    Set m_transform = Nothing
    Set m_GeomFactory = Nothing
    Set m_oCatResMgr = Nothing
    Set m_oErrors = Nothing
    Set m_oLocalizer = Nothing
End Sub

Private Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = definitionProgId
End Function

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParamaters As Variant, ByVal ActiveConnection As Object) As Object
  Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo ErrorHandler
  
  Dim oSymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
  Dim oSymbolDefinition As IMSSymbolEntities.DSymbolDefinition
  
  Set oSymbolDefinition = oSymbolFactory.CreateEntity(definition, ActiveConnection)
  Set oSymbolFactory = Nothing
  IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
  
  ' Set definition progId and codebase
  oSymbolDefinition.ProgId = definitionProgId
  oSymbolDefinition.CodeBase = CodeBase

  ' Give a unique name to the symbol definition
  oSymbolDefinition.name = oSymbolDefinition.ProgId
  
  ' verify if this needed ?
  ' oSymbolDefinition.GeomOption = igSYMBOL_GEOM_DRIVEN_BY_ARG      ' ComputeMatrix will be called if defined

  'return symbol defintion
  Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
  Set oSymbolDefinition = Nothing
  
  Exit Function

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function


Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler
    ' TypeA Handrail Definition
    ' Inputs:
    '          1. "Part"  ( Catalog part )
    '          2. Sketch3dObj
    '          3. Height
    '          4. WithToePlate
    '          5. NoOfMidRails
    '          6. HorizontalOffset
    '          7. HorizontalOffsetDim
    '          8. HandrailOrientation
    '          9. SegmentMaxSpacing
    '          10. SlopedSegmentMaxSpacing
    '          11. TopOfToePlateDim
    '          12. TopOfMidRailDim
    '          13. MidRailSpacing
    '          14. WithPostAtTurn
    '          15. BeginTreatmentType
    '          16. BeginExtensionLength
    '          17. EndTreatmentType
    '          18. EndExtensionLength
    '          19. PostConnectionType
    '          20. IsSystem
    '          21. IsAssembly
    '          22. TopRailSection
    '          23. TopRailSectionReferenceStandard
    '          24. TopRailSectionCP
    '          25. TopRailSectionAngle
    '          26. MidRailSection
    '          27. MidRailSectionReferenceStandard
    '          28. MidRailSectionCP
    '          29. MidRailSectionAngle
    '          30. ToePlateSection
    '          31. ToePlateSectionReferenceStandard
    '          32. ToePlateSectionCP
    '          33. ToePlateSectionAngle
    '          34. PostSection
    '          35. PostSectionReferenceStandard
    '          36. PostSectionCP
    '          37. PostSectionAngle
    
    '
    ' Representations:
    '           Physical
    '           Operation
    
    '       Outputs collection : igCollection_Variable
    '           top rail
    '           mid rails
    '           posts
    '           toe plate
    '           begin treatment
    '           end treatment
    '
    ' Remove all previous Symbol Definition information
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
  
    On Error GoTo ErrorHandler

    ' Set the input to the definition
    Dim iInputs As IMSSymbolEntities.IJDInputs
    Set iInputs = pSymbolDefinition

    Dim iUM As IMSSymbolEntities.IJDUserMethods
    Set iUM = pSymbolDefinition

    Dim libDesc As New DLibraryDescription
    Dim mCookie As Long
    Dim libCookie As Long
    Dim mthCookieChecksketch As Long
    Dim CheckMaterial As Long

    libDesc.name = "mySelfAsLib"
    libDesc.Type = imsLIBRARY_IS_ACTIVEX
    libDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    libDesc.Source = definitionProgId

    pSymbolDefinition.IJDUserMethods.SetLibrary libDesc

    ' Get the lib/method cookie
    libCookie = libDesc.Cookie
    mCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheForPart", libCookie)

        Dim ChecklibDesc As New DLibraryDescription
    Dim ChecklibCookie As Long
    Dim GTZeroCheck As Long
    Dim DegCheck As Long
    Dim mthCookieCheckSection As Long
     
    ChecklibDesc.name = "CMCheckLib"
    ChecklibDesc.Type = imsLIBRARY_IS_ACTIVEX
    ChecklibDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    ChecklibDesc.Source = CheckProgId
    pSymbolDefinition.IJDUserMethods.SetLibrary ChecklibDesc
    ChecklibCookie = ChecklibDesc.Cookie
    
    'Assign cookie values to point to the check functions.  Allows the assignment of check functions to attributes later.
    GTZeroCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("GTZero", ChecklibCookie)
    DegCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("IN360Deg", ChecklibCookie)
    mthCookieChecksketch = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CheckForSketch3d", ChecklibCookie)
    ' Retrieve the cookie of the method IsCrossSectionName
    ' This method controls that the type of the input is a cross sction
    ' don't change the name of these custom check methods
    mthCookieCheckSection = pSymbolDefinition.IJDUserMethods.GetMethodCookie("IsCrossSectionName", ChecklibCookie)
    CheckMaterial = pSymbolDefinition.IJDUserMethods.GetMethodCookie("IsMaterailGrade", ChecklibCookie)

    Dim pIJDInput As IMSSymbolEntities.IJDInput
    Set pIJDInput = New IMSSymbolEntities.DInput
    
    Dim oInput As IMSSymbolEntities.IJDInput
    Set oInput = New IMSSymbolEntities.DInput
    Dim oPC As IMSSymbolEntities.IJDParameterContent
    Set oPC = New IMSSymbolEntities.DParameterContent
    
    oPC.Type = igValue

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDefinition

    oInput.name = "Sketch3d"
    oInput.Description = "Sketch3d output complex curve"
    oInput.Properties = 0
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieChecksketch
    iInputs.SetInput oInput, 2
    oInput.Reset
    
    oInput.name = "Height"
    oInput.Description = "Height of the Handrail from path to top of Top Rail"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 1.067    'TR#42710
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 3
    oInput.Reset
    oPC.Reset

    oInput.name = "WithToePlate"
    oInput.Description = "With or Without Toe/Kick plate"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 1#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 4
    oInput.Reset
    oPC.Reset
    
    oInput.name = "NoOfMidRails"
    oInput.Description = "With One/Many/None mid rails"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 1#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 5
    oInput.Reset
    oPC.Reset
    
    oInput.name = "HorizontalOffset"
    oInput.Description = "Offset position that the handrail will be from the path chosen"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 15#  'TR#42710
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 6
    oInput.Reset
    oPC.Reset
    
    oInput.name = "HorizontalOffsetDim"
    oInput.Description = "Distance that the handrail will be offset from the path chosen"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.0762   'TR#42710
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 7
    oInput.Reset
    oPC.Reset
    
    oInput.name = "HandrailOrientation"
    oInput.Description = "Vertical / Normal"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 8
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SegmentMaxSpacing"
    oInput.Description = "Maximum spacing between posts on straight/arc line"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 1.524
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 9
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SlopedSegmentMaxSpacing"
    oInput.Description = "Maximum spacing between posts  on slope"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 1.829
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 10
    oInput.Reset
    oPC.Reset
    
    oInput.name = "TopOfToePlateDim"
    oInput.Description = "Dimension to the top of the toe plate from Path "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.102
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 11
    oInput.Reset
    oPC.Reset
    
    oInput.name = "TopOfMidRailDim"
    oInput.Description = "Dimension to the top of the Mid rail from top of toe plate "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.5334
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 12
    oInput.Reset
    oPC.Reset
    
    oInput.name = "MidRailSpacing"
    oInput.Description = "Spacing between mid rails"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.2
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 13
    oInput.Reset
    oPC.Reset
    
    oInput.name = "WithPostAtTurn"
    oInput.Description = "With or Without Post at the turn"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 1#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 14
    oInput.Reset
    oPC.Reset
    
    oInput.name = "BeginTreatmentType"
    oInput.Description = "Begin treatment type 0 for rectangular, 1 for rounded"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 2#   'TR#42710
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 15
    oInput.Reset
    oPC.Reset
    
    oInput.name = "BeginExtensionLength"
    oInput.Description = "Extent of the end of the rail (at beginning) "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.6096
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 16
    oInput.Reset
    oPC.Reset
    
    oInput.name = "EndTreatmentType"
    oInput.Description = "End treatment type 0 for rectangular, 1 for rounded"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 2#   'TR#42710
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 17
    oInput.Reset
    oPC.Reset
    
    oInput.name = "EndExtensionLength"
    oInput.Description = "Extent of the end of the rail (at ending) "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.6096
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 18
    oInput.Reset
    oPC.Reset
    
    oInput.name = "PostConnectionType"
    oInput.Description = "connection required to assemble the entire handrail"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 1#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 19
    oInput.Reset
    oPC.Reset
    
    oInput.name = "IsAssembly"
    oInput.Description = "If system placed as a symbol otherwise constructed with SPS members"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 20
    oInput.Reset
    oPC.Reset
    
    oInput.name = "TopRail_SPSSectionName"
    oInput.Description = "TopRailSection from structural cross sections"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "HSS1-7/8x.145"
    oInput.DefaultParameterValue = oPC
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
    iInputs.SetInput oInput, 21
    oInput.Reset
    oPC.Reset

    oInput.name = "TopRail_SPSSectionRefStandard"
    oInput.Description = "TopRail Reference Standard "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "AISC-LRFD-3.1"
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 22
    oInput.Reset
    oPC.Reset

    oInput.name = "TopRailSectionCP"
    oInput.Description = "TopRailSection Cardinal Point"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 5#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 23
    oInput.Reset
    oPC.Reset
    
    oInput.name = "TopRailSectionAngle"
    oInput.Description = "TopRailSection Angle"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
    oPC.Type = igValue
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 24
    oInput.Reset
    oPC.Reset
    
    oInput.name = "MidRail_SPSSectionName"
    oInput.Description = "MidRailSection from structural cross sections"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "HSS1-7/8x.145"
    oInput.DefaultParameterValue = oPC
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
    iInputs.SetInput oInput, 25
    oInput.Reset
    oPC.Reset
'
    oInput.name = "MidRail_SPSSectionRefStandard"
    oInput.Description = "MidRail Reference Standard"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "AISC-LRFD-3.1"
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 26
    oInput.Reset
    oPC.Reset

    oInput.name = "MidRailSectionCP"
    oInput.Description = "MidRailSection Cardinal Point"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 5#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 27
    oInput.Reset
    oPC.Reset
    
    oInput.name = "MidRailSectionAngle"
    oInput.Description = "MidRailSection Angle"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
    oPC.Type = igValue
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 28
    oInput.Reset
    oPC.Reset
    
    oInput.name = "ToePlate_SPSSectionName"
    oInput.Description = "MidRailSection from structural cross sections"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "RS0.125x4"
    oInput.DefaultParameterValue = oPC
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
    iInputs.SetInput oInput, 29
    oInput.Reset
    oPC.Reset

    oInput.name = "ToePlate_SPSSectionRefStandard"
    oInput.Description = "ToePlate Reference Standard"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "Misc"
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 30
    oInput.Reset
    oPC.Reset

    oInput.name = "ToePlateSectionCP"
    oInput.Description = "ToePlateSection Cardinal Point"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 2#   'TR#42710
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 31
    oInput.Reset
    oPC.Reset
    
    oInput.name = "ToePlateSectionAngle"
    oInput.Description = "ToePlateSection Angle"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
    oPC.Type = igValue
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 32
    oInput.Reset
    oPC.Reset
    
    oInput.name = "Post_SPSSectionName"
    oInput.Description = "PostSection from structural cross sections"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "HSS1-7/8x.145"
    oInput.DefaultParameterValue = oPC
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
    iInputs.SetInput oInput, 33
    oInput.Reset
    oPC.Reset

    oInput.name = "Post_SPSSectionRefStandard"
    oInput.Description = "Post Section Reference Standard"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "AISC-LRFD-3.1"
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 34
    oInput.Reset
    oPC.Reset

    oInput.name = "PostSectionCP"
    oInput.Description = "PostSection Cardinal Point"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 5#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 35
    oInput.Reset
    oPC.Reset
    
    oInput.name = "PostSectionAngle"
    oInput.Description = "PostSection Angle"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
    oPC.Type = igValue
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 36
    oInput.Reset
    oPC.Reset
    
    oInput.name = "Primary_SPSMaterial"
    oInput.Description = "Block foundation SPS Materail"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "Steel - Carbon"   'TR#42710
    oInput.DefaultParameterValue = oPC
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, CheckMaterial
    iInputs.SetInput oInput, 37
    oInput.Reset
    oPC.Reset
    
    oInput.name = "Primary_SPSGrade"
    oInput.Description = "Block foundation SPS Grade"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "A"    'TR#42710
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, 38
    oInput.Reset
    oPC.Reset
    
    'Define the representation "Simple Physical"
    Dim pIRep As IMSSymbolEntities.IJDRepresentation
    Set pIRep = New IMSSymbolEntities.DRepresentation

    pIRep.name = "Physical"
    pIRep.Description = "Physical Representation of Handrail"
    pIRep.RepresentationId = SimplePhysical
    mCookie = iUM.GetMethodCookie("Physical", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
     
    Dim pOutputs As IMSSymbolEntities.IJDOutputs
    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable

    Dim output As IMSSymbolEntities.DOutput
    Set output = New IMSSymbolEntities.DOutput

'    output.name = "TopRail"
'    output.Description = "TopRail"
'    pOutputs.SetOutput output
'    output.Reset

'    'Set the representation to definition
    Dim pIReps As IMSSymbolEntities.IJDRepresentations
    Set pIReps = pSymbolDefinition
    pIReps.SetRepresentation pIRep
     
    'Operation Representation
    pOutputs.RemoveAllOutput
    Set output = New DOutput

    pIRep.name = "OperationRepresentation"
    pIRep.Description = "Operational representation"
    pIRep.Properties = igREPRESENTATION_ISVBFUNCTION
    pIRep.RepresentationId = Operation
    mCookie = iUM.GetMethodCookie("OperationRepresentation", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie

    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    output.name = "OperationalEnvelope1"
    output.Description = "Operational Envelope of the Handrail"
    pOutputs.SetOutput output
    output.Reset

    pIReps.SetRepresentation pIRep 'Add representation to definition
    
    ' Centerline Representation
    pOutputs.RemoveAllOutput
    Set output = New DOutput

    pIRep.name = "Centerline"
    pIRep.Description = "Center Line"
    pIRep.Properties = igREPRESENTATION_ISVBFUNCTION
    pIRep.RepresentationId = 512        ' for Centerline
    mCookie = iUM.GetMethodCookie("CenterLine", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie

    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    pIReps.SetRepresentation pIRep 'Add Centerline representation to definition
    
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
    
    Set pIRep = Nothing
    Set pIReps = Nothing
    Set output = Nothing
    Set pOutputs = Nothing


    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
    
End Sub

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputColl As Object, ByRef arrayOfInputs())
End Sub

'=========================================================================
'CREATION OF SIMPLE PHYSICAL REPRESENTATION OF Handrail
'========================================================================
Public Sub Physical(pIRepSCM As IJDRepresentationStdCustomMethod)
    
    Call PhysicalRepresentation(pIRepSCM, SimpleRep)

End Sub

'=========================================================================
'CREATION OF CENTERLINE REPRESENTATION OF Handrail
'========================================================================
Public Sub CenterLine(pIRepSCM As IJDRepresentationStdCustomMethod)
    
    Call PhysicalRepresentation(pIRepSCM, CenterLineRep)
    
End Sub

Private Sub PhysicalRepresentation(oParent As Object, RepType As eRepresentationType)
Const METHOD = "PhysicalRepresentation"
On Error GoTo ErrorHandler

    Dim pIRepSCM As IJDRepresentationStdCustomMethod
    Dim pRepDG As IJDRepresentationDuringGame
    Dim pOC As IJDOutputCollection
    Dim bNeedToRaiseWarning As Boolean
    Dim pOutputParentObject As Object         ' common object used to call subs.  can be OutputCollection or Handrail
    
    Dim oHandrail As ISPSHandrail
    Dim pInputs As IJDInputs

    If TypeOf oParent Is ISPSHandrail Then
        Set oHandrail = oParent
        Set pOutputParentObject = oParent
    Else
        Set pIRepSCM = oParent
        Set pRepDG = pIRepSCM
        Set pInputs = pRepDG.definition.IJDInputs
        Set oHandrail = pRepDG.definition.IJDDefinitionPlayerEx.PlayingSymbol
        Set pOC = pRepDG.outputCollection
        Set pOutputParentObject = pOC
    End If
    
    'assign to meaningful variables from the input array

    Dim Height As Double, SegmentMaxSpacing As Double, SlopedSegmentMaxSpacing As Double, Offset As Double
    Dim TopOfToePlateDim As Double, TopOfMidRailDim As Double, MidRailSpacing As Double, BeginExtensionLength As Double, EndExtensionLength As Double
    Dim TopRailSize As Double, MidRailSize As Double, ToePlateSize As Double, PostSize As Double, HorizOffset As Double
    Dim WithToePlate As Boolean, WithPostAtTurn As Boolean, IsSystem As Boolean, IsAssembly As Boolean
    
    Dim noofmidrails As Integer, HorizontalOffset As Integer, HandrailOrientation As Integer, PostConnectionType As Integer
    Dim BeginTreatmentType As Integer, EndTreatmentType As Integer
        
    Dim MidRailSection As String, MidRailSecStandard As String
    Dim MidRailSectionCP As Integer, MidRailSectionAngle As Double

    Dim ToePlateSection As String, ToePlateSecStandard As String
    Dim ToePlateSectionCP As Integer, ToePlateSectionAngle As Double

    Dim PostSection As String, PostSecStandard As String
    Dim PostSectionCP As Integer, PostSectionAngle As Double

    If Not pInputs Is Nothing Then
        Height = pInputs.GetInputByIndex(3).IJDInputDuringGame.Result.UomValue
        WithToePlate = pInputs.GetInputByIndex(4).IJDInputDuringGame.Result.UomValue
        
        noofmidrails = pInputs.GetInputByIndex(5).IJDInputDuringGame.Result.UomValue
        HorizontalOffset = pInputs.GetInputByIndex(6).IJDInputDuringGame.Result.UomValue
        HorizOffset = pInputs.GetInputByIndex(7).IJDInputDuringGame.Result.UomValue
        
        HandrailOrientation = pInputs.GetInputByIndex(8).IJDInputDuringGame.Result.UomValue
        
        SegmentMaxSpacing = pInputs.GetInputByIndex(9).IJDInputDuringGame.Result.UomValue
        
        SlopedSegmentMaxSpacing = pInputs.GetInputByIndex(10).IJDInputDuringGame.Result.UomValue
        
        TopOfToePlateDim = pInputs.GetInputByIndex(11).IJDInputDuringGame.Result.UomValue
        TopOfMidRailDim = pInputs.GetInputByIndex(12).IJDInputDuringGame.Result.UomValue
        MidRailSpacing = pInputs.GetInputByIndex(13).IJDInputDuringGame.Result.UomValue
    
        WithPostAtTurn = pInputs.GetInputByIndex(14).IJDInputDuringGame.Result.UomValue
        
        BeginTreatmentType = pInputs.GetInputByIndex(15).IJDInputDuringGame.Result.UomValue
        ' treatment type code list value ; 5 - circular 10 - rectangular; 2 - none
        BeginExtensionLength = pInputs.GetInputByIndex(16).IJDInputDuringGame.Result.UomValue
        
        If BeginTreatmentType = 2 Then BeginExtensionLength = 0#
            
        EndTreatmentType = pInputs.GetInputByIndex(17).IJDInputDuringGame.Result.UomValue
        ' treatment type code list value ; 5 - circular 10 - rectangular 2 -none
        EndExtensionLength = pInputs.GetInputByIndex(18).IJDInputDuringGame.Result.UomValue
        
        If (EndTreatmentType = 2) Then EndExtensionLength = 0#
           
        PostConnectionType = pInputs.GetInputByIndex(19).IJDInputDuringGame.Result.UomValue
        
        IsAssembly = pInputs.GetInputByIndex(20).IJDInputDuringGame.Result.UomValue
        
        TopRailSection = pInputs.GetInputByIndex(21).IJDInputDuringGame.Result.String
        TopRailSecStandard = pInputs.GetInputByIndex(22).IJDInputDuringGame.Result.String
        TopRailSectionCP = pInputs.GetInputByIndex(23).IJDInputDuringGame.Result.UomValue
        ToprailSectionAngle = pInputs.GetInputByIndex(24).IJDInputDuringGame.Result.UomValue
        
        MidRailSection = pInputs.GetInputByIndex(25).IJDInputDuringGame.Result.String
        MidRailSecStandard = pInputs.GetInputByIndex(26).IJDInputDuringGame.Result.String
        MidRailSectionCP = pInputs.GetInputByIndex(27).IJDInputDuringGame.Result.UomValue
        MidRailSectionAngle = pInputs.GetInputByIndex(28).IJDInputDuringGame.Result.UomValue
        
        ToePlateSection = pInputs.GetInputByIndex(29).IJDInputDuringGame.Result.String
        ToePlateSecStandard = pInputs.GetInputByIndex(30).IJDInputDuringGame.Result.String
        ToePlateSectionCP = pInputs.GetInputByIndex(31).IJDInputDuringGame.Result.UomValue
        ToePlateSectionAngle = pInputs.GetInputByIndex(32).IJDInputDuringGame.Result.UomValue
        
        PostSection = pInputs.GetInputByIndex(33).IJDInputDuringGame.Result.String
        PostSecStandard = pInputs.GetInputByIndex(34).IJDInputDuringGame.Result.String
        PostSectionCP = pInputs.GetInputByIndex(35).IJDInputDuringGame.Result.UomValue
        PostSectionAngle = pInputs.GetInputByIndex(36).IJDInputDuringGame.Result.UomValue
    
    Else
    
        Dim oAttrColl As Object
        
        Set oAttrColl = GetAttributeCollection(oHandrail, "IJUAHRTypeAProps")
        Height = GetAttributeValue(oAttrColl, "Height")
        WithToePlate = GetAttributeValue(oAttrColl, "WithToePlate")
        
        noofmidrails = GetAttributeValue(oAttrColl, "NoOfMidRails")
        HorizOffset = GetAttributeValue(oAttrColl, "HorizontalOffsetDim")
        
        HandrailOrientation = GetAttributeValue(oAttrColl, "HandrailOrientation")

        SegmentMaxSpacing = GetAttributeValue(oAttrColl, "SegmentMaxSpacing")
        
        SlopedSegmentMaxSpacing = GetAttributeValue(oAttrColl, "SlopedSegmentMaxSpacing")
        
        TopOfToePlateDim = GetAttributeValue(oAttrColl, "TopOfToePlateDim")
        TopOfMidRailDim = GetAttributeValue(oAttrColl, "TopOfMidRailDim")
        MidRailSpacing = GetAttributeValue(oAttrColl, "MidRailSpacing")
    
        WithPostAtTurn = GetAttributeValue(oAttrColl, "WithPostAtTurn")
        ' treatment type code list value ; 5 - circular 10 - rectangular; 2 - none
        BeginExtensionLength = GetAttributeValue(oAttrColl, "BeginExtensionLength")
        

            
        EndExtensionLength = GetAttributeValue(oAttrColl, "EndExtensionLength")
        

        
        IsAssembly = GetAttributeValue(oAttrColl, "IsAssembly")
        
        TopRailSection = GetAttributeValue(oAttrColl, "TopRail_SPSSectionName")
        TopRailSecStandard = GetAttributeValue(oAttrColl, "TopRail_SPSSectionRefStandard")
        TopRailSectionCP = GetAttributeValue(oAttrColl, "TopRailSectionCP")
        ToprailSectionAngle = GetAttributeValue(oAttrColl, "TopRailSectionAngle")
        
        MidRailSection = GetAttributeValue(oAttrColl, "MidRail_SPSSectionName")
        MidRailSecStandard = GetAttributeValue(oAttrColl, "MidRail_SPSSectionRefStandard")
        MidRailSectionCP = GetAttributeValue(oAttrColl, "MidRailSectionCP")
        MidRailSectionAngle = GetAttributeValue(oAttrColl, "MidRailSectionAngle")
        
        ToePlateSection = GetAttributeValue(oAttrColl, "ToePlate_SPSSectionName")
        ToePlateSecStandard = GetAttributeValue(oAttrColl, "ToePlate_SPSSectionRefStandard")
        ToePlateSectionCP = GetAttributeValue(oAttrColl, "ToePlateSectionCP")
        ToePlateSectionAngle = GetAttributeValue(oAttrColl, "ToePlateSectionAngle")
        
        PostSection = GetAttributeValue(oAttrColl, "Post_SPSSectionName")
        PostSecStandard = GetAttributeValue(oAttrColl, "Post_SPSSectionRefStandard")
        PostSectionCP = GetAttributeValue(oAttrColl, "PostSectionCP")
        PostSectionAngle = GetAttributeValue(oAttrColl, "PostSectionAngle")

        Set oAttrColl = GetAttributeCollection(oHandrail, "IJHandrailProperties")

        HorizontalOffset = GetAttributeValue(oAttrColl, "HorizontalOffset")
        BeginTreatmentType = GetAttributeValue(oAttrColl, "BeginTreatmentType")
        EndTreatmentType = GetAttributeValue(oAttrColl, "EndTreatmentType")
        PostConnectionType = GetAttributeValue(oAttrColl, "PostConnectionType")
         If BeginTreatmentType = 2 Then BeginExtensionLength = 0#
         If (EndTreatmentType = 2) Then EndExtensionLength = 0#
    End If

    ' code list; 5 - right, 10 - left
    If HorizontalOffset = 10 Then
        Offset = HorizOffset        ' -HorizOffset
    ElseIf HorizontalOffset = 5 Then
        Offset = HorizOffset
    Else
        Offset = 0#  ' center chosen
    End If
    
    CheckForUndefinedValueAndRaiseError oHandrail, CLng(HandrailOrientation), "HandrailOrientation", 101
    CheckForUndefinedValueAndRaiseError oHandrail, CLng(TopRailSectionCP), "CrossSectionCardinalPoints", 102
    CheckForUndefinedValueAndRaiseError oHandrail, CLng(MidRailSectionCP), "CrossSectionCardinalPoints", 103
    CheckForUndefinedValueAndRaiseError oHandrail, CLng(ToePlateSectionCP), "CrossSectionCardinalPoints", 104
    CheckForUndefinedValueAndRaiseError oHandrail, CLng(PostSectionCP), "CrossSectionCardinalPoints", 105
    CheckForUndefinedValueAndRaiseError oHandrail, CLng(HorizontalOffset), "HandrailOffset", 106
    CheckForUndefinedValueAndRaiseError oHandrail, CLng(BeginTreatmentType), "HandrailEndTreatment", 107
    CheckForUndefinedValueAndRaiseError oHandrail, CLng(EndTreatmentType), "HandrailEndTreatment", 108
    CheckForUndefinedValueAndRaiseError oHandrail, CLng(PostConnectionType), "HandrailConnectionType", 109
    
    bNeedToRaiseWarning = False
    ' This call verifies the existance of XSection in Catalog & access to Sym file. Any problem will be logged
    ' to middle tier error log with error as S_FALSE.
    If (DoesXSectionAndAccessToSymFileExists(definitionProgId, "Structure:HR", TopRailSection, TopRailSecStandard) = False) Then
        bNeedToRaiseWarning = True
    End If
    If (DoesXSectionAndAccessToSymFileExists(definitionProgId, "Structure:HR", MidRailSection, MidRailSecStandard) = False) Then
        bNeedToRaiseWarning = True
    End If
    If (DoesXSectionAndAccessToSymFileExists(definitionProgId, "Structure:HR", ToePlateSection, ToePlateSecStandard) = False) Then
        bNeedToRaiseWarning = True
    End If
    If (DoesXSectionAndAccessToSymFileExists(definitionProgId, "Structure:HR", PostSection, PostSecStandard) = False) Then
        bNeedToRaiseWarning = True
    End If
    
    Dim Sketch As IJDSketch3d
    
    If Not pInputs Is Nothing Then
        Dim pDefPlayerEx As IJDDefinitionPlayerEx
        Dim Arguments As IJDArguments
        Dim pArg1 As IJDArgument
        
        Dim oProxy As IJDProxy
        Dim oSrcPart As iJDObject
        Dim oPartDef As Object
        
        Dim oRep As IJDRepresentation
        Dim oOutputs As IJDOutputs
        
        Set pDefPlayerEx = pRepDG.definition.IJDDefinitionPlayerEx
        Set Arguments = pDefPlayerEx.SymbolInputs
        Set pArg1 = Arguments.Item(2)
        Set Sketch = pArg1.Entity
        
        Set pArg1 = Arguments.Item(1)
        Set oPartDef = pArg1.Entity
        Set oProxy = oPartDef
        Set oSrcPart = oProxy.Source
        Set m_oCatResMgr = oSrcPart.ResourceManager
        
        Set pDefPlayerEx = Nothing
        Set Arguments = Nothing
        Set pArg1 = Nothing
    
        Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
        Set oOutputs = oRep
        oOutputs.RemoveAllOutput
        Set oRep = Nothing
        Set oOutputs = Nothing
    
    Else
        Set Sketch = oHandrail.SketchPath
        Set m_oCatResMgr = GetCatalogResourceManager
    End If

    ' now we need to create handrails from sketch
    
    Dim pSegments As IJElements
    Dim pComplex  As ComplexString3d
    Dim p3dCurve As ComplexString3d

    Set pComplex = Sketch.GetComplexString
    Set p3dCurve = GetOffsetCurve(pComplex, HorizontalOffset, Offset)
    
    p3dCurve.GetCurves pSegments
    Dim pCurve      As IJCurve
    Set pCurve = pSegments(1)
        
    Dim nX As Double, nY As Double, nZ As Double
    Dim curvescope As Geom3dCurveScopeConstants
    pCurve.Normal curvescope, nX, nY, nZ

    Dim sP As Double, eP As Double
        
    pCurve.ParamRange sP, eP
        
    Dim sX As Double, sY As Double, sZ As Double
    Dim tanX As Double, tanY As Double, tanZ As Double
    Dim tan2x As Double, tan2y As Double, tan2z As Double
        
    pCurve.Evaluate sP, sX, sY, sZ, tanX, tanY, tanZ, tan2x, tan2y, tan2z
    
    Dim SegVec As DVector
    Set SegVec = New DVector
    
    SegVec.Set tanX, tanY, tanZ
    
    Dim vecAxis As DVector
    Dim vecPDir As DVector

    Set vecAxis = New DVector
    Set vecPDir = New DVector

    vecAxis.Set 0, 0, 1
    Set vecPDir = vecAxis.Cross(SegVec)

    If vecPDir.Length < 0.00001 Then
        vecAxis.Set 1, 0, 0
        Set vecPDir = vecAxis.Cross(SegVec)

        If vecPDir.Length < 0.00001 Then
            vecAxis.Set 0, 1, 0
            Set vecPDir = vecAxis.Cross(SegVec)
        End If

    End If
        vecPDir.Length = 1

    Dim vecOrient As DVector
    Set vecOrient = New DVector

    Set vecOrient = SegVec.Cross(vecPDir)
'    vecOrient.length = 1

    Dim startpt     As IJDPosition
    Set startpt = New DPosition
    startpt.Set sX, sY, sZ
    
    Dim DirPt       As IJDPosition
    Set DirPt = New DPosition
    Set DirPt = startpt.Offset(vecOrient)
    
    Set pCurve = Nothing
    Set vecOrient = Nothing
    Set vecAxis = Nothing
    Set vecPDir = Nothing
    Set SegVec = Nothing
   
    'TR#51383- use actual orientation only if there is no arc in handrail path
    If ComplexstringcontainsArc(p3dCurve) Then
        HandrailOrientation = 0
    End If
   
    ' get the offset distance for top, mid rail for circular end treatments
    
    ' get top rail section depth and width
    Dim htDelta As Double, htdelta1 As Double
    
    GetCrossSecData TopRailSection, TopRailSecStandard, TopRailSectionWidth, TopRailSectionDepth
    Select Case TopRailSectionCP
        Case 1, 2, 3, 11
            htDelta = TopRailSectionDepth
            htdelta1 = TopRailSectionDepth / 2
        Case 4, 12, 5, 10, 15, 6, 13
            htDelta = TopRailSectionDepth / 2
            htdelta1 = TopRailSectionDepth / 2
        Case 7, 8, 14, 9
            htDelta = 0#
            htdelta1 = TopRailSectionDepth / 2
        Case Else
            htDelta = 0#
            htdelta1 = TopRailSectionDepth / 2
    End Select
    
    'TR45745- Ignore htDelta calculated above and get htDelta again so that top of toprail is maintained in all combinations of CP/Orientatio
    GethDeltaForTopRail TopRailSecStandard, TopRailSection, TopRailSectionCP, ToprailSectionAngle, htDelta, TopRailSectionWidth, TopRailSectionDepth
    
    Dim dCirTOffset As Double
    If BeginTreatmentType = 5 Or EndTreatmentType = 5 Then
    
        Dim h1 As Double
        Dim i As Integer
        h1 = TopOfMidRailDim
        For i = 1 To noofmidrails - 1
            If MidRailSpacing > 0# And h1 > TopOfToePlateDim Then
                h1 = TopOfMidRailDim - MidRailSpacing
            End If
        Next i
        
        dCirTOffset = (Height - htDelta - h1) / 6
        'If the above assuption is not a good one, use the following formular to get offset
        'This value should be consistent with the value of ht1 in CreateCirEndTreatment
        If TopRailSectionDepth > (Height - htDelta - h1) Then
            dCirTOffset = (Height - htDelta - h1) / 2
        ElseIf dCirTOffset < TopRailSectionDepth / 2 Then
            dCirTOffset = TopRailSectionDepth / 4 + (Height - htDelta - h1) / 4
        End If
        TreatmentRadius = dCirTOffset
    End If
    
    'before creating differnt handrail parts check whether given skinning option works for given path & other input
    'if not then change it appropriately. Check this for max dimension of top/mid rail & toe plate
    SkinOption = 7 'Valid values 0 to 7 user need to set appropriate skinning option as per requirement
    Dim maxCSdim As Double
    maxCSdim = GetCrossSectionDiamension(TopRailSection, TopRailSecStandard, MidRailSection, MidRailSecStandard, ToePlateSection, ToePlateSecStandard)
    Call ValidateInputForSkinning(p3dCurve, SkinOption, maxCSdim)
    
    ' add handrail path, physical path (offset curve) and not logical (sketched path )as one of the outputs
    If Not pOC Is Nothing Then
        If RepType <> CenterLineRep Then
            AddPathAsOutputOfHandrail pOC, p3dCurve
        End If
    End If
    ' create top rail
    If RepType = CenterLineRep Then
        CreateTopRail_Ex pOutputParentObject, p3dCurve, DirPt, Height - htDelta, HandrailOrientation, "", "", TopRailSectionCP, ToprailSectionAngle, dCirTOffset, BeginTreatmentType, EndTreatmentType, False
    Else
        CreateTopRail_Ex pOutputParentObject, p3dCurve, DirPt, Height - htDelta, HandrailOrientation, TopRailSection, TopRailSecStandard, TopRailSectionCP, ToprailSectionAngle, dCirTOffset, BeginTreatmentType, EndTreatmentType, False
    End If
    
    ' create mid rail(s)
    Dim midrailHt As Double
    midrailHt = TopOfMidRailDim
    Dim actualmidrails As Integer
    Dim PtStart As DPosition, PtEnd As DPosition
    Set PtStart = New DPosition
    Set PtEnd = New DPosition

    If RepType = CenterLineRep Then
        CreateMidRails_Ex pOutputParentObject, p3dCurve, DirPt, HandrailOrientation, noofmidrails, actualmidrails, MidRailSpacing, TopOfMidRailDim, TopOfToePlateDim, "", "", midrailHt, PtStart, PtEnd, MidRailSectionCP, MidRailSectionAngle, h1, dCirTOffset, BeginTreatmentType, EndTreatmentType, False, TopRailSectionDepth / 2
    Else
        CreateMidRails_Ex pOutputParentObject, p3dCurve, DirPt, HandrailOrientation, noofmidrails, actualmidrails, MidRailSpacing, TopOfMidRailDim, TopOfToePlateDim, MidRailSection, MidRailSecStandard, midrailHt, PtStart, PtEnd, MidRailSectionCP, MidRailSectionAngle, h1, dCirTOffset, BeginTreatmentType, EndTreatmentType, False, TopRailSectionDepth / 2
    End If
    
    ' create toe plate
    If WithToePlate Then
        If RepType = CenterLineRep Then
            CreateToePlate_Ex pOutputParentObject, p3dCurve, DirPt, TopOfToePlateDim, HandrailOrientation, "", "", ToePlateSectionCP, ToePlateSectionAngle, False
        Else
            CreateToePlate_Ex pOutputParentObject, p3dCurve, DirPt, TopOfToePlateDim, HandrailOrientation, ToePlateSection, ToePlateSecStandard, ToePlateSectionCP, ToePlateSectionAngle, False
        End If
    End If
    
    ' create post(s)
    'TR#51383- If orienation is "Perpendicular.." then get appropriate path segment so that posts
    ' are placed at proper location. Here we are first getting path at top rail postion with perpendicular orientation
    ' Then project it to original handrail path to get appropriate start-end points of each segment
    If HandrailOrientation <> 0 Then
        Dim pTopSegments As IJElements
        Dim pTopRailcurve As ComplexString3d
        CreateProjectedHRCurve pTopRailcurve, p3dCurve, HandrailOrientation, Height - htdelta1
        pTopRailcurve.GetCurves pTopSegments
        GetModifiedHRSegments pSegments, pTopSegments, Height - htdelta1
        Set pTopSegments = Nothing
        Set pTopRailcurve = Nothing
    End If
            
    If RepType = CenterLineRep Then
        TopRailSecStandard = ""
        TopRailSection = ""
        CreatePosts pOutputParentObject, pSegments, Height - htdelta1, midrailHt, HandrailOrientation, _
                PtStart, PtEnd, SegmentMaxSpacing, SlopedSegmentMaxSpacing, BeginExtensionLength, EndExtensionLength, _
                "", "", WithPostAtTurn, PostSectionCP, PostSectionAngle, _
                BeginTreatmentType, EndTreatmentType
    Else
        CreatePosts pOutputParentObject, pSegments, Height - htdelta1, midrailHt, HandrailOrientation, _
                PtStart, PtEnd, SegmentMaxSpacing, SlopedSegmentMaxSpacing, BeginExtensionLength, EndExtensionLength, _
                PostSection, PostSecStandard, WithPostAtTurn, PostSectionCP, PostSectionAngle, _
                BeginTreatmentType, EndTreatmentType
    End If
    Set startpt = Nothing
    Set DirPt = Nothing
    Set PtStart = Nothing
    Set PtEnd = Nothing
    Set Sketch = Nothing
    Set pComplex = Nothing
    Set p3dCurve = Nothing
    Set pSegments = Nothing
    
    Set pRepDG = Nothing
    Set pOC = Nothing
    Set pInputs = Nothing
    
    If (bNeedToRaiseWarning = True) Then
        Err.Raise SYMBOL_E_WARNING
    End If
    
    Exit Sub
    
ErrorHandler:
    If (Err.Number = SYMBOL_E_WARNING) Then
        Err.Raise SYMBOL_E_WARNING
    ElseIf (Err.Description = "Undefined Value") Then
        Err.Raise E_FAIL
    Else
        m_oErrors.Add Err.Number, METHOD, Err.Description
        Err.Raise E_FAIL
    End If
End Sub


Private Sub CreatePosts(oOutputParentObject As Object, _
                         ByVal pSegments As IJElements, _
                         PostHeight As Double, _
                         midrailHt As Double, _
                         Orientation As Integer, _
                         ByRef sPt As DPosition, _
                         ByRef ePt As DPosition, _
                         SegmentMaxSpacing As Double, _
                         SlopedSegmentMaxSpacing As Double, _
                         BeginExtensionLength As Double, _
                         EndExtensionLength As Double, _
                         PostSection As String, _
                         PostSecStandard As String, _
                         bPostAtEveryTurn As Boolean, _
                         CSectionCP As Integer, _
                         CSectionAngle As Double, _
                         BeginTreatType As Integer, _
                         EndTreatType As Integer, Optional CalcVolume As Boolean = False, Optional ByRef volume As Double, Optional ByRef VCogX As Double, Optional ByRef VCogY As Double, Optional ByRef VCogZ As Double)


Const METHOD = "CreatePosts"
On Error GoTo ErrorHandler

    Dim outCount                As Long
    Dim bNoCSSymbol             As Boolean
    
    Dim PostCSArea              As Double
    Dim PostPosition As DPosition
    
    Dim Posts As Collection
    Dim post As HandrailPost
    'dummy variable
    Dim Mirror As Boolean
    Mirror = False
    Dim DeltaAngle  As Double
    DeltaAngle = PI
    Dim dPadOffset As Double
    dPadOffset = 0#
    Dim pOC             As IJDOutputCollection
    Dim oSectionOccPost As Object       ' cross section occurrence used by symbol for post
    Dim oSectionDefPost As Object       ' cross-section definition used by members for post
    Dim oSectionObjPost As Object       ' common cross section obj sent to BuildHandrailOutput for post
    
    Dim oSectionOccRail As Object       ' cross section occurrence used by symbol for topRail
    Dim oSectionDefRail As Object       ' cross-section definition used by members for topRail
    Dim oSectionObjRail As Object       ' common cross section obj sent to BuildHandrailOutput for topRail
    
    Dim MaxClearenceAtPostTurn As Double
    Dim MinClearenceAtPostTurn As Double
    Dim PostDistBeforeTurnPercent As Double
    Dim PostDistAfterTurnPercent As Double
    'Presently set this value to be 2Feet max and 9inches min. later on
    'modify this to get it from the IJDInputs. -- Satish N Kota
    MaxClearenceAtPostTurn = 0.6096
    MinClearenceAtPostTurn = 0.2286
        
    outCount = 1
    
    Dim ToprailCSArea           As Double

    If CalcVolume Then
        PostCSArea = GetCSAttribData(PostSection, PostSecStandard, "ISTRUCTCrossSectionDimensions", "Area")
        ToprailCSArea = GetCSAttribData(TopRailSection, TopRailSecStandard, "ISTRUCTCrossSectionDimensions", "Area")
        volume = 0#
        VCogX = 0#
        VCogY = 0#
        VCogZ = 0#
        Set PostPosition = New DPosition

    ElseIf TypeOf oOutputParentObject Is IJDOutputCollection Then
        Set pOC = oOutputParentObject
        If Trim(PostSecStandard) <> "" And Trim(PostSection) <> "" Then
            Set oSectionOccPost = GetCSProfile(Nothing, PostSecStandard, PostSection, m_oCatResMgr)
        End If
        If oSectionOccPost Is Nothing Then
            bNoCSSymbol = True
        Else
            Set oSectionObjPost = oSectionOccPost
        End If
        If Trim(TopRailSecStandard) <> "" And Trim(TopRailSection) <> "" Then
            Set oSectionOccRail = GetCSProfile(Nothing, TopRailSecStandard, TopRailSection, m_oCatResMgr)
        End If
        Set oSectionObjRail = oSectionOccRail

    Else
        Dim xService As SP3DStructGenericTools.CrossSectionServices
        Set xService = New SP3DStructGenericTools.CrossSectionServices

        xService.GetStructureCrossSectionDefinition GetCatalogResourceManager, PostSecStandard, "", PostSection, oSectionDefPost
        Set oSectionObjPost = oSectionDefPost
        If oSectionObjPost Is Nothing Then
            bNoCSSymbol = True
        End If
        xService.GetStructureCrossSectionDefinition GetCatalogResourceManager, TopRailSecStandard, "", TopRailSection, oSectionDefRail
        Set oSectionObjRail = oSectionDefRail

    End If
    
    
    CalcPostPositions Posts, pSegments, BeginExtensionLength, BeginTreatType, _
                    EndExtensionLength, EndTreatType, _
                    SegmentMaxSpacing, SlopedSegmentMaxSpacing, _
                    bPostAtEveryTurn, MinClearenceAtPostTurn, MaxClearenceAtPostTurn, ToprailSectionAngle, CSectionAngle, Orientation, Mirror, DeltaAngle

    DeltaAngle = 0

    For Each post In Posts
    'Post type: 2 for post, 5 for circular treatment, 10 for regular treatment

        Dim tempMirror As Boolean
        If post.Index = Posts.Count Then
            tempMirror = Not Mirror
        Else
            tempMirror = Mirror
        End If

        If post.PostType <> 2 Then
            If CalcVolume Then
                PostPosition.Set post.BasePos.x + midrailHt * post.DirectionVec.x, post.BasePos.y + midrailHt * post.DirectionVec.y, post.BasePos.z + midrailHt * post.DirectionVec.z
                AddTreatmentVolumeInfo ToprailCSArea, PostPosition, post.DirectionVec, PostHeight - midrailHt, post.PostType, volume, VCogX, VCogY, VCogZ
            Else
                CreateTreatment oOutputParentObject, post, pSegments, bNoCSSymbol, Nothing, PostHeight, midrailHt, oSectionObjRail, TopRailSectionCP, Mirror, SkinOption, TopRailSectionDepth, TreatmentRadius
            End If
        Else
            If CalcVolume Then
                PostPosition.Set post.BasePos.x + dPadOffset * post.DirectionVec.x, post.BasePos.y + dPadOffset * post.DirectionVec.y, post.BasePos.z + dPadOffset * post.DirectionVec.z
                AddPostVolumeInfo PostCSArea, PostPosition, post.DirectionVec, PostHeight - dPadOffset, volume, VCogX, VCogY, VCogZ
            Else
                CreatePost oOutputParentObject, post, oSectionObjPost, CSectionCP, PostHeight, dPadOffset, tempMirror, SkinOption, bNoCSSymbol, outCount, DeltaAngle
             End If
       End If
    Next
    
    Set PostPosition = Nothing
    If Not Posts Is Nothing Then
        While Posts.Count > 0
            Posts.Remove Posts.Count
        Wend
    End If
    Set Posts = Nothing

    Dim iIJDObject As iJDObject
    If Not oSectionOccPost Is Nothing Then

        Set iIJDObject = oSectionOccPost
        iIJDObject.Remove
        Set iIJDObject = Nothing
    End If
    If Not oSectionOccRail Is Nothing Then

        Set iIJDObject = oSectionOccRail
        iIJDObject.Remove
        Set iIJDObject = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub
Private Sub InitNewOutput(pOC As IJDOutputCollection, name As String)
Const METHOD = "InitNewOutput"
    On Error GoTo ErrHandler
    
    Dim oRep As IMSSymbolEntities.IJDRepresentation
    Dim oOutputs As IMSSymbolEntities.IJDOutputs
    Dim oOutput As IMSSymbolEntities.IJDOutput
    
    Set oOutput = New DOutput
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
    Set oOutputs = oRep
    oOutput.name = name
    oOutput.Description = name
    oOutputs.SetOutput oOutput
    oOutput.Reset
    Exit Sub
    
ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)

End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    
End Sub

Private Function ISPSHandrailConvertServices_ConnectComponents(ByVal pHandrail As Object) As SPSHandrails.StructHandrailConvertHelperStatus
Const METHOD = "ISPSHandrailConvertServices_ConnectComponents"
On Error GoTo ErrorHandler
    
    ISPSHandrailConvertServices_ConnectComponents = ConnectComponents(pHandrail)
    Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    ISPSHandrailConvertServices_ConnectComponents = StructHandrailConvertHelperStatus_Unexpected
    Err.Raise E_FAIL

End Function

Private Sub ISPSHandrailSymbolServices_EvaluateGeometry(ByVal HandrailBO As SPSHandrails.ISPSHandrail, ByVal PartOccChanged_IN As Boolean, ByVal Sketch3dCurve_IN As Object, ByVal AttributeCol_IN As SPSHandrails.IJDAttributesCol, ByVal AttributesChanged_IN As Boolean, ByVal OccAttributeCol_INOUT As SPSHandrails.IJDAttributesCol, ByVal OccAttributesChanged_IN As Boolean)

End Sub

Private Sub ISPSHandrailSymbolServices_EvaluateLocalWCG(ByVal HandrailBO As SPSHandrails.ISPSHandrail, ByVal OccAttributeCol As SPSHandrails.IJDInfosCol, weight As Double, CogX As Double, CogY As Double, CogZ As Double)
Const METHOD = "ISPSHandrailSymbolServices_EvaluateLocalWCG"
On Error GoTo ErrorHandler

    Call CalcWCG(HandrailBO, OccAttributeCol, weight, CogX, CogY, CogZ)
    Exit Sub
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL

End Sub
'********************************************************************************
'  CREATION OF OPERATIONAL REPRESENTATION OF HANDRAIL TYPEA OBJECT
'********************************************************************************
Public Sub OperationRepresentation(pIRepSCM As IJDRepresentationStdCustomMethod)
    Const METHOD = "Operation"

    Dim pRepDG As IJDRepresentationDuringGame
    Dim pOutputCol As IJDOutputCollection
    Dim pInputs As IJDInputs
    Dim pDefPlayerEx As IJDDefinitionPlayerEx
    Dim Arguments As IJDArguments
    Dim pArg1 As IJDArgument
    Dim Sketch As IJDSketch3d
    Dim oProxy As IJDProxy
    Dim oSrcPart As iJDObject
    Dim oPartDef As Object
    Dim pSegments As IJElements
    Dim pComplex As ComplexString3d
    Dim oGeomServices As IGeometryServices
    Dim oRep        As IJDRepresentation
    Dim oOutputs    As IJDOutputs
    Dim OutputSurfaces As IJElements
    Dim OutputSurfacesPlane As RuledSurface3d ' This is for ruled surface along the handrail
    Dim oCurveObj As iJDObject
    Dim opt1 As IJDPosition
    Dim opt2 As IJDPosition
    Dim ovec1 As IJDVector
    Dim ovec2 As IJDVector
    Dim ovec3 As IJDVector
    Dim oOffsetComplex As ComplexString3d
    Dim oLine As IJLine
    Dim oCurve As IJCurve
    Dim StNorm() As Double
    Dim EdNorm() As Double
    Dim Index       As Integer
    Dim oName       As String
    Dim Height As Double

    Dim TopRailSection As String, TopRailStandard As String
    Dim TopRailCP As Integer, TopRailAngle As Double
    Dim TopRailDepth As Double, TopRailWidth As Double

    Dim dCenX As Double
    Dim dCenY As Double

    Dim x1 As Double, y1 As Double, z1 As Double, x2 As Double, y2 As Double, z2 As Double
    Dim cx As Double, cy As Double, cz As Double
    Dim startParam As Double
    Dim endParam As Double
    Dim startx As Double, starty As Double, startz As Double
    Dim Endx As Double, Endy As Double, Endz As Double
    Dim nX As Double, nY As Double, nZ As Double
    Dim vtanx As Double, vtany As Double, vtanz As Double, vTan2X As Double, vTan2Y As Double, vTan2Z As Double
    Dim etanx As Double, etany As Double, etanz As Double, etan2x As Double, etan2y As Double, etan2z As Double
    Dim xp As Double, yp As Double, zp As Double, xN As Double, yN As Double, zN As Double
    Dim HorizontalOffset As Integer
    Dim dHorizOffset As Double

    Dim p3dCurve As ComplexString3d
    Dim pLine As IJLine
    Dim pCurve As IJCurve
    Dim bCurve As Boolean
    Dim HandrailOrientation As Integer 'Required to get toprail curve
    Dim pComplexTopRailCurve As ComplexString3d 'Required to get toprail curve

    On Error GoTo ErrorHandler

    Set pRepDG = pIRepSCM

    'Get The Output Collection
    Set pOutputCol = pRepDG.outputCollection

    'Get The Inputs of the Symbol for the following objects
    Set pInputs = pRepDG.definition.IJDInputs
    'Height
    Height = pInputs.GetInputByIndex(3).IJDInputDuringGame.Result.UomValue
    'TopRail Details
    TopRailSection = pInputs.GetInputByIndex(21).IJDInputDuringGame.Result.String
    TopRailStandard = pInputs.GetInputByIndex(22).IJDInputDuringGame.Result.String
    TopRailCP = pInputs.GetInputByIndex(23).IJDInputDuringGame.Result.UomValue
    TopRailAngle = pInputs.GetInputByIndex(24).IJDInputDuringGame.Result.UomValue
    If TopRailCP = -1 Then TopRailCP = 5

    'Other Handrail Details
    HorizontalOffset = pInputs.GetInputByIndex(6).IJDInputDuringGame.Result.UomValue
    dHorizOffset = pInputs.GetInputByIndex(7).IJDInputDuringGame.Result.UomValue
    If HorizontalOffset = 15 Then
        dHorizOffset = 0#
    End If
    
    'get orientation of HR which is required to get toprail curve
    HandrailOrientation = pInputs.GetInputByIndex(ORIENTATION_INDEX).IJDInputDuringGame.Result.UomValue
    If HandrailOrientation = -1 Then HandrailOrientation = 0
    
    'Get Definition from the Game.
    Set pDefPlayerEx = pRepDG.definition.IJDDefinitionPlayerEx
    Set Arguments = pDefPlayerEx.SymbolInputs
    Set pArg1 = Arguments.Item(2)

    'Get the Sketch of the Handrail path.
    Set Sketch = pArg1.Entity
    'Get the Catalog Resource Manager
    Set pArg1 = Arguments.Item(1)
    Set oPartDef = pArg1.Entity
    Set oProxy = oPartDef
    Set oSrcPart = oProxy.Source
    Set m_oCatResMgr = oSrcPart.ResourceManager

    Set pDefPlayerEx = Nothing
    Set Arguments = Nothing
    Set pArg1 = Nothing

    'Remove the Presently set outputs and re load them based on present aspect.
    Set oRep = pOutputCol.definition.IJDRepresentations.GetRepresentationByName("OperationRepresentation")
    Set oOutputs = oRep

    oOutputs.RemoveAllOutput

    Set oRep = Nothing
    Set oOutputs = Nothing
    

    'Get the Maximum possible distance covered by the CrossSection based on the selected
    'Cardinal Point and the Angle rotated for
    'TopRail
    GetCenOfHRObjectWRTCP TopRailStandard, TopRailSection, TopRailCP, TopRailAngle, dCenX, dCenY, TopRailWidth, TopRailDepth
    'Get the Complex string from the Sketch3D
    Set p3dCurve = Sketch.GetComplexString

    'TR#51383- use actual orientation only if there is no arc in handrail path
    If ComplexstringcontainsArc(p3dCurve) Then
        HandrailOrientation = 0
    End If
    
    
    Set pComplex = GetOffsetCurve(p3dCurve, HorizontalOffset, dHorizOffset)
    


    Dim hDelta As Double
    Dim wDelta As Double
    'Push the TopRail to some depth based on it's cardinal point
    ' I dont know why it has to be done, but the Physical Aspect has
    'done that so I too am doing it...most probably to honour the
    'actual height of the Handrail.
    Select Case TopRailCP
        Case 1, 2, 3, 11
            hDelta = TopRailDepth
        Case 4, 12, 5, 10, 15, 6, 13
            hDelta = TopRailDepth / 2
        Case 7, 8, 14, 9
            hDelta = 0#
        Case Else
            hDelta = 0#
    End Select
    
    'TR45745- Ignore earlier calculation of hDelta and get hDelta again so that top of toprail ia maintained in all cases
    GethDeltaForTopRail TopRailStandard, TopRailSection, TopRailCP, TopRailAngle, hDelta, TopRailWidth, TopRailDepth
    
    'Set the new height
    Height = Height - hDelta

    'Get the Curves defined by Sketch3D
    If Not pSegments Is Nothing Then
        pSegments.Clear
        Set pSegments = Nothing
    End If
'    pComplex.GetCurves pSegments 'Manish 8/14/03 instead get top curve and use its first point as circle radius

    'Manish 8/14/03- get complex string curve at toprail location. This can be used to create
    'ruled surface between this curve and HR path. Also can be used for creating cylinder along toprail
    Set pComplexTopRailCurve = New ComplexString3d
    CreateProjectedHRCurve pComplexTopRailCurve, pComplex, HandrailOrientation, Height
    pComplexTopRailCurve.GetCurves pSegments

    'Get the Actual Center point at which the Cylinder has to be laid...
    'Consider the center point of action changes with Cardinal Points
    Dim Scope As Geom3dCurveScopeConstants
    Set oCurve = pSegments(1)
    oCurve.ParamRange startParam, endParam
    Dim Normx As Double, Normy As Double, NormZ As Double
    oCurve.Normal Scope, Normx, Normy, NormZ
    Dim xS As Double, yS As Double, zS As Double
    Dim tX As Double, tY As Double, tZ As Double
    Dim t2x As Double, t2y As Double, t2z As Double

    oCurve.Evaluate startParam, xS, yS, zS, tX, tY, tZ, t2x, t2y, t2z
    Set ovec1 = Nothing
    Set ovec2 = Nothing
    Set ovec1 = New AutoMath.DVector
    ovec1.Set tX, tY, tZ
    Set ovec2 = Nothing
    Set ovec2 = New AutoMath.DVector
    Set opt1 = Nothing
    Set opt1 = New AutoMath.DPosition
'    opt1.Set xS, yS, zS+ height 'now we are using top curve so don't add height.
    opt1.Set xS, yS, zS
    
    ''based on cp of toprail, get the appropriate location of center of circle
     If dCenX = 0# And dCenY = 0# Then
     Else
        Dim vZ As IJDVector
        Set vZ = New DVector
        If dCenY <> 0# Then
            vZ.Set 0#, 0#, dCenY
        Else
            vZ.Set 0#, 0#, 1#
        End If
        Set ovec3 = vZ.Cross(ovec1)
        Set ovec2 = ovec1.Cross(ovec3)
            ovec2.Length = Abs(dCenY)
        Set ovec3 = Nothing
        Set vZ = Nothing

        If dCenY <> 0 Then
            Set opt2 = Nothing
            Set opt2 = opt1.Offset(ovec2)
        Else
            ovec2.Set 0#, 0#, 1#
            Set opt2 = opt1
        End If
        Set opt1 = Nothing
        Set ovec3 = Nothing
        If dCenX > 0# Then
            If dCenY >= 0# Then
                Set ovec3 = ovec1.Cross(ovec2)
            ElseIf dCenY < 0# Then
                Set ovec3 = ovec2.Cross(ovec1)
            End If
                ovec3.Length = Abs(dCenX)
            Set opt1 = opt2.Offset(ovec3)
        ElseIf dCenX < 0# Then
            If dCenY >= 0# Then
                Set ovec3 = ovec2.Cross(ovec1)
            ElseIf dCenY < 0# Then
                Set ovec3 = ovec1.Cross(ovec2)
            End If
                ovec3.Length = Abs(dCenX)
            Set opt1 = opt2.Offset(ovec3)
        Else
            Set opt1 = opt2
        End If
    End If

    'Manish 8/14/03- Create ruled surface between HR path and curve along top rail
    Set oGeomServices = m_GeomFactory
    Set OutputSurfacesPlane = m_GeomFactory.RuledSurfaces3d.CreateByCurves(Nothing, pComplexTopRailCurve, pComplex, 0)
    Index = 1
    InitOperationOutput pOutputCol, "Handrail Surface" & CStr(Index)
    pOutputCol.AddOutput "Handrail Surface" & CStr(Index), OutputSurfacesPlane
    
    'move toprail curve to center of circle location and sweep circle along this
    If dCenY <> 0# Then
        CreateProjectedTopRailCurve pComplexTopRailCurve, pComplexTopRailCurve, ovec2
    ElseIf dCenX <> 0# Then
        CreateProjectedTopRailCurve pComplexTopRailCurve, pComplexTopRailCurve, ovec3
    End If
    
    'Create a circle and sweep it through the whole sketch path.
    Dim oCircle As IJCircle
    Dim radius As Double
    Set oCircle = New Circle3d
    radius = (Sqr((TopRailWidth * TopRailWidth) + (TopRailDepth * TopRailDepth))) / 2 + 3 * 0.0254
    oCircle.DefineByCenterNormalRadius opt1.x, opt1.y, opt1.z, tX, tY, tZ, radius
    Set oOffsetComplex = Nothing
    Set oOffsetComplex = New ComplexString3d
    oOffsetComplex.AddCurve oCircle, True
    Set oCircle = Nothing

    'Cover the Top Rail by a cylinder with
    Set oGeomServices = m_GeomFactory
    'Manish 8/14/03- sweeping circle along curve at top rail location. Earlier it was sweeping it
    'along original handrail path. pComplexTopRailCurve used instead of pComplex in call below
    Set OutputSurfaces = oGeomServices.CreateBySingleSweep(Nothing, pComplexTopRailCurve, oOffsetComplex, 0, 0, StNorm, EdNorm, 1)
    Dim oldIndex As Integer
    oldIndex = Index - 1
    'With the outputs generated, add each outputs to the OutputCollection with a name.
    For Index = 1 To OutputSurfaces.Count
        InitOperationOutput pOutputCol, "Top Rail Cylinder" & CStr(Index + oldIndex)
        pOutputCol.AddOutput "Top Rail Cylinder" & CStr(Index + oldIndex), OutputSurfaces.Item(Index)
    Next
    Set pRepDG = Nothing
    Set pOutputCol = Nothing
    Set pInputs = Nothing
    Set pDefPlayerEx = Nothing
    Set Arguments = Nothing
    Set pArg1 = Nothing
    Set Sketch = Nothing
    Set oProxy = Nothing
    Set oSrcPart = Nothing
    Set oPartDef = Nothing
    Set pSegments = Nothing
    Set pComplex = Nothing
    Set oGeomServices = Nothing
    Set oRep = Nothing
    Set oOutputs = Nothing
    Set OutputSurfaces = Nothing
    Set oCurveObj = Nothing
    Set opt1 = Nothing
    Set opt2 = Nothing
    Set ovec1 = Nothing
    Set ovec2 = Nothing
    Set ovec3 = Nothing
    Set oOffsetComplex = Nothing
    Set oLine = Nothing
    Set oCurve = Nothing
    Set p3dCurve = Nothing
    Set pLine = Nothing
    Set pCurve = Nothing
    Set pComplexTopRailCurve = Nothing
    Set OutputSurfacesPlane = Nothing

    Exit Sub
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
Const METHOD = "UserAttributeMgmt_Validate"
On Error GoTo ErrorHandler

' first of all check if the symbol definition has CMCheck methods defined - TBD
    UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_HANDRAIL_ERROR, "ERROR")

    Dim dInputs As IJDInputs
    Dim CurrentInput As IJDInput
    Dim oAttribute As IJDAttribute
    Dim PC As DParameterContent
    Dim bvalid As Boolean
    Dim oSymbolOcc As IJDSymbol
    Set oSymbolOcc = pIJDAttrs
    Dim oSymbolDef As IJDSymbolDefinition
    Dim ErrMessage As String
    Set oSymbolDef = oSymbolOcc.IJDSymbolDefinition(2)
    Set dInputs = oSymbolDef.IJDInputs
    Set PC = New DParameterContent
    
    '#TR-CP·52075 - mkaveti
    Dim strAttribValid As Boolean
    strAttribValid = True

    If (sAttributeName = "Height") Or (sAttributeName = "NoOfMidRails") Or (sAttributeName = "TopOfMidRailDim") Or (sAttributeName = "MidRailSpacing") Then
        ValidateHandrailAttributes pIJDAttrs, sInterfaceName, strAttribValid
        
        If strAttribValid = False Then
            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_HANDRAIL_INVALID_ATTRIB_CHECK, "Invalid Attribute - Check User inputs")
            Exit Function
        End If
    End If
    
    Set oAttribute = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item(sAttributeName)

    If oAttribute.Value <> "" Then
        If oAttribute.AttributeInfo.Type = igString Then    ' check for string type here
        Else
            PC.UomValue = oAttribute.Value
            Set CurrentInput = Nothing
            bvalid = True
            On Error Resume Next
            Set CurrentInput = dInputs.GetInputByName(oAttribute.AttributeInfo.name)
            On Error GoTo ErrorHandler
            If Not CurrentInput Is Nothing Then
                CurrentInput.IJDInputDuringGame.definition = oSymbolDef
                CurrentInput.IJDInputStdCustomMethod.InvokeCMCheck PC, bvalid, ErrMessage
                CurrentInput.IJDInputDuringGame.definition = Nothing
                Set oSymbolOcc = Nothing
                Set oSymbolDef = Nothing
                If bvalid = False Then
'                    UserAttributeMgmt_Validate = "Symbol CMCheck Failed"
                    ' tr 59002
                    UserAttributeMgmt_Validate = ErrMessage

                    Exit Function
                Else
                End If
            End If
        End If
    End If
    
    UserAttributeMgmt_Validate = ""
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_HANDRAIL_ERROR, "ERROR")

    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String
    Dim i As Integer
    Dim pColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim NonStateRO As Long
    If bOnPreLoad = False Then
        ErrStr = UserAttributeMgmt_Validate(pIJDAttrs, pAttrToChange.InterfaceName, pAttrToChange.attrName, varNewAttrValue)
        If Len(ErrStr) > 0 Then
'            IJUserAttributeMgmt_OnAttributeChange = "ERROR::Bad Value"
            IJUserAttributeMgmt_OnAttributeChange = ErrStr
            Exit Function
        End If
    End If

    ' attribute management should be done here
    
    IJUserAttributeMgmt_OnAttributeChange = ""
   
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_HANDRAIL_ERROR, "ERROR")
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String
    
    Set pAttrColl = CollAllDisplayedValues

    For i = 1 To pAttrColl.Count
        Set pAttrDescr = pAttrColl.Item(i)
            ErrStr = IJUserAttributeMgmt_OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrDescr, pAttrDescr.attrValue)
            If Len(ErrStr) > 0 Then
                bOnPreLoad = False
                Exit Function
            End If
    Next
    
    bOnPreLoad = False

    IJUserAttributeMgmt_OnPreLoad = ""

End Function


Private Sub ISPSHandrailSymbolServices_GetHandrailLength(ByVal HandrailBO As SPSHandrails.ISPSHandrail, length As Double)
Const METHOD = "ISPSHandrailSymbolServices_GetHandrailLength"
        On Error GoTo ErrorHandler
        length = GetHandrailSketchPathLength(HandrailBO)
        Exit Sub
ErrorHandler:
        m_oErrors.Add Err.Number, METHOD, Err.Description
        Err.Raise E_FAIL
End Sub

Private Sub ISPSHandrailSymbolServices_GetTopRailRadius(ByVal pIJDPart As SPSHandrails.IJDPart, radius As Double)
Const METHOD = "ISPSHandrailSymbolServices_GetTopRailRadius"
On Error GoTo ErrorHandler

    Dim TopRailSection As String, TopRailStandard As String
    Dim TopRailCP As Integer, TopRailAngle As Double
    Dim TopRailDepth As Double, TopRailWidth As Double
    Dim dCenX As Double, dCenY As Double

    GetTopRailInfo pIJDPart, TopRailSection, TopRailStandard, TopRailCP, TopRailAngle
    
    'Get the Maximum possible distance covered by the CrossSection based on the selected
    'Cardinal Point and the Angle rotated for TopRail
    GetCenOfHRObjectWRTCP TopRailStandard, TopRailSection, TopRailCP, TopRailAngle, dCenX, dCenY, TopRailWidth, TopRailDepth
    
    radius = GetTopRailRadius(TopRailWidth, TopRailDepth)
   
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub CalcWCG(oHandrail As ISPSHandrail, _
                      ByVal PartOccInfoCol As IJDInfosCol, _
                      ByRef weight As Double, _
                      ByRef CogX As Double, ByRef CogY As Double, ByRef CogZ As Double)
    Const METHOD = "CalcWCG"
    On Error GoTo ErrHandler
        
    Dim iMaterial   As IJDMaterial
    Dim material As Variant, grade As Variant
    Dim density     As Variant
    Dim totalVolume As Double
    Dim volume1 As Double
    Dim VCOGX1 As Double
    Dim VCOGY1 As Double
    Dim VCOGZ1 As Double
    Dim volume2 As Double
    Dim VCOGY2 As Double
    Dim VCOGX2 As Double
    Dim VCOGZ2 As Double
    
    Dim OccAttrs As IJDAttributes
    Set OccAttrs = oHandrail
    'Get Material and Grade Values from Occurance Attributes...to calculate the Density[TR 41415-Murali-June 13, 2003]
    material = GetAttribute(OccAttrs, "Primary_SPSMaterial", PartOccInfoCol)
    grade = GetAttribute(OccAttrs, "Primary_SPSGrade", PartOccInfoCol)
    TopRailSection = GetAttribute(OccAttrs, "TopRail_SPSSectionName", PartOccInfoCol)
    TopRailSecStandard = GetAttribute(OccAttrs, "TopRail_SPSSectionRefStandard", PartOccInfoCol)
    TopRailSectionCP = GetAttribute(OccAttrs, "TopRailSectionCP", PartOccInfoCol)
    ToprailSectionAngle = GetAttribute(OccAttrs, "TopRailSectionAngle", PartOccInfoCol)
    GetCrossSecData TopRailSection, TopRailSecStandard, TopRailSectionWidth, TopRailSectionDepth

    'Get the MaterialObject using the material and grade
    Set iMaterial = GetMaterialObject(material, grade)
    Set OccAttrs = Nothing
       
    CalcRailsVolumeCG oHandrail, PartOccInfoCol, TypeA, volume1, VCOGX1, VCOGY1, VCOGZ1
    CalcPostsVolumeCG oHandrail, PartOccInfoCol, volume2, VCOGX2, VCOGY2, VCOGZ2
    
    'Set Density by default to 7800kg/m^3 (Steel Density)[TR 41415-Murali-June 13, 2003]
    density = 7849.04705324047 'TR#52838- use exact density 7849.04705324047
        
    If Not iMaterial Is Nothing Then
        density = iMaterial.density
    End If
    
    totalVolume = volume1 + volume2
    CogX = (VCOGX1 + VCOGX2) / totalVolume
    CogY = (VCOGY1 + VCOGY2) / totalVolume
    CogZ = (VCOGZ1 + VCOGZ2) / totalVolume
    weight = totalVolume * density
    
Exit Sub
ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub


Private Sub CalcPostsVolumeCG(oHandrail As ISPSHandrail, _
                      ByVal PartOccInfoCol As IJDInfosCol, _
                      ByRef volume As Double, _
                      ByRef VCogX As Double, ByRef VCogY As Double, ByRef VCogZ As Double)
    Const METHOD = "CalcPostsVolumeCG"
    On Error GoTo ErrHandler
    
    'Prepare arguments for CreatePosts
    
    'HRpathCurve
    Dim OccAttrs As IJDAttributes
    Dim HRpathCurve As ComplexString3d
    Dim pSegments As IJElements
    Dim HRpath As Sketch3d
    Set OccAttrs = oHandrail
    Set HRpath = oHandrail.SketchPath
    Dim iOffsetType As Integer
    iOffsetType = GetAttribute(OccAttrs, "HorizontalOffset", PartOccInfoCol)
    Dim dOffsetValue As Double
    dOffsetValue = GetAttribute(OccAttrs, "HorizontalOffsetDim", PartOccInfoCol)
    Set HRpathCurve = GetOffsetCurve(HRpath.GetComplexString, _
                               iOffsetType, _
                               dOffsetValue)
  
    
    
    HRpathCurve.GetCurves pSegments
    Set HRpath = Nothing
    
    'Height and htdelta1
    Dim Height As Double
    Dim htdelta1 As Double

    Height = GetAttribute(OccAttrs, "Height", PartOccInfoCol) 'TR#55627

    htdelta1 = TopRailSectionDepth / 2
    
    'TopOfMidRailDim and MidRailSectionDepth
    Dim TopOfMidRailDim As Double
    Dim MidRailSection As String
    Dim MidRailSecStandard As String
    Dim MidRailSectionWidth As Double
    Dim MidRailSectionDepth As Double
    MidRailSection = GetAttribute(OccAttrs, "MidRail_SPSSectionName", PartOccInfoCol)
    MidRailSecStandard = GetAttribute(OccAttrs, "MidRail_SPSSectionRefStandard", PartOccInfoCol)
    TopOfMidRailDim = GetAttribute(OccAttrs, "TopOfMidRailDim", PartOccInfoCol)
    GetCrossSecData MidRailSection, MidRailSecStandard, MidRailSectionWidth, MidRailSectionDepth 'get midrail details also
    
    'Orientation
    Dim Orientation As Integer
    Orientation = GetAttribute(OccAttrs, "HandrailOrientation", PartOccInfoCol) 'TR#55627
    
    'SegmentMaxSpacing
    Dim SegmentMaxSpacing As Double
    Dim SlopedSegmentMaxSpacing As Double
    SegmentMaxSpacing = GetAttribute(OccAttrs, "SegmentMaxSpacing", PartOccInfoCol)
    SlopedSegmentMaxSpacing = GetAttribute(OccAttrs, "SlopedSegmentMaxSpacing", PartOccInfoCol)
    
    'ExtensionLength
    Dim BeginExtensionLength  As Double
    Dim EndExtensionLength As Double
    BeginExtensionLength = GetAttribute(OccAttrs, "BeginExtensionLength", PartOccInfoCol)
    EndExtensionLength = GetAttribute(OccAttrs, "EndExtensionLength", PartOccInfoCol)
    
    'Post Section
    Dim PostSection As String
    Dim PostStandard As String
    Dim PostSectionCP       As Integer
    Dim PostSectionAngle    As Double
    PostSection = GetAttribute(OccAttrs, "Post_SPSSectionName", PartOccInfoCol)
    PostStandard = GetAttribute(OccAttrs, "Post_SPSSectionRefStandard", PartOccInfoCol)
    PostSectionCP = GetAttribute(OccAttrs, "PostSectionCP", PartOccInfoCol)
    PostSectionAngle = GetAttribute(OccAttrs, "PostSectionAngle", PartOccInfoCol)
    
    'With Post at Turn
    Dim WithPostAtTurn As Boolean
    WithPostAtTurn = GetAttribute(OccAttrs, "WithPostAtTurn", PartOccInfoCol)
    
    'Treatment type
    Dim SPSHRBeginTreatment As Integer
    Dim SPSHREndTreatment As Integer
    SPSHRBeginTreatment = GetAttribute(OccAttrs, "BeginTreatmentType", PartOccInfoCol)
    SPSHREndTreatment = GetAttribute(OccAttrs, "EndTreatmentType", PartOccInfoCol)
    
    Set OccAttrs = Nothing
    
    'Call createPosts to get volume and COG for posts
    CreatePosts Nothing, pSegments, Height - htdelta1, TopOfMidRailDim - MidRailSectionDepth / 2, _
                Orientation, Nothing, Nothing, _
                SegmentMaxSpacing, SlopedSegmentMaxSpacing, BeginExtensionLength, EndExtensionLength, _
                PostSection, PostStandard, WithPostAtTurn, PostSectionCP, PostSectionAngle, _
                SPSHRBeginTreatment, SPSHREndTreatment, True, volume, VCogX, VCogY, VCogZ
        
Exit Sub
ErrHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub ISPSHandrailConvertServices_ConvertHandrailToComponents(ByVal oHandrail As Object)
Const METHOD = "ISPSHandrailConvertServices_ConvertHandrailToComponents"
On Error GoTo ErrorHandler

    ' Convert calls same code as usual symbol, but with the Handrail object as argument instead of SymbolCustomMethod.
    ' That code switches accordingly to read occurrence attributes and create membersystems as output objects.

    Call PhysicalRepresentation(oHandrail, SimpleRep)

    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

