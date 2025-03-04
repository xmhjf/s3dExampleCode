VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeASideMntToMem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'******************************************************************
'  Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'  File               TypeASideMount.cls
'  ProgID             SPSHandrailMacros.TypeASideMntToMem
'  Author             NB
'  Creation Date      [Feb, 20 2006]
'  Description
'
'  Change History:
'
'   05-oct-06   SS      TR#107161 - commented out msgbox as we should not be raising any msgboxes from
'                       symbol could. If such error condition happens during sync, ifc it could
'                       stop the process.
'
'   11-Oct-08   SS      CR#38818 - add handrail path (physical and not sketched path which can be called logical)
'                       as its output,locatable but not displayable
'  05-Aug-08   WR       TR-CP-131048 - Added new method to ISPSHandrailSymbolServices to get the top rail radius
'
'   02-Oct-08   SS      CR#148021 - Add Centerline representation to support Drawings for CR112776.
'                       DLL version is bumped and will require a synchronize.
'
'   02-Oct-08   SS      CR#141607 - Removed Detailed Physical from handrail representations as it is same
'                       as SimplePhysical and redundant.

'   14-Oct-08   SS      TR#151855 - Added code to offset the path only when offset is right/center after creating
'                       new sketch based on connected members and attribution.
'   12-Aug-09   GG      DM-169707  - Fixed the weight issues and some COG issues
'   04-Sept-09  GG      DM-171219 Changed the way to calculate weight and COG for handrails. Now the WCG calculation
'                       is independent to the Physical evaluation. The main changes in this module are:
'                       1. Moved CalcWCG from ComputeStuff module with modifications
'                       2. Wrap CreatePosts with CreatePosts_Ex to handle the new option for calculate WCG. The wrapper may be removed in next version update.
'                       3. Created subroutine CalcPostsVolumeCG
'                       4. replaced 22/7 by PI
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
'   10-Mar-10   PW      Added this stub function ISPSHandrailConvertServices_ConnectComponents due to build failure in 0900180004
'                       Function will return unexpected value, developer needs to determine correct workflow.
'   14-May-10   GG      DM 182314. Fixed the following issues:
'                       1.  The SideMountToMember Handrail does not honor horizontal offset consistently;
'                       2.  The SideMountToMember Handrail does not honor different CPs of the supporting Member;
'                       3.  The horizontal offset direction is dependent on the direction of the curvature for curved supporting member.
'   02-June-10  GG      DM 183787. Fixed the splitting issues for SideMntToMem handrail
'   17-Aug-10   GG      DM 186125. Circular end treatment of handrail is not correct if the path is sloped
'   29-Dec-10   GG      TR 191191  The handrail computing fails for some path patterns. Also removed some obsolete subs
'
'*******************************************************************

Const definitionProgId As String = "SPSHandrailMacros.TypeASideMntToMem"

Implements IJDUserSymbolServices
Implements ISPSHandrailSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck
Implements ISPSHandrailConvertServices

Private Const MODULE = "SPSHandrailMacros.TypeASideMntToMem"
Private m_oLocalizer As IJLocalizer
Private TopRailSection As String
Private TopRailSecStandard As String
Private TopRailSectionDepth As Double
Private TopRailSectionWidth As Double
Private TopRailSectionCP As Integer
Private ToprailSectionAngle As Double
Private TreatmentRadius As Double

Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259
Dim bOnPreLoad As Boolean
Dim msgAlreadyDisplayed As Boolean

Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"
    
Private Enum SideMountInputIndex
    SIDEMOUNT_CONN_TYPE_INDEX = 41
    SIDEMOUNT_CONN_VERT_OFFSET_INDEX
End Enum

Private Enum SideMntByMemInputIndex
    SIDEMNTBYMEM_FLOOR_THICKNESS = 43
    SIDEMNTBYMEM_COLUMN_CLEARANCE
    SIDEMNTBYMEM_END_CLEARANCE
    SIDEMNTBYMEM_MIRROR
    SIDEMNTBYMEM_START_COLUMN_OFFSET
    SIDEMNTBYMEM_END_COLUMN_OFFSET
    SIDEMNTBYMEM_OFFSET_REFERENCE
End Enum

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
    ' TypeA1 Handrail Definition
    ' Inputs:
    '          1. "Part"  ( Catalog part )
    '          2. Sketch3dObj
    '          3. Height
    '          4. WithToePlate
    '          5. NoOfMidRails
    '          6. SPSHRPathHorizontalOffset
    '          7. SPSHRPathHorizontalOffsetDim
    '          8. HandrailOrientation
    '          9. SegmentMaxSpacing
    '          10. SlopedSegmentMaxSpacing
    '          11. TopOfToePlateDim
    '          12. TopOfMidRailDim
    '          13. MidRailSpacing
    '          14. WithPostAtTurn
    '          15. SPSHRBeginTreatment
    '          16. SPSHRBeginExtensionLength
    '          17. SPSHREndTreatment
    '          18. SPSHREndExtensionLength
    '          19. IsSystem
    '          20. IsAssembly
    '          21. TopRailSection
    '          22. TopRailSectionReferenceStandard
    '          23. TopRailSectionCP
    '          24. TopRailSectionAngle
    '          25. MidRailSection
    '          26. MidRailSectionReferenceStandard
    '          27. MidRailSectionCP
    '          28. MidRailSectionAngle
    '          29. ToePlateSection
    '          30. ToePlateSectionReferenceStandard
    '          31. ToePlateSectionCP
    '          32. ToePlateSectionAngle
    '          33. PostSection
    '          34. PostSectionReferenceStandard
    '          35. PostSectionCP
    '          36. PostSectionAngle
    '          37. Primary_SPSMaterial
    '          38. Primary_SPSGrade
    '          39. MaximumClearanceAtPostTurn
    '          40. MinimumClearanceAtPostTurn
    '          41. SideMountVerticalOffset
    '          42. FloorThickness
    '          43. ColumnClearance
    '          44. EndClearance
    '          45. Mirror
    '          46. StartColumnOffset
    '          47. EndColumnOffset
    '          48. Offset Reference
    '
    ' Representations:
    '           Physical
    '
    ' Outputs collection : igCollection_Variable
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
    Dim NegativeNum As Long
    
     
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
    NegativeNum = pSymbolDefinition.IJDUserMethods.GetMethodCookie("NegativeNum", ChecklibCookie)

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
    iInputs.SetInput oInput, SKETCHOBJ_INDEX
    oInput.Reset
    
    oInput.name = "Height"
    oInput.Description = "Height of the Handrail from path to top of Top Rail"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 1.0668
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, HEIGHT_INDEX
    oInput.Reset
    oPC.Reset

    oInput.name = "WithToePlate"
    oInput.Description = "With or Without Toe/Kick plate"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 1#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, WITHTOEPLATE_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "NoOfMidRails"
    oInput.Description = "With One/Many/None mid rails"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 1#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, MIDRAILSNO_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SPSHRPathHorizontalOffset"
    oInput.Description = "Offset position that the handrail will be from the path chosen"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 5#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, HORIZONTAL_OFFSET_TYPE_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SPSHRPathHorizontalOffsetDim"
    oInput.Description = "Distance that the handrail will be offset from the path chosen"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.0254
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, HORIZONTAL_OFFSET_DIM_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "HandrailOrientation"
    oInput.Description = "Vertical / Normal"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, ORIENTATION_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SegmentMaxSpacing"
    oInput.Description = "Maximum spacing between posts on straight/arc line"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 2.134
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, MAXSPACING_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SlopedSegmentMaxSpacing"
    oInput.Description = "Maximum spacing between posts  on slope"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 1.829
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, SLOPE_MAXSPACING_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "TopOfToePlateDim"
    oInput.Description = "Dimension to the top of the toe plate from Path "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.102
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, TOPOFTOEPLATE_DIM_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "TopOfMidRailDim"
    oInput.Description = "Dimension to the top of the Mid rail from top of toe plate "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.5842
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, TOPOFMIDRAIL_DIM_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "MidRailSpacing"
    oInput.Description = "Spacing between mid rails"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.33
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, MIDRAIL_SPACING_DIM_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "WithPostAtTurn"
    oInput.Description = "With or Without Post at the turn"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, POST_AT_TURN_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SPSHRBeginTreatment"
    oInput.Description = "Begin treatment type 0 for rectangular, 1 for rounded"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 10#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, BEGIN_TREAT_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SPSHRBeginExtensionLength"
    oInput.Description = "Extent of the end of the rail (at beginning) "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.4572
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, BEGIN_EXT_DIM_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SPSHREndTreatment"
    oInput.Description = "End treatment type 0 for rectangular, 1 for rounded"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 10#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, END_TREAT_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SPSHREndExtensionLength"
    oInput.Description = "Extent of the end of the rail (at ending) "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.4572
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, END_EXT_DIM_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "IsAssembly"
    oInput.Description = "If system placed as a symbol otherwise constructed with SPS members"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, ISASSEMBLY_INDEX
    oInput.Reset
    oPC.Reset

    oInput.name = "IsSystem"
    oInput.Description = "If system placed as a symbol otherwise constructed with SPS members"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, ISSYSTEM_INDEX
    oInput.Reset
    oPC.Reset
'
    oInput.name = "TopRail_SPSSectionName"
    oInput.Description = "TopRailSection from structural cross sections"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "L2-1/2x2-1/2x1/4"
    oInput.DefaultParameterValue = oPC
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
    iInputs.SetInput oInput, TOPRAIL_CSNAME_INDEX
    oInput.Reset
    oPC.Reset

    oInput.name = "TopRail_SPSSectionRefStandard"
    oInput.Description = "TopRail Reference Standard "
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "AISC-LRFD-3.1"
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, TOPRAIL_CSSTD_INDEX
    oInput.Reset
    oPC.Reset

    oInput.name = "TopRailSectionCP"
    oInput.Description = "TopRailSection Cardinal Point"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 1#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, TOPRAIL_CSCP_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "TopRailSectionAngle"
    oInput.Description = "TopRailSection Angle"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
    oPC.Type = igValue
    oPC.UomValue = 3.142857
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, TOPRAIL_CSANGLE_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "MidRail_SPSSectionName"
    oInput.Description = "MidRailSection from structural cross sections"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "RS0.25x2"
    oInput.DefaultParameterValue = oPC
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
    iInputs.SetInput oInput, MIDRAIL_CSNAME_INDEX
    oInput.Reset
    oPC.Reset

    oInput.name = "MidRail_SPSSectionRefStandard"
    oInput.Description = "MidRail Reference Standard"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "Misc"
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, MIDRAIL_CSSTD_INDEX
    oInput.Reset
    oPC.Reset

    oInput.name = "MidRailSectionCP"
    oInput.Description = "MidRailSection Cardinal Point"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 4#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, MIDRAIL_CSCP_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "MidRailSectionAngle"
    oInput.Description = "MidRailSection Angle"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
    oPC.Type = igValue
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, MIDRAIL_CSANGLE_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "ToePlate_SPSSectionName"
    oInput.Description = "MidRailSection from structural cross sections"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "RS0.25x4"
    oInput.DefaultParameterValue = oPC
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
    iInputs.SetInput oInput, TOEPLATE_CSNAME_INDEX
    oInput.Reset
    oPC.Reset

    oInput.name = "ToePlate_SPSSectionRefStandard"
    oInput.Description = "ToePlate Reference Standard"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "Misc"
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, TOEPLATE_CSSTD_INDEX
    oInput.Reset
    oPC.Reset

    oInput.name = "ToePlateSectionCP"
    oInput.Description = "ToePlateSection Cardinal Point"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 7#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, TOEPLATE_CSCP_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "ToePlateSectionAngle"
    oInput.Description = "ToePlateSection Angle"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
    oPC.Type = igValue
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, TOEPLATE_CSANGLE_INDEX
    oInput.Reset
    oPC.Reset
    
    
    oInput.name = "Post_SPSSectionName"
    oInput.Description = "PostSection from structural cross sections"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "L2-1/2x2-1/2x1/4"
    oInput.DefaultParameterValue = oPC
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
    iInputs.SetInput oInput, POST_CSNAME_INDEX
    oInput.Reset
    oPC.Reset

    oInput.name = "Post_SPSSectionRefStandard"
    oInput.Description = "Post Section Reference Standard"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "AISC-LRFD-3.1"
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, POST_CSSTD_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "PostSectionCP"
    oInput.Description = "PostSection Cardinal Point"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 1#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, POST_CSCP_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "PostSectionAngle"
    oInput.Description = "PostSection Angle"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
    oPC.Type = igValue
    oPC.UomValue = 0#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, POST_CSANGLE_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "Primary_SPSMaterial"
    oInput.Description = "Block foundation SPS Materail"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "Steel - Carbon"
    oInput.DefaultParameterValue = oPC
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, CheckMaterial
    iInputs.SetInput oInput, MATERIAL_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "Primary_SPSGrade"
    oInput.Description = "Block foundation SPS Grade"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igString
    oPC.String = "A"
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, GRADE_INDEX
    oInput.Reset
    oPC.Reset
    
     oInput.name = "SPSHRMinClearAtPostTurn"
    oInput.Description = "Minimum clearance between the Post and the turn"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.2286
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, MIN_CLEARANCE_AT_POST_TURN_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SPSHRMaxClearAtPostTurn"
    oInput.Description = "Maximum clearance between the Post and the turn"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 0.61
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, MAX_CLEARANCE_AT_POST_TURN_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SPSHRSideMountConnType"
    oInput.Description = "Type of Connection either with Pad or with Bracket"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.UomValue = 1#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, SIDEMOUNT_CONN_TYPE_INDEX
    oInput.Reset
    oPC.Reset
    
    oInput.name = "SPSHRSideMountConnVertOffset"
    oInput.Description = "Vertical Offset Dimension of Side Mount"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = -0.254
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, SIDEMOUNT_CONN_VERT_OFFSET_INDEX
    oInput.Reset
    oPC.Reset
      
    
    oInput.name = "FloorThickness"
    oInput.Description = "Floor Thickness"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 0.1524
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, SIDEMNTBYMEM_FLOOR_THICKNESS
    oInput.Reset
    oPC.Reset
    
    oInput.name = "ColumnClearance"
    oInput.Description = "Column Clearance"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, NegativeNum
    oPC.Type = igValue
    oPC.UomValue = 0.0762
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, SIDEMNTBYMEM_COLUMN_CLEARANCE
    oInput.Reset
    oPC.Reset
    
    oInput.name = "EndClearance"
    oInput.Description = "End Clearance"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, NegativeNum
    oPC.Type = igValue
    oPC.UomValue = 0
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, SIDEMNTBYMEM_END_CLEARANCE
    oInput.Reset
    oPC.Reset
    
    oInput.name = "Mirror"
    oInput.Description = "Mirror"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 0
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, SIDEMNTBYMEM_MIRROR
    oInput.Reset
    oPC.Reset
    
    oInput.name = "StartColumnOffset"
    oInput.Description = "Start Column Offset"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 0
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, SIDEMNTBYMEM_START_COLUMN_OFFSET
    oInput.Reset
    oPC.Reset

    oInput.name = "EndColumnOffset"
    oInput.Description = "End Column Offset"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oPC.Type = igValue
    oPC.UomValue = 0
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, SIDEMNTBYMEM_END_COLUMN_OFFSET
    oInput.Reset
    oPC.Reset
    
    oInput.name = "OffsetReference"
    oInput.Description = "Offset Reference"
    oInput.Properties = igINPUT_IS_A_PARAMETER
    oInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
    oPC.Type = igValue
    oPC.UomValue = 2#
    oInput.DefaultParameterValue = oPC
    iInputs.SetInput oInput, SIDEMNTBYMEM_OFFSET_REFERENCE
    oInput.Reset
    oPC.Reset
    
    'Define the representation "Symbolic"
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

    'Set the representation to definition
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
Const METHOD = "Physical"
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

    Dim Height As Double
    Dim SegmentMaxSpacing As Double
    Dim SlopedSegmentMaxSpacing As Double
    Dim Offset As Double
    Dim TopOfToePlateDim As Double
    Dim TopOfMidRailDim As Double
    Dim MidRailSpacing As Double
    Dim HorizOffset As Double
    Dim WithToePlate As Boolean
    Dim WithPostAtTurn As Boolean
    Dim noofmidrails As Integer
    Dim HorizontalOffset As Integer
    Dim HandrailOrientation As Integer
    Dim PostConnectionType As Integer
    Dim PadOffset As Double
    Dim BeginTreatmentType As Integer, EndTreatmentType As Integer
    Dim BeginExtensionLength As Double, EndExtensionLength As Double
        
    Dim MidRailSection As String, MidRailSecStandard As String
    Dim MidRailSectionCP As Integer, MidRailSectionAngle As Double

    Dim ToePlateSection As String, ToePlateSecStandard As String
    Dim ToePlateSectionCP As Integer, ToePlateSectionAngle As Double

    Dim PostSection As String, PostSecStandard As String
    Dim PostSectionCP As Integer, PostSectionAngle As Double
    
    Dim MaxClearenceAtPostTurn As Double
    Dim MinClearenceAtPostTurn As Double
    
    Dim PostSecWidth As Double, PostSecDepth As Double
    Dim IsPadRequired As Boolean
    
    Dim FloorThickness As Double
    Dim ColumnClearance As Double
    Dim EndClearance As Double
    Dim Mirror As Boolean
    Dim StartColumnOffset As Double
    Dim EndColumnOffset As Double
    Dim OffsetReference As Long
        
    If Not pInputs Is Nothing Then
        Height = pInputs.GetInputByIndex(HEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
        WithToePlate = pInputs.GetInputByIndex(WITHTOEPLATE_INDEX).IJDInputDuringGame.Result.UomValue
        noofmidrails = pInputs.GetInputByIndex(MIDRAILSNO_INDEX).IJDInputDuringGame.Result.UomValue
        HorizontalOffset = pInputs.GetInputByIndex(HORIZONTAL_OFFSET_TYPE_INDEX).IJDInputDuringGame.Result.UomValue
        HorizOffset = pInputs.GetInputByIndex(HORIZONTAL_OFFSET_DIM_INDEX).IJDInputDuringGame.Result.UomValue
    
        ''CheckForUndefinedValueAndRaiseError pRepDG, CLng(HorizontalOffset), "HandrailOffset", 110
    
        ''         code list; 5 - right, 10 - left
        ''        If HorizontalOffset = 10 Then
        ''            Offset = HorizOffset
        ''        ElseIf HorizontalOffset = 5 Then
        ''            Offset = HorizOffset
        ''        Else
        ''            Offset = 0#  ' center chosen
        ''        End If
    
        HandrailOrientation = pInputs.GetInputByIndex(ORIENTATION_INDEX).IJDInputDuringGame.Result.UomValue
    
'        CheckForUndefinedValueAndRaiseError pRepDG, CLng(HandrailOrientation), "HandrailOrientation", 101
    
        SegmentMaxSpacing = pInputs.GetInputByIndex(MAXSPACING_INDEX).IJDInputDuringGame.Result.UomValue
    
        SlopedSegmentMaxSpacing = pInputs.GetInputByIndex(SLOPE_MAXSPACING_INDEX).IJDInputDuringGame.Result.UomValue
    
        TopOfToePlateDim = pInputs.GetInputByIndex(TOPOFTOEPLATE_DIM_INDEX).IJDInputDuringGame.Result.UomValue
        TopOfMidRailDim = pInputs.GetInputByIndex(TOPOFMIDRAIL_DIM_INDEX).IJDInputDuringGame.Result.UomValue
        MidRailSpacing = pInputs.GetInputByIndex(MIDRAIL_SPACING_DIM_INDEX).IJDInputDuringGame.Result.UomValue
    
        WithPostAtTurn = pInputs.GetInputByIndex(POST_AT_TURN_INDEX).IJDInputDuringGame.Result.UomValue
        
        BeginTreatmentType = pInputs.GetInputByIndex(BEGIN_TREAT_INDEX).IJDInputDuringGame.Result.UomValue
        ' treatment type code list value ; 5 - circular 10 - rectangular; 2 - none
        BeginExtensionLength = pInputs.GetInputByIndex(BEGIN_EXT_DIM_INDEX).IJDInputDuringGame.Result.UomValue
    
        '    CheckForUndefinedValueAndRaiseError pRepDG, CLng(BeginTreatmentType), "HandrailEndTreatment", 107
    
        If BeginTreatmentType = 2 Then BeginExtensionLength = 0#
        
        EndTreatmentType = pInputs.GetInputByIndex(END_TREAT_INDEX).IJDInputDuringGame.Result.UomValue
        ' treatment type code list value ; 5 - circular 10 - rectangular 2 -none
        EndExtensionLength = pInputs.GetInputByIndex(END_EXT_DIM_INDEX).IJDInputDuringGame.Result.UomValue
    
'        CheckForUndefinedValueAndRaiseError pRepDG, CLng(EndTreatmentType), "HandrailEndTreatment", 108
    
        If (EndTreatmentType = 2) Then EndExtensionLength = 0#
            
        TopRailSection = pInputs.GetInputByIndex(TOPRAIL_CSNAME_INDEX).IJDInputDuringGame.Result.String
        TopRailSecStandard = pInputs.GetInputByIndex(TOPRAIL_CSSTD_INDEX).IJDInputDuringGame.Result.String
        TopRailSectionCP = pInputs.GetInputByIndex(TOPRAIL_CSCP_INDEX).IJDInputDuringGame.Result.UomValue
        ToprailSectionAngle = pInputs.GetInputByIndex(TOPRAIL_CSANGLE_INDEX).IJDInputDuringGame.Result.UomValue
'        CheckForUndefinedValueAndRaiseError pRepDG, CLng(TopRailSectionCP), "CrossSectionCardinalPoints", 102
    
'            bNeedToRaiseWarning = False
'            ' This call verifies the existance of XSection in Catalog & access to Sym file. Any problem will be logged
'            ' to middle tier error log with error as S_FALSE.
'            If (DoesXSectionAndAccessToSymFileExists(definitionProgId, "Structure:HR", TopRailSection, TopRailSecStandard) = False) Then
'                bNeedToRaiseWarning = True
'            End If
    
        MidRailSection = pInputs.GetInputByIndex(MIDRAIL_CSNAME_INDEX).IJDInputDuringGame.Result.String
        MidRailSecStandard = pInputs.GetInputByIndex(MIDRAIL_CSSTD_INDEX).IJDInputDuringGame.Result.String
        MidRailSectionCP = pInputs.GetInputByIndex(MIDRAIL_CSCP_INDEX).IJDInputDuringGame.Result.UomValue
        MidRailSectionAngle = pInputs.GetInputByIndex(MIDRAIL_CSANGLE_INDEX).IJDInputDuringGame.Result.UomValue
    
        '    CheckForUndefinedValueAndRaiseError pRepDG, CLng(MidRailSectionCP), "CrossSectionCardinalPoints", 103
    
'        If (DoesXSectionAndAccessToSymFileExists(definitionProgId, "Structure:HR", MidRailSection, MidRailSecStandard) = False) Then
'            bNeedToRaiseWarning = True
'        End If
        
        ToePlateSection = pInputs.GetInputByIndex(TOEPLATE_CSNAME_INDEX).IJDInputDuringGame.Result.String
        ToePlateSecStandard = pInputs.GetInputByIndex(TOEPLATE_CSSTD_INDEX).IJDInputDuringGame.Result.String
        ToePlateSectionCP = pInputs.GetInputByIndex(TOEPLATE_CSCP_INDEX).IJDInputDuringGame.Result.UomValue
        ToePlateSectionAngle = pInputs.GetInputByIndex(TOEPLATE_CSANGLE_INDEX).IJDInputDuringGame.Result.UomValue
   
        '    CheckForUndefinedValueAndRaiseError pRepDG, CLng(ToePlateSectionCP), "CrossSectionCardinalPoints", 104
    
'        If (DoesXSectionAndAccessToSymFileExists(definitionProgId, "Structure:HR", ToePlateSection, ToePlateSecStandard) = False) Then
'            bNeedToRaiseWarning = True
'        End If
    
        PostSection = pInputs.GetInputByIndex(POST_CSNAME_INDEX).IJDInputDuringGame.Result.String
        PostSecStandard = pInputs.GetInputByIndex(POST_CSSTD_INDEX).IJDInputDuringGame.Result.String
        PostSectionCP = pInputs.GetInputByIndex(POST_CSCP_INDEX).IJDInputDuringGame.Result.UomValue
        PostSectionAngle = pInputs.GetInputByIndex(POST_CSANGLE_INDEX).IJDInputDuringGame.Result.UomValue
        
        PadOffset = pInputs.GetInputByIndex(SIDEMOUNT_CONN_VERT_OFFSET_INDEX).IJDInputDuringGame.Result.UomValue
        
        PadOffset = Abs(PadOffset) * -1
            
        MinClearenceAtPostTurn = pInputs.GetInputByIndex(MIN_CLEARANCE_AT_POST_TURN_INDEX).IJDInputDuringGame.Result.UomValue
        MaxClearenceAtPostTurn = pInputs.GetInputByIndex(MAX_CLEARANCE_AT_POST_TURN_INDEX).IJDInputDuringGame.Result.UomValue
        
        
        FloorThickness = pInputs.GetInputByIndex(SIDEMNTBYMEM_FLOOR_THICKNESS).IJDInputDuringGame.Result.UomValue
        ColumnClearance = pInputs.GetInputByIndex(SIDEMNTBYMEM_COLUMN_CLEARANCE).IJDInputDuringGame.Result.UomValue
        EndClearance = pInputs.GetInputByIndex(SIDEMNTBYMEM_END_CLEARANCE).IJDInputDuringGame.Result.UomValue
        Mirror = pInputs.GetInputByIndex(SIDEMNTBYMEM_MIRROR).IJDInputDuringGame.Result.UomValue
        StartColumnOffset = pInputs.GetInputByIndex(SIDEMNTBYMEM_START_COLUMN_OFFSET).IJDInputDuringGame.Result.UomValue
        EndColumnOffset = pInputs.GetInputByIndex(SIDEMNTBYMEM_END_COLUMN_OFFSET).IJDInputDuringGame.Result.UomValue
        OffsetReference = pInputs.GetInputByIndex(SIDEMNTBYMEM_OFFSET_REFERENCE).IJDInputDuringGame.Result.UomValue
        
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

        Set oAttrColl = GetAttributeCollection(oHandrail, "IJUAHRPathOffset")

        HorizontalOffset = GetAttributeValue(oAttrColl, "SPSHRPathHorizontalOffsetType")
        
        HorizOffset = GetAttributeValue(oAttrColl, "SPSHRPathHorizontalOffset")
        
        Set oAttrColl = GetAttributeCollection(oHandrail, "IJUATypeAHREndTreatments")
        BeginTreatmentType = GetAttributeValue(oAttrColl, "SPSHRBeginTreatment")
        EndTreatmentType = GetAttributeValue(oAttrColl, "SPSHREndTreatment")
        ' treatment type code list value ; 5 - circular 10 - rectangular; 2 - none
        BeginExtensionLength = GetAttributeValue(oAttrColl, "SPSHRBeginExtensionLength")
        
        If BeginTreatmentType = 2 Then BeginExtensionLength = 0#
            
        EndExtensionLength = GetAttributeValue(oAttrColl, "SPSHREndExtensionLength")
        
        If (EndTreatmentType = 2) Then EndExtensionLength = 0#
        
        Set oAttrColl = GetAttributeCollection(oHandrail, "IJUAHRByMemberProps")
        
        FloorThickness = GetAttributeValue(oAttrColl, "FloorThickness")
        ColumnClearance = GetAttributeValue(oAttrColl, "ColumnClearance")
        EndClearance = GetAttributeValue(oAttrColl, "EndClearance")
        Mirror = GetAttributeValue(oAttrColl, "Mirror")
        StartColumnOffset = GetAttributeValue(oAttrColl, "StartColumnOffset")
        EndColumnOffset = GetAttributeValue(oAttrColl, "EndColumnOffset")
        OffsetReference = GetAttributeValue(oAttrColl, "OffsetReference")
        
        Set oAttrColl = GetAttributeCollection(oHandrail, "IJUATypeAHR_SideMountConnProps")
        
        PadOffset = GetAttributeValue(oAttrColl, "SPSHRSideMountConnVertOffset")
        PadOffset = Abs(PadOffset) * -1
        
         Set oAttrColl = GetAttributeCollection(oHandrail, "IJUAHRTypeAClearAtPostTurn")
        MinClearenceAtPostTurn = GetAttributeValue(oAttrColl, "SPSHRMinClearAtPostTurn")
        MaxClearenceAtPostTurn = GetAttributeValue(oAttrColl, "SPSHRMaxClearAtPostTurn")
        
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
     CheckForUndefinedValueAndRaiseError oHandrail, CLng(HorizontalOffset), "HandrailOffset", 110
     CheckForUndefinedValueAndRaiseError oHandrail, CLng(BeginTreatmentType), "HandrailEndTreatment", 107
     CheckForUndefinedValueAndRaiseError oHandrail, CLng(EndTreatmentType), "HandrailEndTreatment", 108
 
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
 
'        CheckForUndefinedValueAndRaiseError pRepDG, CLng(PostSectionCP), "CrossSectionCardinalPoints", 105
 

  
     Dim Sketch As IJDSketch3d
 
     If Not pInputs Is Nothing Then
         Dim pDefPlayerEx As IJDDefinitionPlayerEx
         
         Dim Arguments As IJDArguments
         Dim pArg1 As IJDArgument
         
         Dim oProxy As IJDProxy
         Dim oSrcPart As iJDObject
         Dim oPartDef As Object
         
         Dim oRep        As IJDRepresentation
         Dim oOutputs    As IJDOutputs
         
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
 
         '    Dim oRep        As IJDRepresentation
         '    Dim oOutputs    As IJDOutputs
 
         Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
         Set oOutputs = oRep
         oOutputs.RemoveAllOutput
         Set oRep = Nothing
         Set oOutputs = Nothing
     Else
         Set Sketch = oHandrail.SketchPath
         Set m_oCatResMgr = GetCatalogResourceManager
     End If
     Dim tmpComplexString As ComplexString3d
     Dim pSegments As IJElements, pTmpSegments As IJElements
     Dim pComplex  As ComplexString3d, p3dCurve As ComplexString3d
     
     Dim DeltaHOffset As Double
     Dim TempHOffset As Integer
     DeltaHOffset = 0#
     TempHOffset = 0
     createNewSketch3D Sketch, OffsetReference, HorizontalOffset, FloorThickness, ColumnClearance, EndClearance, StartColumnOffset, EndColumnOffset, DeltaHOffset, TempHOffset, tmpComplexString
     
     Set pComplex = tmpComplexString
             
     Dim pCurve      As IJCurve
     Dim pLine       As IJLine
     Dim bCurve      As Boolean
     ' side mounted to member sketch curves must be offset twice
     ' once to bring the sketch in lign with the offset reference
     ' and once to honor the user specified offset
     
     Set tmpComplexString = GetOffsetCurve(pComplex, TempHOffset, DeltaHOffset)
     Set p3dCurve = GetOffsetCurve(tmpComplexString, HorizontalOffset, Offset)
     
     p3dCurve.GetCurves pSegments
             
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
         
     If vecPDir.Length < dtol Then
         vecAxis.Set 1, 0, 0
         Set vecPDir = vecAxis.Cross(SegVec)
         
         If vecPDir.Length < dtol Then
             vecAxis.Set 0, 1, 0
             Set vecPDir = vecAxis.Cross(SegVec)
         End If
         
     End If
        vecPDir.Length = 1
         
     Dim vecOrient As DVector
     Set vecOrient = New DVector
     
     Set vecOrient = SegVec.Cross(vecPDir)
         
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
     If ComplexstringcontainsArc(p3dCurve) Then
         HandrailOrientation = 0
     End If
    
     ' get the offset distance for top, mid rail for circular end treatments
             
 
     Dim htDelta As Double, htdelta1 As Double
     Dim MidRailSectionWidth As Double
     Dim MidRailSectionDepth As Double
     Dim PostSectionWidth As Double
     Dim PostSectionDepth As Double
             
     GetCrossSecData TopRailSection, TopRailSecStandard, TopRailSectionWidth, TopRailSectionDepth
     GetCrossSecData MidRailSection, MidRailSecStandard, MidRailSectionWidth, MidRailSectionDepth 'get midrail details also
     GetCrossSecData PostSection, PostSecStandard, PostSectionWidth, PostSectionDepth 'get post section details
     
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
        
     'Ignore htDelta calculated above and get htDelta again so that top of toprail ia maintained in all combinations of CP/Orientation
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
         ' PostSectionDepth & MidRailSectionDepth  considered for consistency with similar calculations in CreateCirEndTreatment()
         dCirTOffset = (Height - h1 - htDelta) / 6
         'If the above assuption is not a good one, use the following formular to get offset
         'This value should be consistent with the value of ht1 in CreateCirEndTreatment
         If htDelta > (Height - h1 - htDelta) Then
             dCirTOffset = (Height - h1 - htDelta) / 2
         ElseIf dCirTOffset < htDelta Then
             dCirTOffset = htDelta / 2 + (Height - htDelta - h1) / 4
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
         CreateTopRail_Ex pOutputParentObject, p3dCurve, DirPt, Height - htDelta, HandrailOrientation, "", "", TopRailSectionCP, ToprailSectionAngle, dCirTOffset, BeginTreatmentType, EndTreatmentType, Mirror
     Else
         CreateTopRail_Ex pOutputParentObject, p3dCurve, DirPt, Height - htDelta, HandrailOrientation, TopRailSection, TopRailSecStandard, TopRailSectionCP, ToprailSectionAngle, dCirTOffset, BeginTreatmentType, EndTreatmentType, Mirror
     End If
     
     ' create mid rail(s)
     Dim midrailHt As Double
     midrailHt = TopOfMidRailDim
     Dim actualmidrails As Integer
     Dim PtStart As DPosition, PtEnd As DPosition
     Set PtStart = New DPosition
     Set PtEnd = New DPosition
     If RepType = CenterLineRep Then
         CreateMidRails_Ex pOutputParentObject, p3dCurve, DirPt, HandrailOrientation, noofmidrails, actualmidrails, MidRailSpacing, TopOfMidRailDim, TopOfToePlateDim, "", "", midrailHt, PtStart, PtEnd, MidRailSectionCP, MidRailSectionAngle, h1, dCirTOffset, BeginTreatmentType, EndTreatmentType, Mirror, TopRailSectionDepth / 2
     Else
         CreateMidRails_Ex pOutputParentObject, p3dCurve, DirPt, HandrailOrientation, noofmidrails, actualmidrails, MidRailSpacing, TopOfMidRailDim, TopOfToePlateDim, MidRailSection, MidRailSecStandard, midrailHt, PtStart, PtEnd, MidRailSectionCP, MidRailSectionAngle, h1, dCirTOffset, BeginTreatmentType, EndTreatmentType, Mirror, TopRailSectionDepth / 2
     End If
     
     ' create toe plate
     If WithToePlate Then
         If RepType = CenterLineRep Then
             CreateToePlate_Ex pOutputParentObject, p3dCurve, DirPt, TopOfToePlateDim, HandrailOrientation, "", "", ToePlateSectionCP, ToePlateSectionAngle, Mirror
         Else
             CreateToePlate_Ex pOutputParentObject, p3dCurve, DirPt, TopOfToePlateDim, HandrailOrientation, ToePlateSection, ToePlateSecStandard, ToePlateSectionCP, ToePlateSectionAngle, Mirror
         End If
     End If
     
     'htdelta1 to be used only when top rail & post is pipe section. Side mount HR has angle section. so don't lower the post height
     htdelta1 = 0
     
     ' If orienation is "Perpendicular.." then get appropriate path segment so that posts
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
         CreatePosts_Ex pOutputParentObject, pSegments, Height - htdelta1, midrailHt - TopRailSectionDepth / 2, HandrailOrientation, PtStart, PtEnd, SegmentMaxSpacing, SlopedSegmentMaxSpacing, BeginExtensionLength, EndExtensionLength, "", "", WithPostAtTurn, PostSectionCP, PostSectionAngle, BeginTreatmentType, EndTreatmentType, PadOffset, MinClearenceAtPostTurn, MaxClearenceAtPostTurn, PostSectionDepth, MidRailSectionDepth, Mirror
     Else
         CreatePosts_Ex pOutputParentObject, pSegments, Height - htdelta1, midrailHt - TopRailSectionDepth / 2, HandrailOrientation, PtStart, PtEnd, SegmentMaxSpacing, SlopedSegmentMaxSpacing, BeginExtensionLength, EndExtensionLength, PostSection, PostSecStandard, WithPostAtTurn, PostSectionCP, PostSectionAngle, BeginTreatmentType, EndTreatmentType, PadOffset, MinClearenceAtPostTurn, MaxClearenceAtPostTurn, PostSectionDepth, MidRailSectionDepth, Mirror
     End If
     
    
cleanup:
    
    If (bNeedToRaiseWarning = True) Then
        Err.Raise SYMBOL_E_WARNING
    End If
    
Exit Sub
    
ErrorHandler:
    If (Err.Number = SYMBOL_E_WARNING) Then
        Err.Raise SYMBOL_E_WARNING
    Else
        m_oErrors.Add Err.Number, METHOD, Err.Description
        Err.Raise E_FAIL
    End If
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
    Dim OutputSurfacesPlane As RuledSurface3d ' This is for ruled surface along the handrail
    Dim OutputSurfaces As IJElements
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
    
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If
    
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
    CheckForUndefinedValueAndRaiseError pRepDG, CLng(TopRailCP), "CrossSectionCardinalPoints", 102

    'Other Handrail Details
    HorizontalOffset = pInputs.GetInputByIndex(6).IJDInputDuringGame.Result.UomValue
    dHorizOffset = pInputs.GetInputByIndex(7).IJDInputDuringGame.Result.UomValue
    CheckForUndefinedValueAndRaiseError pRepDG, CLng(HorizontalOffset), "HandrailOffset", 110
    If HorizontalOffset = 15 Then
        dHorizOffset = 0#
    End If
    
    'get orientation of HR which is required to get toprail curve
    HandrailOrientation = pInputs.GetInputByIndex(ORIENTATION_INDEX).IJDInputDuringGame.Result.UomValue
    CheckForUndefinedValueAndRaiseError pRepDG, CLng(HandrailOrientation), "HandrailOrientation", 101

    
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
    
    'use actual orientation only if there is no arc in handrail path
    If ComplexstringcontainsArc(Sketch.GetComplexString) Then
        HandrailOrientation = 0
    End If
    Dim DeltaHOffset As Double
    Dim TempHOffset As Integer
    DeltaHOffset = 0#
    TempHOffset = 0
    
    Dim OffsetReference As Long
    OffsetReference = pInputs.GetInputByIndex(SIDEMNTBYMEM_OFFSET_REFERENCE).IJDInputDuringGame.Result.UomValue
    
    Dim FloorThickness As Double
    FloorThickness = pInputs.GetInputByIndex(SIDEMNTBYMEM_FLOOR_THICKNESS).IJDInputDuringGame.Result.UomValue

    Dim ColumnClearance As Double
    ColumnClearance = pInputs.GetInputByIndex(SIDEMNTBYMEM_COLUMN_CLEARANCE).IJDInputDuringGame.Result.UomValue
    
    Dim EndClearance As Double
    EndClearance = pInputs.GetInputByIndex(SIDEMNTBYMEM_END_CLEARANCE).IJDInputDuringGame.Result.UomValue
    
    Dim StartColumnOffset As Double
    StartColumnOffset = pInputs.GetInputByIndex(SIDEMNTBYMEM_START_COLUMN_OFFSET).IJDInputDuringGame.Result.UomValue
    
    Dim EndColumnOffset As Double
    EndColumnOffset = pInputs.GetInputByIndex(SIDEMNTBYMEM_END_COLUMN_OFFSET).IJDInputDuringGame.Result.UomValue
    Dim tmpComplexString As ComplexString3d
    
    createNewSketch3D Sketch, OffsetReference, HorizontalOffset, FloorThickness, ColumnClearance, EndClearance, StartColumnOffset, EndColumnOffset, DeltaHOffset, TempHOffset, tmpComplexString
    
     ' side mounted to member sketch curves must be offset twice
    ' once to bring the sketch in lign with the offset reference
    ' and once to honor the user specified offset
    
    Set p3dCurve = GetOffsetCurve(tmpComplexString, TempHOffset, DeltaHOffset)
    Set pComplex = GetOffsetCurve(p3dCurve, HorizontalOffset, dHorizOffset)
    
    Dim hDelta As Double
    Dim wDelta As Double
    'Push the TopRail to some depth based on it's cardinal point to honour the
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
    
    'Ignore earlier calculation of hDelta and get hDelta again so that top of toprail ia maintained in all cases
    GethDeltaForTopRail TopRailStandard, TopRailSection, TopRailCP, TopRailAngle, hDelta, TopRailWidth, TopRailDepth
    
    'Set the new height
    Height = Height - hDelta
    
    'Get the Curves defined by Sketch3D
    If Not pSegments Is Nothing Then
        pSegments.Clear
        Set pSegments = Nothing
    End If

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

        If dCenY <> 0# Then
            Set opt2 = Nothing
            Set opt2 = opt1.Offset(ovec2)
        Else
            ovec2.Set 0#, 0#, 1#
            Set opt2 = opt1
        End If
        Set opt1 = Nothing
        Set ovec3 = Nothing
        If dCenX > 0# Then
            If dCenY >= 0 Then
                Set ovec3 = ovec1.Cross(ovec2)
            ElseIf dCenY < 0# Then
                Set ovec3 = ovec2.Cross(ovec1)
            End If
                ovec3.Length = Abs(dCenX)
            Set opt1 = opt2.Offset(ovec3)
        ElseIf dCenX < 0 Then
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

    ' Create the Symbol geometry
    'Create ruled surface between HR path and curve along top rail
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

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnAttributeChange = "ERROR"
    
    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String
    Dim i As Integer
    Dim pColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim NonStateRO As Long
    
    If bOnPreLoad = False Then
        ErrStr = UserAttributeMgmt_Validate(pIJDAttrs, pAttrToChange.InterfaceName, pAttrToChange.attrName, varNewAttrValue)
        If Len(ErrStr) > 0 Then
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
IJUserAttributeMgmt_OnPreLoad = "ERROR"
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

Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
Const METHOD = "UserAttributeMgmt_Validate"
On Error GoTo ErrorHandler

' first of all check if the symbol definition has CMCheck methods defined
    UserAttributeMgmt_Validate = "ERROR"

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
     
      
    Dim strAttribValid As Boolean
    strAttribValid = True
    
    If (sAttributeName = "Height") Or (sAttributeName = "NoOfMidRails") Or (sAttributeName = "TopOfMidRailDim") Or (sAttributeName = "MidRailSpacing") Then
    
        ValidateHandrailAttributes pIJDAttrs, sInterfaceName, strAttribValid
        
        If strAttribValid = False Then
            UserAttributeMgmt_Validate = "Invalid Attribute - Check User inputs"
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
                    UserAttributeMgmt_Validate = ErrMessage
                    Exit Function
                Else
                End If
            End If
        End If
    End If
    '***************************************************************************************
    ' Symbol CMCheck passed and now begins the attribute management's validation process
    ' First we check to see if the attribute being modified by the user is on a userDefined
    ' interface in the catalog; If not we return a "SchemaERROR"
    '***************************************************************************************
    
    Dim InterfaceID As Variant
    Dim oAttrObj As IJDAttributeInfo
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
    Dim oAttrCol As IJDInfosCol
    Dim IsInterfaceFound As Boolean
    Dim AttrCount As Long
    Dim AttrType As Long
    
    Set oAttributeMetaData = pIJDAttrs
    IsInterfaceFound = False
    For Each InterfaceID In pIJDAttrs
        Set oInterfaceInfo = Nothing
        Set oInterfaceInfo = oAttributeMetaData.InterfaceInfo(InterfaceID)
        If (oInterfaceInfo.IsHardCoded = False) Then
            If (oInterfaceInfo.name = sInterfaceName) Then
                IsInterfaceFound = True
                Exit For
            End If
        End If
    Next
    
    Set oInterfaceInfo = Nothing
    
    If IsInterfaceFound = False Then
        UserAttributeMgmt_Validate = "SchemaERROR"
        GoTo ErrorHandler
    End If
    Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
    
    ' loop on the attributes on the interface to match the supplied attribute type
    ' We can check for specific attributes to be within a given range of values we choose
        
    For AttrCount = 1 To oAttrCol.Count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        If oAttrObj.name = sAttributeName Then
            If (sAttributeName = "BeginExtensionLength") Or (sAttributeName = "EndExtensionLength") Then
                If (varAttributeValue < 0#) Then
                    UserAttributeMgmt_Validate = "Invalid Attribute Value "
                    Set oAttributeMetaData = Nothing
                    Exit Function
                End If
            End If
        End If
    Next
    
    Set oAttributeMetaData = Nothing
    
    UserAttributeMgmt_Validate = ""
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

Public Sub CreatePosts(pOC As IJDOutputCollection, _
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
                         EndTreatType As Integer, _
                         dPadOffset As Double, _
                         MinClearenceAtPostTurn As Double, _
                         MaxClearenceAtPostTurn As Double, _
                         PostSectionDepth As Double, _
                         MidRailSectionDepth As Double, _
                         Mirror As Boolean)


Const METHOD = "CreatePosts"
On Error GoTo ErrorHandler
    Call CreatePosts_Ex(pOC, _
                        pSegments, _
                        PostHeight, _
                        midrailHt, _
                        Orientation, _
                        sPt, _
                        ePt, _
                        SegmentMaxSpacing, _
                        SlopedSegmentMaxSpacing, _
                        BeginExtensionLength, _
                        EndExtensionLength, _
                        PostSection, _
                        PostSecStandard, _
                        bPostAtEveryTurn, _
                        CSectionCP, _
                        CSectionAngle, _
                        BeginTreatType, _
                        EndTreatType, _
                        dPadOffset, _
                        MinClearenceAtPostTurn, _
                        MaxClearenceAtPostTurn, _
                        PostSectionDepth, _
                        MidRailSectionDepth, _
                        Mirror)
    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub
Private Sub CreatePosts_Ex(oOutputParentObject As Object, _
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
                         EndTreatType As Integer, _
                         dPadOffset As Double, _
                         MinClearenceAtPostTurn As Double, _
                         MaxClearenceAtPostTurn As Double, _
                         PostSectionDepth As Double, _
                         MidRailSectionDepth As Double, Mirror As Boolean, Optional CalcVolume As Boolean = False, Optional ByRef volume As Double, Optional ByRef VCogX As Double, Optional ByRef VCogY As Double, Optional ByRef VCogZ As Double)


Const METHOD = "CreatePosts_Ex"
On Error GoTo ErrorHandler

    Dim outCount                As Long
    Dim bNoCSSymbol             As Boolean

    Dim PostCSArea              As Double
    Dim PostPosition As DPosition
    Dim Posts As Collection
    Dim post As HandrailPost
    Dim DeltaAngle              As Double
    DeltaAngle = PI 'EndTreatment rotation


    Dim pOC             As IJDOutputCollection
    Dim oSectionOccPost As Object       ' cross section occurrence used by symbol for post
    Dim oSectionDefPost As Object       ' cross-section definition used by members for post
    Dim oSectionObjPost As Object       ' common cross section obj sent to BuildHandrailOutput for post
    
    Dim oSectionOccRail As Object       ' cross section occurrence used by symbol for topRail
    Dim oSectionDefRail As Object       ' cross-section definition used by members for topRail
    Dim oSectionObjRail As Object       ' common cross section obj sent to BuildHandrailOutput for topRail
    
    
    Dim MaxPostDistClearance As Double
    Dim MinPostDistClearance As Double
    Dim PostDistBeforeTurnPercent As Double
    Dim PostDistAfterTurnPercent As Double
    
    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If
                    
    MaxPostDistClearance = MaxClearenceAtPostTurn
    MinPostDistClearance = MinClearenceAtPostTurn
    
    outCount = 1

    Dim ToprailCSArea        As Double

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
    'MsgBox "Postcount = " & postcount & ", TreatCount = " & treatcount
    
    If Not Posts Is Nothing Then
        While Posts.Count > 0
            Posts.Remove Posts.Count
        Wend
    End If
    Set Posts = Nothing
    
    Set PostPosition = Nothing
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


Private Sub createNewSketch3D(oldPath As Sketch3d, OffsetReference As Long, HorizontalOffset As Integer, FloorThickness As Double, ColumnClearance As Double, EndClearance As Double, ByRef StartColumnOffset As Double, ByRef EndColumnOffset As Double, ByRef DeltaHOffset As Double, ByRef TempHOffset As Integer, ByRef oComplexString As ComplexString3d) 'newPath As Sketch3d
Const METHOD = "createNewSketch3D"
On Error GoTo ErrorHandler

        Dim oassochelper As IJPointAssocConstraintHelper
    Dim oPointsColl As IJElements
    Dim oCachedPoint As CachedPoint
    
    Dim oPt As IJPoint
    Dim oCurve As IJCurve
    
    Dim oEl As IJElements
    Dim i As Integer, status As Integer, oToggleSide As Integer
    Dim oRangeValues() As Double
    
    Dim Angle As Double, b As Double, c As Double, distance(2) As Double
    
    Dim a As Long
    Dim strPointOnGeometryProgID As String
    Dim bFlag As Boolean
    
    Dim startx As Double, starty As Double, startz As Double
    Dim MidX As Double, MidY As Double, MidZ As Double
    Dim Endx As Double, Endy As Double, Endz As Double
    Dim TempX As Double, TempY As Double, TempZ As Double
    Dim dDistanceFromStart As Double, dDistanceFromEnd As Double
    
        oldPath.CacheTurnPoints oPointsColl
    
    Dim iElements As IJElements
    
    
    Dim tempComplex As ComplexString3d
    Set tempComplex = oldPath.GetComplexString
    Dim Segments  As IJElements
    tempComplex.GetCurves Segments
    Set oCurve = Segments(1)
    oCurve.EndPoints startx, starty, startz, Endx, Endy, Endz
    If Segments.Count > 1 Then
        Set oCurve = Segments(Segments.Count)
        oCurve.EndPoints MidX, MidY, MidZ, Endx, Endy, Endz
    End If
    
    ReDim oRangeValues(2)
    
    If Not ValidateInputPath(Segments) Then
        Set tempComplex = Nothing
        Err.Raise E_FAIL, METHOD, "This Handrail doesn't support the path's inclination"
    End If

    Set oassochelper = New PointAssocConstraintHelper
    Set iElements = New JObjectCollection

    Dim DeltaYOffset As Double
    Dim oMembPartPrism As ISPSMemberPartPrismatic
    DeltaYOffset = 0
    Set oComplexString = tempComplex
    Dim bSetPathAlongMember As Boolean
    bSetPathAlongMember = False
    'Get The information on supporting member's  related columns, and make end offset later
    If Not oPointsColl Is Nothing Then
        distance(1) = EndClearance
        distance(2) = EndClearance
        Set oCachedPoint = New CachedPoint
        Dim lPointIndex As Long
        For lPointIndex = 1 To oPointsColl.Count
            Set oCachedPoint = oPointsColl.Item(lPointIndex)
            If oCachedPoint.IsAssocPoint Then
                oCachedPoint.GetAssocPoint oPt
                oPt.GetPoint TempX, TempY, TempZ
                dDistanceFromStart = (TempX - startx) * (TempX - startx) + (TempY - starty) * (TempY - starty) + (TempZ - startz) * (TempZ - startz)
                dDistanceFromEnd = (TempX - Endx) * (TempX - Endx) + (TempY - Endy) * (TempY - Endy) + (TempZ - Endz) * (TempZ - Endz)
                On Error Resume Next
                    oassochelper.GetInfo oPt, strPointOnGeometryProgID, oEl, c, bFlag, b, a
                On Error GoTo ErrorHandler
                If Not oEl Is Nothing Then
                    On Error Resume Next
                    For i = 1 To 2
                        Set oMembPartPrism = Nothing
                        Set oMembPartPrism = GetMembAssociatedWithConn(oEl, oPt, i) 'To get the supported member to this frame / split connection
                        If Not oMembPartPrism Is Nothing Then Exit For
                    Next
                    On Error GoTo ErrorHandler
                    status = getColumnRangeValues(oEl, oMembPartPrism, oPt, i, oRangeValues) 'To get Rangebox values from the supporting member
                    If status <> 0 Then
                        If dDistanceFromStart < dtol Then
                            If StartColumnOffset = 0 Then
                                StartColumnOffset = (oRangeValues(1) / 2)
                            End If
                            distance(1) = ColumnClearance + StartColumnOffset
                        ElseIf dDistanceFromEnd < dtol Then
                            If EndColumnOffset = 0 Then
                                EndColumnOffset = (oRangeValues(1) / 2)
                            End If
                            distance(2) = ColumnClearance + EndColumnOffset
                        End If
                    End If
                End If
                Set oPt = Nothing
                Set oEl = Nothing
            End If
        Next
        Dim bNeedUpdatePath As Boolean
        bNeedUpdatePath = False
        ' chop the start end off
        If Abs(distance(1)) > dtol Then
            Set oCurve = Segments.Item(1)
            SetNewPosition oCurve, 1, distance(1)
            bNeedUpdatePath = True
        End If
        ' chop the end end off
        If Abs(distance(2)) > dtol Then
            Set oCurve = Segments(Segments.Count)
            SetNewPosition oCurve, 2, distance(2)
            bNeedUpdatePath = True
        End If
        If m_GeomFactory Is Nothing Then
            Set m_GeomFactory = New GeometryFactory
        End If

        If bNeedUpdatePath Then
            Set tempComplex = Nothing
            Set tempComplex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, Segments)
        End If
        
        ' get supporting member's geometrical information through the assoc points and offset the path accordingly
        For lPointIndex = 1 To oPointsColl.Count
            Set oCachedPoint = oPointsColl.Item(lPointIndex)
            If oCachedPoint.IsAssocPoint Then
                oCachedPoint.GetAssocPoint oPt
                On Error Resume Next
                oassochelper.GetInfo oPt, strPointOnGeometryProgID, oEl, c, bFlag, b, a
                On Error GoTo ErrorHandler
                If Not oEl Is Nothing Then
                    Set oMembPartPrism = GetMembAssociatedWithConn(oEl, oPt, i) 'To get the supported member to this frame / split connection
                End If
                If Not oMembPartPrism Is Nothing Then 'to calculate DeltaHOffset with respect to the starting point of the member
                    DeltaHOffset = getSupportingMemCPPoint(OffsetReference, HorizontalOffset, oMembPartPrism, DeltaYOffset, TempHOffset)
                    Exit For
                End If
            End If
        Next
        If Abs(FloorThickness + DeltaYOffset) > dtol Then
            CreateVerticalOffsetCurve oComplexString, tempComplex, FloorThickness + DeltaYOffset
        End If
       
    End If
    Set tempComplex = Nothing

'GoTo cleanup:
    ReleaseOpenObjects oPointsColl
    Set iElements = Nothing
    Set oassochelper = Nothing
    Set oPointsColl = Nothing
    Set oCachedPoint = Nothing
    Set oPt = Nothing
    Set oCurve = Nothing
    Set oEl = Nothing
    Set iElements = Nothing
    Set Segments = Nothing
    Set oMembPartPrism = Nothing
    
Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub ReleaseOpenObjects(oPointsColl As IJElements)
If Not oPointsColl Is Nothing Then
    Dim i As Integer
    Dim oCachedPoint As CachedPoint
    
    For i = 1 To oPointsColl.Count
        Set oCachedPoint = oPointsColl.Item(i)
        oCachedPoint.ClearLineFeatures
    Next
    oPointsColl.Clear
    
    Set oCachedPoint = Nothing
    Set oPointsColl = Nothing
End If

End Sub
Private Function ValidateInputPath(Segments As IJElements) As Boolean
Const METHOD = "ValidateInputPath"
On Error GoTo ErrorHandler
ValidateInputPath = True
        
    Dim stx As Double, sty As Double, stz As Double
    Dim Endx As Double, Endy As Double, Endz As Double
    Dim Vec1 As DVector, Norm As DVector, ZVec As DVector
    
    Dim Angle As Double
    Dim oLine As IJLine
    Dim ii As Integer
    Set Norm = New DVector
    Set ZVec = New DVector
    Norm.Set 0#, 0#, 1#
    ZVec.Set 0#, 0#, 1#
    
    For ii = 1 To Segments.Count
        If TypeOf Segments.Item(ii) Is IJLine Then
           Set oLine = Segments.Item(ii)
           oLine.GetStartPoint stx, sty, stz
           oLine.GetEndPoint Endx, Endy, Endz
           If Endz - stz < 0 Then 'Interchange the points
               oLine.GetEndPoint stx, sty, stz
               oLine.GetStartPoint Endx, Endy, Endz
           End If
           Set oLine = Nothing
           Set Vec1 = New DVector
           Vec1.Set (Endx - stx), (Endy - sty), (Endz - stz)
           Angle = ZVec.Angle(Vec1, Norm)
           Angle = Angle * 180 / PI
           If Angle >= 30 - dtol And Angle <= 90 + dtol Then
               ValidateInputPath = True
           Else
               Set Vec1 = Nothing
               Set Norm = Nothing
               Set ZVec = Nothing
               ValidateInputPath = False
               Exit Function
           End If
        End If
    Next ii

    Set Vec1 = Nothing
    Set Norm = Nothing
    Set ZVec = Nothing
Exit Function
ErrorHandler:
  m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
Private Function getColumnRangeValues(oEl As IJElements, oSuppMemb As ISPSMemberPartPrismatic, oPt As IJPoint, ptCount As Integer, ByRef oRange() As Double) As Integer
Const METHOD = "getColumnRangeValues"
On Error GoTo ErrorHandler

    getColumnRangeValues = False
    Dim oFrameConn As ISPSFrameConnection
    Dim oSplitConn As ISPSSplitMemberConnection
    Dim iHelper As IJStructILCHelper
    Dim status As StructSOCInputHelperStatus
    Dim oMemSystem1 As ISPSMemberSystem
    Dim oMemSystem2 As ISPSMemberSystem
    Dim oMemPartPrism As ISPSMemberPartPrismatic
    Dim oMemCrossSection As ISPSCrossSection
    Dim oRelatedObj1 As Object
    Dim oRelatedObj2 As Object
    Dim parents As IJElements
    Dim secWidth As Double
    Dim secDepth As Double
    Dim BetaAngle As Double
       
    If Not oEl Is Nothing Then
        If oEl.Count > 0 Then
            If TypeOf oEl.Item(1) Is ISPSFrameConnection Then
                Set oFrameConn = oEl.Item(1)
                oFrameConn.InputHelper.GetRelatedObjects oFrameConn, oRelatedObj1, oRelatedObj2
                If oRelatedObj1 Is Nothing And oRelatedObj2 Is Nothing Then
                    getColumnRangeValues = False
                        GoTo cleanup
                End If
                If oRelatedObj1 Is Nothing Then
                    getColumnRangeValues = False
                ElseIf TypeOf oRelatedObj1 Is ISPSMemberSystem Then 'And oRelatedObj1 <> oMPP.MemberSystem Then
                    Set oMemSystem1 = oRelatedObj1
                    Set oMemPartPrism = oMemSystem1.DesignPartAtEnd(SPSMemberAxisStart)
                    If ValidateMemberInclination(oMemPartPrism) Then 'Is member inclination is >45
                        Set oMemCrossSection = oMemPartPrism.CrossSection
                        GetCrossSecData oMemCrossSection.SectionName, oMemCrossSection.SectionStandard, secWidth, secDepth
                        oRange(1) = GetClearanceValue(oSuppMemb, oMemPartPrism)
                        getColumnRangeValues = True
                            GoTo cleanup
                    End If
                End If
                If oRelatedObj2 Is Nothing Then
                    getColumnRangeValues = False
                ElseIf TypeOf oRelatedObj2 Is ISPSMemberSystem Then 'And oRelatedObj1 <> oMPP.MemberSystem Then
                    Set oMemSystem2 = oRelatedObj2
                    Set oMemPartPrism = oMemSystem2.DesignPartAtEnd(SPSMemberAxisStart)
                    If ValidateMemberInclination(oMemPartPrism) Then 'Is member inclination is >45
                        Set oMemCrossSection = oMemPartPrism.CrossSection
                        GetCrossSecData oMemCrossSection.SectionName, oMemCrossSection.SectionStandard, secWidth, secDepth
                        oRange(1) = GetClearanceValue(oSuppMemb, oMemPartPrism)
                        getColumnRangeValues = True
                            GoTo cleanup
                    End If
                End If
                Set oFrameConn = Nothing
            End If 'end of Frame connection check.
            
            If TypeOf oEl.Item(1) Is ISPSSplitMemberConnection Then
                Set oSplitConn = oEl.Item(1)
                Dim oColl As IJElements
                Dim px As Double, py As Double, pz As Double
                Dim oMPPPrev As ISPSMemberPartPrismatic, oMPPNext As ISPSMemberPartPrismatic
                Dim i As Integer
                oPt.GetPoint px, py, pz
                
                Set oColl = oSplitConn.InputObjects
                If Not oColl Is Nothing Then
                    For i = 1 To oColl.Count
                        If TypeOf oColl.Item(i) Is ISPSMemberSystem Then 'And oColl.Item(i) <> oMPP.MemberSystem Then
                            Set oMemSystem1 = oColl.Item(i)
                            oMemSystem1.GetDesignPartsAtPosition px, py, pz, oMPPPrev, oMPPNext
                            If ptCount = 1 Then
                                Set oMemPartPrism = oMPPNext
                            ElseIf ptCount = 2 Then
                                Set oMemPartPrism = oMPPPrev
                            End If
                            Set oMPPNext = Nothing
                            Set oMPPPrev = Nothing
                            If ValidateMemberInclination(oMemPartPrism) Then 'Is member inclination is >45
                                Set oMemCrossSection = oMemPartPrism.CrossSection
                                GetCrossSecData oMemCrossSection.SectionName, oMemCrossSection.SectionStandard, secWidth, secDepth
                                oRange(1) = GetClearanceValue(oSuppMemb, oMemPartPrism)
                                getColumnRangeValues = True
                                Set oSplitConn = Nothing
                                    GoTo cleanup
                            End If
                        End If
                    Next i
                End If
                Set oSplitConn = Nothing
            End If 'end of SPlit connection check.
        End If 'end of elements count check
    End If ' end of elements collection check.

    Set oMemSystem1 = Nothing
    Set oMemSystem2 = Nothing
    Set oMemPartPrism = Nothing
    Set oMemCrossSection = Nothing
    Set oRelatedObj1 = Nothing
    Set oRelatedObj2 = Nothing
    Set parents = Nothing
cleanup:

    Set oMemSystem1 = Nothing
    Set oMemSystem2 = Nothing
    Set oMemPartPrism = Nothing
    Set oMemCrossSection = Nothing
    Set oRelatedObj1 = Nothing
    Set oRelatedObj2 = Nothing
    Set parents = Nothing
    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
Private Function GetClearanceValue(oSupportedMemb As ISPSMemberPartPrismatic, oSupportingMemb As ISPSMemberPartPrismatic) As Double
    
    Dim oFacePos() As IJDPosition
    Dim IntPos As IJDPosition
    Dim oLine As IJLine, oProLine As IJLine
    Dim ii As Integer
    
    Dim vec As DVector
    Dim ZVec As DVector
    Dim NormVec As DVector
    
    Dim Pos As IJDPosition
    
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    
    Dim StPt As IJPoint
    Dim EndPt As IJPoint
    
    Set oLine = New Line3d
    ReDim oFacePos(0 To 3)
    GetBoundingRectangleForMemberEnd oSupportingMemb, SPSMemberAxisEnd, oFacePos()
    
    
    Set IntPos = New DPosition
    IntPos.x = (oFacePos(0).x + oFacePos(2).x) / 2
    IntPos.y = (oFacePos(0).y + oFacePos(2).y) / 2
    IntPos.z = (oFacePos(0).z + oFacePos(2).z) / 2
    
    Dim Rot As Double
    Rot = oSupportingMemb.Rotation.BetaAngle * 180 / PI
    
    If (Rot >= 0 And Rot < 90) Or (Rot >= 180 And Rot < 270) Then
        oLine.DefineBy2Points IntPos.x, IntPos.y, IntPos.z, oFacePos(3).x, oFacePos(3).y, oFacePos(3).z
    Else
        oLine.DefineBy2Points IntPos.x, IntPos.y, IntPos.z, oFacePos(2).x, oFacePos(2).y, oFacePos(2).z
    End If
       
    Set vec = New DVector
    Set ZVec = New DVector
    Set NormVec = New DVector
    
    ZVec.Set 0#, 0#, 1#
    
       
    Set StPt = oSupportedMemb.PointAtEnd(SPSMemberAxisStart)
    Set EndPt = oSupportedMemb.PointAtEnd(SPSMemberAxisEnd)
    
    StPt.GetPoint x1, y1, z1
    EndPt.GetPoint x2, y2, z2
    
    vec.Set (x2 - x1), (y2 - y1), (z2 - z1)
    Set NormVec = vec.Cross(ZVec)
    Set Pos = New DPosition
    
    Pos.Set x1, y1, z1
    Set oProLine = ProjectLineToPlane(oLine, Pos, NormVec)
    
        GetClearanceValue = 2 * oProLine.Length
    For ii = LBound(oFacePos) To UBound(oFacePos)
        Set oFacePos(ii) = Nothing
    Next ii
    Set IntPos = Nothing
    Set oLine = Nothing
    Set oProLine = Nothing
    Set vec = Nothing
    Set ZVec = Nothing
    Set NormVec = Nothing
    Set Pos = Nothing
    Set StPt = Nothing
    Set EndPt = Nothing
  

End Function
Private Sub GetBoundingRectangleForMemberEnd(oMemb As ISPSMemberPartPrismatic, iPort As SPSMemberAxisPortIndex, oFacePos() As IJDPosition)
  Const METHOD = "GetBoundingRectangleForMemberEnd"
    On Error GoTo ErrorHandler
    
    Dim oPos As IJDPosition
    Dim oProfileBO As ISPSCrossSection
    Dim yOffset As Double, zOffset As Double, yOffsetCP As Double, zOffsetCP As Double
    Dim oMat As IJDT4x4
    Dim eX#, eY#, eZ#, sX#, sY#, sZ#, dblLength#
    Dim oVec As New DVector
    Dim cp As Long
    
    Set oPos = New DPosition
    
    oMemb.Rotation.GetTransform oMat 'Get the member coordinate system
    Set oProfileBO = oMemb.CrossSection
    If iPort = SPSMemberAxisEnd Then
        oMemb.Axis.EndPoints sX, sY, sZ, eX, eY, eZ
        oVec.Set eX - sX, eY - sY, eZ - sZ
            dblLength = oVec.Length
    Else
        dblLength = 0#
    End If
    cp = oProfileBO.CardinalPoint
    oProfileBO.GetCardinalPointOffset cp, yOffsetCP, zOffsetCP 'Returns Rad x and y of active CP which is member local -y and z.
    
    oProfileBO.GetCardinalPointOffset 1, yOffset, zOffset 'Returns Rad x and y of CP=1 which is member local -y and z.
    yOffset = yOffset - yOffsetCP ' get the y offset from current CP
    zOffset = zOffset - zOffsetCP ' get the z offset from current CP
    oPos.Set dblLength, -1# * yOffset, zOffset ' point is now in the local coordinate system of the member. The yooset is multiplied by -1 because RAD + x is member local -Y.
    Set oFacePos(0) = oMat.TransformPosition(oPos)
    
    oProfileBO.GetCardinalPointOffset 3, yOffset, zOffset
    yOffset = yOffset - yOffsetCP ' get the y offset from current CP
    zOffset = zOffset - zOffsetCP ' get the z offset from current CP
    oPos.Set dblLength, -1# * yOffset, zOffset ' point is now in the local coordinate system of the member. The yooset is multiplied by -1 because RAD + x is member local -Y.
    Set oFacePos(1) = oMat.TransformPosition(oPos)
    
    oProfileBO.GetCardinalPointOffset 9, yOffset, zOffset
    yOffset = yOffset - yOffsetCP ' get the y offset from current CP
    zOffset = zOffset - zOffsetCP ' get the z offset from current CP
    oPos.Set dblLength, -1# * yOffset, zOffset ' point is now in the local coordinate system of the member. The yooset is multiplied by -1 because RAD + x is member local -Y.
    Set oFacePos(2) = oMat.TransformPosition(oPos)
    
    oProfileBO.GetCardinalPointOffset 7, yOffset, zOffset
    yOffset = yOffset - yOffsetCP ' get the y offset from current CP
    zOffset = zOffset - zOffsetCP ' get the z offset from current CP
    oPos.Set dblLength, -1# * yOffset, zOffset ' point is now in the local coordinate system of the member. The yooset is multiplied by -1 because RAD + x is member local -Y.
    Set oFacePos(3) = oMat.TransformPosition(oPos)
    
    Exit Sub
ErrorHandler:
 m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
    
End Sub
Private Function ProjectLineToPlane(oLine As IJLine, oPlaneRoot As IJDPosition, oPlaneNormal As IJDVector) As IJLine
  Const METHOD = "ProjectLineToPlane"
    On Error GoTo ErrorHandler
    Dim pGeometryFactory As New GeometryFactory
    Dim oProjectedLine As IJLine
    Dim oStartPos As New DPosition, oEndPos As New DPosition
    Dim x#, y#, z#

    oLine.GetStartPoint x, y, z
    oStartPos.Set x, y, z
    
    oLine.GetEndPoint x, y, z
    oEndPos.Set x, y, z
    
    Set oStartPos = ProjectPosToPlane(oStartPos, oPlaneRoot, oPlaneNormal)
    Set oEndPos = ProjectPosToPlane(oEndPos, oPlaneRoot, oPlaneNormal)
    
    Set oProjectedLine = pGeometryFactory.Lines3d.CreateBy2Points(Nothing, oStartPos.x, oStartPos.y, oStartPos.z, oEndPos.x, oEndPos.y, oEndPos.z)
    Set ProjectLineToPlane = oProjectedLine
    Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
Private Function ProjectPosToPlane(oPos As IJDPosition, oPlaneRoot As IJDPosition, oPlaneNormal As IJDVector) As IJDPosition
  Const METHOD = "ProjectPosToPlane"
    On Error GoTo ErrorHandler
    Dim oVec As New DVector
    Dim oProjectedPos As New DPosition
    Dim dist#
    oVec.Set oPlaneRoot.x - oPos.x, oPlaneRoot.y - oPos.y, oPlaneRoot.z - oPos.z
    'normalize the vector, if not done before passing in
        oPlaneNormal.Length = 1
    'get the projected length of oVec along plane normal
    dist = oPlaneNormal.Dot(oVec)
    'prject the position to the plane
    oProjectedPos.x = dist * oPlaneNormal.x + oPos.x
    oProjectedPos.y = dist * oPlaneNormal.y + oPos.y
    oProjectedPos.z = dist * oPlaneNormal.z + oPos.z

    Set ProjectPosToPlane = oProjectedPos
    Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
Private Function getSupportingMemCPPoint(OffsetReference As Long, HorizontalOffset As Integer, oMemPartPrismmatic As ISPSMemberPartPrismatic, ByRef DeltaYOffset As Double, ByRef TempHOffset As Integer) As Double
Const METHOD = "getSupportingMemCPPoint"
On Error GoTo ErrorHandler

    Dim DeltaX As Double, DeltaY As Double
    
    Dim ICrossSection As ISPSCrossSection
    Dim IRotation As ISPSAxisRotation
    Dim uDistance As Double, vDistance As Double
    Dim depth As Double, bFlange As Double
    Dim TempPos As DPosition
    Dim bMirror As Boolean
    Set IRotation = oMemPartPrismmatic.Rotation
    Set ICrossSection = oMemPartPrismmatic.CrossSection
    bMirror = IRotation.Mirror
    Set TempPos = New DPosition
    TempPos.Set 0, 0, 0
    ICrossSection.GetCardinalPointDelta TempPos, 1, 9, uDistance, vDistance 'bottom left to top right
    depth = Abs(vDistance)
    bFlange = Abs(uDistance)
    ICrossSection.GetCardinalPointDelta TempPos, 5, 0, uDistance, vDistance
    If bMirror Then
        uDistance = -uDistance
    End If
    Set TempPos = Nothing
    Set IRotation = Nothing
    Set ICrossSection = Nothing
    DeltaY = depth / 2 - vDistance
    DeltaX = uDistance
    If OffsetReference = 3 Then 'outer
        If HorizontalOffset = 5 Then  'right
            DeltaX = uDistance - bFlange / 2
        ElseIf HorizontalOffset = 10 Then
            DeltaX = uDistance + bFlange / 2
        End If
    ElseIf OffsetReference = 1 Then 'Inner
        If HorizontalOffset = 5 Then  'right
            DeltaX = uDistance + bFlange / 2
        ElseIf HorizontalOffset = 10 Then
            DeltaX = uDistance - bFlange / 2
        End If
    End If
    If DeltaX > 0 Then
        TempHOffset = 10
    ElseIf DeltaX < 0 Then
        TempHOffset = 5
    Else
        TempHOffset = 15
    End If
    DeltaYOffset = DeltaY
    getSupportingMemCPPoint = Abs(DeltaX)
    
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
'This can be removed once middle function for getting ports from splitconnection (ISPSSplitMemberConnection - GetPartPorts) is ready
Private Function GetMembPartFromSplitConn(oSplitConn As ISPSSplitMemberConnection, oPt As IJPoint, ptCount As Integer, ByRef oMemPartPrismmatic As ISPSMemberPartPrismatic) As Boolean
Const METHOD = "GetMembPartFromSplitConn"
On Error GoTo ErrorHandler
        
        GetMembPartFromSplitConn = False
        Dim oCollMembParents As IJElements
        Dim oCollMembParts As IJElements
        
        Dim px As Double, py As Double, pz As Double
        
        Dim oMS1 As ISPSMemberSystem
        Dim oMPPPrev As ISPSMemberPartPrismatic
        Dim oMPPNext As ISPSMemberPartPrismatic
        
        oPt.GetPoint px, py, pz
        
        Dim jj As Integer
        Dim kk As Integer
        Set oCollMembParents = oSplitConn.InputObjects
        If Not oCollMembParents Is Nothing Then
            For jj = 1 To oCollMembParents.Count
                If TypeOf oCollMembParents.Item(jj) Is ISPSMemberSystem Then
                    Set oMS1 = oCollMembParents.Item(jj)
                    oMS1.GetDesignPartsAtPosition px, py, pz, oMPPPrev, oMPPNext
                    If ptCount = 1 Then
                        Set oMemPartPrismmatic = oMPPNext
                        GetMembPartFromSplitConn = True
                        Set oMPPNext = Nothing
                        GoTo cleanup
                    ElseIf ptCount = 2 Then
                        Set oMemPartPrismmatic = oMPPPrev
                        GetMembPartFromSplitConn = True
                        Set oMPPPrev = Nothing
                        GoTo cleanup
                    End If
                End If
            Next
        End If
        Set oCollMembParts = Nothing
cleanup:
    Set oCollMembParents = Nothing
    Set oCollMembParts = Nothing
    Set oMS1 = Nothing
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

'This is to check whether the supporting member's type is column or members inclination is greater than 45 deg
Private Function GetMembAssociatedWithConn(oEl As IJElements, oPt As IJPoint, ptCount As Integer) As ISPSMemberPartPrismatic
Const METHOD = "GetMembAssociatedWithConn"
On Error GoTo ErrorHandler
  
    Dim oMembPartPrismatic As ISPSMemberPartPrismatic
    Dim oSplitAxisPort As ISPSSplitAxisPort
    If Not oEl Is Nothing Then
        If oEl.Count > 0 Then
            If TypeOf oEl.Item(1) Is ISPSSplitAxisPort Then
                Set oSplitAxisPort = oEl.Item(1)
            ElseIf TypeOf oEl.Item(1) Is ISPSFrameConnection Then
                Dim oFrameConn As ISPSFrameConnection
                Set oFrameConn = oEl.Item(1)
                On Error Resume Next
                Set oSplitAxisPort = oFrameConn.DesignPartEndPort
                On Error GoTo ErrorHandler
            End If
            If Not oSplitAxisPort Is Nothing Then
                Set GetMembAssociatedWithConn = oSplitAxisPort.Part
                Set oSplitAxisPort = Nothing
                Set oFrameConn = Nothing
                Exit Function
            End If
            If TypeOf oEl.Item(1) Is ISPSSplitMemberConnection Then
                Dim oSplitConn As ISPSSplitMemberConnection
        
                Set oSplitConn = oEl.Item(1)
                If GetMembPartFromSplitConn(oSplitConn, oPt, ptCount, oMembPartPrismatic) Then
                    Set GetMembAssociatedWithConn = oMembPartPrismatic
                    Set oMembPartPrismatic = Nothing
                End If
                Set oSplitConn = Nothing
            End If
        End If
    End If
Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
Private Function ValidateMemberInclination(obj As ISPSMemberPartPrismatic) As Boolean
Const METHOD = "ValidateMemberInclination"
On Error GoTo ErrorHander
    ValidateMemberInclination = False
    If TypeOf obj Is ISPSMemberPartPrismatic Then
        Dim oMembPartPrism As ISPSMemberPartPrismatic
        Set oMembPartPrism = obj

        Dim stx As Double, sty As Double, stz As Double
        Dim Endx As Double, Endy As Double, Endz As Double


        Dim x As Double, y As Double, z As Double

        Dim StPoint As IJPoint
        Dim EndPoint As IJPoint

        Dim Vec1 As IJDVector
        Dim ZVec As IJDVector
        Dim Norm As IJDVector
        Dim Angle As Double
        Dim Value As Double


        Set Vec1 = New DVector
        Set ZVec = New DVector
        Set Norm = New DVector

        Set StPoint = New Point3d
        Set EndPoint = New Point3d
        Dim tmpPoint As IJPoint
        Set tmpPoint = New Point3d

        ZVec.Set 0#, 0#, 1#
        Norm.Set 0#, 0#, 1#

        Set StPoint = oMembPartPrism.PointAtEnd(SPSMemberAxisStart)
        Set EndPoint = oMembPartPrism.PointAtEnd(SPSMemberAxisEnd)
        

        StPoint.GetPoint stx, sty, stz
        EndPoint.GetPoint Endx, Endy, Endz

        If Endz - stz < 0 Then
            Set tmpPoint = StPoint
            Set StPoint = EndPoint
            Set EndPoint = tmpPoint
            StPoint.GetPoint stx, sty, stz
            EndPoint.GetPoint Endx, Endy, Endz
        Else
        End If
        Vec1.Set (Endx - stx), (Endy - sty), (Endz - stz)


        Angle = ZVec.Angle(Vec1, Norm)
        Angle = Angle * 180 / PI
        If Angle >= 45 - dtol Then
            ValidateMemberInclination = False
        Else
            ValidateMemberInclination = True
        End If
    End If
    Set StPoint = Nothing
    Set EndPoint = Nothing
    Set Vec1 = Nothing
    Set Norm = Nothing
    Set ZVec = Nothing
    Exit Function
ErrorHander:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Private Function TransformPosToGlobal(oMemb As ISPSMemberPartPrismatic, oPos As IJDPosition) As IJDPosition
Const MT = "TransformPosToGlobal"
  On Error GoTo ErrorHandler
    Dim oProfileBO As ISPSCrossSection
    Dim xOffset As Double, yOffset As Double, xOffsetCP As Double, yOffsetCP As Double, xOffsetCP5 As Double, yOffsetCP5 As Double
    Dim oMat1 As New DT4x4, oMat2 As DT4x4
    Dim oVec As New DVector

    
    Set oProfileBO = oMemb.CrossSection
    oProfileBO.GetCardinalPointOffset oProfileBO.CardinalPoint, xOffsetCP, yOffsetCP 'Returns x and y of the current CP in RAD coordinates, which is member negative y and z.
    oProfileBO.GetCardinalPointOffset 5, xOffsetCP5, yOffsetCP5
    xOffset = xOffsetCP5 - xOffsetCP ' x offset of current cp from cp5
    yOffset = yOffsetCP5 - yOffsetCP ' y offset of current cp from cp5
    
    oVec.Set xOffset, yOffset, 0
    oMat1.LoadIdentity
    oMat1.Translate oVec
    
    oMemb.Rotation.GetTransform oMat2
    'from cross section coordinates to member coordinates
    Set oMat2 = CreateCSToMembTransform(oMat2)
    oMat2.MultMatrix oMat1
    Set TransformPosToGlobal = oMat2.TransformPosition(oPos)
  Exit Function
ErrorHandler:
m_oErrors.Add Err.Number, MT, Err.Description
    Err.Raise E_FAIL
End Function
Private Function CreateCSToMembTransform(pI4x4Mem As IJDT4x4, Optional reflect As Boolean = False) As IJDT4x4
 Const MT = "CreateCSToMembTransform"
On Error GoTo ErrorHandler

  Dim tmpI4x4 As New DT4x4
  Dim pI4x4 As New DT4x4


  tmpI4x4.LoadMatrix pI4x4Mem
  pI4x4.LoadIdentity
  'loads 0 0 -1 0 to get a local 2d CS flipped into member local CS where x along axis, Z up Y perp web
  '     -1 0  0 0
  '      0 1  0 0
  '      0 0  0 1
  'Means local x = Member -Y
  '      local y = Member Z
  '      local z = Member -X
  
  pI4x4.IndexValue(0) = 0
  If reflect = True Then
    pI4x4.IndexValue(1) = 1
  Else
    pI4x4.IndexValue(1) = -1
  End If
  pI4x4.IndexValue(5) = 0
  pI4x4.IndexValue(6) = 1
  pI4x4.IndexValue(8) = -1
  pI4x4.IndexValue(10) = 0
  
  tmpI4x4.MultMatrix pI4x4
  Set CreateCSToMembTransform = tmpI4x4
  Set pI4x4 = Nothing
Exit Function
ErrorHandler:
   m_oErrors.Add Err.Number, MT, Err.Description
    Err.Raise E_FAIL
End Function

Private Sub ISPSHandrailSymbolServices_GetHandrailLength(ByVal HandrailBO As SPSHandrails.ISPSHandrail, length As Double)
Const METHOD = "ISPSHandrailSymbolServices_GetHandrailLength"
        On Error GoTo ErrorHandler

        Dim ColumnOffsetCurvesLength As Double
        Dim ColumnClearance As Double
        Dim EndClearance As Double
        Dim StartColumnOffset As Double
        Dim EndColumnOffset As Double
        Dim HRSketchPathLength As Double
        HRSketchPathLength = GetHandrailSketchPathLength(HandrailBO)
        Dim oAttrColl As Object
        Set oAttrColl = GetAttributeCollection(HandrailBO, "IJUAHRByMemberProps")
        ColumnClearance = GetAttributeValue(oAttrColl, "ColumnClearance")
        EndClearance = GetAttributeValue(oAttrColl, "EndClearance")
        StartColumnOffset = GetAttributeValue(oAttrColl, "StartColumnOffset")
        EndColumnOffset = GetAttributeValue(oAttrColl, "EndColumnOffset")
        ColumnOffsetCurvesLength = GetColumnOffsetCurvesLength(HandrailBO, ColumnClearance, EndClearance, StartColumnOffset, EndColumnOffset)
        length = HRSketchPathLength - ColumnOffsetCurvesLength
        
        Set oAttrColl = Nothing
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
       
    CalcRailsVolumeCGForSideMount oHandrail, PartOccInfoCol, volume1, VCOGX1, VCOGY1, VCOGZ1
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
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
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
    Set HRpathCurve = HRpath.GetComplexString
    Dim iOffsetType As Integer
    iOffsetType = GetAttribute(OccAttrs, "SPSHRPathHorizontalOffsetType", PartOccInfoCol)
    Dim dOffsetValue As Double
    dOffsetValue = GetAttribute(OccAttrs, "SPSHRPathHorizontalOffset", PartOccInfoCol)
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
    BeginExtensionLength = GetAttribute(OccAttrs, "SPSHRBeginExtensionLength", PartOccInfoCol)
    EndExtensionLength = GetAttribute(OccAttrs, "SPSHREndExtensionLength", PartOccInfoCol)
    
    'Post Section
    Dim PostSection As String
    Dim PostStandard As String
    Dim PostSectionCP       As Integer
    Dim PostSectionAngle    As Double
    Dim PostSectionWidth As Double
    Dim PostSectionDepth As Double
    PostSection = GetAttribute(OccAttrs, "Post_SPSSectionName", PartOccInfoCol)
    PostStandard = GetAttribute(OccAttrs, "Post_SPSSectionRefStandard", PartOccInfoCol)
    PostSectionCP = GetAttribute(OccAttrs, "PostSectionCP", PartOccInfoCol)
    PostSectionAngle = GetAttribute(OccAttrs, "PostSectionAngle", PartOccInfoCol)
    GetCrossSecData PostSection, PostStandard, PostSectionWidth, PostSectionDepth 'get post section details
    
    'With Post at Turn
    Dim WithPostAtTurn As Boolean
    WithPostAtTurn = GetAttribute(OccAttrs, "WithPostAtTurn", PartOccInfoCol)
    
    'Treatment type
    Dim SPSHRBeginTreatment As Integer
    Dim SPSHREndTreatment As Integer
    SPSHRBeginTreatment = GetAttribute(OccAttrs, "SPSHRBeginTreatment", PartOccInfoCol)
    SPSHREndTreatment = GetAttribute(OccAttrs, "SPSHREndTreatment", PartOccInfoCol)
    
    'Extra fields compare to TypeA
    Dim PadOffset As Double
    PadOffset = GetAttribute(OccAttrs, "SPSHRSideMountConnVertOffset", PartOccInfoCol)
    PadOffset = Abs(PadOffset) * -1
    
    Dim MinClearenceAtPostTurn As Double
    Dim MaxClearenceAtPostTurn  As Double
    MinClearenceAtPostTurn = GetAttribute(OccAttrs, "SPSHRMinClearAtPostTurn", PartOccInfoCol)
    MaxClearenceAtPostTurn = GetAttribute(OccAttrs, "SPSHRMaxClearAtPostTurn", PartOccInfoCol)
       
    Set OccAttrs = Nothing
          
    CreatePosts_Ex Nothing, pSegments, Height - htdelta1, TopOfMidRailDim - MidRailSectionDepth / 2, Orientation, _
                Nothing, Nothing, SegmentMaxSpacing, SlopedSegmentMaxSpacing, BeginExtensionLength, EndExtensionLength, _
                PostSection, PostStandard, WithPostAtTurn, PostSectionCP, PostSectionAngle, _
                SPSHRBeginTreatment, SPSHREndTreatment, PadOffset, MinClearenceAtPostTurn, MaxClearenceAtPostTurn, _
                PostSectionDepth, MidRailSectionDepth, False, True, volume, VCogX, VCogY, VCogZ
    
    
Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL

End Sub
'virtually identical to a somilarly named method in
'ComputeStuff.bas, but side mounted to member handrails must be offset twice.
Private Sub CalcRailsVolumeCGForSideMount(oHandrail As ISPSHandrail, _
                      ByVal PartOccInfoCol As IJDInfosCol, _
                      ByRef volume As Double, _
                      ByRef VCogX As Double, ByRef VCogY As Double, ByRef VCogZ As Double)
    Const METHOD = "CalcRailsVolumeCG"
    On Error GoTo ErrHandler
    
    Dim i As Integer
    
    Dim OccAttrs As IJDAttributes
    
        Dim Length  As Double
    Dim toeplateht As Double
    Dim railwt As Double, totalVolume As Double
    Dim ToprailSec As String, TopRailStandard As String
    Dim MidrailSec As String, MidRailStandard As String
    Dim ToePlateSec As String, ToePlateStandard As String
    Dim PostSec As String, PostStandard As String
    Dim noofmidrails As Integer, midrailno As Integer
    Dim material As Variant, grade As Variant
    Dim TopOfMidRailDim As Double
    Dim MidRailSpacing As Double
    Dim SPSHRBeginTreatment As Integer
    Dim SPSHREndTreatment As Integer
    Dim treatmentPostLength As Double
    Dim Height As Double
    Dim outPutString As String
    Dim PostCSArea As Double
    Dim CSArea As Double
    Dim totalCOG As DPosition
    Dim partCOG As DPosition
    Dim HRpathCurve As ComplexString3d
    Dim HRpath As Sketch3d
    Dim Orientation As Integer
    Dim WithToePlate As Boolean
    
    Set partCOG = New DPosition
    Set totalCOG = New DPosition
    
    Set OccAttrs = oHandrail
     
    Set HRpath = oHandrail.SketchPath
    Dim strOffsetTypeAttrName As String
    Dim strOffsetValAttrName As String
    strOffsetTypeAttrName = "SPSHRPathHorizontalOffsetType"
    strOffsetValAttrName = "SPSHRPathHorizontalOffset"
   
    Dim iOffsetType As Integer
    iOffsetType = GetAttribute(OccAttrs, strOffsetTypeAttrName, PartOccInfoCol)
    Dim dOffsetValue As Double
    dOffsetValue = GetAttribute(OccAttrs, strOffsetValAttrName, PartOccInfoCol)
    
    ' first, offset based on reference
    Dim DeltaHOffset As Double
    Dim TempHOffset As Integer
    DeltaHOffset = 0#
    TempHOffset = 0
    Dim OffsetReference As Long
    OffsetReference = GetAttribute(OccAttrs, "OffsetReference", PartOccInfoCol)
    Dim FloorThickness  As Double
    FloorThickness = GetAttribute(OccAttrs, "FloorThickness", PartOccInfoCol)
    Dim ColumnClearance  As Double
    ColumnClearance = GetAttribute(OccAttrs, "ColumnClearance", PartOccInfoCol)
    Dim EndClearance  As Double
    EndClearance = GetAttribute(OccAttrs, "EndClearance", PartOccInfoCol)
    Dim StartColumnOffset  As Double
    StartColumnOffset = GetAttribute(OccAttrs, "StartColumnOffset", PartOccInfoCol)
    Dim EndColumnOffset  As Double
    EndColumnOffset = GetAttribute(OccAttrs, "EndColumnOffset", PartOccInfoCol)
    
    Dim tmpComplexString As ComplexString3d
    createNewSketch3D HRpath, OffsetReference, iOffsetType, FloorThickness, ColumnClearance, EndClearance, StartColumnOffset, EndColumnOffset, DeltaHOffset, TempHOffset, tmpComplexString
    
    Set HRpathCurve = GetOffsetCurve(tmpComplexString, _
                               TempHOffset, _
                               DeltaHOffset)
    
    Set tmpComplexString = HRpathCurve
    Set HRpathCurve = Nothing
    
    Set HRpathCurve = GetOffsetCurve(tmpComplexString, _
                               iOffsetType, _
                               dOffsetValue)
    Set HRpath = Nothing
    Orientation = GetAttribute(OccAttrs, "HandrailOrientation", PartOccInfoCol) 'TR#55627
    
    'The length attribute is not currently being correctly set by the handrail semantic
    'It simply uses the length of the sketch path which may not be the same as the actual length
    'Planning has not decided yet if this should be fixed for the current handrail symbols
    'or if it is simply a requirement of the pending "Traffic Items as Parts" CR
'    If HandrailType = TypeA Then
'        Length = GetAttribute(OccAttrs, "TotalLength", PartOccInfoCol)
'    Else
'        Length = GetAttribute(OccAttrs, "Length", PartOccInfoCol)
'    End If
    Dim oOffsetpathCurve As IJCurve
    Set oOffsetpathCurve = HRpathCurve
        Length = oOffsetpathCurve.Length
    
    toeplateht = GetAttribute(OccAttrs, "TopOfToePlateDim", PartOccInfoCol)
    
    ToprailSec = GetAttribute(OccAttrs, "TopRail_SPSSectionName", PartOccInfoCol)
    
    TopRailStandard = GetAttribute(OccAttrs, "TopRail_SPSSectionRefStandard", PartOccInfoCol)
    
    MidrailSec = GetAttribute(OccAttrs, "MidRail_SPSSectionName", PartOccInfoCol)
    
    MidRailStandard = GetAttribute(OccAttrs, "MidRail_SPSSectionRefStandard", PartOccInfoCol)
    
    noofmidrails = GetAttribute(OccAttrs, "NoOfMidRails", PartOccInfoCol)
    
    ToePlateSec = GetAttribute(OccAttrs, "ToePlate_SPSSectionName", PartOccInfoCol)
    
    ToePlateStandard = GetAttribute(OccAttrs, "ToePlate_SPSSectionRefStandard", PartOccInfoCol)
   
    TopOfMidRailDim = GetAttribute(OccAttrs, "TopOfMidRailDim", PartOccInfoCol)
        
    MidRailSpacing = GetAttribute(OccAttrs, "MidRailSpacing", PartOccInfoCol)
        
    Height = GetAttribute(OccAttrs, "Height", PartOccInfoCol)
    
    WithToePlate = GetAttribute(OccAttrs, "WithToePlate", PartOccInfoCol)
    
    CSArea = 0#
    totalVolume = 0#
    totalCOG.Set 0#, 0#, 0#
        
    'TopRail
    CSArea = GetCSAttribData(ToprailSec, TopRailStandard, "ISTRUCTCrossSectionDimensions", "Area")
        totalVolume = totalVolume + CSArea * Length
    Set partCOG = GetCOGForTopMidRail(HRpathCurve, Height, Orientation)
        totalCOG.x = totalCOG.x + partCOG.x * CSArea * Length
        totalCOG.y = totalCOG.y + partCOG.y * CSArea * Length
        totalCOG.z = totalCOG.z + partCOG.z * CSArea * Length
    
    'MidRail
    CSArea = GetCSAttribData(MidrailSec, MidRailStandard, "ISTRUCTCrossSectionDimensions", "Area")
        totalVolume = totalVolume + CSArea * Length * noofmidrails
    For i = 1 To noofmidrails
        Set partCOG = GetCOGForTopMidRail(HRpathCurve, (TopOfMidRailDim - (i - 1) * MidRailSpacing), Orientation)
            totalCOG.x = totalCOG.x + partCOG.x * CSArea * Length
            totalCOG.y = totalCOG.y + partCOG.y * CSArea * Length
            totalCOG.z = totalCOG.z + partCOG.z * CSArea * Length
    Next i
    
    'ToePlate
    If WithToePlate Then
        CSArea = GetCSAttribData(ToePlateSec, ToePlateStandard, "ISTRUCTCrossSectionDimensions", "Area")
            totalVolume = totalVolume + CSArea * Length
        
        Set partCOG = GetCOGForTopMidRail(HRpathCurve, toeplateht, Orientation)
            totalCOG.x = totalCOG.x + partCOG.x * CSArea * Length
            totalCOG.y = totalCOG.y + partCOG.y * CSArea * Length
            totalCOG.z = totalCOG.z + partCOG.z * CSArea * Length
    End If
    
    
    volume = totalVolume
    
    VCogX = totalCOG.x
    VCogY = totalCOG.y
    VCogZ = totalCOG.z
    
    ' TR#55627
    Set partCOG = Nothing
    Set totalCOG = Nothing
    Set HRpathCurve = Nothing

Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
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

            
Private Sub SetNewPosition(oCurve As IJCurve, EndIndex As Long, distance As Double)
    Const METHOD = "SetNewPosition"
    On Error GoTo ErrHandler
    Dim dLength As Double, dX As Double, dY As Double, dZ As Double
    Dim oLine As IJLine
    Dim oArc As IJArc
    Dim oEllArc As IJEllipticalArc
        dLength = oCurve.Length
    Select Case EndIndex
    Case 1
        oCurve.PositionFRatio distance / dLength, dX, dY, dZ
        If TypeOf oCurve Is IJLine Then
            Set oLine = oCurve
            oLine.SetStartPoint dX, dY, dZ
        ElseIf TypeOf oCurve Is IJArc Then
            Set oArc = oCurve
            oArc.SetStartPoint dX, dY, dZ
        ElseIf TypeOf oCurve Is IJEllipticalArc Then
            Set oEllArc = oCurve
            oEllArc.SetStartPoint dX, dY, dZ
        End If
    Case 2
        oCurve.PositionFRatio 1 - distance / dLength, dX, dY, dZ
        If TypeOf oCurve Is IJLine Then
            Set oLine = oCurve
            oLine.SetEndPoint dX, dY, dZ
        ElseIf TypeOf oCurve Is IJArc Then
            Set oArc = oCurve
            oArc.SetEndPoint dX, dY, dZ
        ElseIf TypeOf oCurve Is IJEllipticalArc Then
            Set oEllArc = oCurve
            oEllArc.SetEndPoint dX, dY, dZ
        End If
    End Select
    
Exit Sub
ErrHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL

End Sub

Private Sub CreateVerticalOffsetCurve(ByRef pCurve, _
                                ByVal pComplex As ComplexString3d, _
                                Height As Double)
                                
Const METHOD = "CreateProjectedHRCurve"
On Error GoTo ErrorHandler

    Dim oTrans4x4       As IJDT4x4
    Dim oVector         As IJDVector
    Dim oSegments       As IJElements

    If m_GeomFactory Is Nothing Then
        Set m_GeomFactory = New GeometryFactory
    End If
    
    Set oVector = New DVector
    oVector.Set 0, 0, Height
    
    Set oTrans4x4 = New DT4x4
    oTrans4x4.LoadIdentity
    oTrans4x4.Translate oVector

    pComplex.GetCurves oSegments
    
    Set pCurve = New ComplexString3d
    Set pCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, oSegments)
    
    pCurve.Transform oTrans4x4

    Set oSegments = Nothing
    Set oTrans4x4 = Nothing
    Set oVector = Nothing
    
    Exit Sub
    
ErrorHandler:
    Dim oErrors As New IMSErrorLog.JServerErrors
    oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

' This method checks whether the given handrail has supporting member or not.
' If supporting member exists, checks whether the supporting member has related columns or not.
' If the related columns exist for the supporting member, applies the given start/end column offsets to the handrail start/end offset curves.
' Calculates the total length of the handrail start/end offset curves and returns the same.
Private Function GetColumnOffsetCurvesLength(oHandrail As ISPSHandrail, ColumnClearance As Double, EndClearance As Double, StartColumnOffset As Double, EndColumnOffset As Double) As Double
Const METHOD = "GetColumnOffsetCurvesLength"
On Error GoTo ErrorHandler

Dim oAssocHelper As IJPointAssocConstraintHelper
Dim sketchPathStartX As Double, sketchPathStartY As Double, sketchPathStartZ As Double
Dim sketchPathMidX As Double, sketchPathMidY As Double, sketchPathMidZ As Double
Dim sketchPathEndX As Double, sketchPathEndY As Double, sketchPathEndZ As Double
Dim AssocPointX As Double, AssocPointY As Double, AssocPointZ As Double
Dim dStartOffsetCurveDistance As Double, dEndOffsetCurveDistance As Double
Dim oRangeValues(2) As Double
Dim strPointOnGeometryProgID As String
Dim dDistance As Double, dParameter As Double, offsetCurveDistance(2) As Double
Dim status As Integer, i As Integer
Dim lParameter As Long
Dim bParameter As Boolean
Dim oParents As IJElements
Dim oMembPartPrism As ISPSMemberPartPrismatic
Dim oCachedPoint As CachedPoint
Dim oPoint As IJPoint
Dim cacheTurnPoints As IJElements
Dim handrailSketchPathStartCurve As IJCurve
Dim handrailSketchPathEndCurve As IJCurve
Dim handrailSketchPath As IJDSketch3d
Set handrailSketchPath = oHandrail.SketchPath
handrailSketchPath.cacheTurnPoints cacheTurnPoints
Set oAssocHelper = New PointAssocConstraintHelper

Dim handrailComplexString As ComplexString3d
Set handrailComplexString = handrailSketchPath.GetComplexString
Dim handrailSketchPathCurves  As IJElements
handrailComplexString.GetCurves handrailSketchPathCurves

If handrailSketchPathCurves.Count > 0 Then
    Set handrailSketchPathStartCurve = handrailSketchPathCurves(1)
    handrailSketchPathStartCurve.EndPoints sketchPathStartX, sketchPathStartY, sketchPathStartZ, sketchPathEndX, sketchPathEndY, sketchPathEndZ
    If handrailSketchPathCurves.Count > 1 Then
        Set handrailSketchPathEndCurve = handrailSketchPathCurves(handrailSketchPathCurves.Count)
        handrailSketchPathEndCurve.EndPoints sketchPathMidX, sketchPathMidY, sketchPathMidZ, sketchPathEndX, sketchPathEndY, sketchPathEndZ
    End If
End If
    'Get The information on supporting member's  related columns, and make end offset later
    If Not cacheTurnPoints Is Nothing Then
        offsetCurveDistance(1) = EndClearance
        offsetCurveDistance(2) = EndClearance
        Dim lPointIndex As Long
        For lPointIndex = 1 To cacheTurnPoints.Count
            Set oCachedPoint = cacheTurnPoints.Item(lPointIndex)
            If oCachedPoint.IsAssocPoint Then
                oCachedPoint.GetAssocPoint oPoint
                oPoint.GetPoint AssocPointX, AssocPointY, AssocPointZ
                dStartOffsetCurveDistance = (AssocPointX - sketchPathStartX) * (AssocPointX - sketchPathStartX) + (AssocPointY - sketchPathStartY) * (AssocPointY - sketchPathStartY) + (AssocPointZ - sketchPathStartZ) * (AssocPointZ - sketchPathStartZ)
                dEndOffsetCurveDistance = (AssocPointX - sketchPathEndX) * (AssocPointX - sketchPathEndX) + (AssocPointY - sketchPathEndY) * (AssocPointY - sketchPathEndY) + (AssocPointZ - sketchPathEndZ) * (AssocPointZ - sketchPathEndZ)
                oAssocHelper.GetInfo oPoint, strPointOnGeometryProgID, oParents, dDistance, bParameter, dParameter, lParameter
                If Not oParents Is Nothing Then
                    For i = 1 To 2
                        Set oMembPartPrism = Nothing
                        Set oMembPartPrism = GetMembAssociatedWithConn(oParents, oPoint, i) 'To get the supported member to this frame / split connection
                        If Not oMembPartPrism Is Nothing Then Exit For
                    Next
                    status = getColumnRangeValues(oParents, oMembPartPrism, oPoint, i, oRangeValues) 'get Rangebox values from the supporting member
                    'If the supporting member has columns realted to it, then applying the start/end column offsets while calculating start and end curve distances
                    If status <> 0 Then
                        If dStartOffsetCurveDistance < dtol Then
                            If StartColumnOffset = 0 Then
                                StartColumnOffset = (oRangeValues(1) / 2)
                            End If
                            offsetCurveDistance(1) = ColumnClearance + StartColumnOffset
                        ElseIf dEndOffsetCurveDistance < dtol Then
                            If EndColumnOffset = 0 Then
                                EndColumnOffset = (oRangeValues(1) / 2)
                            End If
                            offsetCurveDistance(2) = ColumnClearance + EndColumnOffset
                        End If
                    End If
                End If
                Set oPoint = Nothing
                Set oParents = Nothing
            End If
        Next
    End If
    
    Dim OffsetCurvesTotalLength As Double
    If Abs(offsetCurveDistance(1)) > dtol Then
        OffsetCurvesTotalLength = OffsetCurvesTotalLength + offsetCurveDistance(1)
    End If
    If Abs(offsetCurveDistance(2)) > dtol Then
        OffsetCurvesTotalLength = OffsetCurvesTotalLength + offsetCurveDistance(2)
    End If
    GetColumnOffsetCurvesLength = OffsetCurvesTotalLength
    
    'GoTo cleanup:
        ReleaseOpenObjects cacheTurnPoints
        Set oAssocHelper = Nothing
        Set oParents = Nothing
        Set oMembPartPrism = Nothing
        Set oCachedPoint = Nothing
        Set oPoint = Nothing
        Set cacheTurnPoints = Nothing
        Set handrailSketchPathStartCurve = Nothing
        Set handrailSketchPathEndCurve = Nothing
        Set handrailSketchPath = Nothing
        Set handrailComplexString = Nothing
        Set handrailSketchPathCurves = Nothing
        
    Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
