VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StairTypeA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'******************************************************************
' Copyright (C) 2007, Intergraph Corporation. All rights reserved.
'
'File
'    InclinedLadder.cls
'
'Author
'       January 2002        AP   V3 Created
'
'Description
'
'Notes
'
'History:
' 10 Sept 2004 MD Additional argument "SkinOption" added to CreateProjectionFromCSProfile() calls
' Which be passed to CreateSurface for for appropriate creation of surface as per requirement through
' CreateBySingleSweep() function. For details see Documentaion of CreateBySingleSweep & CreateSurface
'  20-Sep-05 JS    TR#84847 - Shortened the name of the interface IJSPSCommonStairLadderProperties
'                           to IJSPSCommonStairLadderProps (Oracle requirement)
' 26 Sep 2006 AS    DM#106468 - Corrected Operational envelope
' Dec-13-2006 AS TR#111679 - Corrected error reporting in evaluate geometry
' May-05-2007 AS CR#116177 Added custom Mirror copy implementation through ISPSMirrorHelper I/F
' Jun-07-2007 AS TR#120502 Added code to check the values of codelisted atributes
' Aug-10-2007 AS DI -CP·123592. Removed flip condition for gridlines as they are reversed now
' Sept-04-2008 GG DM -CP·149875. Corrected top step pitch and bottom step pitch
' Sept-04-2008 GG removed obsolete code in ISPSMirrorHelper_Mirror to avoid further confusion (TR#148191),
'                 and corrected constant string METHOD for ISPSMirrorHelper_Mirror
' Oct-03-2008 GG DI #150310 removed DetailPhysicalView and added CenterlineView mode, implemented CenterlineView to include side frames, steps and top landing platform
' Dec-09-2008 GG TR-120786    - Don't change the support side if the target top-edge is slab
'
' Jul-29-2009   WR  TR#168745   - Added logic to validate angle in UserAttributeMgmt_Validate since there is a
'                               different logic in SPSOccurenceHelper in the CheckOneAttribute method to validate
'                               both .NET and VB6 symbols.
' Nov-18-2009 GG DM-174662    - Corrected step position problem introduced in DI #150310
'*******************************************************************

Const defProgId As String = "SPSStairMacros.StairTypeA"
Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"
Private Const MODULE = "SPSStairMacros.StairTypeA::"
Const PI As Double = 3.14159265358979
Private Const TOL = 0.0000001

Private m_oErrors As IJEditErrors

Private skinOption As Long
Private m_GeomFactory As IngrGeom3D.GeometryFactory
Private m_oLocalizer As IJLocalizer

Private Enum InputIndex
PART_INDEX = 1
WIDTH_INDEX
ANGLE_INDEX
STEPPITCH_INDEX
HEIGHT_INDEX
NUMSTEPS_INDEX
SPAN_INDEX
LENGTH_INDEX
JUSTIFICATION_INDEX
TOPSUPPORTSIDE_INDEX
SIDEFRAMESECTION_INDEX
SIDEFRAMESECTIONSTANDARD_INDEX
HANDRAILSECTION_INDEX
HANDRAILREFERENCESTANDARD_INDEX
STEPSECTION_INDEX
STEPREFERENCESTANDARD_INDEX
SIDEFRAMESECTIONCP_INDEX
SIDEFRAMESECTIONANGLE_INDEX
HANDRAILSECTIONCP_INDEX
HANDRAILSECTIONANGLE_INDEX
STEPSECTIONCP_INDEX
STEPSECTIONANGLE_INDEX
PRIMARYMATERIAL_INDEX
PRIMARYGRADE_INDEX
PLATFORMTHICKNESS_INDEX
WITHTOPLANDING_INDEX
TOPLANDINGLENGTH_INDEX
POSTHEIGHT_INDEX
HANDRAILPOSTPITCH_INDEX
NUMMIDRAILS_INDEX
ISASSEMBLY_INDEX
ISSYSTEM_INDEX
ENVELOPEHEIGHT_INDEX
End Enum

Implements IJDUserSymbolServices
Implements ISPSStairSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck
Implements ISPSMirrorHelper

Dim bOnPreLoad As Boolean

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
    Set m_oLocalizer = New IMSLocalizer.Localizer
    m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
    Set m_GeomFactory = Nothing
    Set m_oErrors = Nothing
    Set m_oLocalizer = Nothing
End Sub

Private Sub ISPSStairSymbolServices_EvaluateGeometry(ByVal StairBO As SPSStairs.ISPSStair, ByVal bPartOccChanged As Boolean, _
                                                ByVal Top As Object, ByVal bTopChanged As Boolean, _
                                                ByVal bottom As Object, ByVal bBottomChanged As Boolean, _
                                                ByVal Ref As Object, ByVal bRefChanged As Boolean, _
                                                ByVal PartInfoCol As SPSStairs.IJDInfosCol, ByVal bAttributeChanged As Boolean, _
                                                ByVal PartOccInfoCol As SPSStairs.IJDInfosCol, ByVal bOccAttributeChanged As Boolean)
    
Const METHOD = "ISPSStairSymbolServices_EvaluateGeometry"
On Error GoTo ErrorHandler
    
    Dim Angle As Double
    Dim Span As Double
    Dim Height As Double
    Dim OldHeight As Double
    Dim Length As Double
    Dim StepPitch As Double
    Dim NumSteps As Integer
    Dim tempNum As Double
    Dim oPart As IJDPart
    Dim oPartOcc As IJPartOcc
    Dim Matrix As DT4x4
    Dim lJustify As Long
    Dim lSideFrameSectionCP As Long
    Dim lHandRailSectionCP As Long
    Dim lStepSectionCP As Long
    Dim StepSectionName As String
    Dim StepStandardName As String
    Dim StepThickness As Double
    
    Set oPartOcc = StairBO
    oPartOcc.GetPart oPart

    Dim OccAttrs As IJDAttributes
    Dim PartAttrs As IJDAttributes
    Set OccAttrs = StairBO
    Set PartAttrs = oPart
    'now Length is what we called height previously
 
    OldHeight = GetAttribute(OccAttrs, "Height", PartOccInfoCol)
    
    'changes for checking undefined value
    lJustify = GetAttribute(OccAttrs, "Justification", PartOccInfoCol)
    lSideFrameSectionCP = GetAttribute(OccAttrs, "SideFrameSectionCP", PartOccInfoCol)
    lHandRailSectionCP = GetAttribute(OccAttrs, "HandRailSectionCP", PartOccInfoCol)
    lStepSectionCP = GetAttribute(OccAttrs, "StepSectionCP", PartOccInfoCol)
    StepSectionName = GetAttribute(OccAttrs, "Step_SPSSectionName", PartOccInfoCol)
    StepStandardName = GetAttribute(OccAttrs, "Step_SPSSectionRefStandard", PartOccInfoCol)
    StepThickness = GetCSAttribData(StepSectionName, StepStandardName, "ISTRUCTCrossSectionDimensions", "Width")
    
    CheckForUndefinedValueAndRaiseError StairBO, lJustify, "StructAlignment", 71
    CheckForUndefinedValueAndRaiseError StairBO, lSideFrameSectionCP, "CrossSectionCardinalPoints", 81
    CheckForUndefinedValueAndRaiseError StairBO, lHandRailSectionCP, "CrossSectionCardinalPoints", 82
    CheckForUndefinedValueAndRaiseError StairBO, lStepSectionCP, "CrossSectionCardinalPoints", 83

    Dim SetOriginAndSpanValue As Long
    SetOriginAndSpanValue = SetOriginAndSpan(StairBO, Top, bottom, Ref, PartOccInfoCol)
        
    If SetOriginAndSpanValue = E_FAIL Then
        Err.Raise E_FAIL
    ElseIf SetOriginAndSpanValue = STAIR_E_SIDE_REF_MISSING Then
        Err.Raise STAIR_E_SIDE_REF_MISSING
    End If

    Height = GetAttribute(OccAttrs, "Height", PartOccInfoCol)
 
    If Not OldHeight = Height Or bPartOccChanged Or bOccAttributeChanged Or bTopChanged Or bBottomChanged Or bRefChanged Then

        Angle = GetAttribute(OccAttrs, "Angle", PartOccInfoCol)
        Length = Height / Sin(Angle)
        StepPitch = GetAttribute(OccAttrs, "StepPitch", PartOccInfoCol)

        Dim ActualPitch As Double
        Dim bAttrExists As Boolean
        bAttrExists = False
        ActualPitch = GetAttribute(OccAttrs, "ActualPitch", PartOccInfoCol, bAttrExists)
        If bAttrExists Then
            tempNum = Height / StepPitch
            If Abs(tempNum - Int(tempNum)) < TOL Then
                NumSteps = Int(tempNum) - 1
            Else
                NumSteps = Int((tempNum))
            End If
            ActualPitch = Height / (NumSteps + 1)
            Call SetAttribute(OccAttrs, ActualPitch, "ActualPitch", PartOccInfoCol)
        Else
            'take step-thickness into the calculation
            tempNum = (Height - StepThickness) / StepPitch
            NumSteps = Int((tempNum))
        End If
        Call SetAttribute(OccAttrs, NumSteps, "NumSteps", PartOccInfoCol)
        Call SetAttribute(OccAttrs, Length, "Length", PartOccInfoCol)

    End If
    Set oPart = Nothing
    Set oPartOcc = Nothing

    Exit Sub
    
ErrorHandler:
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    ElseIf Err.Number = STAIR_E_SIDE_REF_MISSING Then
        Err.Raise STAIR_E_SIDE_REF_MISSING
    Else
        m_oErrors.Add Err.Number, METHOD, Err.Description
    End If
End Sub

Private Sub ISPSStairSymbolServices_EvaluateLocalWCG(ByVal oStair As SPSStairs.ISPSStair, _
                                                ByVal PartInfoCol As SPSStairs.IJDInfosCol, _
                                                ByRef weight As Double, _
                                                ByRef COGX As Double, ByRef COGY As Double, ByRef COGZ As Double)
Const METHOD = "ISPSStairSymbolServices_EvaluateLocalWCG"
On Error GoTo ErrorHandler

    Call CalcWCG(oStair, PartInfoCol, weight, COGX, COGY, COGZ)

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
Const METHOD = "IJDUserSymbolServices_EditOccurence"

  IJDUserSymbolServices_EditOccurence = False
  
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
Const METHOD = "IJDUserSymbolServices_GetDefinitionName"
  
  IJDUserSymbolServices_GetDefinitionName = defProgId

End Function

'********************************************************************
' Function: IJDUserSymbolServices_InstanciateDefinition()
'
' Description:This instanciates a persistent symbol definition object
' and initializes it for the first time.
'********************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParameters As Variant, ByVal ActiveConnection As Object) As Object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler
    
    'Create a Symbol Definition Object.
    Dim pSymbolFactory As New DSymbolEntitiesFactory
    Dim pSymbolDefinition As IJDSymbolDefinition
    Set pSymbolDefinition = pSymbolFactory.CreateEntity(definition, ActiveConnection)

    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
   
    ' nomenclature
    pSymbolDefinition.ProgId = defProgId
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.name = pSymbolDefinition.ProgId
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED

    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition
    
    Set pSymbolFactory = Nothing
    Set pSymbolDefinition = Nothing

    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
Const METHOD = "IJDUserSymbolServices_InvokeRepresentation"

End Sub

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler

    'Remove all previous Symbol Definition information
    
    pSymbolDefinition.IJDInputs.RemoveAllInput
    pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
    pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
    
    Dim InputsIf As IMSSymbolEntities.IJDInputs
    Set InputsIf = pSymbolDefinition
    
    Dim pUM As IMSSymbolEntities.IJDUserMethods
    Set pUM = pSymbolDefinition

    Dim libDesc As New DLibraryDescription
    Dim mCookie As Long
    Dim libCookie As Long
    Dim DegCheck As Long
    Dim mthCookieCheckSection As Long
    Dim CheckMaterial As Long
    
    libDesc.name = "mySelfAsLib"
    libDesc.Type = imsLIBRARY_IS_ACTIVEX
    libDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    libDesc.Source = defProgId

    pSymbolDefinition.IJDUserMethods.SetLibrary libDesc
    libCookie = libDesc.Cookie
    mCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheForPart", libCookie)

    Dim ChecklibDesc As New DLibraryDescription
    Dim ChecklibCookie As Long
    Dim GTZeroCheck As Long
    Dim LT90DegCheck As Long
'    Dim IsBoolean As Long
    ChecklibDesc.name = "CMCheckLib"
    ChecklibDesc.Type = imsLIBRARY_IS_ACTIVEX
    ChecklibDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    ChecklibDesc.Source = CheckProgId
    pSymbolDefinition.IJDUserMethods.SetLibrary ChecklibDesc
    ChecklibCookie = ChecklibDesc.Cookie

    'Get the lib/method cookie
    GTZeroCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("GTZero", ChecklibCookie)
    LT90DegCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("LT90Deg", ChecklibCookie)
    DegCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("IN360Deg", ChecklibCookie)
    mthCookieCheckSection = pSymbolDefinition.IJDUserMethods.GetMethodCookie("IsCrossSectionName", ChecklibCookie)
    CheckMaterial = pSymbolDefinition.IJDUserMethods.GetMethodCookie("IsMaterailGrade", ChecklibCookie)

'    IsBoolean = pSymbolDefinition.IJDUserMethods.GetMethodCookie("IsBoolean", ChecklibCookie)

    ' Define a new input by new operator
     Dim pIJDInput As IMSSymbolEntities.IJDInput
     Set pIJDInput = New IMSSymbolEntities.DInput

    ' Create a default value
     Dim PC As IMSSymbolEntities.IJDParameterContent
     Set PC = New IMSSymbolEntities.DParameterContent
     
     PC.Type = igValue
    
     Dim oSymbolCache As New CustomCache
     oSymbolCache.SetupCustomCache pSymbolDefinition

     pIJDInput.name = "Width"
     pIJDInput.Description = "Width"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.6
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, WIDTH_INDEX
     pIJDInput.Reset
     PC.Reset
          
     pIJDInput.name = "Angle"
     pIJDInput.Description = "Angle"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, LT90DegCheck
     PC.UomValue = 55#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, ANGLE_INDEX
     pIJDInput.Reset
     PC.Reset
       
     pIJDInput.name = "StepPitch"
     pIJDInput.Description = "StepPitch"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.25
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, STEPPITCH_INDEX
     pIJDInput.Reset
     PC.Reset
       
     pIJDInput.name = "Height"
     pIJDInput.Description = "Height"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HEIGHT_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "NumSteps"
     pIJDInput.Description = "NumSteps"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, NUMSTEPS_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.name = "Span"
     pIJDInput.Description = "Span"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SPAN_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "Length"
     pIJDInput.Description = "Length"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, LENGTH_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.name = "Justification"
     pIJDInput.Description = "Justification"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, JUSTIFICATION_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "TopSupportSide"
     pIJDInput.Description = "TopSupportSide"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, TOPSUPPORTSIDE_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "SideFrame_SPSSectionName"
     pIJDInput.Description = "SideFrame_SPSSectionName"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "C3x4.1"
     pIJDInput.DefaultParameterValue = PC
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
     InputsIf.SetInput pIJDInput, SIDEFRAMESECTION_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "SideFrame_SPSSectionRefStandard"
     pIJDInput.Description = "SideFrame_SPSSectionRefStandard"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "AISC-LRFD-3.1"
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SIDEFRAMESECTIONSTANDARD_INDEX
     pIJDInput.Reset
     PC.Reset
        
     pIJDInput.name = "HandRail_SPSSectionName"
     pIJDInput.Description = "HandRail_SPSSectionName"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "HSS1-11/16x.140"
     pIJDInput.DefaultParameterValue = PC
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
     InputsIf.SetInput pIJDInput, HANDRAILSECTION_INDEX
     pIJDInput.Reset
     PC.Reset
          
     pIJDInput.name = "HandRail_SPSSectionRefStandard"
     pIJDInput.Description = "HandRail_SPSSectionRefStandard"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "AISC-LRFD-3.1"
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HANDRAILREFERENCESTANDARD_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "Step_SPSSectionName"
     pIJDInput.Description = "Step_SPSSectionName"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "RS2x4"
     pIJDInput.DefaultParameterValue = PC
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, mthCookieCheckSection
     InputsIf.SetInput pIJDInput, STEPSECTION_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "Step_SPSSectionRefStandard"
     pIJDInput.Description = "Step_SPSSectionRefStandard"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "AISC-LRFD-3.1"
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, STEPREFERENCESTANDARD_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "SideFrameSectionCP"
     pIJDInput.Description = "SideFrameSectionCP"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 4
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SIDEFRAMESECTIONCP_INDEX
     pIJDInput.Reset
     PC.Reset
           
     pIJDInput.name = "SideFrameSectionAngle"
     pIJDInput.Description = "SideFrameSectionAngle"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
     InputsIf.SetInput pIJDInput, SIDEFRAMESECTIONANGLE_INDEX
     pIJDInput.Reset
     PC.Reset
         
     pIJDInput.name = "HandRailSectionCP"
     pIJDInput.Description = "HandRailSectionCP"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 5#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HANDRAILSECTIONCP_INDEX
     pIJDInput.Reset
     PC.Reset
         
     pIJDInput.name = "HandRailSectionAngle"
     pIJDInput.Description = "HandRailSectionAngle"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
     InputsIf.SetInput pIJDInput, HANDRAILSECTIONANGLE_INDEX
     pIJDInput.Reset
     PC.Reset
         
     pIJDInput.name = "StepSectionCP"
     pIJDInput.Description = "StepSectionCP"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 5#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, STEPSECTIONCP_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "StepSectionAngle"
     pIJDInput.Description = "StepSectionAngle"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, DegCheck
     InputsIf.SetInput pIJDInput, STEPSECTIONANGLE_INDEX
     pIJDInput.Reset
     PC.Reset
       
     pIJDInput.name = "Primary_SPSMaterial"
     pIJDInput.Description = "Primary_SPSMaterial"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Concrete"
     pIJDInput.DefaultParameterValue = PC
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, CheckMaterial
     InputsIf.SetInput pIJDInput, PRIMARYMATERIAL_INDEX
     pIJDInput.Reset
     PC.Reset
        
     pIJDInput.name = "Primary_SPSGrade"
     pIJDInput.Description = "Primary_SPSGrade"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Fc 4000"
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, PRIMARYGRADE_INDEX
     pIJDInput.Reset
     PC.Reset
       
     pIJDInput.name = "PlatformThickness"
     pIJDInput.Description = "PlatformThickness"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 0.02
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, PLATFORMTHICKNESS_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "WithTopLanding"
     pIJDInput.Description = "WithTopLanding"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, WITHTOPLANDING_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.name = "TopLandingLength"
     pIJDInput.Description = "TopLandingLength"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.Type = igValue
     PC.UomValue = 0.6
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, TOPLANDINGLENGTH_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "PostHeight"
     pIJDInput.Description = "PostHeight"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.Type = igValue
     PC.UomValue = 1.2
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, POSTHEIGHT_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "HandRailPostPitch"
     pIJDInput.Description = "HandRailPostPitch"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.Type = igValue
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HANDRAILPOSTPITCH_INDEX
     pIJDInput.Reset
     PC.Reset
          
     pIJDInput.name = "NumMidRails"
     pIJDInput.Description = "NumMidRails"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
'     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.Type = igValue
     PC.UomValue = 3#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, NUMMIDRAILS_INDEX
     pIJDInput.Reset
     PC.Reset
                  
     pIJDInput.name = "IsAssembly"
     pIJDInput.Description = "IsAssembly"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, ISASSEMBLY_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "IsSystem"
     pIJDInput.Description = "IsSystem"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, ISSYSTEM_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "EnvelopeHeight"
     pIJDInput.Description = "EnvelopeHeight"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, ENVELOPEHEIGHT_INDEX
     pIJDInput.Reset
     PC.Reset

 'NEW REP
    Dim pIReps As IMSSymbolEntities.IJDRepresentations
    Set pIReps = pSymbolDefinition
    Dim pIRep As IMSSymbolEntities.IJDRepresentation
    Set pIRep = New IMSSymbolEntities.DRepresentation
    
    pIRep.name = "Physical"
    pIRep.Description = "Physical representation"
    pIRep.Properties = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    pIRep.RepresentationId = SimplePhysical
    mCookie = pUM.GetMethodCookie("Physical", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
       
    Dim pOutputs As IMSSymbolEntities.IJDOutputs
    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    
    Dim output As IMSSymbolEntities.DOutput
    Set output = New IMSSymbolEntities.DOutput
    
    output.name = "LeftSideFrame"
    output.Description = "Left side frame element"
    pOutputs.SetOutput output
    output.Reset
    
    output.name = "RightSideFrame"
    output.Description = "Right side frame element"
    pOutputs.SetOutput output
    output.Reset
    
    pIReps.SetRepresentation pIRep 'Add representation to definition

     'All steps will be generated as variable output arguments
     
    'For Operational Representation.
    pOutputs.RemoveAllOutput
    Set output = New DOutput

    pIRep.name = "OperationalSub"
    pIRep.Description = "Operational representation"
    pIRep.RepresentationId = Operation
    mCookie = pUM.GetMethodCookie("OperationalSub", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie

    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    output.name = "OperationalEnvelope1"
    output.Description = "Operational Envelope of the Stair"
    pOutputs.SetOutput output
    output.Reset

    pIReps.SetRepresentation pIRep 'Add representation to definition
    
    
    ' Centerline Representation
    pOutputs.RemoveAllOutput
    Set output = New DOutput

    pIRep.name = "Centerline"
    pIRep.Description = "Center Line"
    pIRep.Properties = igREPRESENTATION_ISVBFUNCTION
    pIRep.RepresentationId = 512 'need to be replaced with enum for Centerline
    mCookie = pUM.GetMethodCookie("CenterLine", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie

    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    pIReps.SetRepresentation pIRep 'Add Centerline representation to definition
    
'End For Operational Representation.
    
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
    
    'as this symbol def has declared a graphic object as input
    ' GeomOption option will be set to igSYMBOL_GEOM_FIX_TO_ID by the symbol machinerary
    'Because of this the  outputs will be transformed during MDR and the Symbol geometry will
    ' end up in an incorrect location. So resetting the flag - DI226263
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE
    
    Set pUM = Nothing
    Set pIJDInput = Nothing
    Set PC = Nothing
    Set pIRep = Nothing
    Set output = Nothing

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'=========================================================================
'CREATION OF SIMPLE PHYSICAL REPRESENTATION OF Stair
'========================================================================
Public Sub Physical(pIRepSCM As IJDRepresentationStdCustomMethod)
    
    Call PhysicalRepresentation(pIRepSCM, SimpleRep)

End Sub

'=========================================================================
'CREATION OF CENTERLINE REPRESENTATION OF Stair
'========================================================================
Public Sub CenterLine(pIRepSCM As IJDRepresentationStdCustomMethod)
    
    Call PhysicalRepresentation(pIRepSCM, CenterLineRep)
    
End Sub

'Public Sub Physical(pIRepSCM As IJDRepresentationStdCustomMethod)

Private Sub PhysicalRepresentation(pIRepSCM As IJDRepresentationStdCustomMethod, RepType As eRepresentationType)
Const METHOD = "PhysicalRepresentation"
On Error GoTo ErrorHandler

  Dim pRepDG As IJDRepresentationDuringGame
  Dim pOC As IJDOutputCollection
  Dim pInputs As IJDInputs
  Dim i As Integer
  Set pRepDG = pIRepSCM
  Set pOC = pRepDG.outputCollection
  Set pInputs = pRepDG.definition.IJDInputs
    
  Set m_GeomFactory = New IngrGeom3D.GeometryFactory
        
  'Get the inputs
  Dim Span#, Height#, Pitch#, Width#, Justify#
  Dim Angle As Double, Length As Double
  Dim NumSteps As Integer
  Dim SideFrameSection As String, SideFrameReferenceStandard As String, HandRailSection As String, HandRailReferenceStandard As String
  Dim StepSection As String, StepReferenceStandard As String, SideFrameSectionCP As Integer, SideFrameSectionAngle As Double
  Dim HandRailSectionCP As Integer, HandRailSectionAngle As Double, StepSectionCP As Integer, StepSectionAngle As Double
  Dim WithTopLanding As Boolean
  Dim TopLandingLength As Double
  Dim PlatformThickness As Double
  Dim PostHt As Double, HandrailPostPitch As Double, HandRailSpacing As Double
  Dim NumMidRails As Integer
  Dim bNeedToRaiseWarning As Boolean

  skinOption = 7 ' see CreateProjectionFromCSProfile() for more details of skin option
        
  Height = pInputs.GetInputByIndex(HEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
  Length = pInputs.GetInputByIndex(LENGTH_INDEX).IJDInputDuringGame.Result.UomValue

  Span = pInputs.GetInputByIndex(SPAN_INDEX).IJDInputDuringGame.Result.UomValue
  Pitch = pInputs.GetInputByIndex(STEPPITCH_INDEX).IJDInputDuringGame.Result.UomValue
  Width = pInputs.GetInputByIndex(WIDTH_INDEX).IJDInputDuringGame.Result.UomValue
  Angle = pInputs.GetInputByIndex(ANGLE_INDEX).IJDInputDuringGame.Result.UomValue
  NumSteps = pInputs.GetInputByIndex(NUMSTEPS_INDEX).IJDInputDuringGame.Result.UomValue
  Justify = pInputs.GetInputByIndex(JUSTIFICATION_INDEX).IJDInputDuringGame.Result.UomValue
  SideFrameSection = pInputs.GetInputByIndex(SIDEFRAMESECTION_INDEX).IJDInputDuringGame.Result.String
  SideFrameReferenceStandard = pInputs.GetInputByIndex(SIDEFRAMESECTIONSTANDARD_INDEX).IJDInputDuringGame.Result.String
  SideFrameSectionAngle = pInputs.GetInputByIndex(SIDEFRAMESECTIONANGLE_INDEX).IJDInputDuringGame.Result.UomValue
  SideFrameSectionCP = pInputs.GetInputByIndex(SIDEFRAMESECTIONCP_INDEX).IJDInputDuringGame.Result.UomValue
  
  bNeedToRaiseWarning = False
  ' This call verifies the existance of XSection in Catalog & access to Sym file. Any problem will be logged
  ' to middle tier error log with error as S_FALSE.
  If (DoesXSectionAndAccessToSymFileExists(defProgId, "Structure:Stair", SideFrameSection, SideFrameReferenceStandard) = False) Then
    bNeedToRaiseWarning = True
  End If

  StepSection = pInputs.GetInputByIndex(STEPSECTION_INDEX).IJDInputDuringGame.Result.String
  StepReferenceStandard = pInputs.GetInputByIndex(STEPREFERENCESTANDARD_INDEX).IJDInputDuringGame.Result.String
  StepSectionAngle = pInputs.GetInputByIndex(STEPSECTIONANGLE_INDEX).IJDInputDuringGame.Result.UomValue
  StepSectionCP = pInputs.GetInputByIndex(STEPSECTIONCP_INDEX).IJDInputDuringGame.Result.UomValue
  
  If (DoesXSectionAndAccessToSymFileExists(defProgId, "Structure:Stair", StepSection, StepReferenceStandard) = False) Then
    bNeedToRaiseWarning = True
  End If
  
  HandRailSection = pInputs.GetInputByIndex(HANDRAILSECTION_INDEX).IJDInputDuringGame.Result.String
  HandRailReferenceStandard = pInputs.GetInputByIndex(HANDRAILREFERENCESTANDARD_INDEX).IJDInputDuringGame.Result.String
  HandRailSectionAngle = pInputs.GetInputByIndex(HANDRAILSECTIONANGLE_INDEX).IJDInputDuringGame.Result.UomValue
  HandRailSectionCP = pInputs.GetInputByIndex(HANDRAILSECTIONCP_INDEX).IJDInputDuringGame.Result.UomValue
  PostHt = pInputs.GetInputByIndex(POSTHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
  HandrailPostPitch = pInputs.GetInputByIndex(HANDRAILPOSTPITCH_INDEX).IJDInputDuringGame.Result.UomValue
  NumMidRails = pInputs.GetInputByIndex(NUMMIDRAILS_INDEX).IJDInputDuringGame.Result.UomValue
  
  If (DoesXSectionAndAccessToSymFileExists(defProgId, "Structure:Stair", HandRailSection, HandRailReferenceStandard) = False) Then
    bNeedToRaiseWarning = True
  End If
 
  PlatformThickness = pInputs.GetInputByIndex(PLATFORMTHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
  WithTopLanding = pInputs.GetInputByIndex(WITHTOPLANDING_INDEX).IJDInputDuringGame.Result.UomValue
  TopLandingLength = pInputs.GetInputByIndex(TOPLANDINGLENGTH_INDEX).IJDInputDuringGame.Result.UomValue
  Justify = pInputs.GetInputByIndex(JUSTIFICATION_INDEX).IJDInputDuringGame.Result.UomValue
    
    Dim oProxy As IJDProxy
    Dim obj As IJDObject
    Dim oCatResMgr As IUnknown
    Dim pDefPlayerEx As IJDDefinitionPlayerEx
    Set pDefPlayerEx = pRepDG.definition.IJDDefinitionPlayerEx
    Dim Arguments As IJDArguments
    Set Arguments = pDefPlayerEx.SymbolInputs
    Dim oPartDef As Object
    Dim pArg1 As IJDArgument
    Set pArg1 = Arguments.Item(1)
    Set oPartDef = pArg1.Entity
    Set oProxy = oPartDef
    Set obj = oProxy.Source
    Set oCatResMgr = obj.ResourceManager
  
    'Get definition attributes and definition attribute information columns
    'which were used to get definition attribute ActualPitch
    Dim oDefAttrs As IJDAttributes
    Dim oPart As IJDPart
    Dim oAttrColl As IJDInfosCol
    Set oDefAttrs = obj
    GetDefAttribInfoColl obj, oAttrColl
    
    Angle = (PI / 2 - Angle)
  
    Dim Pos As Double
    If Justify = 2 Then
        Pos = -(Width / 2)
    ElseIf Justify = 3 Then
        Pos = (Width / 2)
    End If
   
    Dim oRep As IJDRepresentation
    Dim oOutputs As IJDOutputs
    Dim oOutput As IJDOutput
    Set oOutput = New DOutput
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
    Set oOutputs = oRep
    oOutputs.RemoveAllOutput
    'remove outputs
    
    Dim ULC As New DPosition   'Upper Left
    Dim URC As New DPosition   'Upper Right
      
    Dim LLC As New DPosition   'Lower Left
    Dim LRC As New DPosition   'Lower Left
        
    ULC.x = -Width / 2 - Pos
    ULC.y = 0#
    ULC.z = 0#
    
    URC.x = Width / 2 - Pos
    URC.y = 0#
    URC.z = 0#
    
    LLC.x = -Width / 2 - Pos
    LLC.y = -Tan(Angle) * Height
    LLC.z = -Height
    
    LRC.x = Width / 2 - Pos
    LRC.y = -Tan(Angle) * Height
    LRC.z = -Height
    
    Dim Line1 As IngrGeom3D.Line3d
    Dim pCSProfileObj As Object
    Dim pProjections As IJElements
    Dim pObj As IJSurface
 
    Set pCSProfileObj = GetCSProfile(pOC, SideFrameReferenceStandard, SideFrameSection, oCatResMgr)
'Creation of Sideframes and TopLanding (Using same section as of SideFrames for TopLanding )
    Dim Depth As Double
    Dim SecWidth As Double
    Dim TmpDepth As Double
    Dim SideWidthY As Double
    Dim SideWidthZ As Double
    
    If Not pCSProfileObj Is Nothing Then
        Depth = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "ISTRUCTCrossSectionDimensions", "Depth")
        SecWidth = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "ISTRUCTCrossSectionDimensions", "Width")
        TmpDepth = (Depth / 2) / Cos(Angle) ' to get to the center axis of the member and place steps on that axis
    End If
    SideWidthY = 2 * TmpDepth
    SideWidthZ = 2 * TmpDepth / Tan(Angle)
    
    Dim StNorm() As Double
    Dim EndNorm() As Double
    ReDim StNorm(1 To 6) As Double
    ReDim EndNorm(1 To 6) As Double
    
    StNorm(4) = 0#
    StNorm(5) = 1#
    StNorm(6) = 0#
    
    EndNorm(4) = 0#
    EndNorm(5) = 0#
    EndNorm(6) = 1#
    
    Dim pOutCurve As ComplexString3d
    Dim CurveElems As IJElements
     
    If WithTopLanding Then
        Dim complex As IngrGeom3D.ComplexString3d
        Dim iElements As IJElements
        Dim j As Integer

        LLC.y = (-Tan(Angle) * Height) - TopLandingLength
        LRC.y = (-Tan(Angle) * Height) - TopLandingLength
        
        EndNorm(1) = ULC.x
        EndNorm(2) = ULC.y - TopLandingLength
        EndNorm(3) = ULC.z
       
          
        Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y, ULC.z, ULC.x, (ULC.y - TopLandingLength), ULC.z)
        Set iElements = New JObjectCollection ' IMSElements.DynElements
        iElements.Add Line1
        Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
        Set iElements = Nothing

        StNorm(1) = ULC.x
        StNorm(2) = ULC.y - TopLandingLength
        StNorm(3) = ULC.z
                
        EndNorm(1) = LLC.x
        EndNorm(2) = LLC.y
        EndNorm(3) = LLC.z
             
        Line1.DefineBy2Points ULC.x, (ULC.y - TopLandingLength), ULC.z, LLC.x, LLC.y, LLC.z
        complex.AddCurve Line1, True
        
        If pCSProfileObj Is Nothing Then
            complex.GetCurves CurveElems
            Set pOutCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, CurveElems)
            InitNewOutput pOC, "LeftSideFrame"
            pOC.AddOutput "LeftSideFrame", pOutCurve
        ElseIf RepType = CenterLineRep Then
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, (ULC.y - TopLandingLength), ULC.z, ULC.x, (ULC.y - TopLandingLength), ULC.z - SideWidthZ)
            InitNewOutput pOC, "LeftSideFrame" & "_1"
            pOC.AddOutput "LeftSideFrame" & "_1", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, (ULC.y - TopLandingLength), ULC.z - SideWidthZ, LLC.x, LLC.y + SideWidthY, LLC.z)
            InitNewOutput pOC, "LeftSideFrame" & "_2"
            pOC.AddOutput "LeftSideFrame" & "_2", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, LLC.x, LLC.y + SideWidthY, LLC.z, LLC.x, LLC.y, LLC.z)
            InitNewOutput pOC, "LeftSideFrame" & "_3"
            pOC.AddOutput "LeftSideFrame" & "_3", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, LLC.x, LLC.y, LLC.z, ULC.x, (ULC.y - TopLandingLength), ULC.z)
            InitNewOutput pOC, "LeftSideFrame" & "_4"
            pOC.AddOutput "LeftSideFrame" & "_4", Line1
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, complex, pCSProfileObj, SideFrameSectionCP, SideFrameSectionAngle, 0, Nothing, EndNorm, skinOption)
            For j = 1 To pProjections.Count
                Set pObj = pProjections(j)
                InitNewOutput pOC, "LeftSideFrame" & j
                pOC.AddOutput "LeftSideFrame" & j, pObj
            Next j
        End If
 
        Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, URC.y, URC.z, URC.x, (URC.y - TopLandingLength), URC.z)
        Set iElements = New JObjectCollection ' IMSElements.DynElements
        iElements.Add Line1
        Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
        Set iElements = Nothing

        StNorm(1) = URC.x
        StNorm(2) = URC.y - TopLandingLength
        StNorm(3) = URC.z

        EndNorm(1) = LRC.x
        EndNorm(2) = LRC.y
        EndNorm(3) = LLC.z

        Line1.DefineBy2Points URC.x, (URC.y - TopLandingLength), URC.z, LRC.x, LRC.y, LLC.z
        complex.AddCurve Line1, True
        
        If pCSProfileObj Is Nothing Then
            complex.GetCurves CurveElems
            Set pOutCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, CurveElems)
            InitNewOutput pOC, "RightSideFrame"
            pOC.AddOutput "RightSideFrame", pOutCurve
        ElseIf RepType = CenterLineRep Then
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, (URC.y - TopLandingLength), URC.z, URC.x, (URC.y - TopLandingLength), URC.z - SideWidthZ)
            InitNewOutput pOC, "RightSideFrame" & "_1"
            pOC.AddOutput "RightSideFrame" & "_1", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, (URC.y - TopLandingLength), URC.z - SideWidthZ, LRC.x, LRC.y + SideWidthY, LLC.z)
            InitNewOutput pOC, "RightSideFrame" & "_2"
            pOC.AddOutput "RightSideFrame" & "_2", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, LRC.x, LRC.y + SideWidthY, LLC.z, LRC.x, LRC.y, LLC.z)
            InitNewOutput pOC, "RightSideFrame" & "_3"
            pOC.AddOutput "RightSideFrame" & "_3", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, LRC.x, LRC.y, LLC.z, URC.x, (URC.y - TopLandingLength), URC.z)
            InitNewOutput pOC, "RightSideFrame" & "_4"
            pOC.AddOutput "RightSideFrame" & "_4", Line1
            
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, complex, pCSProfileObj, SideFrameSectionCP, SideFrameSectionAngle, 1, Nothing, EndNorm, skinOption)
            For j = 1 To pProjections.Count
                Set pObj = pProjections(j)
                InitNewOutput pOC, "RightSideFrame" & j
                pOC.AddOutput "RightSideFrame" & j, pObj
            Next j
        End If

        Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y, ULC.z, URC.x, URC.y, URC.z)
        
        If pCSProfileObj Is Nothing Or RepType = CenterLineRep Then
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y, ULC.z, URC.x, URC.y, URC.z)
            InitNewOutput pOC, "FrontTopLanding"
            pOC.AddOutput "FrontTopLanding", Line1
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, pCSProfileObj, SideFrameSectionCP, SideFrameSectionAngle, 0, Nothing, Nothing, skinOption)

'            tr 76254 - add all elements in a loop instead of just the first one
'            as with skinoption 7 we will receive many surfaces.
'            Set pObj = pProjections(1)
'            InitNewOutput pOC, "FrontTopLanding"
'            pOC.AddOutput "FrontTopLanding", pObj
            
            For j = 1 To pProjections.Count
                Set pObj = pProjections(j)
                InitNewOutput pOC, "FrontTopLanding" & j
                pOC.AddOutput "FrontTopLanding" & j, pObj
            Next j
            
        End If

        Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, (ULC.y - TopLandingLength), ULC.z, URC.x, (URC.y - TopLandingLength), URC.z)
        If pCSProfileObj Is Nothing Or RepType = CenterLineRep Then
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, (ULC.y - TopLandingLength), ULC.z, URC.x, (URC.y - TopLandingLength), URC.z)
            InitNewOutput pOC, "BackTopLanding"
            pOC.AddOutput "BackTopLanding", Line1
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, pCSProfileObj, SideFrameSectionCP, SideFrameSectionAngle, 1, Nothing, Nothing, skinOption)
            
'            tr 76254 - add all elements in a loop instead of just the first one
'            as with skinoption 7 we will receive many surfaces.
'            Set pObj = pProjections(1)
'            InitNewOutput pOC, "BackTopLanding"
'            pOC.AddOutput "BackTopLanding", pObj
             For j = 1 To pProjections.Count
                Set pObj = pProjections(j)
                InitNewOutput pOC, "BackTopLanding" & j
                pOC.AddOutput "BackTopLanding" & j, pObj
            Next j
       End If

        If TopLandingLength > 1.5 Then
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y, ULC.z, URC.x, (URC.y - TopLandingLength), URC.z)
            If pCSProfileObj Is Nothing Or RepType = CenterLineRep Then
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y, ULC.z, URC.x, (URC.y - TopLandingLength), URC.z)
                InitNewOutput pOC, "MiddleFrame"
                pOC.AddOutput "MiddleFrame", Line1
            Else
                Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, pCSProfileObj, SideFrameSectionCP, SideFrameSectionAngle, 0, Nothing, Nothing, skinOption)
                
'            tr 76254 - add all elements in a loop instead of just the first one
'            as with skinoption 7 we will receive many surfaces.
'                Set pObj = pProjections(1)
'                InitNewOutput pOC, "MiddleFrame"
'                pOC.AddOutput "MiddleFrame", pObj
                
                For j = 1 To pProjections.Count
                    Set pObj = pProjections(j)
                    InitNewOutput pOC, "MiddleFrame" & j
                    pOC.AddOutput "MiddleFrame" & j, pObj
                Next j

            End If
        End If

        'Platform on the top landing no schema right now
        Dim projvect As New DVector, upvect As New DVector
        Dim TmpPos As New DPosition
        TmpPos.x = 0 - Pos
        TmpPos.y = 0#
        TmpPos.z = 0#
        projvect.x = 0#
        projvect.y = -TopLandingLength
        projvect.z = 0#

        upvect.x = 0#
        upvect.y = 0#
        upvect.z = 1#

        If RepType = CenterLineRep Then
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, URC.y, URC.z, URC.x, (URC.y - TopLandingLength), URC.z)
            InitNewOutput pOC, "Platform" & "_1"
            pOC.AddOutput "Platform" & "_1", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, (URC.y - TopLandingLength), URC.z, ULC.x, (URC.y - TopLandingLength), URC.z)
            InitNewOutput pOC, "Platform" & "_2"
            pOC.AddOutput "Platform" & "_2", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, (URC.y - TopLandingLength), URC.z, ULC.x, URC.y, URC.z)
            InitNewOutput pOC, "Platform" & "_3"
            pOC.AddOutput "Platform" & "_3", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, URC.y, URC.z, URC.x, URC.y, URC.z)
            InitNewOutput pOC, "Platform" & "_4"
            pOC.AddOutput "Platform" & "_4", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, URC.y, URC.z - SideWidthZ, URC.x, (URC.y - TopLandingLength), URC.z - SideWidthZ)
            InitNewOutput pOC, "Platform" & "_5"
            pOC.AddOutput "Platform" & "_5", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, (URC.y - TopLandingLength), URC.z - SideWidthZ, ULC.x, (URC.y - TopLandingLength), URC.z - SideWidthZ)
            InitNewOutput pOC, "Platform" & "_6"
            pOC.AddOutput "Platform" & "_6", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, (URC.y - TopLandingLength), URC.z - SideWidthZ, ULC.x, URC.y, URC.z - SideWidthZ)
            InitNewOutput pOC, "Platform" & "_7"
            pOC.AddOutput "Platform" & "_7", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, URC.y, URC.z - SideWidthZ, URC.x, URC.y, URC.z - SideWidthZ)
            InitNewOutput pOC, "Platform" & "_8"
            pOC.AddOutput "Platform" & "_8", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, URC.y, URC.z - SideWidthZ, URC.x, URC.y, URC.z)
            InitNewOutput pOC, "Platform" & "_9"
            pOC.AddOutput "Platform" & "_9", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, URC.y, URC.z - SideWidthZ, ULC.x, URC.y, URC.z)
            InitNewOutput pOC, "Platform" & "_10"
            pOC.AddOutput "Platform" & "_10", Line1
        Else
            CreateRectangularProjection pOC, TmpPos, projvect, upvect, Width, PlatformThickness, "Platform", 1
        End If

    Else
    
        StNorm(1) = ULC.x
        StNorm(2) = ULC.y
        StNorm(3) = ULC.z
                    
        EndNorm(1) = LLC.x
        EndNorm(2) = LLC.y
        EndNorm(3) = LLC.z
        
        Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y, ULC.z, LLC.x, LLC.y, LLC.z)
        If pCSProfileObj Is Nothing Then
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y, ULC.z, LLC.x, LLC.y, LLC.z)
            InitNewOutput pOC, "LeftSideFrame"
            pOC.AddOutput "LeftSideFrame", Line1
        ElseIf RepType = CenterLineRep Then
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, (ULC.y), ULC.z, ULC.x, (ULC.y), ULC.z - SideWidthZ)
            InitNewOutput pOC, "LeftSideFrame" & "_1"
            pOC.AddOutput "LeftSideFrame" & "_1", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, (ULC.y), ULC.z - SideWidthZ, LLC.x, LLC.y + SideWidthY, LLC.z)
            InitNewOutput pOC, "LeftSideFrame" & "_2"
            pOC.AddOutput "LeftSideFrame" & "_2", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, LLC.x, LLC.y + SideWidthY, LLC.z, LLC.x, LLC.y, LLC.z)
            InitNewOutput pOC, "LeftSideFrame" & "_3"
            pOC.AddOutput "LeftSideFrame" & "_3", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, LLC.x, LLC.y, LLC.z, ULC.x, (ULC.y), ULC.z)
            InitNewOutput pOC, "LeftSideFrame" & "_4"
            pOC.AddOutput "LeftSideFrame" & "_4", Line1
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, pCSProfileObj, SideFrameSectionCP, SideFrameSectionAngle, 0, StNorm, EndNorm, skinOption)
            For i = 1 To pProjections.Count
                Set pObj = pProjections(i)
                InitNewOutput pOC, "LeftSideFrame" & CStr(i)
                pOC.AddOutput "LeftSideFrame" & CStr(i), pObj
            Next i
        End If
        
        StNorm(1) = URC.x
        StNorm(2) = URC.y
        StNorm(3) = URC.z
        EndNorm(1) = LRC.x
        EndNorm(2) = LRC.y
        EndNorm(3) = LLC.z
        
        Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, URC.y, URC.z, LRC.x, LRC.y, LLC.z)
        If pCSProfileObj Is Nothing Then
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, URC.y, URC.z, LRC.x, LRC.y, LLC.z)
            InitNewOutput pOC, "RightSideFrame"
            pOC.AddOutput "RightSideFrame", Line1
        ElseIf RepType = CenterLineRep Then
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, (URC.y), URC.z, URC.x, (URC.y), URC.z - SideWidthZ)
            InitNewOutput pOC, "RightSideFrame" & "_1"
            pOC.AddOutput "RightSideFrame" & "_1", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, (URC.y), URC.z - SideWidthZ, LRC.x, LRC.y + SideWidthY, LLC.z)
            InitNewOutput pOC, "RightSideFrame" & "_2"
            pOC.AddOutput "RightSideFrame" & "_2", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, LRC.x, LRC.y + SideWidthY, LLC.z, LRC.x, LRC.y, LLC.z)
            InitNewOutput pOC, "RightSideFrame" & "_3"
            pOC.AddOutput "RightSideFrame" & "_3", Line1
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, LRC.x, LRC.y, LLC.z, URC.x, (URC.y), URC.z)
            InitNewOutput pOC, "RightSideFrame" & "_4"
            pOC.AddOutput "RightSideFrame" & "_4", Line1
            
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, pCSProfileObj, SideFrameSectionCP, SideFrameSectionAngle, 1, StNorm, EndNorm, skinOption)
            For i = 1 To pProjections.Count
                Set pObj = pProjections(i)
                InitNewOutput pOC, "RightSideFrame" & CStr(i)
                pOC.AddOutput "RightSideFrame" & CStr(i), pObj
            Next i
        End If
        
    End If
    
        Set pObj = Nothing
        
        Dim StPt As New DPosition
        Dim endpt As New DPosition
        Set StPt = ULC.Clone
        Set endpt = URC.Clone
    
        If Not pCSProfileObj Is Nothing Then
            Set obj = pCSProfileObj
            obj.Remove
            Set pCSProfileObj = Nothing
        End If
        Set pCSProfileObj = GetCSProfile(pOC, StepReferenceStandard, StepSection, oCatResMgr)
        Dim StepDepth As Double
        Dim StepSpan As Double
        If Not pCSProfileObj Is Nothing Then
            'here the StepDepth is really the width of the step section
            StepDepth = GetCSAttribData(StepSection, StepReferenceStandard, "ISTRUCTCrossSectionDimensions", "Width")
            StepSpan = GetCSAttribData(StepSection, StepReferenceStandard, "ISTRUCTCrossSectionDimensions", "Depth")
            'only support step section CP5 as default CP
            StepDepth = StepDepth / 2
            StepSpan = StepSpan / 2
        End If
    
        ReDim StNorm(1 To 3) As Double
        ReDim EndNorm(1 To 3) As Double
        
                'Manish TR#41148  starts here -'''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Currently steps are placed with assumed CP=7 (top-left)for
        'stringer beams (side frame members). So whenever CP is other than 7 width and
        ' y direction location of steps need to be modified.
        ' Implemented assuming steps are inside-inside so it works fine for channel sections
        ' for other sections it needs to be modified. Currently cross section type
        ' is not available so it is not easily possible.
        'stringer beam rotation angle not handled. It is difficult to handle it. Also rotation angle
        'other than 0 is not practicle case. Actually rotation of stringer beam need to be disabled
        ' in property page
        '
        On Error Resume Next
        
        Dim WidthCorrection As Double
        Dim CGx As Double, CGy As Double
        Dim xp As Double, yp As Double
        Dim DepthCorrection As Double
        
        'initialize following values to 0 because some c/s may not have these values
        CGx = 0
        CGy = 0
        xp = 0
        yp = 0
        
        WidthCorrection = 0
        DepthCorrection = 0
        
        'Right now, only default CP (CP7) is supported. Comment out the following code
        'Select Case SideFrameSectionCP
        '    Case 1, 4, 7, 12 'left edge
        '        WidthCorrection = 0
        '    Case 2, 5, 8 'along half width
        '        WidthCorrection = SecWidth / 2
        '    Case 3, 6, 9, 13 'right edge
        '        WidthCorrection = SecWidth
        '    Case 10, 11, 14 'along centroid in width direction
        '        CGx = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "ISTRUCTCrossSectionDesignProperties", "CentroidX")
        '        WidthCorrection = CGx
        '    Case 15 'shear center
        '        xp = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "IJUAL", "xp")
        '        WidthCorrection = xp
        'End Select
        
        'adjust step width WidthCorrection value
        'StPt.x = StPt.x + WidthCorrection
        'endpt.x = endpt.x - WidthCorrection
        
        'now get depth correction
        
        'Right now, only default CP (CP7) is supported. Comment out the following code
        'Select Case SideFrameSectionCP
        '    Case 7, 8, 9, 14  'top edge
        '        DepthCorrection = 0
        '    Case 4, 5, 6  'along half depth
        '        DepthCorrection = (Depth / 2) / Sin(Angle)
        '    Case 1, 2, 3, 11 'bottom  edge
        '        DepthCorrection = Depth / Sin(Angle)
        '    Case 10, 12, 13  'along centroid in depth direction
        '        CGy = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "ISTRUCTCrossSectionDesignProperties", "CentroidY")
        '        DepthCorrection = CGy / Sin(Angle)
        '    Case 15  'shear center
        '        yp = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "IJUAL", "yp")
        '        DepthCorrection = yp / Sin(Angle)
        'End Select
        
        'TmpDepth = TmpDepth - DepthCorrection
        
        Err.Clear
        On Error GoTo ErrorHandler
        'Manish - TR#41148 ends here''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Creation of Steps

        Dim bAttrExists As Boolean
        Dim ActualPitch As Double
        bAttrExists = False
        ActualPitch = GetAttribute(oDefAttrs, "ActualPitch", oAttrColl, bAttrExists)
        Set oDefAttrs = Nothing
        Set oAttrColl = Nothing
        'The occurrence attributes are not available in the method, recalculate actual pitch the same way as in ISPSStairSymbolServices_EvaluateGeometry
        If bAttrExists Then
            ActualPitch = Height / (NumSteps + 1)
            Pitch = ActualPitch
        End If
        If Not pCSProfileObj Is Nothing Then
            Depth = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "ISTRUCTCrossSectionDimensions", "Depth")
            TmpDepth = (Depth / 2) / Sin(Angle) ' to get to the center axis of the member and place steps on that axis
        End If
        For i = 1 To NumSteps '+ 1

            StPt.y = -Tan(Angle) * ((Pitch * i) - TmpDepth + StepDepth)
            endpt.y = StPt.y
            
            If WithTopLanding Then
                StPt.y = StPt.y - TopLandingLength
                endpt.y = endpt.y - TopLandingLength
            End If
            StPt.z = -(i * Pitch + StepDepth)
            endpt.z = StPt.z
            
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y, StPt.z, endpt.x, endpt.y, endpt.z)
            
            If pCSProfileObj Is Nothing Then
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y, StPt.z, endpt.x, endpt.y, endpt.z)
                InitNewOutput pOC, "Step" & i
                pOC.AddOutput "Step" & i, Line1
            ElseIf RepType = CenterLineRep Then
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y - StepSpan, StPt.z, endpt.x, endpt.y - StepSpan, endpt.z)
                InitNewOutput pOC, "Step" & i & "_1"
                pOC.AddOutput "Step" & i & "_1", Line1
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y + StepSpan, StPt.z, endpt.x, endpt.y + StepSpan, endpt.z)
                InitNewOutput pOC, "Step" & i & "_2"
                pOC.AddOutput "Step" & i & "_2", Line1
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y + StepSpan, StPt.z, StPt.x, StPt.y - StepSpan, StPt.z)
                InitNewOutput pOC, "Step" & i & "_3"
                pOC.AddOutput "Step" & i & "_3", Line1
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, endpt.x, endpt.y - StepSpan, endpt.z, endpt.x, endpt.y + StepSpan, endpt.z)
                InitNewOutput pOC, "Step" & i & "_4"
                pOC.AddOutput "Step" & i & "_4", Line1
            Else
                skinOption = 6 ' do not need to create caps for steps tr 76254
                Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, pCSProfileObj, StepSectionCP, StepSectionAngle, 0, Nothing, Nothing, skinOption)
                
                For j = 1 To pProjections.Count
                    Set pObj = pProjections(j)
                    InitNewOutput pOC, "Step" & i & j
                    pOC.AddOutput "Step" & i & j, pObj
                Next j

            End If
             
        Next i
        
        Set endpt = Nothing
        Set StPt = Nothing
     
'Creation of Handrails and posts

        If Not pCSProfileObj Is Nothing Then
            Set obj = pCSProfileObj
            obj.Remove
            Set pCSProfileObj = Nothing
        End If
        Set pCSProfileObj = GetCSProfile(pOC, HandRailReferenceStandard, HandRailSection, oCatResMgr)
        ULC.x = ULC.x - (SecWidth / 2)
        URC.x = URC.x + (SecWidth / 2)
        LLC.x = LLC.x - (SecWidth / 2)
        LRC.x = LRC.x + (SecWidth / 2)
        
        If Not pCSProfileObj Is Nothing Then
            Depth = GetCSAttribData(HandRailSection, HandRailReferenceStandard, "ISTRUCTCrossSectionDimensions", "Depth")
        End If
        
        If RepType = CenterLineRep Then
            If Not pCSProfileObj Is Nothing Then
                Set obj = pCSProfileObj
                obj.Remove
                Set pCSProfileObj = Nothing
            End If
        End If
        CreatePosts pOC, pCSProfileObj, Angle, ULC, LLC, URC, LRC, Length, _
                    HandRailSectionCP, HandRailSectionAngle, HandrailPostPitch, _
                    PostHt, WithTopLanding, TopLandingLength, Height, Depth
        
        If WithTopLanding Then
            ULC.y = 0#
            URC.y = 0#
            CreateHandRailsWithToplanding pOC, pCSProfileObj, ULC, LLC, URC, LRC, _
                            HandRailSectionCP, HandRailSectionAngle, _
                            PostHt, TopLandingLength, HandRailSpacing, NumMidRails, Depth
        Else
            CreateHandrails pOC, pCSProfileObj, ULC, LLC, URC, LRC, _
                            HandRailSectionCP, HandRailSectionAngle, _
                            PostHt, TopLandingLength, HandRailSpacing, NumMidRails, Depth
        End If
        
    Set pObj = Nothing
    
    If Not pCSProfileObj Is Nothing Then
        Set obj = pCSProfileObj
        obj.Remove
        Set pCSProfileObj = Nothing
    End If
    
    Set pProjections = Nothing
    Set Line1 = Nothing
    Set pRepDG = Nothing
    Set pOC = Nothing
    Set pInputs = Nothing
    
    If bNeedToRaiseWarning Then
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

Private Sub CreatePosts(pOC As IJDOutputCollection, ByVal CSObject As Object, _
                        ByVal Angle As Double, ByVal ULC As DPosition, ByVal LLC As DPosition, _
                        ByVal URC As DPosition, ByVal LRC As DPosition, ByVal Length As Double, _
                        cp As Integer, RotAngle As Double, ByVal HandrailPostPitch As Double, _
                        ByVal PostHt As Double, ByVal WithTopLanding As Boolean, TopLandingLength As Double, _
                        ByVal Height As Double, Depth As Double)
Const METHOD = "CreatePosts"
On Error GoTo ErrorHandler

        Dim NumPosts As Integer
        NumPosts = Int(((Length - 0.6) / HandrailPostPitch) + 1)
        Dim NewSpacing As Double
        NewSpacing = Length / NumPosts
        Dim pProjections As IJElements
        Dim pObj As IJSurface
        Dim StPt As New DPosition
        Dim endpt As New DPosition
        Dim i As Integer
        Dim j As Integer
        Dim DepthCorr As Double
        Dim Line1 As IngrGeom3D.Line3d
    
        If WithTopLanding Then
            
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y - 0.1, ULC.z, ULC.x, ULC.y - 0.1, ULC.z + PostHt)
            If CSObject Is Nothing Then
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y - 0.1, ULC.z, ULC.x, ULC.y - 0.1, ULC.z + PostHt)
                InitNewOutput pOC, "LPost"
                pOC.AddOutput "LPost", Line1
            Else
                Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
'            tr 76254 - add all elements in a loop instead of just the first one
'            as with skinoption 7 we will receive many surfaces.
'                Set pObj = pProjections(1)
'                InitNewOutput pOC, "LPost" & i
'                pOC.AddOutput "LPost" & i, pObj
                
                For j = 1 To pProjections.Count
                    Set pObj = pProjections(j)
                    InitNewOutput pOC, "LPost" & i & j
                    pOC.AddOutput "LPost" & i & j, pObj
                Next j

            End If
            
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y - TopLandingLength + 0.1, ULC.z, ULC.x, ULC.y - TopLandingLength + 0.1, ULC.z + PostHt)
            If CSObject Is Nothing Then
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, ULC.x, ULC.y - TopLandingLength + 0.1, ULC.z, ULC.x, ULC.y - TopLandingLength + 0.1, ULC.z + PostHt)
                InitNewOutput pOC, "LPost"
                pOC.AddOutput "LPost", Line1
            Else
                Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
'            tr 76254 - add all elements in a loop instead of just the first one
'            as with skinoption 7 we will receive many surfaces.
'                Set pObj = pProjections(1)
'                InitNewOutput pOC, "LPost" & i
'                pOC.AddOutput "LPost" & i, pObj
                
                For j = 1 To pProjections.Count
                    Set pObj = pProjections(j)
                    InitNewOutput pOC, "LPost" & i & j
                    pOC.AddOutput "LPost" & i & j, pObj
                Next j
            End If
                        
            Select Case cp
                Case 7, 8, 9, 14  'top edge
                    DepthCorr = Depth
                Case 4, 5, 6  'along half depth
                Case 1, 2, 3, 11 'bottom  edge
                    DepthCorr = -(Depth)
                Case 10, 12, 13  'along centroid in depth direction
                Case 15  'shear center
            End Select
            
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x + DepthCorr, URC.y - 0.1, URC.z, URC.x + DepthCorr, URC.y - 0.1, URC.z + PostHt)
            If CSObject Is Nothing Then
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, URC.y - 0.1, URC.z, URC.x, URC.y - 0.1, URC.z + PostHt)
                InitNewOutput pOC, "RPost"
                pOC.AddOutput "RPost", Line1
            Else
                Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
'            tr 76254 - add all elements in a loop instead of just the first one
'            as with skinoption 7 we will receive many surfaces.
'                Set pObj = pProjections(1)
'                InitNewOutput pOC, "RPost" & i
'                pOC.AddOutput "RPost" & i, pObj
                
                 For j = 1 To pProjections.Count
                    Set pObj = pProjections(j)
                    InitNewOutput pOC, "RPost" & i & j
                    pOC.AddOutput "RPost" & i & j, pObj
                Next j
                
            End If
                
            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x + DepthCorr, URC.y - TopLandingLength + 0.1, URC.z, URC.x + DepthCorr, URC.y - TopLandingLength + 0.1, URC.z + PostHt)
            If CSObject Is Nothing Then
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, URC.x, URC.y - TopLandingLength + 0.1, URC.z, URC.x, URC.y - TopLandingLength + 0.1, URC.z + PostHt)
                InitNewOutput pOC, "RPost"
                pOC.AddOutput "RPost", Line1
            Else
                Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
'            tr 76254 - add all elements in a loop instead of just the first one
'            as with skinoption 7 we will receive many surfaces.
'                Set pObj = pProjections(1)
'                InitNewOutput pOC, "RPost" & i
'                pOC.AddOutput "RPost" & i, pObj
                
                For j = 1 To pProjections.Count
                    Set pObj = pProjections(j)
                    InitNewOutput pOC, "RPost" & i & j
                    pOC.AddOutput "RPost" & i & j, pObj
                Next j
              
            End If

            ULC.y = ULC.y - TopLandingLength
            URC.y = URC.y - TopLandingLength
            LLC.y = (-Tan(Angle) * Height) - TopLandingLength
            LRC.y = (-Tan(Angle) * Height) - TopLandingLength
        Else
        End If
        
        Set StPt = ULC.Clone
        
        StPt.z = StPt.z - Cos(Angle) * 0.3
        StPt.y = StPt.y - Sin(Angle) * 0.3
        Set endpt = StPt.Clone
        endpt.z = endpt.z + PostHt
                  
         Dim StNorm() As Double
         ReDim StNorm(1 To 6) As Double
        
         StNorm(4) = 0#
         StNorm(5) = -1 / Sin(Angle)
         StNorm(6) = 1 / Cos(Angle)
            
        'Intermediate Posts
        For i = 1 To NumPosts
            If i <> 1 Then
                StPt.y = StPt.y - (Sin(Angle) * NewSpacing)
                StPt.z = StPt.z - (Cos(Angle) * NewSpacing)
                Set endpt = StPt.Clone
                endpt.z = endpt.z + PostHt
            End If
            
            StNorm(1) = StPt.x
            StNorm(2) = StPt.y
            StNorm(3) = StPt.z

            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y, StPt.z, endpt.x, endpt.y, endpt.z)
            If CSObject Is Nothing Then
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y, StPt.z, endpt.x, endpt.y, endpt.z)
                InitNewOutput pOC, "LPost" & i
                pOC.AddOutput "LPost" & i, Line1
            Else
                Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, CSObject, cp, RotAngle, 0, StNorm, Nothing, skinOption)
                For j = 1 To pProjections.Count
                    Set pObj = pProjections(j)
                    InitNewOutput pOC, "LPost" & i & j
                    pOC.AddOutput "LPost" & i & j, pObj
                Next j
            End If
            
            Next i
            
        Set StPt = URC.Clone
        StPt.z = StPt.z - Cos(Angle) * 0.3
        StPt.y = StPt.y - Sin(Angle) * 0.3
                'now get depth correction
    
        Select Case cp
            Case 7, 8, 9, 14  'top edge
                StPt.x = StPt.x + Depth
            Case 4, 5, 6  'along half depth
            Case 1, 2, 3, 11 'bottom  edge
                StPt.x = StPt.x - Depth
            Case 10, 12, 13  'along centroid in depth direction
            Case 15  'shear center
        End Select
        
        Set endpt = StPt.Clone
        endpt.z = endpt.z + PostHt
            
        For i = 1 To NumPosts
            If i <> 1 Then
                StPt.y = StPt.y - (Sin(Angle) * NewSpacing)
                StPt.z = StPt.z - (Cos(Angle) * NewSpacing)
                Set endpt = StPt.Clone
                endpt.z = endpt.z + PostHt
            End If
            
            StNorm(1) = StPt.x
            StNorm(2) = StPt.y
            StNorm(3) = StPt.z

            Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y, StPt.z, endpt.x, endpt.y, endpt.z)
            If CSObject Is Nothing Then
                Set Line1 = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y, StPt.z, endpt.x, endpt.y, endpt.z)
                InitNewOutput pOC, "RPost" & i
                pOC.AddOutput "RPost" & i, Line1
            Else
                Set pProjections = CreateProjectionFromCSProfile(Nothing, Line1, CSObject, cp, RotAngle, 0, StNorm, Nothing, skinOption)
                For j = 1 To pProjections.Count
                    Set pObj = pProjections(j)
                    InitNewOutput pOC, "RPost" & i & j
                    pOC.AddOutput "RPost" & i & j, pObj
                Next j
            End If
            
        Next i
          
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub CreateHandrails(pOC As IJDOutputCollection, ByVal CSObject As Object, _
                         ByVal ULC As DPosition, ByVal LLC As DPosition, _
                        ByVal URC As DPosition, ByVal LRC As DPosition, _
                        ByVal cp As Integer, ByVal RotAngle As Double, _
                        ByVal PostHt As Double, ByVal TopLandingLength As Double, _
                        ByVal HandRailSpacing As Double, ByVal NumMidRails As Integer, ByVal Depth As Double)
Const METHOD = "CreateHandRails"
On Error GoTo ErrorHandler
           
        Dim pProjections As IJElements
        Dim iElements As IJElements
        Dim pObj As IJSurface
        Dim StPt As New DPosition
        Dim endpt As New DPosition
        Dim i As Integer, j As Integer
        Dim oLine As IngrGeom3D.Line3d
        Dim complex As IngrGeom3D.ComplexString3d

  
        HandRailSpacing = PostHt / (NumMidRails + 1)
 
        Set StPt = ULC.Clone
        Set endpt = LLC.Clone
        Dim StNorm() As Double
        Dim EndNorm() As Double
        ReDim StNorm(1 To 3) As Double
        ReDim EndNorm(1 To 3) As Double
        Dim pOutCurve As ComplexString3d
        Dim CurveElems As IJElements
        Dim DepthCorr As Double
        
        Select Case cp
            Case 6, 8, 13, 14
                DepthCorr = -Depth / 2
            Case 1
                DepthCorr = Depth
            Case 2, 4, 11, 12
                DepthCorr = Depth / 2
            Case 3, 5, 7, 10, 15
            Case 9
                DepthCorr = -Depth
        End Select
          
        '1
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x + DepthCorr, StPt.y, StPt.z + PostHt, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt)
        Set iElements = New JObjectCollection ' IMSElements.DynElements
        iElements.Add oLine
        Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
        Set iElements = Nothing
        '2
        If NumMidRails >= 1 Then
            oLine.DefineBy2Points endpt.x + DepthCorr, endpt.y, endpt.z + PostHt, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (NumMidRails))
            complex.AddCurve oLine, True
            '3
            oLine.DefineBy2Points endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (NumMidRails)), StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (NumMidRails))
            complex.AddCurve oLine, True
            '4
            oLine.DefineBy2Points StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (NumMidRails)), StPt.x + DepthCorr, StPt.y, StPt.z + PostHt
            complex.AddCurve oLine, True
        End If
        
        If CSObject Is Nothing Then
            complex.GetCurves CurveElems
            Set pOutCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, CurveElems)
            InitNewOutput pOC, "LHandRail1"  ' tr 76254
            pOC.AddOutput "LHandRail1", pOutCurve
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, complex, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
        'need to loop for numbers for WCG
            For i = 1 To pProjections.Count
                Set pObj = pProjections(i)
                InitNewOutput pOC, "LHandRail1" & i    ' tr 76254
                pOC.AddOutput "LHandRail1" & i, pObj
                Set complex = Nothing
                Set oLine = Nothing
            Next i
        End If
                                
        'MiddleRails on Left
        Set StPt = ULC.Clone
        Set endpt = LLC.Clone
        If NumMidRails >= 2 Then
            For i = 2 To NumMidRails
                Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (i - 1)), endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (i - 1)))
                If CSObject Is Nothing Then
                    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y, StPt.z + PostHt - (HandRailSpacing * (i - 1)), endpt.x, endpt.y, endpt.z + PostHt - (HandRailSpacing * (i - 1)))
                    InitNewOutput pOC, "LHandRail1" & i
                    pOC.AddOutput "LHandRail1" & i, oLine
                Else
                    Set pProjections = CreateProjectionFromCSProfile(Nothing, oLine, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
                    
                    ' tr 76254
'                    Set pObj = pProjections(1)
'                    InitNewOutput pOC, "LHandRail1" & i
'                    pOC.AddOutput "LHandRail1" & i, pObj
                    
                    For j = 1 To pProjections.Count
                        Set pObj = pProjections(j)
                        InitNewOutput pOC, "LHandRail1" & i & j
                        pOC.AddOutput "LHandRail1" & i & j, pObj
                    Next j

                End If
            Next i
        End If
        
        Set StPt = URC.Clone
        Set endpt = LRC.Clone
        DepthCorr = 0
        Select Case cp
            Case 4, 8, 12, 14
                DepthCorr = Depth / 2
            Case 1, 5, 9, 10, 15
            Case 2, 6, 11, 13
                DepthCorr = -Depth / 2
            Case 3
                DepthCorr = -Depth
            Case 7
                DepthCorr = Depth
        End Select
        
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x + DepthCorr, StPt.y, StPt.z + PostHt, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt)
        Set iElements = New JObjectCollection 'IMSElements.DynElements
        iElements.Add oLine
        Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
        Set iElements = Nothing
        
        If NumMidRails >= 1 Then
             oLine.DefineBy2Points endpt.x + DepthCorr, endpt.y, endpt.z + PostHt, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (NumMidRails))
             complex.AddCurve oLine, True
             
             oLine.DefineBy2Points endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (NumMidRails)), StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (NumMidRails))
             complex.AddCurve oLine, True
        
             oLine.DefineBy2Points StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (NumMidRails)), StPt.x + DepthCorr, StPt.y, StPt.z + PostHt
             complex.AddCurve oLine, True
        End If
    
        If CSObject Is Nothing Then
            complex.GetCurves CurveElems
            Set pOutCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, CurveElems)
            InitNewOutput pOC, "RHandRail1"
            pOC.AddOutput "RHandRail1", pOutCurve
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, complex, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
            For i = 1 To pProjections.Count
                Set pObj = pProjections(i)
                InitNewOutput pOC, "RHandRail1" & i
                pOC.AddOutput "RHandRail1" & i, pObj
            Next i
        End If
    
        'MiddleRails on Right
        If NumMidRails >= 2 Then
            For i = 2 To NumMidRails
                Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (i - 1)), endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (i - 1)))
                If CSObject Is Nothing Then
                    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x, StPt.y, StPt.z + PostHt - (HandRailSpacing * (i - 1)), endpt.x, endpt.y, endpt.z + PostHt - (HandRailSpacing * (i - 1)))
                    InitNewOutput pOC, "RHandRail1" & i
                    pOC.AddOutput "RHandRail1" & i, oLine
                Else
                    Set pProjections = CreateProjectionFromCSProfile(Nothing, oLine, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
                    ' tr 76254
'                    Set pObj = pProjections(1)
'                    InitNewOutput pOC, "RHandRail1" & i
'                    pOC.AddOutput "RHandRail1" & i, pObj

                    For j = 1 To pProjections.Count
                        Set pObj = pProjections(j)
                        InitNewOutput pOC, "RHandRail1" & i & j
                        pOC.AddOutput "RHandRail1" & i & j, pObj
                    Next j

                End If
            Next i
       End If
        
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub CreateHandRailsWithToplanding(pOC As IJDOutputCollection, ByVal CSObject As Object, _
                         ByVal ULC As DPosition, ByVal LLC As DPosition, _
                        ByVal URC As DPosition, ByVal LRC As DPosition, _
                        cp As Integer, RotAngle As Double, _
                        ByVal PostHt, ByVal TopLandingLength As Double, _
                        ByVal HandRailSpacing As Double, ByVal NumMidRails As Integer, ByVal Depth As Double)
Const METHOD = "CreateHandRailsWithToplanding"
On Error GoTo ErrorHandler
           
        Dim pProjections As IJElements
        Dim iElements As IJElements
        Dim pObj As IJSurface
        Dim StPt As New DPosition
        Dim endpt As New DPosition
        Dim i As Integer
        Dim j As Integer
        Dim outindex As Long
        Dim oLine As IngrGeom3D.Line3d
        Dim complex As IngrGeom3D.ComplexString3d
        
        Dim pOutCurve As ComplexString3d
        Dim CurveElems As IJElements
        Dim DepthCorr As Double

  
        Select Case cp
            Case 6, 8, 13, 14
                DepthCorr = -Depth / 2
            Case 1
                DepthCorr = Depth
            Case 2, 4, 11, 12
                DepthCorr = Depth / 2
            Case 3, 5, 7, 10, 15
            Case 9
                DepthCorr = -Depth
        End Select

        HandRailSpacing = PostHt / (NumMidRails + 1)

        Set StPt = ULC.Clone
        Set endpt = ULC.Clone
        endpt.y = endpt.y - TopLandingLength

        '1
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x + DepthCorr, StPt.y, StPt.z + PostHt, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt)
        Set iElements = New JObjectCollection 'IMSElements.DynElements
        iElements.Add oLine
        Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
        Set iElements = Nothing

        '2
        Set StPt = ULC.Clone
        StPt.y = StPt.y - TopLandingLength
        Set endpt = LLC.Clone
        oLine.DefineBy2Points StPt.x + DepthCorr, StPt.y, StPt.z + PostHt, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt
        complex.AddCurve oLine, True
        
        If CSObject Is Nothing Then
            complex.GetCurves CurveElems
            Set pOutCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, CurveElems)
            InitNewOutput pOC, "LeftHandRail"
            pOC.AddOutput "LeftHandRail", pOutCurve
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, complex, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
            For j = 1 To pProjections.Count
                Set pObj = pProjections(j)
                InitNewOutput pOC, "LeftHandRail" & j
                pOC.AddOutput "LeftHandRail" & j, pObj
            Next j
        End If

        If NumMidRails >= 1 Then
        DepthCorr = 0
            
        '3
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (NumMidRails)))
        Set iElements = New JObjectCollection 'IMSElements.DynElements
        iElements.Add oLine
        Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
        Set iElements = Nothing
        '        oLine.DefineBy2Points EndPt.x, EndPt.y, EndPt.z + PostHt, EndPt.x, EndPt.y, EndPt.z + PostHt - (HandRailSpacing * (NumMidRails - 1))
        '        complex.AddCurve oLine, True

        '4
        oLine.DefineBy2Points endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (NumMidRails)), StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (NumMidRails))
        complex.AddCurve oLine, True

        '5
        oLine.DefineBy2Points StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (NumMidRails)), StPt.x + DepthCorr, StPt.y + TopLandingLength, StPt.z + PostHt - (HandRailSpacing * (NumMidRails))
        complex.AddCurve oLine, True

        '6
        oLine.DefineBy2Points StPt.x + DepthCorr, StPt.y + TopLandingLength, StPt.z + PostHt - (HandRailSpacing * (NumMidRails)), StPt.x + DepthCorr, StPt.y + TopLandingLength, StPt.z + PostHt
        complex.AddCurve oLine, True

        If CSObject Is Nothing Then
            complex.GetCurves CurveElems
            Set pOutCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, CurveElems)
            InitNewOutput pOC, "LeftHandRail"
            pOC.AddOutput "LeftHandRail", pOutCurve
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, complex, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
            outindex = 1
            For j = 1 To pProjections.Count
                Set pObj = pProjections(j)
                InitNewOutput pOC, "LeftHandRail" & outindex
                pOC.AddOutput "LeftHandRail" & Trim$(Str$(outindex)), pObj
                outindex = outindex + 1
            Next j
        End If
            Set complex = Nothing
            Set oLine = Nothing
        End If

'        MiddleRails on Left

        DepthCorr = 0
         Select Case cp
            Case 6, 8, 13, 14
                DepthCorr = -Depth / 2
            Case 1
                DepthCorr = Depth
            Case 2, 4, 11, 12
                DepthCorr = Depth / 2
            Case 3, 5, 7, 10, 15
            Case 9
                DepthCorr = -Depth
        End Select

        If NumMidRails >= 2 Then
            For i = 2 To NumMidRails
                
                Set StPt = ULC.Clone
                Set endpt = ULC.Clone
                endpt.y = endpt.y - TopLandingLength
                
                Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (i - 1)), endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (i - 1)))
                Set iElements = New JObjectCollection 'IMSElements.DynElements
                iElements.Add oLine
                Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
                Set iElements = Nothing

                Set StPt = ULC.Clone
                Set endpt = LLC.Clone
                StPt.y = StPt.y - TopLandingLength

                oLine.DefineBy2Points StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (i - 1)), endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (i - 1))
                complex.AddCurve oLine, True

            If CSObject Is Nothing Then
                complex.GetCurves CurveElems
                Set pOutCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, CurveElems)
                InitNewOutput pOC, "LHandRail"
                pOC.AddOutput "LHandRail", pOutCurve
            Else
                Set pProjections = CreateProjectionFromCSProfile(Nothing, complex, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
                For j = 1 To pProjections.Count
                    Set pObj = pProjections(j)
                    InitNewOutput pOC, "LHandRail" & j
                    pOC.AddOutput "LHandRail" & j, pObj
                Next j
            End If

            Next i
       End If


        Set StPt = URC.Clone
        Set endpt = URC.Clone
        endpt.y = endpt.y - TopLandingLength
        DepthCorr = 0
        'TopRail
        Select Case cp
            
            Case 1, 5, 9, 10, 15
                'Nothing
            Case 2, 6, 11, 13
                DepthCorr = -Depth / 2
            Case 3
                DepthCorr = -Depth
            Case 4, 8, 12, 14
                DepthCorr = Depth / 2
            Case 7
                DepthCorr = Depth

        End Select
        
        '1
        Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x + DepthCorr, StPt.y, StPt.z + PostHt, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt)
        Set iElements = New JObjectCollection 'IMSElements.DynElements
        iElements.Add oLine
        Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
        Set iElements = Nothing

        '2
        Set StPt = URC.Clone
        StPt.y = StPt.y - TopLandingLength
        Set endpt = LRC.Clone
        oLine.DefineBy2Points StPt.x + DepthCorr, StPt.y, StPt.z + PostHt, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt
        complex.AddCurve oLine, True
        If CSObject Is Nothing Then
            complex.GetCurves CurveElems
            Set pOutCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, CurveElems)
            InitNewOutput pOC, "RightHandRail"
            pOC.AddOutput "RightHandRail", pOutCurve
        Else
            Set pProjections = CreateProjectionFromCSProfile(Nothing, complex, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
            For j = 1 To pProjections.Count
                Set pObj = pProjections(j)
                InitNewOutput pOC, "RightHandRail" & j
                pOC.AddOutput "RightHandRail" & j, pObj
            Next j
        End If
        
        If NumMidRails >= 1 Then
            DepthCorr = 0
            'BottomRail
            Select Case cp
            Case 4, 5, 6, 10, 13, 15
                'Nothing
            Case 1, 2, 3, 11
                DepthCorr = -Depth
            Case 7, 8, 9, 14
                DepthCorr = Depth
            End Select
            '3
            Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt, endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (NumMidRails)))
            Set iElements = New JObjectCollection 'IMSElements.DynElements
            iElements.Add oLine
            Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
            Set iElements = Nothing

            '4
            oLine.DefineBy2Points endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (NumMidRails)), StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (NumMidRails))
            complex.AddCurve oLine, True

            '5
            oLine.DefineBy2Points StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (NumMidRails)), StPt.x + DepthCorr, StPt.y + TopLandingLength, StPt.z + PostHt - (HandRailSpacing * (NumMidRails))
            complex.AddCurve oLine, True

            '6
            oLine.DefineBy2Points StPt.x + DepthCorr, StPt.y + TopLandingLength, StPt.z + PostHt - (HandRailSpacing * (NumMidRails)), StPt.x + DepthCorr, StPt.y + TopLandingLength, StPt.z + PostHt
            complex.AddCurve oLine, True

            If CSObject Is Nothing Then
                complex.GetCurves CurveElems
                Set pOutCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, CurveElems)
                InitNewOutput pOC, "RightHandRail"
                pOC.AddOutput "RightHandRail", pOutCurve
            Else
                Set pProjections = CreateProjectionFromCSProfile(Nothing, complex, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
                For j = 1 To pProjections.Count
                    Set pObj = pProjections(j)
                    InitNewOutput pOC, "RightHandRail" & j
                    pOC.AddOutput "RightHandRail" & j, pObj
                Next j
            End If

        End If

        'MiddleRails on Right
        DepthCorr = 0
        'MiddleRail
        Select Case cp
            Case 1, 5, 9, 10, 15
                'Nothing
            Case 2, 6, 11, 13
                DepthCorr = -Depth / 2
            Case 3
                DepthCorr = -Depth
            Case 4, 8, 12, 14
                DepthCorr = Depth / 2
            Case 7
                DepthCorr = Depth

        End Select
        
        If NumMidRails >= 2 Then
            For i = 2 To NumMidRails
                
                Set StPt = URC.Clone
                Set endpt = URC.Clone
                endpt.y = endpt.y - TopLandingLength
        
                Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (i - 1)), endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (i - 1)))
                Set iElements = New JObjectCollection ' IMSElements.DynElements
                iElements.Add oLine
                Set complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
                Set iElements = Nothing

                Set StPt = URC.Clone
                Set endpt = LRC.Clone
                StPt.y = StPt.y - TopLandingLength

                oLine.DefineBy2Points StPt.x + DepthCorr, StPt.y, StPt.z + PostHt - (HandRailSpacing * (i - 1)), endpt.x + DepthCorr, endpt.y, endpt.z + PostHt - (HandRailSpacing * (i - 1))
                complex.AddCurve oLine, True

                
                If CSObject Is Nothing Then
                    complex.GetCurves CurveElems
                    Set pOutCurve = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, CurveElems)
                    InitNewOutput pOC, "RightHandRail"
                    pOC.AddOutput "RightHandRail", pOutCurve
                Else
                    Set pProjections = CreateProjectionFromCSProfile(Nothing, complex, CSObject, cp, RotAngle, 0, Nothing, Nothing, skinOption)
                    For j = 1 To pProjections.Count
                        Set pObj = pProjections(j)
                        InitNewOutput pOC, "RightHandRail" & j
                        pOC.AddOutput "RightHandRail" & j, pObj
                    Next j
                End If
            Next i
        End If

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub


Private Sub InitNewOutput(pOC As IJDOutputCollection, name As String)
Const METHOD = "InitNewOutput"
On Error GoTo ErrorHandler
    
    Dim oRep As IJDRepresentation
    Dim oOutputs As IJDOutputs
    Dim oOutput As IJDOutput
    Set oOutput = New DOutput
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
    Set oOutputs = oRep

    oOutput.name = name
    oOutput.Description = name
    oOutputs.SetOutput oOutput
    oOutput.Reset
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub
 
Private Sub CreateRectangularProjection(pOC As IJDOutputCollection, centerPos As IJDPosition, _
                                        projvect As IJDVector, _
                                        upvect As IJDVector, _
                                        RectWidth As Double, _
                                        RectHeight As Double, _
                                        stepOrBar As String, _
                                        stepNumber As Integer)
Const METHOD = "CreateRectangularProjection"
On Error GoTo ErrorHandler

    Dim pProj3d As Projection3d
    Dim Tmpvect As IJDVector, tmpprojvect As IJDVector
    
    Set pProj3d = CreateUnitRectProjection(pOC, RectWidth, RectHeight, projvect.Length) 'Create a unit length projection
    
    Set tmpprojvect = projvect.Clone 'So we dont alter input vector
    tmpprojvect.Length = 1#
    upvect.Length = 1#
    Set Tmpvect = tmpprojvect.Cross(upvect)
    Tmpvect.Length = 1#
    
    Dim m_pIJDT4x4 As IJDT4x4
    Set m_pIJDT4x4 = New DT4x4
    m_pIJDT4x4.IndexValue(0) = Tmpvect.x
    m_pIJDT4x4.IndexValue(1) = Tmpvect.y
    m_pIJDT4x4.IndexValue(2) = Tmpvect.z
        
    m_pIJDT4x4.IndexValue(4) = upvect.x
    m_pIJDT4x4.IndexValue(5) = upvect.y
    m_pIJDT4x4.IndexValue(6) = upvect.z
    
    m_pIJDT4x4.IndexValue(8) = tmpprojvect.x
    m_pIJDT4x4.IndexValue(9) = tmpprojvect.y
    m_pIJDT4x4.IndexValue(10) = tmpprojvect.z
    
    m_pIJDT4x4.IndexValue(12) = centerPos.x
    m_pIJDT4x4.IndexValue(13) = centerPos.y
    m_pIJDT4x4.IndexValue(14) = centerPos.z
    
    pProj3d.Transform m_pIJDT4x4
    
    pOC.AddOutput stepOrBar & Trim$(Str$(stepNumber)), pProj3d
    
    Set tmpprojvect = Nothing
    Set pProj3d = Nothing
    Set Tmpvect = Nothing

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Function CreateUnitRectProjection(pOC As IJDOutputCollection, RectWidth As Double, RectHeight As Double, projlength As Double) As IJProjection
Const METHOD = "CreateUnitRectProjection"
On Error GoTo ErrorHandler

    Dim pts(15) As Double
    InitRectCurvePoints pts, RectWidth, RectHeight 'Build points in local XY plane at the centroid of the rectangle
    Dim m_pIJLineString As IJLineString

    Set m_pIJLineString = New LineString3d
    m_pIJLineString.SetPoints 5, pts    'Init Points
    
    Set CreateUnitRectProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                                    m_pIJLineString, 0#, 0#, 1#, projlength, True)
    
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Private Sub InitRectCurvePoints(pts() As Double, Width As Double, Height As Double)
Const METHOD = "InitRectCurvePoints"
On Error GoTo ErrorHandler

    'Build points in local XY plane at the centroid of the rectangle
    pts(0) = -(Width / 2#)
    pts(1) = -(Height / 2#)
    pts(2) = 0
    
    pts(3) = Width / 2#
    pts(4) = -(Height / 2#)
    pts(5) = 0
    
    pts(6) = Width / 2#
    pts(7) = Height / 2#
    pts(8) = 0
    
    pts(9) = -(Width / 2#)
    pts(10) = Height / 2#
    pts(11) = 0
    
    pts(12) = pts(0) 'Same as first point for closed shape
    pts(13) = pts(1)
    pts(14) = pts(2)

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

'Graphics for Operational Representation
Public Sub OperationalSub(pIRepSCM As IJDRepresentationStdCustomMethod)
On Error GoTo ErrorHandler
Const METHOD = "OperationalSub"

    Dim pRepDG As IJDRepresentationDuringGame
    Set pRepDG = pIRepSCM
    Dim pOC As IJDOutputCollection
    Set pOC = pRepDG.outputCollection
    Dim pInputs As IJDInputs
    Set pInputs = pRepDG.definition.IJDInputs
    Dim Span#, Height#, Width#, Justify#
    Dim Angle As Double, Length As Double
    Dim WithTopLanding As Boolean
    Dim TopLandingLength As Double
    Dim PostHt As Double, EnvHeight As Double
    Dim SideFrameReferenceStandard As String, SideFrameSection As String
    Dim oOutputName As String
    Dim stepNumber As Integer
 
    Height = pInputs.GetInputByIndex(HEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
    Length = pInputs.GetInputByIndex(LENGTH_INDEX).IJDInputDuringGame.Result.UomValue
    Span = pInputs.GetInputByIndex(SPAN_INDEX).IJDInputDuringGame.Result.UomValue
     
    Width = pInputs.GetInputByIndex(WIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    Angle = pInputs.GetInputByIndex(ANGLE_INDEX).IJDInputDuringGame.Result.UomValue
    PostHt = pInputs.GetInputByIndex(POSTHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue

    SideFrameSection = pInputs.GetInputByIndex(SIDEFRAMESECTION_INDEX).IJDInputDuringGame.Result.String
    SideFrameReferenceStandard = pInputs.GetInputByIndex(SIDEFRAMESECTIONSTANDARD_INDEX).IJDInputDuringGame.Result.String
    WithTopLanding = pInputs.GetInputByIndex(WITHTOPLANDING_INDEX).IJDInputDuringGame.Result.UomValue
    TopLandingLength = pInputs.GetInputByIndex(TOPLANDINGLENGTH_INDEX).IJDInputDuringGame.Result.UomValue
    EnvHeight = pInputs.GetInputByIndex(ENVELOPEHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
   
    Justify = pInputs.GetInputByIndex(JUSTIFICATION_INDEX).IJDInputDuringGame.Result.UomValue
    
    Dim xVect As IJDVector
    Dim yVect As IJDVector
    Dim zVect As IJDVector
    Set xVect = New DVector
    Set yVect = New DVector
    Set zVect = New DVector
    Dim TotalWidth As Double, TotalHeight As Double
    Dim Pos As Double
    Dim oGeom3DFactory As New GeometryFactory
    Dim oLineStrings3d As ILineStrings3d
    Dim oLineString As IJLineString
    Dim oProjection3d As IProjections3d
    Dim oProjection As IJProjection
    Dim pts(15) As Double
    Dim dLandingLength As Double
    
    Pos = 0#
    dLandingLength = 0#
    
    If Justify = 2 Then
        Pos = -(Width / 2)
    ElseIf Justify = 3 Then
        Pos = (Width / 2)
    End If
    
    Dim Depth As Double
    Dim SecWidth As Double
        
    Depth = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "ISTRUCTCrossSectionDimensions", "Depth")
    SecWidth = GetCSAttribData(SideFrameSection, SideFrameReferenceStandard, "ISTRUCTCrossSectionDimensions", "Width")

    'output name aand step number hardcoded
    oOutputName = "OperationalEnvelope"
    stepNumber = 1
    
    'we will create the side of the stair and then project it along thw width to cretae the envelope
    ' all the points for the side will be in local xy, assuming width to be along the z -axis.
    If WithTopLanding Then
        'dLandingLength value is 0 w/o landing
        dLandingLength = TopLandingLength
    End If
    
    'Create automation I/Fs for geometry objects
    Set oLineStrings3d = oGeom3DFactory.LineStrings3d
    Set oProjection3d = oGeom3DFactory.Projections3d
    
    'Total width is 2 times the HR section width plus the stair width.
    'Heinht is just envelope height as the origin is on the top of the stair step and the operation
    ' envelop is calculated from there.
    TotalWidth = Width + 2 * SecWidth
    TotalHeight = EnvHeight

    'We start at the top of the stair and go down along the stair. Origin at the top right- both w and w/o landing
    '                       ______
    '                   -         |
    '               -       ______|
    '           -       -
    '       -       -
    '   |       -
    '   |   -
    
    pts(0) = 0#
    pts(1) = 0#
    pts(2) = 0#
    pts(3) = -Length * Cos(Angle)
    pts(4) = -Length * Sin(Angle)
    pts(5) = 0#
    Set oLineString = oLineStrings3d.CreateByPoints(Nothing, 2, pts)
    
    oLineString.AddPoint -Length * Cos(Angle), -Length * Sin(Angle) + TotalHeight, 0#
    oLineString.AddPoint 0#, TotalHeight, 0#
    
    If WithTopLanding Then
        oLineString.AddPoint dLandingLength, TotalHeight, 0#
        oLineString.AddPoint dLandingLength, 0#, 0#
    End If
    oLineString.AddPoint 0#, 0#, 0#     'Close out the line string
    
    'Now project it along z
    Set oProjection = oProjection3d.CreateByCurve(Nothing, _
                                                    oLineString, 0#, 0#, 1#, TotalWidth, True)
    
    ' x-axis is the y-axis on the projection
    xVect.x = 0#
    xVect.y = 1#
    xVect.z = 0#
    xVect.Length = 1#
    
    'Global Y axis is the z axis on the projection
    yVect.x = 0#
    yVect.y = 0#
    yVect.z = 1#
    yVect.Length = 1#
    
    'Global z axis is the x axis on the projection
    zVect.x = 1#
    zVect.y = 0#
    zVect.z = 0#
    zVect.Length = 1#
    
    'Now apply the xform to correctly orient the envelope
    Dim oT4x4 As IJDT4x4
    Set oT4x4 = New DT4x4
    oT4x4.IndexValue(0) = xVect.x
    oT4x4.IndexValue(1) = xVect.y
    oT4x4.IndexValue(2) = xVect.z
        
    oT4x4.IndexValue(4) = yVect.x
    oT4x4.IndexValue(5) = yVect.y
    oT4x4.IndexValue(6) = yVect.z
    
    oT4x4.IndexValue(8) = zVect.x
    oT4x4.IndexValue(9) = zVect.y
    oT4x4.IndexValue(10) = zVect.z
    
    'Taransfrom the x-axis to move the origin relative to the side and to include the justify
    oT4x4.IndexValue(12) = -TotalWidth / 2# - Pos
    'Trasnform he y-axis to include the landinglength
    oT4x4.IndexValue(13) = 0# - dLandingLength
    oT4x4.IndexValue(14) = 0#
    
    oProjection.Transform oT4x4
    
    pOC.AddOutput oOutputName & Trim$(Str$(stepNumber)), oProjection

    Set oGeom3DFactory = Nothing
    Set pRepDG = Nothing
    Set pOC = Nothing
    Set pInputs = Nothing
    Set xVect = Nothing
    Set yVect = Nothing
    Set zVect = Nothing
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)

End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)

End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_STAIRMACROS_ERROR, "ERROR")
    
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

    IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_STAIRMACROS_ERROR, "ERROR")

    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String
    
    Set pAttrColl = CollAllDisplayedValues
         
    For i = 1 To pAttrColl.Count
        Set pAttrDescr = pAttrColl.Item(i)
            ErrStr = IJUserAttributeMgmt_OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrDescr, pAttrDescr.AttrValue)
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

' first of all check if the symbol definition has CMCheck methods defined - TBD
    UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_STAIRMACROS_ERROR, "ERROR")

    Dim dInputs As IJDInputs
    Dim CurrentInput As IJDInput
    Dim oAttribute As IJDAttribute
    Dim PC As DParameterContent
    Dim bvalid As Boolean
    Dim oSymbolOcc As IJDSymbol
    
    If Not pIJDAttrs Is Nothing Then
        Set oSymbolOcc = pIJDAttrs
        Dim oSymbolDef As IJDSymbolDefinition
        Dim ErrMessage As String
        Set oSymbolDef = oSymbolOcc.IJDSymbolDefinition(2)
        Set dInputs = oSymbolDef.IJDInputs
        Set PC = New DParameterContent
        
        Set oAttribute = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item(sAttributeName)
    
        If oAttribute.Value <> "" Then
            If oAttribute.AttributeInfo.Type = igString Then    ' check for string type here
            Else
                PC.UomValue = oAttribute.Value
                Set CurrentInput = Nothing
                bvalid = True
                On Error Resume Next
                Set CurrentInput = dInputs.GetInputByName(oAttribute.AttributeInfo.name)
                If Not CurrentInput Is Nothing Then
                    CurrentInput.IJDInputDuringGame.definition = oSymbolDef
                    CurrentInput.IJDInputStdCustomMethod.InvokeCMCheck PC, bvalid, ErrMessage
                    CurrentInput.IJDInputDuringGame.definition = Nothing
                    Set oSymbolOcc = Nothing
                    Set oSymbolDef = Nothing
                    If bvalid = False Then
'                    UserAttributeMgmt_Validate = "Symbol CMCheck Failed"
                    UserAttributeMgmt_Validate = ErrMessage
                    Exit Function
                    Else
                    End If
                End If
                On Error GoTo ErrorHandler
            End If
        End If
        
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
            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_STAIRMACROS_SCHEMAERROR, "SchemaERROR")
            GoTo ErrorHandler
        End If
        Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
        
        ' loop on the attributes on the interface to match the supplied attribute type
        For AttrCount = 1 To oAttrCol.Count
            Set oAttrObj = oAttrCol.Item(AttrCount)
            If oAttrObj.name = sAttributeName Then
                If (sAttributeName = "PlatformThickness") Or (sAttributeName = "NumMidRails") Or (sAttributeName = "EnvelopeHeight") Then
                    If (varAttributeValue <= 0#) Then
                        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_STAIRMACTOS_INVALID_ATTRIB_VALUE, "Invalid Attribute Value ")
                        Set oAttributeMetaData = Nothing
                        Exit Function
                    End If
                End If
            End If
        Next
        
        Set oAttributeMetaData = Nothing
        
    Else
        If sInterfaceName = "IJSPSCommonStairLadderProps" And sAttributeName = "Width" Then
            If (varAttributeValue <= 0) Or (varAttributeValue > 100) Then
                UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_STAIRMACROS_INVALID_WIDTH, "Invalid Width value")
                Exit Function
            End If
        End If
        If sInterfaceName = "IJSPSCommonStairLadderProps" And sAttributeName = "Span" Then
            If (varAttributeValue > 100) Then
                UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_STAIRMACROS_INVALID_OFFSET, "Invalid offset value")
                Exit Function
            End If
        End If
        If sInterfaceName = "IJSPSCommonStairLadderProps" And sAttributeName = "StepPitch" Then
            If (varAttributeValue <= 0) Or (varAttributeValue > 100) Then
                UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_STAIRMACROS_INVALID_PITCH, "Invalid pitch value")
                Exit Function
            End If
        End If
        If sInterfaceName = "IJSPSCommonStairLadderProps" And sAttributeName = "Angle" Then
            Dim dValue As Double
            dValue = Round(varAttributeValue, 4)
            If (varAttributeValue < 0.4363) Or (varAttributeValue > 1.1345) Then  ' less than 25 deg or greater than 65 deg
                UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_STAIRMACTOS_INVALID_ATTRIB_VALUE, "Invalid angle value")
                Exit Function
            End If
        End If
    End If
    UserAttributeMgmt_Validate = ""
    
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

Private Sub ISPSMirrorHelper_Mirror(ByVal oPlane As IJPlane, ByVal oMatrix As IJDT4x4, ByVal bCopy As Boolean, ByVal oSource As Object, ByVal oTarget As Object)
    Const METHOD = "ISPSMirrorHelper_Mirror"
    On Error GoTo ErrorHandler
    
    'check if source and target objects are valid
    If oSource Is Nothing Or oTarget Is Nothing Then
        Exit Sub
    End If
    
    'source vars
    Dim oPartOccSource As IJPartOcc
    Dim oPartSource As IJDPart
    Dim oAttrsOccSource As IJDAttributes
    Dim oAttrsPartSource As IJDAttributes
    Dim oStairSource As ISPSStair
    
    'target vars
    Dim oPartOccTarget As IJPartOcc
    Dim oPartTarget As IJDPart
    Dim oAttrsOccTarget As IJDAttributes
    Dim oAttrsPartTarget As IJDAttributes
    Dim lMirrorBehavior As Long
    Dim oInfosColTargetPart As IJDInfosCol
    Dim oInfosColTargetOcc As IJDInfosCol
    Dim oStairTarget As ISPSStair
        
    Set oStairSource = oSource
    Set oPartOccSource = oSource
    'get occ attrs
    Set oAttrsOccSource = oPartOccSource
    
    If Not oPartOccSource Is Nothing Then
        'Now get the part definition attrs
        oPartOccSource.GetPart oPartSource
        Set oAttrsPartSource = oPartSource
    End If
     
    Set oStairTarget = oTarget
    'get occ and part from the target
    Set oPartOccTarget = oTarget
    'get occ attrs
    Set oAttrsOccTarget = oPartOccTarget
    
    If Not oPartOccTarget Is Nothing Then
        'Now get the part definition attrs
        oPartOccTarget.GetPart oPartTarget
        Set oAttrsPartTarget = oPartTarget
    End If
    
    'let's get the mirror behavior from the part attrs- 5 maybe mirrored, 50 replacement part, 15 can not mirror
    GetOccAttribInfoColl oSource, False, oInfosColTargetPart
    lMirrorBehavior = GetAttribute(oAttrsPartTarget, "MirrorBehaviorOption", oInfosColTargetPart)
    
    'now lets start massaging the occ attrs of the target
    GetOccAttribInfoColl oTarget, True, oInfosColTargetOcc
    
    'if we are replacing part, need to set the hoopopening attr back to def attr as the part occ mirror
    ' copies the original part's occ attrs over to the replaced part and we don't want the hoopopening to be copied
    If lMirrorBehavior = 50 Then
        Dim lHoopOpeningPart As Long
        Dim lHoopOpeningOcc As Long
        lHoopOpeningPart = GetAttribute(oAttrsPartTarget, "HoopOpening", oInfosColTargetPart)
        lHoopOpeningOcc = GetAttribute(oAttrsOccTarget, "HoopOpening", oInfosColTargetOcc)
        
        If Not lHoopOpeningPart = lHoopOpeningOcc Then
            SetAttribute oAttrsOccTarget, lHoopOpeningPart, "HoopOpening", oInfosColTargetOcc
        End If
    End If

    Dim bTopSupportSide As Boolean
    'gg 11/14/2008 TR-120786 Don't change the support side if the target top-edge is slab
    'The best way to unify the mirror copy behaviors of Slab with other objects are to consider the handness of the models from G&T.
    'The current fix can only be considered as a short term fix and it looks like working for the current code base.
    'If for some reason the fix fails in future, there's a suggestion to look more details on slab object to determine the support side direction:
    'Broad plan:
    '1.  If edge is from Slab, get the 2 faces to which the edge belongs.
    '2.  Get the face whose normal is most aligned along the x(out) vector of the ladder xform matrix.
    '3.  Check if the xvector is along or opposite to the face normal. This is important as this is what we expect on the mirrored ladder too.
    '4. Repeat 1,2 on the resulting ladder and slab to get the face.
    '5. Now, get the mirrored ladder's x-vector and see if it is along or opposite to the resulting face. Tweak the side flag to fix it.

    If Not ConnectedToSlab(oStairTarget) Then
        bTopSupportSide = GetAttribute(oAttrsOccTarget, "TopSupportSide", oInfosColTargetOcc)
        bTopSupportSide = Not bTopSupportSide
        SetAttribute oAttrsOccTarget, bTopSupportSide, "TopSupportSide", oInfosColTargetOcc
    End If

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description

End Sub

Private Function ConnectedToSlab(oStair As ISPSStair) As Boolean
    Const METHOD = "ConnectedToSlab"
    On Error GoTo ErrorHandler
    ConnectedToSlab = False
    Dim oPort As IJPort
    Dim oConnectable As IJConnectable
    If TypeOf oStair.TopEdge Is IJPort Then
        Set oPort = oStair.TopEdge
    End If
    If Not oPort Is Nothing Then
        Set oConnectable = oPort.Connectable
        If Not oConnectable Is Nothing Then
            If TypeOf oConnectable Is SPSSlabEntity Then
                ConnectedToSlab = True
            End If
        End If
    End If
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description

End Function

