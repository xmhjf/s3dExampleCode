VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LadderTypeA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************************************
'  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
'
'  Project: \Symbols\Ladders\LadderMacros.vbp
'  File:  \Symbols\Ladder\VerticalLadder.cls
'
'  Description:  Macro for Drawing Ladder given a Height.
'                Symbol inputs correspond to the Vertical Ladder in catalog.
'
'  Author: Structure Team
'  History:
'  CE  12/05/99   Updates for real vertical ladder geometry
'  JM 02/04/2000 Added check functions to allow the attributes value to be checked for
'                           Validity before being commited.
'  JM 03/09/2000 Started implimenting LadderSymbolService which in essence moved the semantic
'                           into the symbol.  (For setting Origin, height, and other driven OAs.)
'  JM 03/10/2000 Added Weight and CG
'  JM 04/04/2000 Angle input is now used.  CMCheck for between 75 and 90 degrees.
'  JM 04/18/2000 Changed EvaluateGeometry.  It now tells you what has caused the recompute so that
'                           the recompute can be as costless as possible.
'
'  Jan 2002 AP    V3 Additions new Functionality
'  M        02/02/2003 TR-CP·52149  Hoop geometry calculations for all ladder widths
'  JS  09/20/2005  TR#84847 - Shortened the name of the interface IJSPSCommonStairLadderProperties
'                           to IJSPSCommonStairLadderProps (Oracle requirement)
'  AS  Dec-13-2006 TR#111679 - Corrected error reporting in evaluate geometry
'  AS  Apr-17-2007 CR#69885  - Added support for Side Step Ladder. Added new routines to calculate
'                           hoop and vertical strap geometry in such cases.
'  AS  May-05-2007 CR#116079 Added custom Mirror copy implementation through ISPSMirrorHelper I/F
'  AS  Jun-07-2007 TR#120500 Added code to check the values of codelisted atributes
'  AS Aug-10-2007   DI -CP·123592. Removed flip condition for gridlines as they are reversed now
'  WR Jul-11-2008   TR-CP-144430 - Changing vertical offset of ladder results in incorrect geometry.
'                                  Vertical offset was not taken into account when creating the ladder.
'                                  Added code to account for the vertical offset.
'  GG Jul-31-2008   TR-135471    - Always change the support side for mirror copy no matter with or without safety hoop
'  GG Aug-14-2008   TR-148460    - Align the top of the rung with the top reference surface.
'  GG Aug-15-2008   TR-148191    - Fixed the E/W mirror copy and cleaned up the code in ISPSMirrorHelper_Mirror.
'  GG Sept-09-2008  TR-148134    - Fixed the First Step Pitch and Step Number calculation.
'  GG Dec-09-2008   TR-120786    - Don't change the support side if the target top-edge is slab
'  GG Mar-31-2009   DM#162555    - To make sure the first time weight calculation gives correct weight.
'  GG May-08-2009   DM#162555    - Fixed name mismatch for FlareSafetyHoop.
'******************************************************************

Const PI As Double = 3.14159265358979

Private Const MODULE = "SPSLadderMacros.LadderTypeA"
Const defProgId As String = "SPSLadderMacros.LadderTypeA"
Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"

Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259
Private Const REL_OUT_DATE = &H80000

Private METHOD As String
Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_GeomFactory As IngrGeom3D.GeometryFactory
Private m_pIJDT4x4 As IJDT4x4
Private m_pIJLineString As IJLineString
Private m_complex As IngrGeom3D.ComplexString3d
Private m_Hoopcomplex As IngrGeom3D.ComplexString3d
Private m_Flarecomplex As IngrGeom3D.ComplexString3d
Private m_HoopArc As Arc3d
Private m_FlareArc As Arc3d
Private m_oLocalizer As IJLocalizer

'Careful we can't just change these enums and have everthing work below....Chad
Private Enum InputIndex
PART_INDEX = 1
WIDTH_INDEX
ANGLE_INDEX
STEPPITCH_INDEX
SUPPORTLEGPITCH_INDEX '5
SUPPORTLEGWIDTH_INDEX
SUPPORTLEGTHICKNESS_INDEX
SIDEFRAMEWIDTH_INDEX
SIDEFRAMETHICKNESS_INDEX
STEPDIAMETER_INDEX '10
VLDIM1_INDEX
VLDIM2_INDEX
VLDIM3_INDEX
WALLOFFSET_INDEX '15
SPAN_INDEX
HEIGHT_INDEX
LENGTH_INDEX
WITHWALLSUPPORTS_INDEX
NUMSTEPS_INDEX
STEPPROTRUSION_INDEX '20
WITHSAFETYHOOP_INDEX
HOOPPITCH_INDEX
BOTTOMHOOPLEVEL_INDEX
HOOPCLEARANCE_INDEX '25
HOOPRADIUS_INDEX
HOOPPLATETHICKNESS_INDEX
HOOPPLATEWIDTH_INDEX
HOOPBENDRADIUS_INDEX
HOOPOPENING_INDEX
SHDIM1_INDEX
SHDIM2_INDEX '30
SHDIM3_INDEX
FLARECLEARANCE_INDEX     'New additions
FLARERADIUS_INDEX     'New additions
HOOPFLAREBENDRADIUS_INDEX
FLARESHDIM1_INDEX
FLARESHDIM2_INDEX '35
FLARESHDIM3_INDEX
HOOPFLAREHEIGHT_INDEX
HOOPFLAREMAXHEIGHT_INDEX
VERTICALSTRAPWIDTH_INDEX '40
VERTICALSTRAPTHICKNESS_INDEX
VERTICALSTRAPCOUNT_INDEX
TOPEXTENSION_INDEX
BOTTOMEXTENSION_INDEX
JUSTIFICATION_INDEX
TOPSUPPORTSIDE_INDEX
ISASSEMBLY_INDEX
ENVELOPEHEIGHT_INDEX
PRIMARYMATERIAL_INDEX
PRIMARYGRADE_INDEX
End Enum

Implements IJDUserSymbolServices
Implements ISPSLadderSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck
Implements ISPSMirrorHelper

Private Const TOL = 0.0000001

Dim bOnPreLoad As Boolean

Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)

End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)

End Sub

Private Sub ISPSLadderSymbolServices_EvaluateGeometry(ByVal LadderBo As SPSLadders.ISPSLadder, _
                                                    ByVal bPartOccChanged As Boolean, _
                                                    ByVal Top As Object, _
                                                    ByVal bTopChanged As Boolean, _
                                                    ByVal bottom As Object, _
                                                    ByVal bBottomChanged As Boolean, _
                                                    ByVal Ref As Object, _
                                                    ByVal bRefChanged As Boolean, _
                                                    ByVal PartInfoCol As SPSLadders.IJDInfosCol, _
                                                    ByVal bAttributeChanged As Boolean, _
                                                    ByVal PartOccInfoCol As SPSLadders.IJDInfosCol, _
                                                    ByVal bOccAttributeChanged As Boolean)
Const METHOD = "ISPSLadderSymbolServices_EvaluateGeometry"
On Error GoTo ErrorHandler
    
    Dim Height As Double
    Dim Length As Double
    Dim OldHeight As Double
    Dim Span As Double
    Dim TopExt As Double
    Dim BottomExt As Double
    Dim StepPitch As Double
    Dim StepDia As Double
    Dim FirstStepPitch As Double
    Dim NumSteps As Integer
    Dim tempNum As Double
    Dim oPart As IJDPart
    Dim oPartOcc As IJPartOcc
    Dim Angle As Double
    
    'TR#CP52149 - mkaveti
    Dim ladderWidth As Double
    Dim HoopBendRadius As Double
    Dim SideFrameThick As Double
    Dim OldHoopRadius#, HoopRadius#, HoopClearance#
    Dim FlareRadius#, FlareClearance#
    Dim ShDim1#, ShDim2#, ShDim3#, FlareShDim1#, FlareShDim2#, FlareShDim3#
    
    Dim lJustify As Long
        
    Set oPartOcc = LadderBo
    oPartOcc.GetPart oPart
    Dim OccAttrs As IJDAttributes
    Dim PartAttrs As IJDAttributes
    Set OccAttrs = LadderBo
    Set PartAttrs = oPart

    OldHeight = GetAttribute(OccAttrs, "Height", PartOccInfoCol)
    
    'changes for checking undefined value
    lJustify = GetAttribute(OccAttrs, "Justification", PartOccInfoCol)
    CheckForUndefinedValueAndRaiseError LadderBo, lJustify, "StructAlignment", 71
    
    Dim SetOriginAndSpanValue As Long
    SetOriginAndSpanValue = SetOriginAndSpan(LadderBo, Top, bottom, Ref, PartOccInfoCol)
        
    If SetOriginAndSpanValue = E_FAIL Then
        Err.Raise E_FAIL
    ElseIf SetOriginAndSpanValue = LADDER_E_SIDE_REF_MISSING Then
        Err.Raise LADDER_E_SIDE_REF_MISSING
    End If
        
    TopExt = GetAttribute(OccAttrs, "TopExtension", PartOccInfoCol)
    BottomExt = GetAttribute(OccAttrs, "BottomExtension", PartOccInfoCol)
    Height = GetAttribute(OccAttrs, "Height", PartOccInfoCol)
    
    Angle = GetAttribute(OccAttrs, "Angle", PartOccInfoCol)
   
    If Angle <> 1.5707963267949 Then
        Angle = 1.5707963267949
        Call SetAttribute(OccAttrs, Angle, "Angle", PartOccInfoCol)
    End If
    
    'TR#CP52149 - mkaveti
    'Take the existing HoopRadius
    OldHoopRadius = GetAttribute(OccAttrs, "HoopRadius", PartOccInfoCol)
    
    'Get the parameters to calculate the new HoopRadius
    ShDim1 = GetAttribute1(PartAttrs, "ShDim1", PartInfoCol)
    ShDim2 = GetAttribute1(PartAttrs, "ShDim2", PartInfoCol)
    ShDim3 = GetAttribute1(PartAttrs, "ShDim3", PartInfoCol)
    FlareShDim1 = GetAttribute1(PartAttrs, "FlareShDim1", PartInfoCol)
    FlareShDim2 = GetAttribute1(PartAttrs, "FlareShDim2", PartInfoCol)
    FlareShDim3 = GetAttribute1(PartAttrs, "FlareShDim3", PartInfoCol)
    ladderWidth = GetAttribute(OccAttrs, "Width", PartOccInfoCol)
    HoopBendRadius = GetAttribute1(PartAttrs, "HoopBendRadius", PartInfoCol)
    SideFrameThick = GetAttribute(OccAttrs, "SideFrameThickness", PartOccInfoCol)
    
    'Calculate New HoopRadius
    HoopRadius = ShDim2 + (ladderWidth / 2) + (2 * HoopBendRadius) + SideFrameThick

    'Check whether the old HoopRadius is same or different to update the other
    '   depending parameters in property page   - mkaveti
    If OldHoopRadius <> HoopRadius Then
        OldHoopRadius = HoopRadius
        HoopClearance = HoopRadius + (ShDim3 - ShDim1)
        FlareRadius = FlareShDim2 + (ladderWidth / 2) + (2 * HoopBendRadius) + SideFrameThick
        FlareClearance = FlareRadius + (ShDim3 - FlareShDim1)

        Call SetAttribute(OccAttrs, OldHoopRadius, "HoopRadius", PartOccInfoCol)
        Call SetAttribute(OccAttrs, HoopClearance, "HoopClearance", PartOccInfoCol)
        Call SetAttribute(OccAttrs, FlareRadius, "FlareRadius", PartOccInfoCol)
        Call SetAttribute(OccAttrs, FlareClearance, "FlareClearance", PartOccInfoCol)
    
    End If
    
    If Not OldHeight = Height Or bPartOccChanged Or bOccAttributeChanged Or bTopChanged Or bBottomChanged Or bRefChanged Then
            Length = Height - BottomExt
            Call SetAttribute(OccAttrs, Length, "Length", PartOccInfoCol)
            StepPitch = GetAttribute(OccAttrs, "StepPitch", PartOccInfoCol)
            StepDia = GetAttribute(OccAttrs, "StepDiameter", PartOccInfoCol)
            
            tempNum = (Length - StepDia - TopExt) / StepPitch
            If tempNum = Int(tempNum) Then
                NumSteps = Int(tempNum)                 'If the pitch divides evenly into the Length don't add extra
            Else                                        'step.  Otherwise you will get a step at height 0.
                NumSteps = Int((tempNum) + 1)
                If NumSteps <= 0 Then
                    NumSteps = 1
                End If
            End If
            FirstStepPitch = (Length - TopExt) - (StepPitch * (NumSteps - 1))

            'set the part occurrence with any values that have been driven.
            Call SetAttribute(OccAttrs, NumSteps, "NumSteps", PartOccInfoCol)
            Call SetAttribute(OccAttrs, FirstStepPitch, "FirstStepPitch", PartOccInfoCol)
    End If
    
    Set oPart = Nothing
    Set oPartOcc = Nothing
    Set OccAttrs = Nothing
    Set PartAttrs = Nothing
    
    Exit Sub
    
ErrorHandler:
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    ElseIf Err.Number = LADDER_E_SIDE_REF_MISSING Then
        Err.Raise LADDER_E_SIDE_REF_MISSING
    Else
        m_oErrors.Add Err.Number, METHOD, Err.Description
    End If
End Sub

Private Sub ISPSLadderSymbolServices_EvaluateLocalWCG(ByVal oLadder As SPSLadders.ISPSLadder, _
                                                    ByVal PartInfoCol As SPSLadders.IJDInfosCol, _
                                                    ByRef weight As Double, _
                                                    ByRef COGX As Double, _
                                                    ByRef COGY As Double, _
                                                    ByRef COGZ As Double)
Const METHOD = "ISPSLadderSymbolServices_EvaluateLocalWCG"
On Error GoTo ErrorHandler

    Call CalcWCG(oLadder, PartInfoCol, weight, COGX, COGY, COGZ)

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub
  
Private Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = defProgId
End Function

'********************************************************************
' Routine: IJDUserSymbolServices_InstanciateDefinition
'
' Description:This instanciates a persistent symbol definition object
' and initializes it for the first time.
'********************************************************************
Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParameters As Variant, ByVal ActiveConnection As Object) As Object
Const METHOD = "InstanciateDefinition"
On Error GoTo ErrorHandler

    'Create a Symbol Definition Object.
    Dim pSymbolFactory As New DSymbolEntitiesFactory
    Dim pSymbolDefinition As IJDSymbolDefinition
    
    Set pSymbolDefinition = pSymbolFactory.CreateEntity(definition, ActiveConnection)
    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
   
    'Now Let's do some nomenclature stuff
    pSymbolDefinition.ProgId = defProgId
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.Name = pSymbolDefinition.ProgId
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED

    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition
    Set pSymbolFactory = Nothing
    Set pSymbolDefinition = Nothing
    
    Exit Function

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function


Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())

End Sub

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler

        ' Remove all previous Symbol Definition information
        
        pSymbolDefinition.IJDInputs.RemoveAllInput
        pSymbolDefinition.IJDRepresentations.RemoveAllRepresentation
        pSymbolDefinition.IJDRepresentationEvaluations.RemoveAllRepresentationEvaluations
        Dim InputsIf As IMSSymbolEntities.IJDInputs
        Set InputsIf = pSymbolDefinition
        Dim pUM As IMSSymbolEntities.IJDUserMethods
        Set pUM = pSymbolDefinition
    
        ' set the CMCacheForPart method in the definition
        Dim libDesc As New DLibraryDescription
        Dim mCookie As Long
        Dim libCookie As Long
        
        libDesc.Name = "mySelfAsLib"
        libDesc.Type = imsLIBRARY_IS_ACTIVEX
        libDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
        libDesc.Source = defProgId
        pSymbolDefinition.IJDUserMethods.SetLibrary libDesc
        ' Get the lib/method cookie
        libCookie = libDesc.Cookie
        mCookie = pSymbolDefinition.IJDUserMethods.GetMethodCookie("CMCacheForPart", libCookie)
        
        
        Dim ChecklibDesc As New DLibraryDescription
        Dim ChecklibCookie As Long
        Dim GTZeroCheck As Long
        Dim LT90DegCheck As Long
'        Dim IsBoolean As Long
        Dim BT90And75Check As Long
        Dim NegativeNum As Long
        Dim CheckMaterial As Long
        
        ChecklibDesc.Name = "CMCheckLib"
        ChecklibDesc.Type = imsLIBRARY_IS_ACTIVEX
        ChecklibDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
        ChecklibDesc.Source = CheckProgId
        pSymbolDefinition.IJDUserMethods.SetLibrary ChecklibDesc
        ChecklibCookie = ChecklibDesc.Cookie
 
        
'Assign cookie values to point to the check functions.  Allows the assignment of check functions to attributes later.
        GTZeroCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("GTZero", ChecklibCookie)
        LT90DegCheck = pSymbolDefinition.IJDUserMethods.GetMethodCookie("LT90Deg", ChecklibCookie)
'        IsBoolean = pSymbolDefinition.IJDUserMethods.GetMethodCookie("IsBoolean", ChecklibCookie)
        BT90And75Check = pSymbolDefinition.IJDUserMethods.GetMethodCookie("BT90And75Check", ChecklibCookie)
        NegativeNum = pSymbolDefinition.IJDUserMethods.GetMethodCookie("NegativeNum", ChecklibCookie)
            CheckMaterial = pSymbolDefinition.IJDUserMethods.GetMethodCookie("IsMaterailGrade", ChecklibCookie)
'************/

' Define a new input by new operator
     Dim pIJDInput As IMSSymbolEntities.IJDInput
     Set pIJDInput = New IMSSymbolEntities.DInput
' Create a default value
     Dim PC As IMSSymbolEntities.DParameterContent
     Set PC = New IMSSymbolEntities.DParameterContent
     PC.Type = igValue
    
     Dim oSymbolCache As New CustomCache
     oSymbolCache.SetupCustomCache pSymbolDefinition
          
     pIJDInput.Name = "Width"
     pIJDInput.Description = "Width"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
'Assign the Attribute a check function.
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.035
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, WIDTH_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "Angle"
     pIJDInput.Description = "Angle"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, BT90And75Check
     PC.UomValue = 90
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, ANGLE_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.Name = "StepPitch"
     pIJDInput.Description = "StepPitch"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.25
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, STEPPITCH_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.Name = "SupportLegPitch"
     pIJDInput.Description = "SupportLegPitch"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.5
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SUPPORTLEGPITCH_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.Name = "SupportLegWidth"
     pIJDInput.Description = "SupportLegWidth"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.065
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SUPPORTLEGWIDTH_INDEX
     pIJDInput.Reset
     PC.Reset
            
     pIJDInput.Name = "SupportLegThickness"
     pIJDInput.Description = "SupportLegThickness"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.016
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SUPPORTLEGTHICKNESS_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.Name = "SideFrameWidth"
     pIJDInput.Description = "SideFrameWidth"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.065
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SIDEFRAMEWIDTH_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.Name = "SideFrameThickness"
     pIJDInput.Description = "SideFrameThickness"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.016
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SIDEFRAMETHICKNESS_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.Name = "StepDiameter"
     pIJDInput.Description = "StepDiameter"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.04
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, STEPDIAMETER_INDEX
     pIJDInput.Reset
     PC.Reset
       
     pIJDInput.Name = "VlDim1"
     pIJDInput.Description = "VlDim1"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.03
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, VLDIM1_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.Name = "VlDim2"
     pIJDInput.Description = "VlDim2"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.01
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, VLDIM2_INDEX
     pIJDInput.Reset
     PC.Reset
        
     pIJDInput.Name = "VlDim3"
     pIJDInput.Description = "VlDim3"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.165
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, VLDIM3_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "WallOffset"
     pIJDInput.Description = "WallOffset"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.135
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, WALLOFFSET_INDEX
     pIJDInput.Reset
     PC.Reset
         
     pIJDInput.Name = "Span"
     pIJDInput.Description = "Span"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SPAN_INDEX
     pIJDInput.Reset
     PC.Reset
     
       
     pIJDInput.Name = "Height"
     pIJDInput.Description = "Height"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HEIGHT_INDEX
     pIJDInput.Reset
     PC.Reset
    
     pIJDInput.Name = "Length"
     pIJDInput.Description = "Length"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, LENGTH_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "WithWallSupports"
     pIJDInput.Description = "WithWallSupports"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
'     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, IsBoolean
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, WITHWALLSUPPORTS_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "NumSteps"
     pIJDInput.Description = "NumSteps"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, NUMSTEPS_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "StepProtrusion"
     pIJDInput.Description = "StepProtrusion"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.005
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, STEPPROTRUSION_INDEX
     pIJDInput.Reset
     PC.Reset
          
    'Hoop parameters
         
     pIJDInput.Name = "WithSafetyHoop"
     pIJDInput.Description = "WithSafetyHoop"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, WITHSAFETYHOOP_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "HoopPitch"
     pIJDInput.Description = "HoopPitch"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 2
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HOOPPITCH_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "BottomHoopLevel"
     pIJDInput.Description = "BottomHoopLevel"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 2.2
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, BOTTOMHOOPLEVEL_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "HoopClearance"
     pIJDInput.Description = "HoopClearance"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.75
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HOOPCLEARANCE_INDEX
     pIJDInput.Reset
     PC.Reset
         
     pIJDInput.Name = "HoopRadius"
     pIJDInput.Description = "HoopRadius"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.7
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HOOPRADIUS_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "HoopPlateThickness"
     pIJDInput.Description = "HoopPlateThickness"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.016
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HOOPPLATETHICKNESS_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "HoopPlateWidth"
     pIJDInput.Description = "HoopPlateWidth"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.065
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HOOPPLATEWIDTH_INDEX
     pIJDInput.Reset
     PC.Reset
    
     
     pIJDInput.Name = "HoopBendRadius"
     pIJDInput.Description = "HoopBendRadius"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.05
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HOOPBENDRADIUS_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "HoopOpening"
     pIJDInput.Description = "HoopOpening"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HOOPOPENING_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "ShDim1"
     pIJDInput.Description = "ShDim1"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.07
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SHDIM1_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "ShDim2"
     pIJDInput.Description = "ShDim2"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.15
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SHDIM2_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "ShDim3"
     pIJDInput.Description = "ShDim3"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.3
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SHDIM3_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "FlareClearance"
     pIJDInput.Description = "FlareClearance"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, FLARECLEARANCE_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "FlareRadius"
     pIJDInput.Description = "FlareRadius"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, FLARERADIUS_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "HoopFlareBendRadius"
     pIJDInput.Description = "HoopFlareBendRadius"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.05
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HOOPFLAREBENDRADIUS_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "FlareShDim1"
     pIJDInput.Description = "FlareShDim1"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.07
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, FLARESHDIM1_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "FlareShDim2"
     pIJDInput.Description = "FlareShDim2"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.3
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, FLARESHDIM2_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "FlareShDim3"
     pIJDInput.Description = "FlareShDim3"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.3
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, FLARESHDIM3_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "HoopFlareHeight"
     pIJDInput.Description = "HoopFlareHeight"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 2#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HOOPFLAREHEIGHT_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "HoopFlareMaxHeight"
     pIJDInput.Description = "HoopFlareMaxHeight"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 4#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HOOPFLAREMAXHEIGHT_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "VerticalStrapWidth"
     pIJDInput.Description = "VerticalStrapWidth"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.065
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, VERTICALSTRAPWIDTH_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "VerticalStrapThickness"
     pIJDInput.Description = "VerticalStrapThickness"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.016
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, VERTICALSTRAPTHICKNESS_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "VerticalStrapCount"
     pIJDInput.Description = "VerticalStrapCount"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 4#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, VERTICALSTRAPCOUNT_INDEX
     pIJDInput.Reset
     PC.Reset
          
     pIJDInput.Name = "TopExtension"
     pIJDInput.Description = "TopExtension"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, NegativeNum
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
       InputsIf.SetInput pIJDInput, TOPEXTENSION_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.Name = "BottomExtension"
     pIJDInput.Description = "BottomExtension"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, NegativeNum
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, BOTTOMEXTENSION_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "Justification"
     pIJDInput.Description = "Justification"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, JUSTIFICATION_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "TopSupportSide"
     pIJDInput.Description = "TopSupportSide"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, TOPSUPPORTSIDE_INDEX
     pIJDInput.Reset
     PC.Reset
          
     pIJDInput.Name = "IsAssembly"
     pIJDInput.Description = "IsAssembly"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, ISASSEMBLY_INDEX
     pIJDInput.Reset
     PC.Reset

     pIJDInput.Name = "EnvelopeHeight"
     pIJDInput.Description = "EnvelopeHeight"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igValue
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, ENVELOPEHEIGHT_INDEX
     pIJDInput.Reset
     PC.Reset
          
     pIJDInput.Name = "Primary_SPSMaterial"
     pIJDInput.Description = "Primary_SPSMaterial"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Concrete"
     pIJDInput.DefaultParameterValue = PC
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, CheckMaterial
     InputsIf.SetInput pIJDInput, PRIMARYMATERIAL_INDEX
     pIJDInput.Reset
     PC.Reset
        
     pIJDInput.Name = "Primary_SPSGrade"
     pIJDInput.Description = "Primary_SPSGrade"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.Type = igString
     PC.String = "Fc 4000"
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, PRIMARYGRADE_INDEX
     pIJDInput.Reset
     PC.Reset
       

 'NEW REP
    Dim pIReps As IMSSymbolEntities.IJDRepresentations
    Set pIReps = pSymbolDefinition
    Dim pIRep As IMSSymbolEntities.IJDRepresentation
    Set pIRep = New IMSSymbolEntities.DRepresentation
    
    pIRep.Name = "Physical"
    pIRep.Description = "Physical representation"
    'pIRep.Properties = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    pIRep.RepresentationId = SimplePhysical
    mCookie = pUM.GetMethodCookie("Physical", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
          
    Dim pOutputs As IMSSymbolEntities.IJDOutputs
    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    
    Dim output As IMSSymbolEntities.IJDOutput
    Set output = New IMSSymbolEntities.DOutput
    
    pIReps.SetRepresentation pIRep 'Add representation to definition
            
    pIRep.Name = "DetailPhysical"
    pIRep.Description = "DetailPhysical representation"
    'pIRep.Properties = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    pIRep.RepresentationId = DetailPhysical
    mCookie = pUM.GetMethodCookie("Physical", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
    Set pOutputs = pIRep
    pIReps.SetRepresentation pIRep 'Add representation to definition
    

'For Operational Representation.
    pOutputs.RemoveAllOutput
    Set output = New DOutput

    pIRep.Name = "OperationalSub"
    pIRep.Description = "Operational representation"
    pIRep.Properties = igREPRESENTATION_ISVBFUNCTION
    pIRep.RepresentationId = Operation
    mCookie = pUM.GetMethodCookie("OperationalSub", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie

    Set pOutputs = pIRep
    pOutputs.Property = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    output.Name = "OperationalEnvelope1"
    output.Description = "Operational Envelope of the Vertical Ladder"
    pOutputs.SetOutput output
    output.Reset

    pIReps.SetRepresentation pIRep 'Add representation to definition
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED

    'as this symbol def has declared a graphic object as input
    ' GeomOption option will be set to igSYMBOL_GEOM_FIX_TO_ID by the symbol machinerary
    'Because of this the  outputs will be transformed during MDR and the Symbol geometry will
    ' end up in an incorrect location. So resetting the flag - DI226263
    pSymbolDefinition.GeomOption = igSYMBOL_GEOM_FREE


'End For Operational Representation.


    Set pIRep = Nothing
    Set pIReps = Nothing
    Set output = Nothing
    Set pOutputs = Nothing
    
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
    Dim Height As Double
    Dim Width As Double, Length As Double, HoopClearance As Double, FlareClearance As Double
    Dim SideFrameThk As Double, SupportThk As Double, HoopRadius As Double, WallOffset As Double
    Dim WithSafetyHoop As Boolean
    Dim EnvHeight As Double, FlareRadius As Double, BottomHoopLevel As Double, HoopPlateThickness As Double
    Dim TotalWidth As Double, TotalHeight As Double, HoopPitch As Double, HoopPlateWidth As Double
    Dim Justify#
    Set m_GeomFactory = New IngrGeom3D.GeometryFactory
    Set m_pIJLineString = New LineString3d
    
    Length = pInputs.GetInputByIndex(LENGTH_INDEX).IJDInputDuringGame.Result.UomValue
    Height = pInputs.GetInputByIndex(HEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
    Width = pInputs.GetInputByIndex(WIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    SideFrameThk = pInputs.GetInputByIndex(SIDEFRAMETHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
    SupportThk = pInputs.GetInputByIndex(SUPPORTLEGTHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
    WithSafetyHoop = pInputs.GetInputByIndex(WITHSAFETYHOOP_INDEX).IJDInputDuringGame.Result.UomValue
    EnvHeight = pInputs.GetInputByIndex(ENVELOPEHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
    HoopRadius = pInputs.GetInputByIndex(HOOPRADIUS_INDEX).IJDInputDuringGame.Result.UomValue
    FlareRadius = pInputs.GetInputByIndex(FLARERADIUS_INDEX).IJDInputDuringGame.Result.UomValue
    BottomHoopLevel = pInputs.GetInputByIndex(BOTTOMHOOPLEVEL_INDEX).IJDInputDuringGame.Result.UomValue
    HoopPitch = pInputs.GetInputByIndex(HOOPPITCH_INDEX).IJDInputDuringGame.Result.UomValue
    HoopClearance = pInputs.GetInputByIndex(HOOPCLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
    FlareClearance = pInputs.GetInputByIndex(FLARECLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
    HoopPlateWidth = pInputs.GetInputByIndex(HOOPPLATEWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    HoopPlateThickness = pInputs.GetInputByIndex(HOOPPLATETHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
    Justify = pInputs.GetInputByIndex(JUSTIFICATION_INDEX).IJDInputDuringGame.Result.UomValue
    WallOffset = pInputs.GetInputByIndex(WALLOFFSET_INDEX).IJDInputDuringGame.Result.UomValue
    
    Dim pos As Double
    If Justify = 2 Then
        pos = ((Width / 2) + SideFrameThk)
    ElseIf Justify = 3 Then
        pos = -((Width / 2) + SideFrameThk)
    End If
    Dim TmpPos As IJDPosition
    Dim projvect As IJDVector
    Dim upvect As IJDVector
    Set TmpPos = New DPosition
    Set projvect = New DVector
    Set upvect = New DVector
    
    Dim TmpHt As Double
    
    projvect.x = 0#
    projvect.y = 0#
       
    upvect.Set 0, 1, 0
     
    If WithSafetyHoop Then
        TmpHt = BottomHoopLevel + HoopPitch + (HoopPlateWidth * 2)
        TmpPos.Set -pos, ((FlareClearance + HoopPlateThickness) / 2) + (WallOffset / 2), -(Length)
        projvect.z = TmpHt
    
        TotalWidth = FlareRadius * 2
        TotalHeight = FlareClearance + (HoopPlateThickness / 2) + WallOffset
    Else
        TmpPos.Set -pos, (0.762 / 2) + (WallOffset / 2), -(Length)
        projvect.z = Length + EnvHeight
        TotalWidth = Width + SideFrameThk * 2 + SupportThk * 2
        TotalHeight = 0.762 + WallOffset
    End If

    CreateRectangularProjection pOC, TmpPos, projvect, upvect, TotalWidth, TotalHeight, "OperationalEnvelope", 1
     
    If WithSafetyHoop Then
        TmpPos.Set -pos, ((HoopClearance + HoopPlateThickness) / 2) + (WallOffset / 2), -Length + TmpHt
        projvect.z = Length + EnvHeight - TmpHt
        TotalWidth = HoopRadius * 2
        TotalHeight = HoopClearance + (HoopPlateThickness / 2) + WallOffset
        CreateRectangularProjection pOC, TmpPos, projvect, upvect, TotalWidth, TotalHeight, "OperationalEnvelope", 2
    End If
        
    Set pRepDG = Nothing
    Set pOC = Nothing
    Set pInputs = Nothing
    Set TmpPos = Nothing
    Set projvect = Nothing
    Set upvect = Nothing
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Public Sub Physical(pIRepSCM As IJDRepresentationStdCustomMethod)
On Error GoTo ErrorHandler
Const METHOD = "Physical"
    Dim pRepDG As IJDRepresentationDuringGame
    Set pRepDG = pIRepSCM
    Dim pOC As IJDOutputCollection
    Set pOC = pRepDG.outputCollection
        
    Dim pInputs As IJDInputs
    Set pInputs = pRepDG.definition.IJDInputs
    
    Set m_GeomFactory = New IngrGeom3D.GeometryFactory
    Set m_pIJDT4x4 = New DT4x4
    Set m_pIJLineString = New LineString3d
    
    'Get the inputs
    Dim Length#, Span#, Height#, Pitch#, Width#, SideFrameWidth#, StepDia#, SideFrameThk#, WallOffset#
    Dim SupportWidth#, SupportThk#, SupportLegLength#, VlDim#(1 To 5), SupportPitch#, StepProtrusion#
    Dim WithSupp#, bWithSupports As Boolean
    Dim NumSteps As Integer
    Dim i As Integer
    Dim Angle As Double
    
    Dim WithSafetyHoop As Boolean
    Dim HoopPitch#, BottomHoopLevel#, HoopWidth#, HoopClearance#, CagePitch#
    Dim HoopPlateThickness#, HoopPlateWidth#, HoopBendRadius#, CageThickness#, ShDim1#, ShDim2#, ShDim3#, HoopHtAboveDeck#
    Dim HoopRadius As Double
    Dim hoopOpening As Integer
    Dim topRungHeight As Double
    
    Length = pInputs.GetInputByIndex(LENGTH_INDEX).IJDInputDuringGame.Result.UomValue
    Span = pInputs.GetInputByIndex(SPAN_INDEX).IJDInputDuringGame.Result.UomValue
    Height = pInputs.GetInputByIndex(HEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
    Pitch = pInputs.GetInputByIndex(STEPPITCH_INDEX).IJDInputDuringGame.Result.UomValue
    Width = pInputs.GetInputByIndex(WIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    SideFrameWidth = pInputs.GetInputByIndex(SIDEFRAMEWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    SideFrameThk = pInputs.GetInputByIndex(SIDEFRAMETHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
    StepDia = pInputs.GetInputByIndex(STEPDIAMETER_INDEX).IJDInputDuringGame.Result.UomValue
    bWithSupports = pInputs.GetInputByIndex(WITHWALLSUPPORTS_INDEX).IJDInputDuringGame.Result.UomValue
    'for hoop
    WithSafetyHoop = pInputs.GetInputByIndex(WITHSAFETYHOOP_INDEX).IJDInputDuringGame.Result.UomValue
    HoopPitch = pInputs.GetInputByIndex(HOOPPITCH_INDEX).IJDInputDuringGame.Result.UomValue
    BottomHoopLevel = pInputs.GetInputByIndex(BOTTOMHOOPLEVEL_INDEX).IJDInputDuringGame.Result.UomValue
    HoopRadius = pInputs.GetInputByIndex(HOOPRADIUS_INDEX).IJDInputDuringGame.Result.UomValue
    HoopClearance = pInputs.GetInputByIndex(HOOPCLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
    HoopPlateThickness = pInputs.GetInputByIndex(HOOPPLATETHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
    HoopPlateWidth = pInputs.GetInputByIndex(HOOPPLATEWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    HoopBendRadius = pInputs.GetInputByIndex(HOOPBENDRADIUS_INDEX).IJDInputDuringGame.Result.UomValue
    hoopOpening = pInputs.GetInputByIndex(HOOPOPENING_INDEX).IJDInputDuringGame.Result.UomValue

    ShDim1 = pInputs.GetInputByIndex(SHDIM1_INDEX).IJDInputDuringGame.Result.UomValue
    ShDim2 = pInputs.GetInputByIndex(SHDIM2_INDEX).IJDInputDuringGame.Result.UomValue
    ShDim3 = pInputs.GetInputByIndex(SHDIM3_INDEX).IJDInputDuringGame.Result.UomValue
  
    SupportWidth = pInputs.GetInputByIndex(SUPPORTLEGWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    SupportThk = pInputs.GetInputByIndex(SUPPORTLEGTHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
    SupportPitch = pInputs.GetInputByIndex(SUPPORTLEGPITCH_INDEX).IJDInputDuringGame.Result.UomValue
    NumSteps = pInputs.GetInputByIndex(NUMSTEPS_INDEX).IJDInputDuringGame.Result.UomValue
    StepProtrusion = pInputs.GetInputByIndex(STEPPROTRUSION_INDEX).IJDInputDuringGame.Result.UomValue
    'For some reason cannot do VLDIM1_INDEX + i????
    For i = 1 To 4 'Starts at 17, and VlDim is 1 relative
      VlDim(i) = pInputs.GetInputByIndex(VLDIM1_INDEX + i - 1).IJDInputDuringGame.Result.UomValue
    Next
    WallOffset = pInputs.GetInputByIndex(WALLOFFSET_INDEX).IJDInputDuringGame.Result.UomValue
    Angle = pInputs.GetInputByIndex(ANGLE_INDEX).IJDInputDuringGame.Result.UomValue
   
    'For HoopFlare
    Dim FlareRadius#, FlareClearance#, HoopFlareBendRadius#, FlareShDim1#, FlareShDim2#, FlareShDim3#
    Dim HoopFlareHeight#, HoopFlareMaxHeight#, VerticalStrapWidth#, VerticalStrapThickness#, VerticalStrapCount#
    Dim TopExt#, BottomExt#, Justify#
    
    FlareClearance = pInputs.GetInputByIndex(FLARECLEARANCE_INDEX).IJDInputDuringGame.Result.UomValue
    FlareRadius = pInputs.GetInputByIndex(FLARERADIUS_INDEX).IJDInputDuringGame.Result.UomValue
    HoopFlareBendRadius = pInputs.GetInputByIndex(HOOPFLAREBENDRADIUS_INDEX).IJDInputDuringGame.Result.UomValue
    FlareShDim1 = pInputs.GetInputByIndex(FLARESHDIM1_INDEX).IJDInputDuringGame.Result.UomValue
    FlareShDim2 = pInputs.GetInputByIndex(FLARESHDIM2_INDEX).IJDInputDuringGame.Result.UomValue
    FlareShDim3 = pInputs.GetInputByIndex(FLARESHDIM3_INDEX).IJDInputDuringGame.Result.UomValue
    HoopFlareHeight = pInputs.GetInputByIndex(HOOPFLAREHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
    HoopFlareMaxHeight = pInputs.GetInputByIndex(HOOPFLAREMAXHEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
    VerticalStrapWidth = pInputs.GetInputByIndex(VERTICALSTRAPWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
    VerticalStrapThickness = pInputs.GetInputByIndex(VERTICALSTRAPTHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
    VerticalStrapCount = pInputs.GetInputByIndex(VERTICALSTRAPCOUNT_INDEX).IJDInputDuringGame.Result.UomValue
    TopExt = pInputs.GetInputByIndex(TOPEXTENSION_INDEX).IJDInputDuringGame.Result.UomValue
    Justify = pInputs.GetInputByIndex(JUSTIFICATION_INDEX).IJDInputDuringGame.Result.UomValue
    
    BottomExt = pInputs.GetInputByIndex(BOTTOMEXTENSION_INDEX).IJDInputDuringGame.Result.UomValue
   
    '------------------- TR-CP-144430 --------------------
    Dim vertOffset As Double
    vertOffset = (Height - TopExt - Span)
    
    Dim topSupportLocation As Double
    topSupportLocation = -(vertOffset + TopExt)
    '-----------------------------------------------------
    
    Dim oRep As IJDRepresentation
    Dim oOutputs As IJDOutputs
    Dim oOutput As IJDOutput
    Set oOutput = New DOutput
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
    Set oOutputs = oRep
    oOutputs.RemoveAllOutput
     
     
    '***** Set ladder postiion based on justification*****
    'Remove Support leg legth from catalog.
    'SupportLegLength will be Function of Walloffset distance and SideframeWidth
    Dim pos As Double
    If Justify = 2 Then
        pos = ((Width / 2) + SideFrameThk)
    ElseIf Justify = 3 Then
        pos = -((Width / 2) + SideFrameThk)
    End If
    
    '***** Place the left side frame *****
   Dim ULC(3) As Double   'Upper Left Center
   Dim URC(3) As Double   '      Right

   ULC(0) = 0 - (Width / 2#) - SideFrameThk / 2# - pos
   ULC(1) = WallOffset + SideFrameWidth / 2#     'SupportLegLength+ VlDim(1) - SideFrameWidth / 2#
   ULC(2) = 0
  '   DrawCylinder pOC, SideFrameWidth, Height, LLC(0), LLC(1), LLC(2), _
  '                NormalX, NormalY, NormalZ, "LeftSideFrame", 1
  Dim TmpPos As New DPosition
  Dim projvect As New DVector, upvect As New DVector
  TmpPos.x = ULC(0)
  TmpPos.y = ULC(1)
  TmpPos.z = ULC(2)
  
  projvect.x = 0
  projvect.y = (Length) * Tan(PI / 2 - Angle)
  projvect.z = -Length
 
  upvect.x = 0
  upvect.y = 1#
  upvect.z = 1 * Tan(PI / 2 - Angle)
  
  oOutput.Reset
  oOutput.Name = "LeftSideFrame1"
  oOutput.Description = "LeftSideFrame1"
  oOutput.Key = "LeftSideFrame1"
  oOutputs.SetOutput oOutput
  CreateRectangularProjection pOC, TmpPos, projvect, upvect, SideFrameThk, SideFrameWidth, "LeftSideFrame", 1
  
  
  '***** Place the right side frame *****
  URC(0) = 0 + Width / 2# + SideFrameThk / 2# - pos
  URC(1) = ULC(1) 'HorizSuppOffset + VlDim(0) - SideFrameWidth / 2#
  URC(2) = ULC(2) 'VertOffset - Length
  
  TmpPos.x = URC(0)
  TmpPos.y = URC(1)
  TmpPos.z = URC(2)

  oOutput.Reset
  oOutput.Name = "RightSideFrame1"
  oOutput.Description = "RightSideFrame1"
  oOutput.Key = "RightSideFrame1"
  oOutputs.SetOutput oOutput
  CreateRectangularProjection pOC, TmpPos, projvect, upvect, SideFrameThk, SideFrameWidth, "RightSideFrame", 1
  
  '***** Place the rungs *****
  Dim stepNumber As Integer, numberOfSteps As Integer, nSupports As Integer
  numberOfSteps = NumSteps
  'numberOfSteps = Int(((Length + VertOffset - VlDim(5)) / Pitch) + 1)
  
  'Account for vertical offset as well [TR-CP-144430].  This will allow the wall supports to be place correctly.
  nSupports = (Length - (TopExt + vertOffset)) / SupportPitch 'Need to consider reduced space on side frame!
  If nSupports <= 0 Then
    nSupports = 1
  End If
  
  'Set start pos for steps
  TmpPos.x = ULC(0) - SideFrameThk / 2# - StepProtrusion 'Account for outer welding extent 5mm from drawing, no ref data for this value yet
  'Align the top of the rung with the top reference surface.
  TmpPos.y = ULC(1)  '+ 0.127 * Cos(Angle)
  TmpPos.z = ULC(2) - TopExt - StepDia / 2 '- 0.127 * Sin(Angle) 'Drop from top of side frame to top step
  
  topRungHeight = TmpPos.z
  
  'Vect length determines proj length and it runs across the width of the ladder
  projvect.x = Width + 2 * (SideFrameThk + StepProtrusion) 'Twice sidethk and outer weld extent 5mm
  projvect.y = 0#
  projvect.z = 0#
  
  'Set UpVect to rotate the step bar on edge to 45 deg from Z
  'Should be normalized but CreateRectangularProjection() will do it...
  upvect.x = 0#
  upvect.y = 1#
  upvect.z = 1#
    
  For stepNumber = 1 To numberOfSteps
    oOutput.Reset
    oOutput.Name = "Step" & stepNumber
    oOutput.Description = "Ladder Step" & stepNumber
    oOutput.Key = "Step" & stepNumber
    oOutputs.SetOutput oOutput
    CreateCircularProjection pOC, TmpPos, projvect, upvect, StepDia, "Step", stepNumber

    TmpPos.y = TmpPos.y + Pitch * Cos(Angle)
    TmpPos.z = TmpPos.z - Pitch * Sin(Angle)
  Next stepNumber
  
  '***** Place the Wall support *****
  If (bWithSupports) And WallOffset <> 0 Then
    SupportLegLength = WallOffset + (SideFrameWidth / 2#)
    
    upvect.x = 0# 'up is Z again
    upvect.y = 0#
    upvect.z = 1#
    
    'Left Side Top to bottom project from wall y=0 out
    TmpPos.x = ULC(0) - SideFrameThk / 2# - SupportThk / 2#
    TmpPos.y = 0
    TmpPos.z = ULC(2) - 0.01016 - SupportWidth / 2# - (TopExt + vertOffset) '[TR-CP-144430]
    
    projvect.x = 0
    projvect.y = SupportLegLength
    projvect.z = 0#
    For i = 1 To nSupports
        oOutput.Reset
        oOutput.Name = "LeftSupportLeg" & i
        oOutput.Description = "Left Support Leg" & i
        oOutputs.SetOutput oOutput
        projvect.y = SupportLegLength - TmpPos.z / Tan(Angle)
        CreateRectangularProjection pOC, TmpPos, projvect, upvect, SupportThk, SupportWidth, "LeftSupportLeg", i
        TmpPos.z = TmpPos.z - SupportPitch  'Center of end of support bar
    Next

    'Right Side
    TmpPos.x = URC(0) + SideFrameThk / 2# + SupportThk / 2#
    TmpPos.y = 0
    TmpPos.z = URC(2) - 0.01016 - SupportWidth / 2# - (TopExt + vertOffset) '[TR-CP-144430]
        For i = 1 To nSupports
            oOutput.Reset
            oOutput.Name = "RightSupportLeg" & i
            oOutput.Description = "Right Support Leg" & i
            oOutputs.SetOutput oOutput
            projvect.y = SupportLegLength - TmpPos.z / Tan(Angle)
            CreateRectangularProjection pOC, TmpPos, projvect, upvect, SupportThk, SupportWidth, "RightSupportLeg", i
            TmpPos.z = TmpPos.z - SupportPitch 'Center of end of support bar
        Next
    ElseIf WallOffset = 0 Then
'        Call SetAttribute(OccAttrs, False, "WallOffset", PartOccInfoCol)
    End If  'If (bWithSupport)
   
'***** Place the hoop *****
  If (WithSafetyHoop) Then
    Dim weldClearance As Double
    weldClearance = 3 * StepDia / 2    ' This is just thumb-rule. Maybe this should form yet another parameter in Catalog.
    TmpPos.x = ULC(0) - SideFrameThk / 2#
    TmpPos.y = ULC(1) + (0.127 + weldClearance) * Cos(Angle)
    TmpPos.z = ULC(2) - (0.127 + weldClearance) * Sin(Angle)
  
    projvect.x = 0
    projvect.y = (HoopPlateWidth) * Cos(Angle)
    projvect.z = -(HoopPlateWidth) * Sin(Angle)
  
    upvect.x = 0
    upvect.y = 1#
    upvect.z = 1 * Tan(PI / 2 - Angle)
    
    Dim ShDim(1 To 4) As Double
    Dim FlareShDim(1 To 4) As Double
    
    Dim HoopExternalWidth#, HoopCurveOuterRadius#, HoopCurveInnerRadius#, HoopWidth_Resultant#
  
    ShDim(1) = ShDim1
    
    'TR#CP52149 - mkaveti - Calculate the HoopRadius and HoopClearance based on Ladderwidth
    HoopRadius = ShDim2 + (Width / 2) + (2 * HoopBendRadius) + SideFrameThk
    HoopClearance = HoopRadius + (ShDim3 - ShDim1)
    
    Dim DimX As Double, DimY As Double      'Refer figure in InitHoopCurve (Need to change figure)
    DimX = HoopRadius
'    DimY = HoopClearance - HoopRadius
    DimY = HoopRadius
    ShDim(2) = HoopRadius - (Width / 2) - 2 * HoopBendRadius - SideFrameThk
'    ShDim(3) = DimY - ShDim(1) ' + HoopBendRadius
    ShDim(3) = ShDim3

    Dim nHoops As Integer
'    Dim ShDim4 As Double
    'now check with MaxHoopPitch  TO DO
    nHoops = (Length - BottomHoopLevel) / HoopPitch
        
    If nHoops <= 0 Then
      nHoops = 1
    End If
    Dim d As Double
    Dim d1 As Double
    Dim d2 As Double
    Dim countHoops As Integer
    Dim numHoopsOpen As Integer
    
    d = TmpPos.z
     
    'Start placing hoops from Top to Bottom except the Bottom hoop.
    If nHoops > 1 Then
        For i = 1 To nHoops
            oOutput.Reset
            oOutput.Name = "SafetyHoop" & i
            oOutput.Description = "SafetyHoop" & i
            oOutputs.SetOutput oOutput
            '******Change for HoopOpening TR-CP·79181
            'hoops above the top rung. These may need side opening
            'If TmpPos.z >= topRungHeight Then
            
            'Changed from hoops above the top rung to hoops above the top support.
            If TmpPos.z >= topSupportLocation Then '[TR-CP-144430]
                CreateProjectionOfCurve pOC, TmpPos, projvect, upvect, HoopBendRadius, ShDim(), DimX, DimY, HoopPlateThickness, "SafetyHoop", i, hoopOpening, Width + 2 * SideFrameThk
                numHoopsOpen = numHoopsOpen + 1
            Else
                'hoops below the top rung. Treat them as ordinary hoops, even for side step ladder
                CreateProjectionOfCurve pOC, TmpPos, projvect, upvect, HoopBendRadius, ShDim(), DimX, DimY, HoopPlateThickness, "SafetyHoop", i
            End If
            d1 = TmpPos.z
            TmpPos.z = TmpPos.z - HoopPitch * Sin(Angle)
            TmpPos.y = TmpPos.y + HoopPitch * Cos(Angle)
        Next
        
        'Now place the Bottom Hoop
        FlareShDim(1) = FlareShDim1
        FlareShDim(2) = FlareShDim2
        FlareShDim(3) = FlareShDim3
        
        'TR#CP52149 - mkaveti
        'Calculate the FlareRadius and FlareClearance based on the Ladderwidth and other Initialized inputs
        FlareRadius = FlareShDim2 + (Width / 2) + 2 * HoopBendRadius + SideFrameThk
        FlareClearance = FlareRadius + (FlareShDim3 - FlareShDim1)
        
        DimX = FlareRadius
'        DimY = FlareClearance - FlareRadius
        DimY = FlareRadius
        FlareShDim(2) = FlareRadius - (Width / 2) - 2 * HoopBendRadius - SideFrameThk
'        FlareShDim(3) = DimY - ShDim(1) '- FlareShDim(1) + HoopBendRadius
        FlareShDim(3) = ShDim(3)
        
        oOutput.Reset
        oOutput.Name = "FlareSafetyHoop" & nHoops + 1
        oOutput.Description = "FlareSafetyHoop" & nHoops + 1
        oOutputs.SetOutput oOutput
        CreateProjectionOfCurve pOC, TmpPos, projvect, upvect, HoopBendRadius, FlareShDim(), DimX, HoopClearance, HoopPlateThickness, "FlareSafetyHoop", nHoops + 1
    End If
  
  End If ' If (WithSafetyHoop) Then
 
  'Vertical Strap Code
  If (WithSafetyHoop) And nHoops > 1 Then
    'Vertical Strap from 1st to (n-1) hoop
    Dim Rad As Double
    Dim Ang As Double
    Dim RotAng As Double
    Dim x As Double, y As Double, z As Double
    Dim StrapPos() As DPosition
    ReDim StrapPos(VerticalStrapCount) As DPosition
    Dim FlarePos() As DPosition
    ReDim FlarePos(VerticalStrapCount) As DPosition

    Dim TmpVector As DVector
    Set TmpVector = New DVector

    'Vertical straps
    For i = 1 To VerticalStrapCount
        Set StrapPos(i) = New DPosition
        Set FlarePos(i) = New DPosition
    Next i

    Dim oSegments As IJElements
    m_Hoopcomplex.GetCurves oSegments
    Set m_HoopArc = oSegments(1)
    Set oSegments = Nothing

    Ang = m_HoopArc.SweepAngle
    Rad = m_HoopArc.Radius

    RotAng = Ang / (VerticalStrapCount - 1)
    m_HoopArc.GetCenterPoint x, y, z
    Dim theta As Double
    Dim zPosSideOpeningStrap As Double
    Dim lengthSideOpeningStrap As Double
    
    For i = 1 To VerticalStrapCount
        If i = 1 Then
            'first strap
            theta = (VerticalStrapWidth / 2) / Rad
            TmpVector.x = Cos(theta + (RotAng * (i - 1))) * (Rad - VerticalStrapThickness - VerticalStrapThickness / 2)
            TmpVector.y = Sin(theta + (RotAng * (i - 1))) * (Rad - VerticalStrapThickness - VerticalStrapThickness / 2)
            StrapPos(i).x = TmpVector.x + x
            StrapPos(i).y = TmpVector.y + y
        ElseIf i = VerticalStrapCount Then
            'last strap
            theta = (VerticalStrapWidth / 2) / Rad
            TmpVector.x = Cos(theta + (RotAng * (i - 1))) * (Rad - VerticalStrapThickness - VerticalStrapThickness / 2)
            TmpVector.y = Sin(theta + (RotAng * (i - 1))) * (Rad - VerticalStrapThickness - VerticalStrapThickness / 2)
            StrapPos(i).x = TmpVector.x + x
            StrapPos(i).y = TmpVector.y + y
        Else
            'any strap in between
            TmpVector.x = Cos((RotAng * (i - 1))) * (Rad - VerticalStrapThickness - VerticalStrapThickness / 2)
            TmpVector.y = Sin((RotAng * (i - 1))) * (Rad - VerticalStrapThickness - VerticalStrapThickness / 2)
            StrapPos(i).x = TmpVector.x + x
            StrapPos(i).y = TmpVector.y + y
        End If
    Next i

    zPosSideOpeningStrap = d - numHoopsOpen * HoopPitch
    TmpPos.z = d 'this point on first hoop

    projvect.x = 0
    projvect.y = Length * Tan(PI / 2 - Angle)
    projvect.z = -((nHoops - 1) * HoopPitch) - (HoopPlateWidth)
    lengthSideOpeningStrap = projvect.z - numHoopsOpen * (-HoopPitch)
   
    upvect.x = 0
    upvect.y = 1#
    upvect.z = 1 * Tan(PI / 2 - Angle)

    '******Change for HoopOpening TR-CP·79181
    For i = 1 To VerticalStrapCount
        TmpPos.x = StrapPos(i).x
        TmpPos.y = StrapPos(i).y
        
        If hoopOpening <> 1 Then
            If hoopOpening = 2 And 1 = i Then
                TmpPos.z = zPosSideOpeningStrap
                projvect.z = lengthSideOpeningStrap
            ElseIf hoopOpening = 3 And i = VerticalStrapCount Then
                TmpPos.z = zPosSideOpeningStrap
                projvect.z = lengthSideOpeningStrap
            Else
                TmpPos.z = d
                projvect.z = -((nHoops - 1) * HoopPitch) - (HoopPlateWidth)
            End If
        End If
        
        upvect.x = 1 * Cos((RotAng * (i - 1)))
        upvect.y = 1 * Sin((RotAng * (i - 1)))
        
        oOutput.Reset
        oOutput.Name = "VerticalStap" & i
        oOutput.Description = "VerticalStap" & i
        oOutputs.SetOutput oOutput
        CreateRectangularProjection pOC, TmpPos, projvect, upvect, VerticalStrapWidth, VerticalStrapThickness, "VerticalStap", i
    Next i

    'reset for flare straps
    TmpPos.z = d
    projvect.z = -((nHoops - 1) * HoopPitch) - (HoopPlateWidth)

    'Now place Vertical Straps upto the Flare Hoop

    Dim FlRad As Double
    Dim FlAng As Double
    Dim FlRotAng As Double
    m_Flarecomplex.GetCurves oSegments
    Set m_FlareArc = oSegments(1)
    FlAng = m_FlareArc.SweepAngle
    FlRad = m_FlareArc.Radius
    FlRotAng = FlAng / (VerticalStrapCount - 1)
    Dim tmpprojvect As New DVector
    Dim TmpVec As DVector
    Set TmpVec = New DVector
    Dim TmpVec1 As DVector
    Set TmpVec1 = New DVector
    TmpPos.z = d + projvect.z ' -((nHoops - 1) * HoopPitch) - (3 * HoopPlateWidth) ' d1 + VlDim(4)  'this point on last hoop

    For i = 1 To VerticalStrapCount
        TmpPos.x = StrapPos(i).x
        TmpPos.y = StrapPos(i).y

        upvect.x = 1 * Cos((RotAng * (i - 1)))
        upvect.y = 1 * Sin((RotAng * (i - 1)))
        upvect.z = 1 * Tan(PI / 2 - Angle)

        Set TmpVec1 = projvect.Cross(upvect)

        TmpVector.x = Cos((FlRotAng * (i - 1))) * (FlRad - VerticalStrapThickness - VerticalStrapThickness / 2)
        TmpVector.y = Sin((FlRotAng * (i - 1))) * (FlRad - VerticalStrapThickness - VerticalStrapThickness / 2)
        FlarePos(i).x = TmpVector.x + x
        If i = 1 Or VerticalStrapCount Then
            FlarePos(i).y = TmpVector.y + y + (VerticalStrapThickness + VerticalStrapThickness / 2)
        Else
            FlarePos(i).y = TmpVector.y + y
        End If
        FlarePos(i).z = TmpPos.z - HoopPitch

        projvect.Set FlarePos(i).x - StrapPos(i).x, FlarePos(i).y - StrapPos(i).y, FlarePos(i).z - TmpPos.z

        Set tmpprojvect = projvect.Clone

        Set TmpVec = TmpVec1.Cross(tmpprojvect)

        oOutput.Reset
        oOutput.Name = "VerticalStap" & (VerticalStrapCount + i)
        oOutput.Description = "VerticalStap" & (VerticalStrapCount + i)
        oOutputs.SetOutput oOutput
        CreateRectangularProjection pOC, TmpPos, projvect, TmpVec, VerticalStrapWidth, VerticalStrapThickness, "VerticalStap", VerticalStrapCount + i

    Next i

    End If
    Set pRepDG = Nothing
    Set pOC = Nothing
    Set pInputs = Nothing
    Set TmpPos = Nothing
    Set projvect = Nothing
    Set upvect = Nothing
    
    Erase StrapPos
    Erase FlarePos

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

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

Private Function CreateUnitRectProjection(pOC As IJDOutputCollection, RectWidth As Double, RectHeight As Double, projlength As Double) As IJProjection
Const METHOD = "CreateUnitRectProjection"
On Error GoTo ErrorHandler
    
    Dim pts(15) As Double
    InitRectCurvePoints pts, RectWidth, RectHeight 'Build points in local XY plane at the centroid of the rectangle
    
    m_pIJLineString.SetPoints 5, pts    'Init Points
    
    Set CreateUnitRectProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                                    m_pIJLineString, 0#, 0#, 1#, projlength, True)
    
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

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
    Dim tmpVect As IJDVector, tmpprojvect As IJDVector
    
    Set pProj3d = CreateUnitRectProjection(pOC, RectWidth, RectHeight, projvect.Length) 'Create a unit length projection
    
    Set tmpprojvect = projvect.Clone 'So we dont alter input vector
    tmpprojvect.Length = 1#
    upvect.Length = 1#
    Set tmpVect = tmpprojvect.Cross(upvect)
    tmpVect.Length = 1#
    
    m_pIJDT4x4.IndexValue(0) = tmpVect.x
    m_pIJDT4x4.IndexValue(1) = tmpVect.y
    m_pIJDT4x4.IndexValue(2) = tmpVect.z
        
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
    Set tmpVect = Nothing

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Function CreateUnitCircProjection(pOC As IJDOutputCollection, diameter As Double, projlength As Double) As IJProjection
Const METHOD = "CreateUnitCircProjection"
On Error GoTo ErrorHandler
    
    Dim pts(15) As Double
    Dim pCirlce As Circle3d
    Set pCirlce = New Circle3d
    pCirlce.DefineByCenterNormalRadius 0, 0, 0, 0, 0, 1, diameter / 2

    Set CreateUnitCircProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                                      pCirlce, 0#, 0#, 1#, projlength, True)

    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Private Sub CreateCircularProjection(pOC As IJDOutputCollection, centerPos As IJDPosition, _
                                        projvect As IJDVector, _
                                        upvect As IJDVector, _
                                        diameter As Double, _
                                        stepOrBar As String, _
                                        stepNumber As Integer)
Const METHOD = "CreateCircularProjection"
On Error GoTo ErrorHandler

    Dim pProj3d As Projection3d
    Dim tmpVect As IJDVector, tmpprojvect As IJDVector
    
    Set pProj3d = CreateUnitCircProjection(pOC, diameter, projvect.Length) 'Create a unit length projection
    
    Set tmpprojvect = projvect.Clone 'So we dont alter input vector
    tmpprojvect.Length = 1#
    upvect.Length = 1#
    Set tmpVect = tmpprojvect.Cross(upvect)
    tmpVect.Length = 1#
    
    m_pIJDT4x4.IndexValue(0) = tmpVect.x
    m_pIJDT4x4.IndexValue(1) = tmpVect.y
    m_pIJDT4x4.IndexValue(2) = tmpVect.z
        
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
    Set tmpVect = Nothing

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub CreateProjectionOfCurve(pOC As IJDOutputCollection, _
                                    centerPos As IJDPosition, _
                                    projvect As IJDVector, _
                                    upvect As IJDVector, _
                                    ByVal HoopBendRadius As Double, _
                                    ShDim() As Double, _
                                    ByVal DimX As Double, _
                                    ByVal DimY As Double, _
                                    ByVal HoopPlateThickness As Double, _
                                    ByRef strHoop As String, _
                                    ByVal hoopNumber As Integer, _
                                    Optional hoopOpening As Integer, _
                                    Optional ladderWidth As Double)
On Error GoTo ErrorHandler
METHOD = "CreateProjectionOfCurve"
    
    Dim pProj3d As Projection3d
    Dim tmpVect As IJDVector, tmpprojvect As IJDVector
    
    Set pProj3d = CreateUnitProjection(pOC, HoopBendRadius, ShDim(), DimX, DimY, HoopPlateThickness, projvect.Length, strHoop, hoopNumber, hoopOpening, ladderWidth) 'Create a unit length projection
    If pProj3d Is Nothing Then GoTo ErrorHandler

    Set tmpprojvect = projvect.Clone 'So we dont alter input vector
    tmpprojvect.Length = 1#
    upvect.Length = 1#
    Set tmpVect = tmpprojvect.Cross(upvect)
    tmpVect.Length = 1#

    m_pIJDT4x4.IndexValue(0) = tmpVect.x
    m_pIJDT4x4.IndexValue(1) = tmpVect.y
    m_pIJDT4x4.IndexValue(2) = tmpVect.z
        
    m_pIJDT4x4.IndexValue(4) = -upvect.x
    m_pIJDT4x4.IndexValue(5) = -upvect.y
    m_pIJDT4x4.IndexValue(6) = -upvect.z
    
    m_pIJDT4x4.IndexValue(8) = tmpprojvect.x
    m_pIJDT4x4.IndexValue(9) = tmpprojvect.y
    m_pIJDT4x4.IndexValue(10) = tmpprojvect.z
    
    m_pIJDT4x4.IndexValue(12) = centerPos.x
    m_pIJDT4x4.IndexValue(13) = centerPos.y
    m_pIJDT4x4.IndexValue(14) = centerPos.z

    m_complex.Transform m_pIJDT4x4
    If hoopNumber = 1 Then
        m_Hoopcomplex.Transform m_pIJDT4x4
    ElseIf strHoop = "FlareSafetyHoop" Then
        m_Flarecomplex.Transform m_pIJDT4x4
    End If
    pProj3d.Transform m_pIJDT4x4

    pOC.AddOutput strHoop & Trim$(Str$(hoopNumber)), pProj3d  '(stepOrBar & Trim$(Str$(stepNumber))), pProj3d
    Set tmpprojvect = Nothing
    Set pProj3d = Nothing
    Set tmpVect = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise E_FAIL
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Sub

Private Function CreateUnitProjection(pOC As IJDOutputCollection, _
                                      HoopBendRadius As Double, _
                                      ShDim() As Double, _
                                      DimX As Double, _
                                      DimY As Double, _
                                      HoopPlateThickness As Double, _
                                      projlength As Double, strHoop As String, hoopNumber As Integer, _
                                      Optional hoopOpening As Integer = 1, _
                                      Optional ladderWidth As Double) As IJProjection
    METHOD = "CreateUnitProjection"
    On Error GoTo ErrorHandler
 
    If "FlareSafetyHoop" = strHoop Or hoopOpening = 1 Or hoopOpening = 0 Then
        InitHoopCurve ShDim(1), HoopBendRadius, ShDim(2), ShDim(3), DimX, DimY, HoopPlateThickness, strHoop, hoopNumber 'Build points in local XY plane starting at the Left (outer) corner.
    
    ElseIf hoopOpening = 2 Then
        'left opening
        InitHoopCurve_LeftOpening ShDim(1), HoopBendRadius, ShDim(2), ShDim(3), DimX, DimY, HoopPlateThickness, strHoop, hoopNumber, ladderWidth 'Build points in local XY plane starting at the Left (outer) corner.
    
    ElseIf hoopOpening = 3 Then
        'right opening
        InitHoopCurve_RightOpening ShDim(1), HoopBendRadius, ShDim(2), ShDim(3), DimX, DimY, HoopPlateThickness, strHoop, hoopNumber, ladderWidth 'Build points in local XY plane starting at the Left (outer) corner.
    End If
    

    Set CreateUnitProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                                    m_complex, 0#, 0#, 1#, projlength, True)
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Private Sub InitHoopCurve(ShDim1 As Double, _
                          HoopBendRad As Double, _
                          ShDim2 As Double, _
                          ShDim3 As Double, _
                          DimX As Double, _
                          DimY As Double, _
                          HoopPlateThickness As Double, strHoop As String, hoopNumber As Integer)
    
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double
    Dim z As Double
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim iElements As IJElements
                                          '                        /\+Y
    x1 = 0#                               '  2 |                   |
    y1 = 0                                '    |                   |
    x2 = 0                                '  1 |.(0,0)             .------> +X
    y2 = ShDim1
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, x1, y1, z, x2, y2, z)
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    iElements.Add oLine
    Set m_complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
 

    x1 = x2                     'Start Point           3__
    y1 = y2                     'Start Point             . \
    x2 = x1 - HoopBendRad       'Center Point            2 | 1
    y2 = y2                     'Center Point
    x3 = x2                     'End Point
    y3 = y2 + HoopBendRad       'End Point
    Set oArc = m_GeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, x2, y2, z, x1, y1, z, x3, y3, z)
    m_complex.AddCurve oArc, True
    
    x1 = x3                     '               2______1
    y1 = y3
    x2 = x3 - ShDim2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    x1 = x2                     'Start Point        1________
    y1 = y2                     'Start Point        / .2    . \             ^ +Y
    x2 = x1                     'Center Point     3|           |            |
    y2 = y2 - HoopBendRad       'Center Point      |           |            |
    x3 = x1 - HoopBendRad       'End Point         |           |            .--------> +X
    y3 = y2                     'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
    
    x1 = x3                     '                  1|
    y1 = y3                     '                   |
    x2 = x1                     '                   |
    y2 = y1 - ShDim3            '                  2|
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    x1 = x2                   'Start Point                   |--DimX---|
    y1 = y2                   'Start Point      1 |                    | 3       -
    x2 = x1 + DimX                  'Center       -                  -           | DimY
    y2 = y1 - DimX                  'Center         -              -             |
    x3 = x1 + 2 * DimX               'End Point         - __2__ -                -
    y3 = y1                               'End Point
    oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
'    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
    
    If strHoop = "SafetyHoop" And hoopNumber = 1 Then
        Set iElements = New JObjectCollection ' IMSElements.DynElements
        iElements.Add oArc
        Set m_Hoopcomplex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    End If
       
    x1 = x3                                              '                  2|
    y1 = y3                                              '                   |
    x2 = x1                                              '                   |
    y2 = y1 + ShDim3                                     '                  1|
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    x1 = x2                     'Start Point                         ______3__
    y1 = y2                     'Start Point                        /      2. \ 1
    x2 = x1 - HoopBendRad       'Center Point                      |           |
    y2 = y1                     'Center Point                                  |
    x3 = x1 - HoopBendRad       'End Point                                     |
    y3 = y1 + HoopBendRad       'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True

    x1 = x3                     '               2______1
    y1 = y3
    x2 = x1 - ShDim2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
   
    x1 = x2                     'Start Point        __1
    y1 = y2                     'Start Point       / .2
    x2 = x1                     'Center Point     3|
    y2 = y1 - HoopBendRad       'Center Point
    x3 = x1 - HoopBendRad       'End Point
    y3 = y1 - HoopBendRad       'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
  
    x1 = x3#                              '  1 |
    y1 = y3                               '    |
    x2 = x1                               '  2 |
    y2 = y3 - ShDim1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True


' Now return along the same lines with an offset of HoopPlateThickness
    
    
    Dim HoopBendInnerRad, DimX_Inner, DimY_Inner As Double
    HoopBendInnerRad = HoopBendRad - HoopPlateThickness
    DimX_Inner = DimX - HoopPlateThickness
    DimY_Inner = DimY - HoopPlateThickness
    
    x1 = x2                              '  1--2
    y1 = y2
    x2 = x1 + HoopPlateThickness
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    x1 = x2                               '  2 |
    y1 = y2                               '    |
    x2 = x1                               '  1 |
    y2 = y1 + ShDim1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    x1 = x2                                 'Start Point        __3
    y1 = y2                                 'Start Point       / .2
    x2 = x1 + HoopBendInnerRad              'Center Point     1|
    y2 = y1                                 'Center Point
    x3 = x2                                 'End Point
    y3 = y2 + HoopBendInnerRad              'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True

    x1 = x3                     '               1______2
    y1 = y3
    x2 = x1 + ShDim2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
 
    x1 = x2                     'Start Point         ______1__
    y1 = y2                     'Start Point        /      2. \ 3
    x2 = x1                     'Center Point      |           |
    y2 = y1 - HoopBendInnerRad  'Center Point      |
    x3 = x2 + HoopBendInnerRad  'End Point         |
    y3 = y2                     'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
    
    x1 = x3                                              '                  1|
    y1 = y3                                              '                   |
    x2 = x1                                              '                   |
    y2 = y1 - ShDim3                                     '                  2|
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
                              '                            |-DimX_inner-|
    x1 = x2                         'Start Point
    y1 = y2                   'Start Point      3 |                     | 1      -
    x2 = x1 - DimX_Inner             'Center       -                  -           |DimY_Inner
    y2 = y1 - DimX_Inner             'Center         -              -             |
    x3 = x1 - 2 * DimX_Inner         'End Point         - __2__ -                -
    y3 = y1                                 'End Point
    oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
'    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
    If strHoop = "FlareSafetyHoop" Then
        Set iElements = New JObjectCollection 'IMSElements.DynElements
        iElements.Add oArc
        Set m_Flarecomplex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    End If
    
    x1 = x3                     '                  2|
    y1 = y3                     '                   |
    x2 = x1                     '                   |
    y2 = y1 + ShDim3            '                  1|
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    x1 = x2                     'Start Point        3________
    y1 = y2                     'Start Point        / .2    . \             ^ +Y
    x2 = x1 + HoopBendInnerRad  'Center Point     1|           |            |
'   y2 = y2                     'Center Point      |           |            |
    x3 = x2                     'End Point         |           |            .--------> +X
    y3 = y2 + HoopBendInnerRad  'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
    
    x1 = x3                     '               1______2
    y1 = y3
    x2 = x3 + ShDim2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    x1 = x2                     'Start Point           1__
    y1 = y2                     'Start Point             . \
'   x2 = x1                     'Center Point            2 | 3
    y2 = y2 - HoopBendInnerRad  'Center Point
    x3 = x2 + HoopBendInnerRad  'End Point
    y3 = y2                     'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
                                           '                        /\+Y
    x1 = x3                                '  1 |                   |
    y1 = y3                                '    |                   |
    x2 = x1                                '  2 |                   .------> +X
    y2 = y1 - ShDim1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    x1 = x2                                '   1-------2 (0,0)
    y1 = y2
    x2 = 0
    y2 = 0
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    Set oArc = Nothing
    Set oLine = Nothing
    
End Sub

'Private Sub CreateProjectionbyPoints(pOC As IJDOutputCollection, centerPos As IJDPosition, _
'                                        projvect As IJDVector, _
'                                        upvect As IJDVector, _
'                                        diameter As Double, _
'                                        stepOrBar As String, _
'                                        stepNumber As Integer)
'    Dim Proj As IJElements
'    set proj=  m_GeomFactory.Projections3d.CreateByCurve(
'End Sub

'Private Sub InitHoopCurve(ShDim1 As Double, _
'                          HoopBendRad As Double, _
'                          ShDim2 As Double, _
'                          ShDim3 As Double, _
'                          Dimx As Double, _
'                          DimY As Double, _
'                          HoopPlateThickness As Double)
'
'    Dim x1 As Double, y1 As Double
'    Dim x2 As Double, y2 As Double
'    Dim x3 As Double, y3 As Double
'    Dim z As Double
'    Dim oLine As IngrGeom3D.Line3d
'    Dim oArc As IngrGeom3D.Arc3d
'    Dim iElements As IJElements
'                                          '                        /\+Y
'    x1 = 0#                               '  2 |                   |
'    y1 = 0                                '    |                   |
'    x2 = 0                                '  1 |.(0,0)             .------> +X
'    y2 = ShDim1
'    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, x1, y1, z, x2, y2, z)
'    Set iElements = New IMSElements.DynElements
'    iElements.Add oLine
'    Set m_complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
'    Set iElements = Nothing
'
'
'    x1 = x2                     'Start Point           3__
'    y1 = y2                     'Start Point             . \
'    x2 = x1 - HoopBendRad       'Center Point            2 | 1
'   'y2 = y2                     'Center Point
'    x3 = x2                     'End Point
'    y3 = y2 + HoopBendRad       'End Point
'    Set oArc = m_GeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, x2, y2, z, x1, y1, z, x3, y3, z)
'    m_complex.AddCurve oArc, True
'
'    x1 = x3                     '               2______1
'    y1 = y3
'    x2 = x3 - ShDim2
'    y2 = y1
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'    x1 = x2                     'Start Point        1________
'    y1 = y2                     'Start Point        / .2    . \             ^ +Y
'    x2 = x1                     'Center Point     3|           |            |
'    y2 = y2 - HoopBendRad       'Center Point      |           |            |
'    x3 = x1 - HoopBendRad       'End Point         |           |            .--------> +X
'    y3 = y2                     'End Point
'    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
'    m_complex.AddCurve oArc, True
'
'    x1 = x3                     '                  1|
'    y1 = y3                     '                   |
'    x2 = x1                     '                   |
'    y2 = y1 - ShDim3            '                  2|
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'    x1 = x2                         'Start Point                   |--DimX---|
'    y1 = y2                   'Start Point      1 |                    | 3       -
'    x2 = x1 + Dimx                   'Center       -                  -           | DimY
'    y2 = y1 - DimY                      'Center         -              -             |
'    x3 = x1 + 2 * Dimx                      'End Point         - __2__ -               -
'    y3 = y1                                 'End Point
'    oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
''    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
'
'    m_complex.AddCurve oArc, True
'
'    x1 = x3                                              '                  2|
'    y1 = y3                                              '                   |
'    x2 = x1                                              '                   |
'    y2 = y1 + ShDim3                                     '                  1|
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'    x1 = x2                     'Start Point                         ______3__
'    y1 = y2                     'Start Point                        /      2. \ 1
'    x2 = x1 - HoopBendRad       'Center Point                      |           |
'    y2 = y1                     'Center Point                                  |
'    x3 = x1 - HoopBendRad       'End Point                                     |
'    y3 = y1 + HoopBendRad       'End Point
'    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
'    m_complex.AddCurve oArc, True
'
'    x1 = x3                     '               2______1
'    y1 = y3
'    x2 = x1 - ShDim2
'    y2 = y1
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'
'    x1 = x2                     'Start Point        __1
'    y1 = y2                     'Start Point       / .2
'    x2 = x1                     'Center Point     3|
'    y2 = y1 - HoopBendRad       'Center Point
'    x3 = x1 - HoopBendRad       'End Point
'    y3 = y1 - HoopBendRad       'End Point
'    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
'    m_complex.AddCurve oArc, True
'
'    x1 = x3#                              '  1 |
'    y1 = y3                               '    |
'    x2 = x1                               '  2 |
'    y2 = y3 - ShDim1
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'
'' Now return along the same lines with an offset of HoopPlateThickness
'
'
'    Dim HoopBendInnerRad, DimX_Inner, DimY_Inner As Double
'    HoopBendInnerRad = HoopBendRad - HoopPlateThickness
'    DimX_Inner = Dimx - HoopPlateThickness
'    DimY_Inner = DimY - HoopPlateThickness
'
'    x1 = x2                              '  1--2
'    y1 = y2
'    x2 = x1 + HoopPlateThickness
'    y2 = y1
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'    x1 = x2                               '  2 |
'    y1 = y2                               '    |
'    x2 = x1                               '  1 |
'    y2 = y1 + ShDim1
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'    x1 = x2                                 'Start Point        __3
'    y1 = y2                                 'Start Point       / .2
'    x2 = x1 + HoopBendInnerRad              'Center Point     1|
'    y2 = y1                                 'Center Point
'    x3 = x2                                 'End Point
'    y3 = y2 + HoopBendInnerRad              'End Point
'    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
'    m_complex.AddCurve oArc, True
'
'    x1 = x3                     '               1______2
'    y1 = y3
'    x2 = x1 + ShDim2
'    y2 = y1
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'    x1 = x2                     'Start Point         ______1__
'    y1 = y2                     'Start Point        /      2. \ 3
'    x2 = x1                     'Center Point      |           |
'    y2 = y1 - HoopBendInnerRad  'Center Point      |
'    x3 = x2 + HoopBendInnerRad  'End Point         |
'    y3 = y2                     'End Point
'    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
'    m_complex.AddCurve oArc, True
'
'    x1 = x3                                              '                  1|
'    y1 = y3                                              '                   |
'    x2 = x1                                              '                   |
'    y2 = y1 - ShDim3                                     '                  2|
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'                                    '                            |-DimX_inner-|
'    x1 = x2                         'Start Point
'    y1 = y2                    'Start Point      3 |                     | 1      -
'    x2 = x1 - DimX_Inner             'Center       -                  -           |DimY_Inner
'    y2 = y1 - DimY_Inner                'Center         -              -             |
'    x3 = x1 - 2 * DimX_Inner               'End Point         - __2__ -                -
'    y3 = y1                                 'End Point
'    oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
'    'oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
'    m_complex.AddCurve oArc, True
'
'    x1 = x3                     '                  2|
'    y1 = y3                     '                   |
'    x2 = x1                     '                   |
'    y2 = y1 + ShDim3            '                  1|
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'    x1 = x2                     'Start Point        3________
'    y1 = y2                     'Start Point        / .2    . \             ^ +Y
'    x2 = x1 + HoopBendInnerRad  'Center Point     1|           |            |
''   y2 = y2                     'Center Point      |           |            |
'    x3 = x2                     'End Point         |           |            .--------> +X
'    y3 = y2 + HoopBendInnerRad  'End Point
'    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
'    m_complex.AddCurve oArc, True
'
'    x1 = x3                     '               1______2
'    y1 = y3
'    x2 = x3 + ShDim2
'    y2 = y1
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'    x1 = x2                     'Start Point           1__
'    y1 = y2                     'Start Point             . \
''   x2 = x1                     'Center Point            2 | 3
'    y2 = y2 - HoopBendInnerRad  'Center Point
'    x3 = x2 + HoopBendInnerRad  'End Point
'    y3 = y2                     'End Point
'    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
'    m_complex.AddCurve oArc, True
'                                           '                        /\+Y
'    x1 = x3                                '  1 |                   |
'    y1 = y3                                '    |                   |
'    x2 = x1                                '  2 |                   .------> +X
'    y2 = y1 - ShDim1
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'    x1 = x2                                '   1-------2 (0,0)
'    y1 = y2
'    x2 = 0
'    y2 = 0
'    oLine.DefineBy2Points x1, y1, z, x2, y2, z
'    m_complex.AddCurve oLine, True
'
'    Set oArc = Nothing
'    Set oLine = Nothing
'
'End Sub

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
    Set m_oLocalizer = New IMSLocalizer.Localizer
    m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
    Set m_SymbolFactory = Nothing
    Set m_GeomFactory = Nothing
    Set m_pIJDT4x4 = Nothing
    Set m_pIJLineString = Nothing
    Set m_oErrors = Nothing
    Set m_oLocalizer = Nothing
End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_LADDERMACROS_ERROR, "ERROR")
        
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
IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_LADDERMACROS_ERROR, "ERROR")
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String
    
    Set pAttrColl = CollAllDisplayedValues
    For i = 1 To pAttrColl.Count
        Set pAttrDescr = pAttrColl.Item(i)
        
        If pAttrDescr.attrName = "Angle" Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        ElseIf pAttrDescr.attrName = "HoopPlateThickness" Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        ElseIf pAttrDescr.attrName = "HoopPlateWidth" Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        ElseIf pAttrDescr.attrName = "HoopFlareBendRadius" Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        ElseIf pAttrDescr.attrName = "HoopFlareHeight" Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        'TR#CP52149 - mkaveti - Make the HoopRadius, HoopClearance, FlareRadius and FlareClearance readonly
        ElseIf pAttrDescr.attrName = "HoopRadius" Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        ElseIf pAttrDescr.attrName = "HoopClearance" Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        ElseIf pAttrDescr.attrName = "FlareRadius" Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        ElseIf pAttrDescr.attrName = "FlareClearance" Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
        End If
            
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
    UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_ERROR, "ERROR")

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
                Set CurrentInput = dInputs.GetInputByName(oAttribute.AttributeInfo.Name)
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
                If (oInterfaceInfo.Name = sInterfaceName) Then
                    IsInterfaceFound = True
                    Exit For
                End If
            End If
        Next
        
        Set oInterfaceInfo = Nothing
        
        If IsInterfaceFound = False Then
            UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_SCHEMAERROR, "SchemaERROR")
            GoTo ErrorHandler
        End If
        Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
        
        ' loop on the attributes on the interface to match the supplied attribute type
        For AttrCount = 1 To oAttrCol.Count
            Set oAttrObj = oAttrCol.Item(AttrCount)
            
            If oAttrObj.Name = sAttributeName Then
            
                If (sAttributeName = "EnvelopeHeight") Or (sAttributeName = "WallOffset") Or (sAttributeName = "BottomHoopLevel") Then
                    If (varAttributeValue <= 0#) Then
                        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_ATTRIBVAL, "Invalid Attribute Value ")
                        Set oAttributeMetaData = Nothing
                        Exit Function
                    End If
                End If
                If (sAttributeName = "VerticalStrapWidth") Or (sAttributeName = "VerticalStrapThickness") Or (sAttributeName = "VerticalStrapCount") Then
                    If (varAttributeValue <= 0#) Then
                        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_ATTRIBVAL, "Invalid Attribute Value ")
                        Set oAttributeMetaData = Nothing
                        Exit Function
                    End If
                End If
                If (sAttributeName = "FlareRadius") Or (sAttributeName = "FlareClearance") Then
                    If (varAttributeValue <= 0#) Then
                        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_ATTRIBVAL, "Invalid Attribute Value ")
                        Set oAttributeMetaData = Nothing
                        Exit Function
                    End If
                End If
                If (sAttributeName = "HoopRadius") Or (sAttributeName = "HoopClearance") Then
                    If (varAttributeValue <= 0#) Then
                        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_ATTRIBVAL, "Invalid Attribute Value ")
                        Set oAttributeMetaData = Nothing
                        Exit Function
                    End If
                End If
                If (sAttributeName = "HoopPitch") Then
                    If (varAttributeValue <= 0#) Then
                        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_ATTRIBVAL, "Invalid Attribute Value ")
                        Set oAttributeMetaData = Nothing
                        Exit Function
                    End If
                End If
                If (sAttributeName = "VlDim1") Or (sAttributeName = "VlDim2") Or (sAttributeName = "VlDim3") Then
                'Attribute not used  currently so do not check its validity
'                    If (varAttributeValue <= 0#) Then
'                        UserAttributeMgmt_Validate = "Invalid Attribute Value "
'                        Set oAttributeMetaData = Nothing
'                        Exit Function
'                    End If
                End If
                If (sAttributeName = "StepProtrusion") Or (sAttributeName = "StepDiameter") Then
                    If (varAttributeValue <= 0#) Then
                        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_ATTRIBVAL, "Invalid Attribute Value ")
                        Set oAttributeMetaData = Nothing
                        Exit Function
                    End If
                End If
                If (sAttributeName = "SideFrameWidth") Or (sAttributeName = "SideFrameThickness") Then
                    If (varAttributeValue <= 0#) Then
                        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_ATTRIBVAL, "Invalid Attribute Value ")
                        Set oAttributeMetaData = Nothing
                        Exit Function
                    End If
                End If
                If (sAttributeName = "SupportLegWidth") Or (sAttributeName = "SupportLegThickness") Or (sAttributeName = "SupportLegPitch") Then
                    If (varAttributeValue <= 0#) Then
                        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_ATTRIBVAL, "Invalid Attribute Value ")
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
                UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_WIDTH, "Invalid Width value")
                Exit Function
            End If
        End If
        If sInterfaceName = "IJSPSCommonStairLadderProps" And sAttributeName = "Span" Then
            If (varAttributeValue > 100) Then
                UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_OFFSETVAL, "Invalid offset value")
                Exit Function
            End If
        End If
        If sInterfaceName = "IJSPSCommonStairLadderProps" And sAttributeName = "StepPitch" Then
            If (varAttributeValue <= 0) Or (varAttributeValue > 100) Then
                UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_LADDERMACROS_INVALID_PITCHVAL, "Invalid pitch value")
                Exit Function
            End If
        End If
    End If
    
    UserAttributeMgmt_Validate = ""
    
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

Private Sub InitHoopCurve_RightOpening(ShDim1 As Double, _
                          HoopBendRad As Double, _
                          ShDim2 As Double, _
                          ShDim3 As Double, _
                          DimX As Double, _
                          DimY As Double, _
                          HoopPlateThickness As Double, strHoop As String, hoopNumber As Integer, ladderWidth As Double)
    
    Const METHOD = "InitHoopCurve_RightOpening"
    On Error GoTo ErrorHandler
    
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double
    Dim z As Double
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim iElements As IJElements
    Dim hoopBotX As Double
    Dim hoopBotY As Double
                           
    'right side opening. The order of the segments for the hoop is shown below.
    '      D|----|C
    '      |      |
    '    E|       B|y
    '     |        |
    '     |        A----x
    '     |
    '     |                               H|---|I
    '     |                              |
    '    F|                             |G
    '       |                        |
    '          |                  |
    '            |             |
    '               |      |
    '                 ____
                           
                           
    
    
    'segment AB
    x1 = ladderWidth
    y1 = 0
    x2 = x1
    y2 = ShDim1
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, x1, y1, z, x2, y2, z)
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    iElements.Add oLine
    Set m_complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
 
    'segment BC
    x1 = x2                     'Start Point
    y1 = y2                     'Start Point
    x2 = x1 + HoopBendRad       'Center Point
    y2 = y2                     'Center Point
    x3 = x2                     'End Point
    y3 = y2 + HoopBendRad       'End Point
    Set oArc = m_GeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, x2, y2, z, x1, y1, z, x3, y3, z)
    m_complex.AddCurve oArc, True
    
    'segment CD
    x1 = x3                     '
    y1 = y3
    x2 = x3 + ShDim2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    'segment DE
    x1 = x2                     'Start Point
    y1 = y2                     'Start Point
    x2 = x1                     'Center Point
    y2 = y2 - HoopBendRad       'Center Point
    x3 = x1 + HoopBendRad       'End Point
    y3 = y2                     'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
    
    'segment EF
    x1 = x3
    y1 = y3
    x2 = x1
    y2 = y1 - ShDim3
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    'segment FG - the actual hoop arc
    x1 = x2                   'Start Point
    y1 = y2
    x2 = x1 - DimX            'Second Point
    y2 = y1 - DimX
    x3 = x2 - DimX * Sin(PI / 4#)     'End Point
    y3 = y2 + DimX * (1 - Cos(PI / 4#))
    oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
    m_complex.AddCurve oArc, True
    
    'save x2,y2 for later use
    hoopBotX = x2
    hoopBotY = y2
    
    If strHoop = "SafetyHoop" And hoopNumber = 1 Then
        Set iElements = New JObjectCollection ' IMSElements.DynElements
        'redefine arc for the flare hoop
        oArc.DefineBy3Points x1, y1, z, x2, y2, z, x2 - DimX, y1, z
        iElements.Add oArc
        Set m_Hoopcomplex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    End If
       
    'segment GH
    x1 = x3                     'Start Point
    y1 = y3                     'Start Point
    x2 = x1 - HoopBendRad       'Center Point
    y2 = y1                     'Center Point
    x3 = x1 - HoopBendRad       'End Point
    y3 = y1 + HoopBendRad       'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True

    'segment HI
    x1 = x3
    y1 = y3
    x2 = x1 - ShDim2 / 2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    
    ' Now return along the same lines with an offset of HoopPlateThickness
    Dim HoopBendInnerRad, HoopBendOuterRad, DimX_Inner, DimY_Inner As Double
    HoopBendInnerRad = HoopBendRad - HoopPlateThickness
    HoopBendOuterRad = HoopBendRad + HoopPlateThickness
    DimX_Inner = DimX - HoopPlateThickness
    DimY_Inner = DimY - HoopPlateThickness
    
    x1 = x2
    y1 = y2
    x2 = x1
    y2 = y1 + HoopPlateThickness
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    'segment IH
    x1 = x2
    y1 = y2
    x2 = x1 + ShDim2 / 2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
 
    'segment HG
    x1 = x2                     'Start Point
    y1 = y2                     'Start Point
    x2 = x1                     'Center Point
    y2 = y1 - HoopBendOuterRad  'Center Point
    x3 = x2 + HoopBendOuterRad  'End Point
    y3 = y2                     'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
    
    'segment GF
    x1 = x3                         'Start Point
    y1 = y3                   'Start Point
    x2 = hoopBotX             'Center
    y2 = hoopBotY + HoopPlateThickness             'Center
    x3 = x2 + DimX_Inner         'End Point
    y3 = y2 + DimX_Inner                                 'End Point
    oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
    m_complex.AddCurve oArc, True
    If strHoop = "FlareSafetyHoop" Then
        Set iElements = New JObjectCollection 'IMSElements.DynElements
        'redefine arc for flare hoop
        x2 = x1 + DimX_Inner             'Center
        y2 = y1 - DimX_Inner             'Center
        iElements.Add oArc
        Set m_Flarecomplex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    End If
    
    'segment FE
    x1 = x3
    y1 = y3
    x2 = x1
    y2 = y1 + ShDim3
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    'segment ED
    x1 = x2                     'Start Point
    y1 = y2                     'Start Point
    x2 = x1 - HoopBendInnerRad  'Center Point
    y2 = y1                     'Center Point
    x3 = x2                     'End Point
    y3 = y2 + HoopBendInnerRad  'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
    
    'segment DC
    x1 = x3
    y1 = y3
    x2 = x3 - ShDim2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    'segment CB
    x1 = x2                     'Start Point
    y1 = y2                     'Start Point
    x2 = x1                     'Center Point
    y2 = y2 - HoopBendInnerRad  'Center Point
    x3 = x2 - HoopBendInnerRad  'End Point
    y3 = y2                     'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
                                           
    'segment BA
    x1 = x3
    y1 = y3
    x2 = x1
    y2 = y1 - ShDim1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    x1 = x2
    y1 = y2
    x2 = ladderWidth
    y2 = 0
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    Set oArc = Nothing
    Set oLine = Nothing
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Sub

Private Sub InitHoopCurve_LeftOpening(ShDim1 As Double, _
                          HoopBendRad As Double, _
                          ShDim2 As Double, _
                          ShDim3 As Double, _
                          DimX As Double, _
                          DimY As Double, _
                          HoopPlateThickness As Double, strHoop As String, hoopNumber As Integer, ladderWidth As Double)
    
    Const METHOD = "InitHoopCurve_LeftOpening"
    On Error GoTo ErrorHandler
    
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim x3 As Double, y3 As Double
    Dim z As Double
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim iElements As IJElements
    Dim hoopBotX As Double
    Dim hoopBotY As Double
                           
    'left side opening. The order of the segments for the hoop is shown below.
    '                       C|----|D
    '                       |      |
    '              y      B|        |
    '              |       |         |E
    '              --x    A|         |
    '                                |
    '   I--H                         |
    '      -                         |
    '       -|G                      |F
    '         |                     |
    '          |                  |
    '            |             |
    '               |      |
    '                 ____
    
    'segment AB
    x1 = 0#
    y1 = 0
    x2 = 0
    y2 = ShDim1
    Set oLine = m_GeomFactory.Lines3d.CreateBy2Points(Nothing, x1, y1, z, x2, y2, z)
    Set iElements = New JObjectCollection ' IMSElements.DynElements
    iElements.Add oLine
    Set m_complex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    Set iElements = Nothing
 
    'segment BC
    x1 = x2                     'Start Point
    y1 = y2                     'Start Point
    x2 = x1 - HoopBendRad       'Center Point
    y2 = y2                     'Center Point
    x3 = x2                     'End Point
    y3 = y2 + HoopBendRad       'End Point
    Set oArc = m_GeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, x2, y2, z, x1, y1, z, x3, y3, z)
    m_complex.AddCurve oArc, True
    
    'segment CD
    x1 = x3
    y1 = y3
    x2 = x3 - ShDim2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    'segment DE
    x1 = x2                     'Start Point
    y1 = y2                     'Start Point
    x2 = x1                     'Center Point
    y2 = y2 - HoopBendRad       'Center Point
    x3 = x1 - HoopBendRad       'End Point
    y3 = y2                     'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
      
    'segment EF
    x1 = x3                     '
    y1 = y3                     '
    x2 = x1                     '
    y2 = y1 - ShDim3            '
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    'segment FG
    ' To leave enough space for user to slide on the side, we need the cage to stop short of full semi-circle
    ' we will leave short at 45deg angle
    x1 = x2                   'Start Point
    y1 = y2                   'Start Point
    x2 = x1 + DimX                  'Center
    y2 = y1 - DimX                  'Center
    hoopBotX = x2
    hoopBotY = y2
    'x3 = x1 + 2 * DimX                    'End Point
    'y3 = y1                               'End Point
    x3 = x2 + DimX * Sin(PI / 4#)          'End Point
    y3 = y2 + DimX * (1# - Cos(PI / 4#))    'End Point
    oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
    m_complex.AddCurve oArc, True
    
    If strHoop = "SafetyHoop" And hoopNumber = 1 Then
        Set iElements = New JObjectCollection ' IMSElements.DynElements
        'redefine x3, y3 for the flare hoop
        oArc.DefineBy3Points x1, y1, z, x2, y2, z, x1 + 2 * DimX, y1, z
        iElements.Add oArc
        Set m_Hoopcomplex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    End If
       
    'segment GH
    x1 = x3                     'Start Point
    y1 = y3                     'Start Point
    x2 = x1 + HoopBendRad       'Center Point
    y2 = y1                     'Center Point
    x3 = x2       'End Point
    y3 = y1 + HoopBendRad       'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True

    'segment HI
    x1 = x3
    y1 = y3
    x2 = x1 + ShDim2 / 2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    ' Now return along the same lines with an offset of HoopPlateThickness
    Dim HoopBendInnerRad, HoopBendOuterRad, DimX_Inner, DimY_Inner As Double
    HoopBendInnerRad = HoopBendRad - HoopPlateThickness
    HoopBendOuterRad = HoopBendRad + HoopPlateThickness
    DimX_Inner = DimX - HoopPlateThickness
    DimY_Inner = DimY - HoopPlateThickness
    
    'Adjust coordinates for plate thickness
    x1 = x2                              '  1--2
    y1 = y2
    x2 = x1
    y2 = y1 + HoopPlateThickness
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    'segment IH
    x1 = x2
    y1 = y2
    x2 = x1 - ShDim2 / 2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    'segment HG
    x1 = x2                                 'Start Point
    y1 = y2                                 'Start Point
    x2 = x1               'Center Point
    y2 = y1 - HoopBendOuterRad                                'Center Point
    x3 = x2 - HoopBendOuterRad                                'End Point
    y3 = y2               'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True

    'segment GF
    x1 = x3
    y1 = y3
    x2 = hoopBotX
    y2 = hoopBotY + HoopPlateThickness
    x3 = x2 - DimX_Inner
    y3 = y2 + DimX_Inner                                 'End Point
    oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
    m_complex.AddCurve oArc, True
    If strHoop = "FlareSafetyHoop" Then
        Set iElements = New JObjectCollection 'IMSElements.DynElements
        'redefine x1, y1 for the flare hoop
        x1 = x2 + DimX_Inner
        y1 = y2 + DimX_Inner
        oArc.DefineBy3Points x1, y1, z, x2, y2, z, x3, y3, z
        iElements.Add oArc
        Set m_Flarecomplex = m_GeomFactory.ComplexStrings3d.CreateByCurves(Nothing, iElements)
    End If
    
    x1 = x3                     '                  2|
    y1 = y3                     '                   |
    x2 = x1                     '                   |
    y2 = y1 + ShDim3            '                  1|
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    x1 = x2                     'Start Point        3________
    y1 = y2                     'Start Point        / .2    . \             ^ +Y
    x2 = x1 + HoopBendInnerRad  'Center Point     1|           |            |
'   y2 = y2                     'Center Point      |           |            |
    x3 = x2                     'End Point         |           |            .--------> +X
    y3 = y2 + HoopBendInnerRad  'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
    
    x1 = x3                     '               1______2
    y1 = y3
    x2 = x3 + ShDim2
    y2 = y1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    x1 = x2                     'Start Point           1__
    y1 = y2                     'Start Point             . \
'   x2 = x1                     'Center Point            2 | 3
    y2 = y2 - HoopBendInnerRad  'Center Point
    x3 = x2 + HoopBendInnerRad  'End Point
    y3 = y2                     'End Point
    oArc.DefineByCenterStartEnd x2, y2, z, x1, y1, z, x3, y3, z
    m_complex.AddCurve oArc, True
                                           '                        /\+Y
    x1 = x3                                '  1 |                   |
    y1 = y3                                '    |                   |
    x2 = x1                                '  2 |                   .------> +X
    y2 = y1 - ShDim1
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True

    x1 = x2                                '   1-------2 (0,0)
    y1 = y2
    x2 = 0
    y2 = 0
    oLine.DefineBy2Points x1, y1, z, x2, y2, z
    m_complex.AddCurve oLine, True
    
    Set oArc = Nothing
    Set oLine = Nothing
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Sub

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
    Dim oLadderSource As ISPSLadder
    
    'target vars
    Dim oPartOccTarget As IJPartOcc
    Dim oPartTarget As IJDPart
    Dim oAttrsOccTarget As IJDAttributes
    Dim oAttrsPartTarget As IJDAttributes
    Dim lMirrorBehavior As Long
    Dim oInfosColTargetPart As IJDInfosCol
    Dim oInfosColTargetOcc As IJDInfosCol
    Dim oLadderTarget As ISPSLadder
    
    Set oLadderSource = oSource
    Set oPartOccSource = oSource
    'get occ attrs
    Set oAttrsOccSource = oPartOccSource
    
    If Not oPartOccSource Is Nothing Then
        'Now get the part definition attrs
        oPartOccSource.GetPart oPartSource
        Set oAttrsPartSource = oPartSource
    End If
     
    Set oLadderTarget = oTarget
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
    'gg 7/31/2008 TR-135471 Always change the support side
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
    If Not ConnectedToSlab(oLadderTarget) Then
        bTopSupportSide = GetAttribute(oAttrsOccTarget, "TopSupportSide", oInfosColTargetOcc)
        bTopSupportSide = Not bTopSupportSide
        SetAttribute oAttrsOccTarget, bTopSupportSide, "TopSupportSide", oInfosColTargetOcc
    End If
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description

End Sub

Private Function ConnectedToSlab(oLadder As ISPSLadder) As Boolean
    Const METHOD = "ConnectedToSlab"
    On Error GoTo ErrorHandler
    ConnectedToSlab = False
    Dim oPort As IJPort
    Dim oConnectable As IJConnectable
    If TypeOf oLadder.TopEdge Is IJPort Then
        Set oPort = oLadder.TopEdge
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
