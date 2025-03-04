VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InclinedLadder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'*************************************************************************************
'  Copyright (C) 1998, Intergraph Corporation.  All rights reserved.
'
'  File:  InclinedLadder.cls
'
'  Description: Geometry creation macro for the "Inclined" Ladder object type.
'
'  History:
'  Chad Edwards 12/05/99  Created
'  Jeff Machamer 02/04/2000  Added check functions to allow the attributes value to be checked for
'                           Validity before being commited.
'  Jeff Machamer 03/09/2000  Started implimenting LadderSymbolService which in essence moved the semantic
'                           into the symbol.  (For setting Origin, height, and other driven OAs.)
'  Jeff Machamer 03/10/2000  Added Weight and CG
'  Jeff Machamer 03/23/2000 Added Handrails
'  Jeff Machamer 03/24/2000 Added curves to handrails.  Bottom Vertical Bar and it's support "drift" a bit, not sure what is causing it.
'                           Still needs to curve one more connection.
'  Jeff Machamer 03/29/2000 All of the handrails are correct with bends.
'  Jeff Machamer 03/29/2000 Offset the HR away from the edge of the ladder and insert the small pads.  Lengthened the support
'                           bars to protrude to bottom of support pads.
'  Jeff Machamer 04/18/2000 Changed EvaluateGeometry.  It now tells you what has caused the recompute so that
'                           the recompute can be as costless as possible.
'
'  Jan 2002  Aniket Patil   V3 Additions new Functionality
'  20-Sep-05 J. Schwartz    TR#84847 - Shortened the name of the interface IJSPSCommonStairLadderProperties
'                           to IJSPSCommonStairLadderProps (Oracle requirement)
'*************************************************************************************

Const defProgId As String = "SPSStairMacros.InclinedLadder"
Const CheckProgId As String = "SPSValidateArgs.CheckFunctions"
Private Const MODULE = "SPSStairMacros.StairTypeB"
Const PI As Double = 3.14159265358979
Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259

Private m_SymbolFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
Private m_GeomFactory As IngrGeom3D.GeometryFactory
Private m_pIJDT4x4 As IJDT4x4
Private m_pIJLineString As IJLineString
Private m_pIJCircle As IJCircle
Private m_oLocalizer As IJLocalizer

Private Enum InputIndex
PART_INDEX = 1
WIDTH_INDEX
ANGLE_INDEX
STEPPITCH_INDEX
SIDEFRAMETHICKNESS_INDEX '5
SIDEFRAMEWIDTH_INDEX
SFDIM1_INDEX
SFDIM2_INDEX
SFDIM3_INDEX
SFDIM4_INDEX '10
SFDIM5_INDEX
SFDIM6_INDEX
SFDIM7_INDEX
SFDIM8_INDEX
SFDIM9_INDEX '15
SFDIM10_INDEX
LOWERCONNECTPLATETHK_INDEX
LCDIM1_INDEX
LCDIM2_INDEX
LCDIM3_INDEX '20
LCDIM4_INDEX
UPPERCONNECTPLATETHK_INDEX
UCDIM1_INDEX
UCDIM2_INDEX
UCDIM3_INDEX '25
STEPTHICKNESS_INDEX
STEPWIDTH_INDEX
STEPFLARE_INDEX ' not used need to be removed
HANDRAILDIA_INDEX
HRDIM1_INDEX '30
HRDIM2_INDEX
HRDIM3_INDEX
HRDIM4_INDEX
HRDIM5_INDEX
HRDIM6_INDEX '35
HRDIM7_INDEX
HRDIM8_INDEX
HRMAXPITCH_INDEX
DUSTPLATETHK_INDEX
DPDIM1_INDEX '40
DPDIM2_INDEX
DPDIM3_INDEX
DPDIM4_INDEX
HEIGHT_INDEX
WITHDUSTPLATE_INDEX '45
NUMSTEPS_INDEX
SPAN_INDEX
LENGTH_INDEX
JUSTIFICATION_INDEX
EDGEDIRECTION_INDEX
End Enum

Implements IJDUserSymbolServices
Implements ISPSStairSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck

Dim bOnPreLoad As Boolean


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
    Dim FirstStepPitch As Double
    Dim NumSteps As Integer
    Dim tempNum As Double
    Dim oPart As IJDPart
    Dim oPartOcc As IJPartOcc
    Dim Matrix As DT4x4
 

    Set oPartOcc = StairBO
    oPartOcc.GetPart oPart

    Dim OccAttrs As IJDAttributes
    Dim PartAttrs As IJDAttributes
    Set OccAttrs = StairBO
    Set PartAttrs = oPart
    'now Length is what we called height previously
    'for Inclined ladder Height=Span
    OldHeight = GetAttribute(OccAttrs, "Height", PartOccInfoCol)

    Call SetOriginAndSpan(StairBO, Top, bottom, Ref, PartOccInfoCol)
    Span = GetAttribute(OccAttrs, "Span", PartOccInfoCol)

    Height = Span

    If Not OldHeight = Height Or bPartOccChanged Or bOccAttributeChanged Or bTopChanged Or bBottomChanged Or bRefChanged Then

        Angle = GetAttribute(OccAttrs, "Angle", PartOccInfoCol)
        Length = Height / Sin(Angle)

        StepPitch = GetAttribute(OccAttrs, "StepPitch", PartOccInfoCol)
            tempNum = Height / StepPitch
            If tempNum = Int(tempNum) Then
                NumSteps = Int(tempNum)                 'If the pitch divides evenly into the height don't add extra
            Else                                        'step.  Otherwise you will get a step at height 0.
                NumSteps = Int((tempNum) + 1)
                If NumSteps <= 0 Then
                    NumSteps = 1
                End If
            End If
            FirstStepPitch = Height - (StepPitch * (NumSteps - 1))

        'set the part occurrence with any values that have been driven.
        Call SetAttribute(OccAttrs, NumSteps, "NumSteps", PartOccInfoCol)
        Call SetAttribute(OccAttrs, FirstStepPitch, "FirstStepPitch", PartOccInfoCol)
        Call SetAttribute(OccAttrs, Length, "Length", PartOccInfoCol)
        Call SetAttribute(OccAttrs, Height, "Height", PartOccInfoCol) 'for Inclined ladder Height=Span

    End If
    Set oPart = Nothing
    Set oPartOcc = Nothing
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
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

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
Const METHOD = "IJDUserSymbolServices_EditOccurence"
  IJDUserSymbolServices_EditOccurence = False
End Function

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
Const METHOD = "IJDUserSymbolServices_GetDefinitionName"
  IJDUserSymbolServices_GetDefinitionName = defProgId
End Function

'********************************************************************
' Function: IJDUserSymbolServices_InstanciateDefinition()
'
' Description:This instanciates a persistent symbol definition object
' and initializes it for the first time.
'********************************************************************
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParameters As Variant, ByVal ActiveConnection As Object) As Object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler

    'Create a Symbol Definition Object.
    Dim pSymbolFactory As New DSymbolEntitiesFactory
    Dim pSymbolDefinition As IJDSymbolDefinition
    Set pSymbolDefinition = pSymbolFactory.CreateEntity(definition, ActiveConnection)

    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
   
    'nomenclature
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


Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
Const METHOD = "IJDUserSymbolServices_InvokeRepresentation"

End Sub

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pSymbolDefinition As IJDSymbolDefinition)
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
            
     pIJDInput.name = "SideFrameThickness"
     pIJDInput.Description = "SideFrameThickness"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.009
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SIDEFRAMETHICKNESS_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SideFrameWidth"
     pIJDInput.Description = "SideFrameWidth"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.125
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SIDEFRAMEWIDTH_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SfDim1"
     pIJDInput.Description = "SfDim1"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.076
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SFDIM1_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SfDim2"
     pIJDInput.Description = "SfDim2"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.039
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SFDIM2_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SfDim3"
     pIJDInput.Description = "SfDim3"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.109
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SFDIM3_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SfDim4"
     pIJDInput.Description = "SfDim4"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.049
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SFDIM4_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SfDim5"
     pIJDInput.Description = "SfDim5"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.25
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SFDIM5_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SfDim6"
     pIJDInput.Description = "SfDim6"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.15
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SFDIM6_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SfDim7"
     pIJDInput.Description = "SfDim7"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.128
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SFDIM7_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SfDim8"
     pIJDInput.Description = "SfDim8"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.1
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SFDIM8_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SfDim9"
     pIJDInput.Description = "SfDim9"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.085
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SFDIM9_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "SfDim10"
     pIJDInput.Description = "SfDim10"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.035
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, SFDIM10_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "LowerConnectPlateThk"
     pIJDInput.Description = "LowerConnectPlateThk"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.01
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, LOWERCONNECTPLATETHK_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "LcDim1"
     pIJDInput.Description = "LcDim1"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.06
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, LCDIM1_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "LcDim2"
     pIJDInput.Description = "LcDim2"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.09
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, LCDIM2_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "LcDim3"
     pIJDInput.Description = "LcDim3"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.03
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, LCDIM3_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "LcDim4"
     pIJDInput.Description = "LcDim4"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.03
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, LCDIM4_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "UpperConnectPlateThk"
     pIJDInput.Description = "UpperConnectPlateThk"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.01
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, UPPERCONNECTPLATETHK_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "UcDim1"
     pIJDInput.Description = "UcDim1"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.05
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, UCDIM1_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "UcDim2"
     pIJDInput.Description = "UcDim2"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.15
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, UCDIM2_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "UcDim3"
     pIJDInput.Description = "UcDim3"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.025
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, UCDIM3_INDEX
     pIJDInput.Reset
     PC.Reset
       
     pIJDInput.name = "StepThickness"
     pIJDInput.Description = "StepThickness"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.023
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, STEPTHICKNESS_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "StepWidth"
     pIJDInput.Description = "StepWidth"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, GTZeroCheck
     PC.UomValue = 0.18
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, STEPWIDTH_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "StepFlare"
     pIJDInput.Description = "StepFlare"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.03
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, STEPFLARE_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "HandRailDia"
     pIJDInput.Description = "HandRailDia"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.02
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HANDRAILDIA_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "HrDim1"
     pIJDInput.Description = "HrDim1"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.152
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HRDIM1_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "HrDim2"
     pIJDInput.Description = "HrDim2"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.596
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HRDIM2_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "HrDim3"
     pIJDInput.Description = "HrDim3"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.96
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HRDIM3_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "HrDim4"
     pIJDInput.Description = "HrDim4"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.6
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HRDIM4_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "HrDim5"
     pIJDInput.Description = "HrDim5"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.72
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HRDIM5_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "HrDim6"
     pIJDInput.Description = "HrDim6"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.397
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HRDIM6_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "HrDim7"
     pIJDInput.Description = "HrDim7"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.45
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HRDIM7_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "HrDim8"
     pIJDInput.Description = "HrDim8"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.5
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HRDIM8_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "HrMaxPitch"
     pIJDInput.Description = "HrMaxPitch"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1.2
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, HRMAXPITCH_INDEX
     pIJDInput.Reset
     PC.Reset
      
     pIJDInput.name = "DustPlateThk"
     pIJDInput.Description = "DustPlateThk"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.005
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, DUSTPLATETHK_INDEX
     pIJDInput.Reset
     PC.Reset
          
     pIJDInput.name = "DpDim1"
     pIJDInput.Description = "DpDim1"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.06
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, DPDIM1_INDEX
     pIJDInput.Reset
     PC.Reset
       
     pIJDInput.name = "DpDim2"
     pIJDInput.Description = "DpDim2"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.1
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, DPDIM2_INDEX
     pIJDInput.Reset
     PC.Reset
     
     pIJDInput.name = "DpDim3"
     pIJDInput.Description = "DpDim1"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.105
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, DPDIM3_INDEX
     pIJDInput.Reset
     PC.Reset
  
     pIJDInput.name = "DpDim4"
     pIJDInput.Description = "DpDim4"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 0.17
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, DPDIM4_INDEX
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
     
     pIJDInput.name = "WithDustPlate"
     pIJDInput.Description = "Add Dust Plate (True/False)"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
'     pIJDInput.IJDInputStdCustomMethod.SetCMCheck ChecklibCookie, IsBoolean
     PC.UomValue = 0#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, WITHDUSTPLATE_INDEX
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
     
     pIJDInput.name = "EdgeDirection"
     pIJDInput.Description = "EdgeDirection"
     pIJDInput.Properties = igINPUT_IS_A_PARAMETER
     PC.UomValue = 1#
     pIJDInput.DefaultParameterValue = PC
     InputsIf.SetInput pIJDInput, EDGEDIRECTION_INDEX
     pIJDInput.Reset
     PC.Reset
     
    pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
    
    
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
     pSymbolDefinition.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED
     'All steps will be generated as variable output arguments
     
     'Need to declare any fixed handrail pieces
     
    pIReps.SetRepresentation pIRep 'Add representation to definition
           
    pIRep.name = "DetailPhysical"
    pIRep.Description = "DetailPhysical representation"
    'pIRep.Properties = igCOLLECTION_VARIABLE ' declare that the number of outputs is variable
    pIRep.RepresentationId = DetailPhysical
    mCookie = pUM.GetMethodCookie("Physical", libCookie)
    pIRep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mCookie
    Set pOutputs = pIRep
     pIReps.SetRepresentation pIRep 'Add representation to definition
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

Public Sub Physical(pIRepSCM As IJDRepresentationStdCustomMethod)
Const METHOD = "Physical"
On Error GoTo ErrorHandler

  Dim pRepDG As IJDRepresentationDuringGame
  Dim pOC As IJDOutputCollection
  Dim pInputs As IJDInputs
  Dim i As Integer
    
  Set pRepDG = pIRepSCM
  Set pOC = pRepDG.outputCollection
  Set pInputs = pRepDG.definition.IJDInputs
    
  Set m_GeomFactory = New IngrGeom3D.GeometryFactory
  Set m_pIJDT4x4 = New DT4x4
  Set m_pIJLineString = New LineString3d
        
  'Get the inputs
  Dim Span#, Length#, Height#, Pitch#, Width#, Angle#, SideFrameWidth#, SideFrameThk#, StepThk#, StepWidth#
  Dim SFDim#(1 To 10)
  Dim HRDim#(1 To 8), HRDiam#, HRPitch#, Justify#
  Dim NumSteps As Integer
  Length = pInputs.GetInputByIndex(LENGTH_INDEX).IJDInputDuringGame.Result.UomValue
  Span = pInputs.GetInputByIndex(SPAN_INDEX).IJDInputDuringGame.Result.UomValue
  Height = pInputs.GetInputByIndex(HEIGHT_INDEX).IJDInputDuringGame.Result.UomValue
  Pitch = pInputs.GetInputByIndex(STEPPITCH_INDEX).IJDInputDuringGame.Result.UomValue
  Width = pInputs.GetInputByIndex(WIDTH_INDEX).IJDInputDuringGame.Result.UomValue
  Angle = pInputs.GetInputByIndex(ANGLE_INDEX).IJDInputDuringGame.Result.UomValue
'  Length = Length * Sin(Angle)
  StepThk = pInputs.GetInputByIndex(STEPTHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
  StepWidth = pInputs.GetInputByIndex(STEPWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
  SideFrameThk = pInputs.GetInputByIndex(SIDEFRAMETHICKNESS_INDEX).IJDInputDuringGame.Result.UomValue
  SideFrameWidth = pInputs.GetInputByIndex(SIDEFRAMEWIDTH_INDEX).IJDInputDuringGame.Result.UomValue
  HRDiam = pInputs.GetInputByIndex(HANDRAILDIA_INDEX).IJDInputDuringGame.Result.UomValue
  HRPitch = pInputs.GetInputByIndex(HRMAXPITCH_INDEX).IJDInputDuringGame.Result.UomValue
  NumSteps = pInputs.GetInputByIndex(NUMSTEPS_INDEX).IJDInputDuringGame.Result.UomValue
  Justify = pInputs.GetInputByIndex(JUSTIFICATION_INDEX).IJDInputDuringGame.Result.UomValue
   
    Dim Pos As Double
    If Justify = 2 Then
        Pos = (Width / 2) + SideFrameThk
    ElseIf Justify = 3 Then
        Pos = -((Width / 2) + SideFrameThk)
    End If

    For i = 1 To 10         'NOTE SFDims start at
      SFDim(i) = pInputs.GetInputByIndex(i + SFDIM1_INDEX - 1).IJDInputDuringGame.Result.UomValue
    Next
    
    For i = 1 To 8          'NOTE HRDims start at
      HRDim(i) = pInputs.GetInputByIndex(i + HRDIM1_INDEX - 1).IJDInputDuringGame.Result.UomValue
    Next
  
    Dim oRep As IJDRepresentation
    Dim oOutputs As IJDOutputs
    Dim oOutput As IJDOutput
    Set oOutput = New DOutput
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
    Set oOutputs = oRep
    Dim a As Integer
    Dim sname As String
    For a = oOutputs.Count To 1 Step -1
        sname = oOutputs.Item(a).name
        If Left(sname, 4) = "Step" Or _
           Left(sname, 2) = "HR" Then
            Call oOutputs.Remove(sname)
        End If
    Next a
    Set oOutput = Nothing
    Set oRep = Nothing
    Set oOutputs = Nothing

  'Declare misc keypoints
  Dim ULP As IJDPosition, URP As IJDPosition
  Set ULP = New DPosition 'Upper Left Root Position
  Set URP = New DPosition 'Upper Right Root Position

  ULP.x = -Width / 2# - Pos 'Y,Z=0
  URP.x = Width / 2# - Pos
  
  CreateSideFrames pOC, ULP, URP, Height, Width, SideFrameThk, SideFrameWidth, Angle, SFDim()

  CreateSteps pOC, ULP, Height, Width, Pitch, StepThk, StepWidth, Angle, NumSteps
  
  CreateHandrails pOC, Pos, Height, Width, SideFrameThk, SideFrameWidth, Angle, HRDiam, SFDim(), HRDim()
     
    Set m_pIJLineString = Nothing
    Set pRepDG = Nothing
    Set pOC = Nothing
    Set pInputs = Nothing

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub CreateSideFrames(pOC As IJDOutputCollection, _
                             ULP As IJDPosition, URP As IJDPosition, _
                             Height#, Width#, SideFrameThk#, SideFrameWidth#, Angle#, SFDim#())
Const METHOD = "CreateSideFrames"
On Error GoTo ErrorHandler

  Dim pIJProjection As IJProjection
  Dim TmpPos As IJDPosition
   
  Set TmpPos = ULP.Clone()
  TmpPos.x = TmpPos.x - SideFrameThk 'ULP is on inside of side frame
  
  CreateFrameLineString m_pIJLineString, TmpPos, Height, SideFrameWidth, Angle, SFDim
  Set pIJProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                       m_pIJLineString, 1#, 0, 0, SideFrameThk, True)
  pOC.AddOutput "LeftSideFrame", pIJProjection

  'We can use exact URP point since on inside of frame
  CreateFrameLineString m_pIJLineString, URP, Height, SideFrameWidth, Angle, SFDim
  Set pIJProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                       m_pIJLineString, 1#, 0, 0, SideFrameThk, True)
  pOC.AddOutput "RightSideFrame", pIJProjection
  Set TmpPos = Nothing
  Set pIJProjection = Nothing

  Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub CreateFrameLineString(pIJLineString As IJLineString, RootPoint As IJDPosition, Height#, _
                                  SideFrameWidth#, Angle#, SFDim#())
Const METHOD = "CreateFrameLineString"
On Error GoTo ErrorHandler

    Dim pts(21) As Double, L1#, L2#, NinetyDegInRad#, ThirtyFiveDegInRad#
    Dim Pos(6) As IJDPosition
    Dim i As Integer, j As Integer
    NinetyDegInRad = PI / 2#
    ThirtyFiveDegInRad# = 35# * (PI / 180#)
    For i = 0 To 5
      Set Pos(i) = New DPosition
      Pos(i).x = RootPoint.x
    Next
    
    L1 = (SideFrameWidth - (Cos(Angle) * SFDim(10))) / Cos(NinetyDegInRad - Angle) 'SFDim(7)
    L2 = (SideFrameWidth - (Cos(Angle) * SFDim(3))) / Cos(NinetyDegInRad - Angle) 'SFDim(1)

    Pos(1).z = SFDim(10)
    Pos(2).y = L1
    Pos(2).z = Pos(1).z
    
    Pos(5).z = -(Height)   'Removed the Hard coded Gap Aniket
    Pos(5).y = (Height) / Tan(Angle)
    
    Pos(4).y = Pos(5).y + L2
    Pos(4).z = Pos(5).z
    
    Pos(3).y = Pos(4).y
    Pos(3).z = Pos(4).z + SFDim(3)

    j = 0
    For i = 0 To 15 Step 3
      pts(i) = Pos(j).x
      pts(i + 1) = Pos(j).y
      pts(i + 2) = Pos(j).z
      j = j + 1
    Next
    'First and last point the same
    pts(18) = pts(0)
    pts(19) = pts(1)
    pts(20) = pts(2)
    
    pIJLineString.SetPoints 7, pts
    For i = 0 To 5
        Set Pos(i) = Nothing
    Next i
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub CreateSteps(pOC As IJDOutputCollection, ULP As IJDPosition, _
                        Height#, Width#, Pitch#, StepThk#, StepWidth#, Angle#, NumSteps As Integer)
Const METHOD = "CreateSteps"
On Error GoTo ErrorHandler

  Dim pIJProjection As IJProjection
  Dim rootPos As IJDPosition
  Dim DirVect As New DVector
  Dim nSteps As Integer, StepNum As Integer
   
  nSteps = NumSteps
    
  'get first position
  Set rootPos = ULP.Clone
  Dim oRep As IJDRepresentation
  Dim oOutputs As IJDOutputs
  Dim oOutput As IJDOutput
  Set oOutput = New DOutput
  Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName("Physical")
  Set oOutputs = oRep

  For StepNum = 1 To nSteps
    oOutput.name = "Step" & StepNum
    oOutput.Description = "Step" & StepNum
    oOutputs.SetOutput oOutput
    oOutput.Reset
        
    If StepNum <> 1 Then 'skip first offset
      DirVect.x = 0#
      DirVect.y = (StepNum - 1) * (Pitch / Tan(Angle))
      DirVect.z = -(StepNum - 1) * Pitch
      Set rootPos = ULP.Offset(DirVect)
    End If
    Set pIJProjection = CreateStepProjection(pOC, rootPos, StepWidth, StepThk, Width) 'Use ladder width
    pOC.AddOutput "Step" & StepNum, pIJProjection
  Next
  Set pIJProjection = Nothing
  Set rootPos = Nothing
  Set pIJProjection = Nothing
  Set DirVect = Nothing

  Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Function CreateStepProjection(pOC As IJDOutputCollection, rootPos As IJDPosition, _
                                      StepWidth#, StepThk#, StepLength#) As IJProjection
Const METHOD = "CreateStepProjection"
On Error GoTo ErrorHandler

    Dim pts(18) As Double
    InitRectCurvePoints pts, rootPos, StepWidth, StepThk
    
    m_pIJLineString.SetPoints 5, pts
    
    Set CreateStepProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                                    m_pIJLineString, 1#, 0#, 0#, StepLength, True)
    
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function


Private Sub InitRectCurvePoints(pts() As Double, rootPos As IJDPosition, Width As Double, Height As Double)
Const METHOD = "InitRectCurvePoints"
On Error GoTo ErrorHandler:

    'Build points in local at the upper right of the rectangle
    pts(0) = rootPos.x 'upper right
    pts(1) = rootPos.y
    pts(2) = rootPos.z

    pts(3) = rootPos.x 'upper left
    pts(4) = rootPos.y + Width
    pts(5) = rootPos.z

    pts(6) = rootPos.x 'lower left
    pts(7) = rootPos.y + Width
    pts(8) = rootPos.z - Height

    pts(9) = rootPos.x 'lower right
    pts(10) = rootPos.y
    pts(11) = rootPos.z - Height

    pts(12) = pts(0) 'Same as first point for closed shape
    pts(13) = pts(1)
    pts(14) = pts(2)

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub InitTrapCurvePoints(pts() As Double, rootPos As IJDPosition, Width As Double, Height As Double, Angle As Double)
Const METHOD = "InitTrapCurvePoints"
On Error GoTo ErrorHandler:

    'Build points in local at the upper right
    pts(0) = rootPos.x 'upper right
    pts(1) = rootPos.y
    pts(2) = rootPos.z

    pts(3) = rootPos.x 'upper left
    pts(4) = rootPos.y + Width * Cos(Angle)
    pts(5) = rootPos.z - Width * Sin(Angle)

    pts(6) = rootPos.x 'lower left
    pts(7) = pts(4) - Height * Cos(PI / 2 - Angle)
    pts(8) = pts(5) - Height * Sin(PI / 2 - Angle)

    pts(9) = rootPos.x 'lower right
    pts(10) = rootPos.y - Height * Cos(PI / 2 - Angle)
    pts(11) = rootPos.z - Height * Sin(PI / 2 - Angle)

    pts(12) = pts(0) 'Same as first point for closed shape
    pts(13) = pts(1)
    pts(14) = pts(2)
    
    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub CreateHandrails(pOC As IJDOutputCollection, Pos As Double, _
                            Height#, Width#, SideFrameThk#, SideFrameWidth#, Angle#, HRDiam#, SFDim#(), HRDim#())
Const METHOD = "CreateHandrails"
On Error GoTo ErrorHandler

    Dim TotWidth As Double
    TotWidth = Width / 2 + 0.02 + SideFrameThk + HRDiam / 2

    Dim BarProjection As Projection3d
    Dim BottomOfSupport As DPosition
    Dim TopOfSupport As DPosition
    Dim TmpPos As New DPosition  'Point to begin the first bar at.
    Dim tmpPos2 As New DPosition
    
    Set BottomOfSupport = New DPosition
    Set TopOfSupport = New DPosition
    Set TmpPos = New DPosition
    Set tmpPos2 = New DPosition
    
    Dim vNormal As DVector
    Set vNormal = New DVector
    Dim vUp As DVector
    Set vUp = New DVector
    Dim vOut As DVector
    Set vOut = New DVector
    Dim vLadder As DVector
    Set vLadder = New DVector
    vLadder.Set 0, Cos(Angle), -Sin(Angle)
    vOut.Set 1, 0, 0
    vUp.Set 0, 0, 1
    Set vNormal = vOut.Cross(vLadder)
    
    TmpPos.y = Cos(Angle) * SFDim(6)
    TmpPos.z = -Sin(Angle) * SFDim(6)
    
    vNormal.Length = SideFrameWidth
    
    Set tmpPos2 = TmpPos.Offset(vNormal)
    
    vNormal.Length = HRDim(8)
    
    BottomOfSupport.Set 0, tmpPos2.y, tmpPos2.z
    
    Set TopOfSupport = BottomOfSupport.Offset(vNormal)
    
    'Bars supporting Hand Rails
    Dim pts(18) As Double
    Dim rootPos As IJDPosition
    Set rootPos = New DPosition
    Dim CreateSupportProjection As Projection3d
    Set CreateSupportProjection = New Projection3d
    
    Dim a As Long
    For a = 1 To (((Height / Sin(Angle)) - 0.25) / 1.2)
        If a <> 1 Then
            tmpPos2.y = tmpPos2.y + Cos(Angle) * 1.2
            tmpPos2.z = tmpPos2.z - Sin(Angle) * 1.2
        End If
        
            tmpPos2.x = TotWidth - Pos
            CreateCylindricalProjection pOC, HRDiam, HRDim(8) + 0.05, tmpPos2.x, tmpPos2.y - Cos(PI / 2 - Angle) * 0.05, tmpPos2.z - Sin(PI / 2 - Angle) * 0.05, vNormal.x, vNormal.y, vNormal.z, "HRSupportR", a
            tmpPos2.x = -(TotWidth + Pos)
            CreateCylindricalProjection pOC, HRDiam, HRDim(8) + 0.05, tmpPos2.x, tmpPos2.y - Cos(PI / 2 - Angle) * 0.05, tmpPos2.z - Sin(PI / 2 - Angle) * 0.05, vNormal.x, vNormal.y, vNormal.z, "HRSupportL", a
        
            'Blocks support Bars.
                       
            rootPos.Set Width / 2 + SideFrameThk - Pos, tmpPos2.y - Cos(Angle) * 0.05, tmpPos2.z + Sin(Angle) * 0.05
            InitTrapCurvePoints pts, rootPos, 0.1, 0.05, Angle
            
            m_pIJLineString.SetPoints 5, pts
            
            Set CreateSupportProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                                        m_pIJLineString, 1#, 0#, 0#, 0.01, True)
            InitNewOutput pOC, "HRSupportProjectionL"
            pOC.AddOutput "HRSupportProjectionL" & a, CreateSupportProjection
            rootPos.Set -Width / 2 - SideFrameThk - Pos, tmpPos2.y - Cos(Angle) * 0.05, tmpPos2.z + Sin(Angle) * 0.05
            InitTrapCurvePoints pts, rootPos, 0.1, 0.05, Angle
            m_pIJLineString.SetPoints 5, pts
            
            Set CreateSupportProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                                        m_pIJLineString, -1#, 0#, 0#, 0.01, True)
            InitNewOutput pOC, "HRSupportProjectionR"
            pOC.AddOutput "HRSupportProjectionR" & a, CreateSupportProjection
        
    Next a
    Dim LastFullSupport As DPosition
    Set LastFullSupport = New DPosition
    vNormal.Length = HRDim(8)
    Set LastFullSupport = tmpPos2.Offset(vNormal)
    
    'HandRails
    Dim dist As Double
    
    dist = BottomOfSupport.DistPt(tmpPos2)
    
    vNormal.Length = HRDim(8)
    Set tmpPos2 = BottomOfSupport.Offset(vNormal)
    If (((Height / Sin(Angle)) - 0.25) / 1.2) > 1 Then
        tmpPos2.x = (TotWidth - Pos)
        CreateCylindricalProjection pOC, HRDiam, dist, tmpPos2.x, tmpPos2.y, tmpPos2.z, TmpPos.x, TmpPos.y, TmpPos.z, "HRHandrailR", 1
        tmpPos2.x = -(TotWidth + Pos)
        CreateCylindricalProjection pOC, HRDiam, dist, tmpPos2.x, tmpPos2.y, tmpPos2.z, TmpPos.x, TmpPos.y, TmpPos.z, "HRHandrailL", 1
    End If
    
    '****************Upper Ladder*********************
    'Vertical Bar setup
    Dim TopOfVert As DPosition
    Dim BottomOfVert As DPosition
    Set TopOfVert = New DPosition
    Set BottomOfVert = New DPosition
    TopOfVert.Set 0, 0.05, 0.965
    BottomOfVert.Set 0, 0.05, 0.965 - HRDim(5)
    
    Dim vTemp As DVector
    Set vTemp = New DVector
    Dim pTemp As IJDPosition
    Set pTemp = New DPosition
    pTemp.Set 0, TopOfVert.y + HRDim(4), TopOfVert.z
    Dim sAngle
    vTemp.Set 0, pTemp.y - TopOfSupport.y, pTemp.z - TopOfSupport.z
    vNormal.Length = 1
    vTemp.Length = 1
    Dim x As Double
    x = vTemp.dot(vNormal)
    x = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    
    TopOfSupport.x = 0
    'Angle between normal and connector of top of first support and torus
    sAngle = x - Atn(HRDim(4) / TopOfSupport.DistPt(pTemp))
    
    'Upper most Torus
    
    Dim oTorus As Torus3d
    Set oTorus = New Torus3d
    TopOfVert.x = -(TotWidth + Pos)
    Set oTorus = m_GeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(Nothing, TopOfVert.x, TopOfVert.y + HRDim(4), TopOfVert.z, _
                                        1, 0, 0, _
                                        0, -1, 0, _
                                        HRDim(4), HRDiam, sAngle - Angle + PI, False)
    InitNewOutput pOC, "HRUpperTorusR1"
    pOC.AddOutput "HRUpperTorusR1", oTorus
    TopOfVert.x = (TotWidth - Pos)
    Set oTorus = m_GeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(Nothing, TopOfVert.x, TopOfVert.y + HRDim(4), TopOfVert.z, _
                                        1, 0, 0, _
                                        0, -1, 0, _
                                        HRDim(4), HRDiam, sAngle - Angle + PI, False)
    InitNewOutput pOC, "HRUpperTorusL1"
    pOC.AddOutput "HRUpperTorusL1", oTorus
    
    'connect Top Of first Support with Torus
    Dim DirToTorus As DVector 'from Top of First support
    Set DirToTorus = New DVector
    Dim PointOnTorus As DPosition
    Set PointOnTorus = New DPosition
    PointOnTorus.Set 0, TopOfVert.y + HRDim(4) + (HRDim(4) * Cos(sAngle - Angle)), TopOfVert.z + (HRDim(4) * Sin(sAngle - Angle))
    
    DirToTorus.Set 0, PointOnTorus.y - TopOfSupport.y, PointOnTorus.z - TopOfSupport.z
    
    TopOfSupport.x = 0
    PointOnTorus.x = 0
    dist = TopOfSupport.DistPt(PointOnTorus)
    TopOfSupport.x = -(TotWidth + Pos)
    CreateCylindricalProjection pOC, HRDiam, dist, TopOfSupport.x, TopOfSupport.y, TopOfSupport.z, DirToTorus.x, DirToTorus.y, DirToTorus.z, "HR94degR", 1
    TopOfSupport.x = (TotWidth - Pos)
    CreateCylindricalProjection pOC, HRDiam, dist, TopOfSupport.x, TopOfSupport.y, TopOfSupport.z, DirToTorus.x, DirToTorus.y, DirToTorus.z, "HR94degL", 1
    
    'Connect bottom of Vert with Support Bar
    
    Dim PntOnSup As DPosition
    Set PntOnSup = New DPosition
    Dim VecToSup As DVector
    Set VecToSup = New DVector
    Dim CenterOfTorus As DPosition
    Set CenterOfTorus = New DPosition
    Dim Length As Double
    
    vNormal.Length = HRDim(8) - HRDim(7)
    Set PntOnSup = BottomOfSupport.Offset(vNormal)
    VecToSup.Set 0, BottomOfVert.y - PntOnSup.y, BottomOfVert.z - PntOnSup.z
    dist = VecToSup.Length
    vUp.Length = 1
    VecToSup.Length = 1
    sAngle = vUp.dot(VecToSup)
    sAngle = Atn(-sAngle / Sqr(-sAngle * sAngle + 1)) + 2 * Atn(1)
    
    CenterOfTorus.Set 0, BottomOfVert.y + HRDim(4), BottomOfVert.z + Tan(sAngle / 2) * HRDim(4)
    CenterOfTorus.x = -(TotWidth + Pos)
    Set oTorus = m_GeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(Nothing, CenterOfTorus.x, CenterOfTorus.y, CenterOfTorus.z, _
                                        -1, 0, 0, _
                                        0, -1, 0, _
                                        HRDim(4), HRDiam, 2 * PI - sAngle, False)
    InitNewOutput pOC, "HRUpperTorusR2"
    pOC.AddOutput "HRUpperTorusR2", oTorus
    CenterOfTorus.x = (TotWidth - Pos)
    Set oTorus = m_GeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(Nothing, CenterOfTorus.x, CenterOfTorus.y, CenterOfTorus.z, _
                                        -1, 0, 0, _
                                        0, -1, 0, _
                                        HRDim(4), HRDiam, 2 * PI - sAngle, False)
    InitNewOutput pOC, "HRUpperTorusL2"
    pOC.AddOutput "HRUpperTorusL2", oTorus
    
    'Draw top Vert
    BottomOfVert.x = -(TotWidth + Pos)
    CreateCylindricalProjection pOC, HRDiam, HRDim(5) - Tan(sAngle / 2) * HRDim(4), BottomOfVert.x, BottomOfVert.y, BottomOfVert.z + Tan(sAngle / 2) * HRDim(4), 0, 0, 1, "HRVertBarR", 1
    BottomOfVert.x = (TotWidth - Pos)
    CreateCylindricalProjection pOC, HRDiam, HRDim(5) - Tan(sAngle / 2) * HRDim(4), BottomOfVert.x, BottomOfVert.y, BottomOfVert.z + Tan(sAngle / 2) * HRDim(4), 0, 0, 1, "HRVertBarL", 1
    
    'draw connector between vert and first support bar.
    PntOnSup.x = -(TotWidth + Pos)
    dist = dist - Tan(sAngle / 2) * HRDim(4)
    CreateCylindricalProjection pOC, HRDiam, dist, PntOnSup.x, PntOnSup.y, PntOnSup.z, VecToSup.x, VecToSup.y, VecToSup.z, "HR94degR", 2
    PntOnSup.x = (TotWidth - Pos)
    CreateCylindricalProjection pOC, HRDiam, dist, PntOnSup.x, PntOnSup.y, PntOnSup.z, VecToSup.x, VecToSup.y, VecToSup.z, "HR94degL", 2
    
    '************Lower Ladder*******************
    
    'Last Support
    Dim LastSupportPos As DPosition
    Set LastSupportPos = New DPosition
    Dim TopOfLastSup As DPosition
    Set TopOfLastSup = New DPosition
    Dim EndRailPos As DPosition
    Set EndRailPos = New DPosition
    
    Dim L2 As Double  'This is the equation used in FrameLinestring..   Have to use the same.
    L2 = (SideFrameWidth - (Cos(Angle) * SFDim(3))) / Cos(PI / 2 - Angle) 'SFDim(1)
    
    'Calculate the bottom of the little support from the origin position.
    LastSupportPos.y = ((Height - 0.02) / Tan(Angle)) + L2 - (Cos(Angle) * SFDim(5))
    LastSupportPos.z = -Height + SFDim(3) + (Sin(Angle) * SFDim(5)) + 0.02
    LastSupportPos.x = -(TotWidth + Pos)
    vNormal.Length = HRDim(1)
    Set TopOfLastSup = LastSupportPos.Offset(vNormal)
    
    'Between Short Support and Bottom Vert.
    CenterOfTorus.Set 0, TopOfLastSup.y - HRDim(4), TopOfLastSup.z + (HRDim(4) * Tan(Angle / 2))
    CenterOfTorus.x = -(TotWidth + Pos)
    Set oTorus = m_GeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(Nothing, CenterOfTorus.x, CenterOfTorus.y, CenterOfTorus.z, _
                                        1, 0, 0, _
                                        0, 1, 0, _
                                        HRDim(4), HRDiam, -Angle, False)
    InitNewOutput pOC, "HRLowerTorusR1"
    pOC.AddOutput "HRLowerTorusR1", oTorus
    CenterOfTorus.x = (TotWidth - Pos)
    Set oTorus = m_GeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(Nothing, CenterOfTorus.x, CenterOfTorus.y, CenterOfTorus.z, _
                                        1, 0, 0, _
                                        0, 1, 0, _
                                        HRDim(4), HRDiam, -Angle, False)
    InitNewOutput pOC, "HRLowerTorusL1"
    pOC.AddOutput "HRLowerTorusL1", oTorus
    
    'Short Support
    TopOfLastSup.x = -(TotWidth + Pos)
    LastSupportPos.x = -(TotWidth + Pos)
    CreateCylindricalProjection pOC, HRDiam, TopOfLastSup.DistPt(LastSupportPos) - HRDim(4) * Tan(Angle / 2) + 0.05, LastSupportPos.x, LastSupportPos.y - Cos(PI / 2 - Angle) * 0.05, LastSupportPos.z - Sin(PI / 2 - Angle) * 0.05, vNormal.x, vNormal.y, vNormal.z, "HRLastSupportR", 1
    LastSupportPos.x = (TotWidth - Pos)
    TopOfLastSup.x = (TotWidth - Pos)
    CreateCylindricalProjection pOC, HRDiam, TopOfLastSup.DistPt(LastSupportPos) - HRDim(4) * Tan(Angle / 2) + 0.05, LastSupportPos.x, LastSupportPos.y - Cos(PI / 2 - Angle) * 0.05, LastSupportPos.z - Sin(PI / 2 - Angle) * 0.05, vNormal.x, vNormal.y, vNormal.z, "HRLastSupportL", 1
    
    'Blocks supporting last support '.05 is the height of the block  .1 is the width.
    rootPos.Set Width / 2 + SideFrameThk - Pos, LastSupportPos.y - Cos(Angle) * 0.05, LastSupportPos.z + Sin(Angle) * 0.05
    InitTrapCurvePoints pts, rootPos, 0.1, 0.05, Angle
    
    m_pIJLineString.SetPoints 5, pts
    
    Set CreateSupportProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                                m_pIJLineString, 1#, 0#, 0#, 0.01, True)
    InitNewOutput pOC, "HRSupportProjectionL"
    pOC.AddOutput "HRSupportProjectionL" & a, CreateSupportProjection
    rootPos.Set -Width / 2 - SideFrameThk - Pos, LastSupportPos.y - Cos(Angle) * 0.05, LastSupportPos.z + Sin(Angle) * 0.05
    InitTrapCurvePoints pts, rootPos, 0.1, 0.05, Angle
    m_pIJLineString.SetPoints 5, pts
    
    Set CreateSupportProjection = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                                m_pIJLineString, -1#, 0#, 0#, 0.01, True)
    InitNewOutput pOC, "HRSupportProjectionR"
    pOC.AddOutput "HRSupportProjectionR" & a, CreateSupportProjection

    'curve at top of bottom vert
    EndRailPos.Set TopOfLastSup.x, TopOfLastSup.y, TopOfLastSup.z + (HRDim(8) - HRDim(1)) / Sin((PI / 2) - Angle)
       
    CenterOfTorus.Set 0, EndRailPos.y - HRDim(4), EndRailPos.z - HRDim(4) * Tan((PI / 2 - Angle) / 2)
    CenterOfTorus.x = -(TotWidth + Pos)
    Set oTorus = m_GeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(Nothing, CenterOfTorus.x, CenterOfTorus.y, CenterOfTorus.z, _
                                        -1, 0, 0, _
                                        0, 1, 0, _
                                        HRDim(4), HRDiam, -PI / 2 + Angle, False)
    InitNewOutput pOC, "HRLowerTorusR2"
    pOC.AddOutput "HRLowerTorusR2", oTorus
    CenterOfTorus.x = (TotWidth - Pos)
    Set oTorus = m_GeomFactory.Tori3d.CreateByAxisMajorMinorRadiusSweep(Nothing, CenterOfTorus.x, CenterOfTorus.y, CenterOfTorus.z, _
                                        -1, 0, 0, _
                                        0, 1, 0, _
                                        HRDim(4), HRDiam, -PI / 2 + Angle, False)
    InitNewOutput pOC, "HRLowerTorusL2"
    pOC.AddOutput "HRLowerTorusL2", oTorus
    
    'Bottom Vert
    TopOfLastSup.x = -(TotWidth + Pos)
    EndRailPos.x = -(TotWidth + Pos)
    dist = EndRailPos.DistPt(TopOfLastSup) - HRDim(4) * Tan(Angle / 2) - HRDim(4) * Tan((PI / 2 - Angle) / 2)
    CreateCylindricalProjection pOC, HRDiam, dist, TopOfLastSup.x, TopOfLastSup.y, TopOfLastSup.z + HRDim(4) * Tan(Angle / 2), 0, 0, 1, "HRBottomVertR", 1
    TopOfLastSup.x = (TotWidth - Pos)
    CreateCylindricalProjection pOC, HRDiam, dist, TopOfLastSup.x, TopOfLastSup.y, TopOfLastSup.z + HRDim(4) * Tan(Angle / 2), 0, 0, 1, "HRBottomVertL", 1
    
    'Between vert and last full support
    EndRailPos.x = (TotWidth - Pos)
    LastFullSupport.x = (TotWidth - Pos)
    dist = LastFullSupport.DistPt(EndRailPos) - HRDim(4) * Tan((PI / 2 - Angle) / 2)
    CreateCylindricalProjection pOC, HRDiam, dist, LastFullSupport.x, LastFullSupport.y, LastFullSupport.z, TmpPos.x, TmpPos.y, TmpPos.z, "HRHandrailL", 2
    LastFullSupport.x = -(TotWidth + Pos)
    CreateCylindricalProjection pOC, HRDiam, dist, LastFullSupport.x, LastFullSupport.y, LastFullSupport.z, TmpPos.x, TmpPos.y, TmpPos.z, "HRHandrailR", 2

    Exit Sub
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub CreateCylindricalProjection(pOC As IJDOutputCollection, ByVal cylRadius As Double, ByVal cylHeight As Double, _
                 ByVal CenterX As Double, ByVal CenterY As Double, ByVal CenterZ As Double, _
                 ByVal NormalX As Double, ByVal NormalY As Double, ByVal NormalZ As Double, _
                 ByVal stepOrBar As String, _
                 ByVal stepNumber As Integer)

Const METHOD = "CreateCylindricalProjection"
On Error GoTo ErrorHandler

    Dim projCyl As IngrGeom3D.Projection3d
    Dim pCircle As Circle3d
    Set pCircle = New Circle3d
    
    InitNewOutput pOC, stepOrBar & stepNumber
    
    pCircle.DefineByCenterNormalRadius CenterX, CenterY, CenterZ, _
                                           NormalX, NormalY, NormalZ, cylRadius

    Set projCyl = m_GeomFactory.Projections3d.CreateByCurve(Nothing, _
                                         pCircle, NormalX, NormalY, NormalZ, cylHeight, False)

    pOC.AddOutput (stepOrBar & Trim$(Str$(stepNumber))), projCyl

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


Private Function FindIntersection2(oLine1 As IJLine, oLine2 As IJLine) As DPosition
Const METHOD = "FindIntersection2"
On Error GoTo ErrorHandler
    
    Dim v1 As DVector
    Dim v2 As DVector
    Dim StartPoint As DPosition
    Dim EndPoint As DPosition
    Dim Line1Point As DPosition
    Dim Dot_prod As Double
        
    Set v1 = New DVector
    Set v2 = New DVector
    Set StartPoint = New DPosition
    Set EndPoint = New DPosition
    Set Line1Point = New DPosition
    Dim x As Double, y As Double, z As Double
    
    oLine1.GetStartPoint x, y, z
    Line1Point.Set x, y, z

    oLine2.GetStartPoint x, y, z
    StartPoint.Set x, y, z
    oLine2.GetEndPoint x, y, z
    EndPoint.Set x, y, z
    
    v1.Set EndPoint.x - StartPoint.x, EndPoint.y - StartPoint.y, EndPoint.z - StartPoint.z
    v2.Set Line1Point.x - StartPoint.x, Line1Point.y - StartPoint.y, Line1Point.z - StartPoint.z
    
    v1.Length = 1
    Dot_prod = v1.dot(v2)
    
    Dim Intersection As DPosition
    Set Intersection = New DPosition

    Intersection.x = StartPoint.x + (Dot_prod * v1.x)
    Intersection.y = StartPoint.y + (Dot_prod * v1.y)
    Intersection.z = StartPoint.z + (Dot_prod * v1.z)
    Set FindIntersection2 = Intersection
    Set v1 = Nothing
    Set v2 = Nothing
    Set StartPoint = Nothing
    Set EndPoint = Nothing
    Set Line1Point = Nothing
    Set Intersection = Nothing
    
    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

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
    End If
    
    UserAttributeMgmt_Validate = ""

    Exit Function
    
ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
End Function

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
