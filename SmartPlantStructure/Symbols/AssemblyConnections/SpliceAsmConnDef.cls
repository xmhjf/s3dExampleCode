VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpliceAsmConnDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : SpliceAsmConnDef.cls
'
'Author : RP
'
'Description :
'    Macro for creating splice   connection
'
'History:
'
' 08/14/03  RP   Substituted interface name in place of guid
'                in IJDPropertyDescriptions->AddProperty. The guid of
'                a virtual interface may change during bulkload
' 02/02/04  JS   Added code to verify the assembly connection
'                and check if an AC already exists between
'                the supported or supporting members TR#52040
' 02/05/04  JS   Changed all generic reference collection relations to
'                specific relations. This includes the relations
'                to the split axis end ports, split axis along ports
'                and the cutback planes (if a plane exists).
' 08/27/04  MH   CMMigrate
' 06/13/06  RP   Changes due to impact from curved members(DI#84001)
' 27/Sep/06 AS   TR#106632 Remove empty CMSetInputs and CMRemoveInputs on Aggregator
' 07/Jun/07 MH   TR 121116 added CheckForUndefinedValueAndRaiseError
' 12/May/14 NK   TR-CP-253696  ATP - 26 minidumps at 'SPSACMacros' module : Adding checks to avoid access of Null objects
' 18/06/14  knukala CR-CP-223345  Rules should not replace root selector inputs
'********************************************************************

Option Explicit

Private Const MODULE = "SpliceAsmConnDef"

Const m_ItemProgId As String = "SPSACMacros.SpliceAsmConnDef"
Const m_ItemName As String = "SPSACMacros.SpliceAsmConnDef"
Private Const strSourceFile As String = "SpliceAsmConnDef.cls"
Private m_oLocalizer As IJLocalizer
Implements ISPSFACInputHelper
' Declaration of the User Symbol Services interface
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJUserAttributeMgmtParent
Implements IJStructCustomFoulCheck
Public Sub DefinitionInputs(pIH As IJDInputsHelper)
  On Error GoTo ErrorHandler
  
'  pIH.SetInput "Supported1Port"
'  pIH.SetInput "Supported2Port"
  
  Exit Sub
ErrorHandler:
  pIH.ReportError
End Sub



Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(ByRef pDefinition As IJDSymbolDefinition)
  Const MT = "IJDUserSymbolServices_InitializeSymbolDefinition"
  On Error GoTo ErrorHandler


  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  ' Define the inputs -
  Dim pIH As IJDInputsHelper
  Set pIH = New InputHelper
  pIH.Definition = pDefinition
'  pIH.InitAs m_FamilyProgid
  DefinitionInputs pIH
  
  ' Aggregator Type
  Dim pAD As IJDAggregatorDescription
  Set pAD = pDefinition
  pAD.AggregatorClsid = "{E43FD681-1B37-4CC1-BD94-F399F43F946F}"     'CStructAssemblyConnection
  pAD.SetCMSetInputs -1, -1
  pAD.SetCMRemoveInputs -1, -1
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAggregator"
  pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
  Set pAD = Nothing
  
  ' tr 74803
  Dim pCADefinition As IJCADefinition
  Set pCADefinition = pDefinition
  Let pCADefinition.CopyBackwardFlag = igCOPY_BACKWARD_TRIM
  Set pCADefinition = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  'listens to IJUASPSSpliceAsmConn
  'V6 change- this connections has too many IJUAs now. Listen to IJDAttributes from now as it takes care of all IJUAs
  'pAPDs.AddProperty "SpliceAsmConnProperties", 1, "IJUASPSSpliceAsmConn", "CMComputeSpliceAsmConnProperties", imsCOOKIE_ID_USS_LIB
  pAPDs.AddProperty "SpliceAsmConnProperties", 1, "IJDAttributes", "CMComputeSpliceAsmConnProperties", imsCOOKIE_ID_USS_LIB
  Set pAPDs = Nothing
  
  ' Define the members
  Dim pMemberDescriptions As IJDMemberDescriptions
  Dim pMemberDescription As IJDMemberDescription
  Dim pPropertyDescriptions As IJDPropertyDescriptions
  
  Set pMemberDescriptions = pDefinition
  ' Remove all the previous member descriptions
  pMemberDescriptions.RemoveAll
  '_________________________________________________________________________________________________________________________________________
  'Top Flange Plate SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnTopFlangePlate", 1, "CMConstructSpliceAsmConnTopFlangePlate", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnTopFlangePlate"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnTopFlangePlate"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnTopFlangePlate"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "SpliceTopFlangePlateSize", 1, "IJUASPSPlatePartDim", "CMComputeSpliceTopFlangePlateSize", imsCOOKIE_ID_USS_LIB
  'outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "SpliceTopFlangePlateThickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeSpliceTopFlangePlateThickness", imsCOOKIE_ID_USS_LIB
  'outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "SpliceTopFlangePlatePosition", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceTopFlangePlatePosition", imsCOOKIE_ID_USS_LIB
  '_________________________________________________________________________________________________________________________________________
  'Bottom Flange Plate SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnBotFlangePlate", 2, "CMConstructSpliceAsmConnBotFlangePlate", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnBotFlangePlate"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnBotFlangePlate"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnBotFlangePlate"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "SpliceBotFlangePlateSize", 1, "IJUASPSPlatePartDim", "CMComputeSpliceBotFlangePlateSize", imsCOOKIE_ID_USS_LIB
  'outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "SpliceBotFlangePlateThickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeSpliceBotFlangePlateThickness", imsCOOKIE_ID_USS_LIB
  'outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "SpliceBotFlangePlatePosition", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceBotFlangePlatePosition", imsCOOKIE_ID_USS_LIB
  '_________________________________________________________________________________________________________________________________________
  'End Plate1 SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnEndPlate1", 3, "CMConstructSpliceAsmConnEndPlate1", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnEndPlate1"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnEndPlate1"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnEndPlate1"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "SpliceEndPlate1Size", 1, "IJUASPSPlatePartDim", "CMComputeSpliceEndPlate1Size", imsCOOKIE_ID_USS_LIB
  'outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "SpliceEndPlate1Thickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeSpliceEndPlate1Thickness", imsCOOKIE_ID_USS_LIB
  'outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "SpliceEndPlate1Position", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceEndPlate1Position", imsCOOKIE_ID_USS_LIB
    '_________________________________________________________________________________________________________________________________________
  'End Plate2 SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnEndPlate2", 4, "CMConstructSpliceAsmConnEndPlate2", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnEndPlate2"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnEndPlate2"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnEndPlate2"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "SpliceEndPlate2Size", 1, "IJUASPSPlatePartDim", "CMComputeSpliceEndPlate2Size", imsCOOKIE_ID_USS_LIB
  'outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "SpliceEndPlate2Thickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeSpliceEndPlate2Thickness", imsCOOKIE_ID_USS_LIB
  'outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "SpliceEndPlate2Position", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceEndPlate2Position", imsCOOKIE_ID_USS_LIB
  
  
  '_________________________________________________________________________________________________________________________________________
  'Top Flange Fill Plate1 SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnTopFlangeFillPlate1", 5, "CMConstructSpliceAsmConnTopFlangeFillPlate1", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnTopFlangeFillPlate1"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnTopFlangeFillPlate1"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnTopFlangeFillPlate1"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "SpliceTopFlangeFillPlate1Size", 1, "IJUASPSPlatePartDim", "CMComputeSpliceTopFlangeFillPlate1Size", imsCOOKIE_ID_USS_LIB
  'outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "SpliceTopFlangeFillPlate1Thickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeSpliceTopFlangeFillPlate1Thickness", imsCOOKIE_ID_USS_LIB
  'outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "SpliceTopFlangeFillPlate1Position", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceTopFlangeFillPlate1Position", imsCOOKIE_ID_USS_LIB
    '_________________________________________________________________________________________________________________________________________
  'Top Flange Fill Plate2 SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnTopFlangeFillPlate2", 6, "CMConstructSpliceAsmConnTopFlangeFillPlate2", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnTopFlangeFillPlate2"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnTopFlangeFillPlate2"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnTopFlangeFillPlate2"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "SpliceTopFlangeFillPlate2Size", 1, "IJUASPSPlatePartDim", "CMComputeSpliceTopFlangeFillPlate2Size", imsCOOKIE_ID_USS_LIB
  'outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "SpliceTopFlangeFillPlate2Thickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeSpliceTopFlangeFillPlate2Thickness", imsCOOKIE_ID_USS_LIB
  'outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "SpliceTopFlangeFillPlate2Position", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceTopFlangeFillPlate2Position", imsCOOKIE_ID_USS_LIB

  '_________________________________________________________________________________________________________________________________________
  'Bottom Flange Fill Plate1 SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnBotFlangeFillPlate1", 7, "CMConstructSpliceAsmConnBotFlangeFillPlate1", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnBotFlangeFillPlate1"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnBotFlangeFillPlate1"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnBotFlangeFillPlate1"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "SpliceBotFlangeFillPlate1Size", 1, "IJUASPSPlatePartDim", "CMComputeSpliceBotFlangeFillPlate1Size", imsCOOKIE_ID_USS_LIB
  'outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "SpliceBotFlangeFillPlate1Thickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeSpliceBotFlangeFillPlate1Thickness", imsCOOKIE_ID_USS_LIB
  'outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "SpliceBotFlangeFillPlate1Position", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceBotFlangeFillPlate1Position", imsCOOKIE_ID_USS_LIB
  
  '_________________________________________________________________________________________________________________________________________
  'Bottom Flange Fill Plate2 SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnBotFlangeFillPlate2", 8, "CMConstructSpliceAsmConnBotFlangeFillPlate2", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnBotFlangeFillPlate2"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnBotFlangeFillPlate2"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnBotFlangeFillPlate2"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "SpliceBotFlangeFillPlate2Size", 1, "IJUASPSPlatePartDim", "CMComputeSpliceBotFlangeFillPlate2Size", imsCOOKIE_ID_USS_LIB
  'outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "SpliceBotFlangeFillPlate2Thickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeSpliceBotFlangeFillPlate2Thickness", imsCOOKIE_ID_USS_LIB
  'outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "SpliceBotFlangeFillPlate2Position", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceBotFlangeFillPlate2Position", imsCOOKIE_ID_USS_LIB
  
  '_________________________________________________________________________________________________________________________________________
  'Left Web Plate SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnLeftWebPlate", 9, "CMConstructSpliceAsmConnLeftWebPlate", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnLeftWebPlate"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnLeftWebPlate"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnLeftWebPlate"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "SpliceLeftWebPlateSize", 1, "IJUASPSPlatePartDim", "CMComputeSpliceLeftWebPlateSize", imsCOOKIE_ID_USS_LIB
  'outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "SpliceLeftWebPlateThickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeSpliceLeftWebPlateThickness", imsCOOKIE_ID_USS_LIB
  'outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "SpliceLeftWebPlatePosition", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceLeftWebPlatePosition", imsCOOKIE_ID_USS_LIB
  
  '_________________________________________________________________________________________________________________________________________
  'Right Web Plate SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnRightWebPlate", 10, "CMConstructSpliceAsmConnRightWebPlate", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnRightWebPlate"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnRightWebPlate"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnRightWebPlate"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJUASPSPlatePartDim.
  pPropertyDescriptions.AddProperty "SpliceRightWebPlateSize", 1, "IJUASPSPlatePartDim", "CMComputeSpliceRightWebPlateSize", imsCOOKIE_ID_USS_LIB
  'outputs IJStructPlate.
  pPropertyDescriptions.AddProperty "SpliceLeftWebPlateThickness", 2, "{274EE192-A0A5-44E7-B536-8D44A08FA64F}", "CMComputeSpliceRightWebPlateThickness", imsCOOKIE_ID_USS_LIB
  'outputs IJSmartOccurrence.
  pPropertyDescriptions.AddProperty "SpliceRightWebPlatePosition", 3, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceRightWebPlatePosition", imsCOOKIE_ID_USS_LIB
  
  '_________________________________________________________________________________________________________________________________________
  'Cutback Plane
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnCutbackPlane", 11, "CMConstructSpliceAsmConnCutbackPlane", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID

  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnCutbackPlane"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnCutbackPlane"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnCutbackPlane"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJPlane.
  pPropertyDescriptions.AddProperty "ComputeSpliceAsmConnCutbackPlane", 1, "{4317C6B3-D265-11D1-9558-0060973D4824}", "CMComputeSpliceAsmConnCutbackPlane", imsCOOKIE_ID_USS_LIB
  
  '_______________________________________________________________________________________________________________________________________
  'Cutback1 SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnCutback1", 12, "CMConstructSpliceAsmConnCutback1", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnCutback1"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnCutback1"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnCutback1"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJSmartOccurrence. This is just to make sure that the parent computes before the child.
  pPropertyDescriptions.AddProperty "ComputeSpliceAsmConnCutback1", 1, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceAsmConnCutback1", imsCOOKIE_ID_USS_LIB
  '_______________________________________________________________________________________________________________________________________
  'Cutback2 SO
  Set pMemberDescription = pMemberDescriptions.AddMember("SpliceAsmConnCutback2", 13, "CMConstructSpliceAsmConnCutback2", imsCOOKIE_ID_USS_LIB)
  pMemberDescription.RelationshipClsid = CONST_CAToMemberRelationCLSID
  pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalSpliceAsmConnCutback2"
  pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputSpliceAsmConnCutback2"
  pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateSpliceAsmConnCutback2"
  Set pPropertyDescriptions = pMemberDescription
  'outputs IJSmartOccurrence. This is just to make sure that the parent computes before the child.
  pPropertyDescriptions.AddProperty "ComputeSpliceAsmConnCutback2", 1, "{A2A655C0-E2F5-11D4-9825-00104BD1CC25}", "CMComputeSpliceAsmConnCutback2", imsCOOKIE_ID_USS_LIB
  '___________________________________________________________________________________________________________________________________
  Set pMemberDescriptions = Nothing
  Set pMemberDescription = Nothing
  Set pPropertyDescriptions = Nothing

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'Aggregator custom methods.......................................................
Public Sub CMFinalConstructAggregator(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAggregator"
On Error GoTo ErrorHandler
    
  
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMComputeSpliceAsmConnProperties(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeSpliceAsmConnProperties"
On Error GoTo ErrorHandler

  
Exit Sub
ErrorHandler:     HandleError MODULE, METHOD
End Sub
Public Sub CMMigrateAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateAggregator"
    On Error GoTo ErrorHandler
    
    MigrateTheAggregator pAggregatorDescription, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub
'Top Flange Plate SO custom methods............................................................................................
Public Sub CMConditionalSpliceAsmConnTopFlangePlate(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnTopFlangePlate"
  On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim oSectionAttrbs As IJDAttributes
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim spliceType As Long
    Dim secName1 As String, secName2 As String
    Dim rot1 As Double, rot2 As Double
    Dim rotTol As Double
    Dim secType1 As String, secType2 As String
    Dim strError As String
    Dim mirror1 As Boolean, mirror2 As Boolean
    Dim oLine1 As IJLine, oLine2 As IJLine
    
    'get the assemly connection
    Set oSmartOccCAO = pMemberDescription.CAO
    'get the interface to app connection
    Set oStructConn = pMemberDescription.CAO
    'get the ports
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    
    If (oSupped1Port Is Nothing) Or (oSupped2Port Is Nothing) Then
        Dim iJDObject As iJDObject
        Set iJDObject = oStructConn
        iJDObject.Remove
        
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_INVALID_INPUTS_SPLICE, oStructConn, Nothing
        Err.Raise SPS_MACRO_WARNING

    End If
    
    Set oSupped1Part = oSupped1Port.Part
    Set oSupped2Part = oSupped2Port.Part

    ' Verify we do not have an assembly connection already attached to the end ports
    '   because if we do then this asssembly connection needs to have its relations
    '   to the ports severed and the assembly connection added to the ToDo list
    If IsAssemblyConnectionInConflictWithAnother(oStructConn) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_TWOACEXIST_ONEDISABLED, oStructConn, Nothing
        Err.Raise SPS_MACRO_WARNING
        
    End If
    
    

    'check the section type of Supported1 and Supported2
    If Not (ValidSectionType(oSupped1Part) And ValidSectionType(oSupped2Part)) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_UNKNOWNSEC_ENCOUNT, oStructConn, Nothing
        Err.Raise SPS_MACRO_WARNING
       
    End If
    
    
    secType1 = oSupped1Part.CrossSection.Definition.Type
    secType2 = oSupped2Part.CrossSection.Definition.Type
    mirror1 = oSupped1Part.Rotation.Mirror
    mirror2 = oSupped2Part.Rotation.Mirror
    'splice connection is not currently supported for circular cross sections
    If secType1 = "HSSC" Or secType1 = "CS" Or secType1 = "PIPE" Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_SPLICE_NOT_SUPPORTED, oStructConn, Nothing
        Err.Raise SPS_MACRO_WARNING
        
    End If
    If secType2 = "HSSC" Or secType2 = "CS" Or secType2 = "PIPE" Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_SPLICE_NOT_SUPPORTED, oStructConn, Nothing
        Err.Raise SPS_MACRO_WARNING
    End If
    
    
    secName1 = oSupped1Part.CrossSection.SectionName
    secName2 = oSupped2Part.CrossSection.SectionName
    
    
    rot1 = oSupped1Part.Rotation.BetaAngle
    rot2 = oSupped2Part.Rotation.BetaAngle
    rotTol = Abs(rot1 - rot2)
    
    'Splice connection is not suported if supported1 and supported2 have diffrent sections
    If (secName1 <> secName2) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_SPLICE_NOT_SUPPORTED_CS, oStructConn, Nothing
        Err.Raise SPS_MACRO_WARNING
        
    End If
    'Splice connection is not suported if supported1 and supported2 have diffrent rotation or mirror
    If (mirror1 <> mirror2) Or (rotTol > distTol) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_SPLICE_NOT_SUPPORED_ROTATION, oStructConn, Nothing
        Err.Raise SPS_MACRO_WARNING
        
    End If
    'check if supported1 and supported 2 are collinear and end connected. Return error if they are not
    ' get CP5 lines of both members
    
    'get start and end point of supping member
    Dim bStartEnd1 As Boolean, bStartEnd2 As Boolean
    If oSupped1Port.portIndex = SPSMemberAxisStart Then
        bStartEnd1 = True
    End If
    
    If oSupped2Port.portIndex = SPSMemberAxisStart Then
        bStartEnd2 = True
    End If
    
    'find out which one is closer to connected end of supped
    'get the tangent  cp5 line at connected end of members
    
    'check if they are colinear and end connected
    Set oLine1 = GetTangentLineAtCP(oSupped1Part, 5, bStartEnd1)
    Set oLine2 = GetTangentLineAtCP(oSupped2Part, 5, bStartEnd2)
    If Not IsMemberAxesColinearAndEndToEnd(oLine1, oLine2) Then
        SPSToDoErrorNotify ACToDoMsgCodelist, TDL_ACMACROS_MEMBS_NOT_COLINEAR, oStructConn, Nothing
        Err.Raise SPS_MACRO_WARNING
        
    End If
    
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    'copy definition attribute    s to occurrent attributes
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn")) Then 'copy only when occ attributes are empty
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSSpliceAsmConn")
    End If
    'get splice type, flange and web plate, flange plate only etc..
    spliceType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("SpliceWith").Value
    
    CheckForUndefinedValueAndRaiseError oSmartOccCAO, spliceType, "StructACSpliceTypes", 22

    ' there is no top flange for L , 2L
    If secType1 = "L" Or secType1 = "2L" Then
        bIsNeeded = False
    End If
    
    If spliceType <> 1 And spliceType <> 2 Then
        bIsNeeded = False
    End If
        
  Exit Sub
ErrorHandler:
    'set the warning error code so that downstream objects are computed.
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = SPS_MACRO_WARNING Then
        Err.Raise SPS_MACRO_WARNING
    Else
        Err.Raise ReportError(Err, strSourceFile, MT).Number
    End If

End Sub

Public Sub CMConstructSpliceAsmConnTopFlangePlate(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnTopFlangePlate"
    On Error GoTo ErrorHandler
    Dim oPlate As IJStructCustomPlatePart
    Dim oFactory As IJStructCustomPlatePartFactory
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oRefColl   As IJDReferencesCollection
    Dim oSmartOccCAO      As IJSmartOccurrence
    Dim oSymbolOfCAO                As IJDSymbol
    Dim oMembObjects As IJDMemberObjects
    Dim oDesignParent As IJDesignParent
    Dim PlateThickness As Double
    Dim oAttrbs As IJDAttributes
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oMembObjects = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped2Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    
    Set oSuppedPart = oSupped2Port.Part
    
    Set oFactory = New StructCustomPlatePartFactory
    'Create the custom Plate.
    Set oPlate = oFactory.CreateCustomPlatePart(pResourceManager)
    Set oAttrbs = oSmartOccCAO.ItemObject
    On Error Resume Next
        Dim materialProxy As CollectionProxy
        'the interface below is added in V6. It may not exist in older catalog
        'New code and old catalog is unusual. The resume next above takes care of this though
        Set materialProxy = oAttrbs.CollectionOfAttributes("IJUASPSPlateMaterial")
    On Error GoTo ErrorHandler
    
    If Not materialProxy Is Nothing Then
        Dim strMaterial As String, strGrade As String
        Dim pIJAttribsCAO As IJDAttributes
        strMaterial = materialProxy.Item("SPSMaterial").Value
        strGrade = materialProxy.Item("SPSGrade").Value
        'set this as occurrence attribute
        Set pIJAttribsCAO = oSmartOccCAO
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSMaterial").Value = strMaterial
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSGrade").Value = strGrade
        
        SetPlateMaterial oPlate, strMaterial, strGrade
        
    Else
        'use hard coded material, as done prior to V6
        SetPlateMaterial oPlate, "Steel - Carbon", "A"
    End If
    'get default from the parent
    PlateThickness = oAttrbs.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Thickness1").Value
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, PlateThickness
    'set the found thickness on attrbs
'    oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value = PlateThickness
    
    Set oSmartOcc = oPlate
    ' Initialize the SmartOccurrence
    oSmartOcc.RootSelectorClass = "GenericRectPlatePart"
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(ReferencesCollection, pResourceManager)

    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    'create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSmartOcc, oRefColl
    Set pObj = oPlate 'return the created plate object to the custom assembly
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnTopFlangePlate(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnTopFlangePlate"
  On Error GoTo ErrorHandler
  Dim pIJAttribsCAO As IJDAttributes
  Dim oSmartOccCAO As IJSmartOccurrence
  Dim oCustomPlate As StructCustomPlatePart
  
  Set oSmartOccCAO = pMemberDesc.CAO
  Set pIJAttribsCAO = oSmartOccCAO
  On Error Resume Next 'suppress the error that VB throws when QI fails
  Set oCustomPlate = pMemberDesc.object
  On Error GoTo ErrorHandler
  
  If Not oCustomPlate Is Nothing Then
    'copy plate category and type from parent to child
      oCustomPlate.CustomPlatePartCategory = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Category1").Value
      oCustomPlate.CustomPlatePartType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Type1").Value
  End If

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceTopFlangePlateSize(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceTopFlangePlateSize"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim sizeRule As Long
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSectionAttrbs As IJDAttributes
    Dim secDepth As Double, SecWidth As Double
    Dim depth As Double, Width As Double
    Dim oPlate As IJStructPlate
    Dim PlateThickness As Double
    Dim varThickness As Variant
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim gageDist As Double, distFromGage As Double
  
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped1Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    
        Set oSuppedPart = oSupped1Port.Part
        Set oSectionAttrbs = oSuppedPart.CrossSection.definition
        
        secDepth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
        SecWidth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
        On Error Resume Next ' some sections don't support the interface IStructFlangedBoltGage
        gageDist = oSectionAttrbs.CollectionOfAttributes("IStructFlangedBoltGage").Item("gw").Value
        On Error GoTo ErrorHandler

    
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSSpliceAsmConn")
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1"), oAttrbs.CollectionOfAttributes("IJUASPSAsmConnPlatePart1")
    End If
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1"), oAttrbs.CollectionOfAttributes("IJUASPSAsmConnPlatePart1")
    End If
      

    depth = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Length1").Value
    Width = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Width1").Value
    PlateThickness = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Thickness1").Value
    sizeRule = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("SizingRule").Value
    distFromGage = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("GageDistance1").Value

    CheckForUndefinedValueAndRaiseError oSmartOccCAO, sizeRule, "StructACSizingRule", 20

    Set oSmartOccChild = pPropertyDescriptions.object
    Set pIJAttribsChild = oSmartOccChild
    
    If Not pIJAttribsChild Is Nothing Then
        pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value = depth
        If sizeRule <> 2 Then ' Width is  calculated from gage
            If gageDist > 0 Then ' some sections don't have gage distance specified in the catalog. In that case we use the width
                pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = gageDist + 2 * distFromGage
            Else
                pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = Width
            End If
        Else 'user defined.
            pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = Width
        
        End If
        Set oPlate = oSmartOccChild
        On Error Resume Next
            Dim materialProxy As CollectionProxy
            'the interface below is added in V6. It may not exist in older catalog
            'New code and old catalog is unusual. The resume next above takes care of this though
            Set materialProxy = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial")
        On Error GoTo ErrorHandler
        If Not materialProxy Is Nothing Then
            If IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial")) Then
                Dim strMaterial As String, strGrade As String
                
                strMaterial = materialProxy.Item("SPSMaterial").Value
                strGrade = materialProxy.Item("SPSGrade").Value
                SetPlateMaterial oPlate, strMaterial, strGrade
            Else
                Dim oIJDMaterial As IJDMaterial
                Dim oPlateMaterial As IJStructMaterial
                Set oPlateMaterial = oPlate
                Set oIJDMaterial = oPlateMaterial.StructMaterial
                'copy material from  child to parent
                'this can only happen to a connection which was migrated from v5 to v6
                pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSMaterial").Value = oIJDMaterial.MaterialType
                pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSGrade").Value = oIJDMaterial.MaterialGrade
            End If
        End If
        
        SetPlateDimensions oPlate, PlateThickness
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Thickness1").Value = PlateThickness
    End If
    
  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
Public Sub CMComputeSpliceTopFlangePlateThickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceTopFlangePlateThickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceTopFlangePlatePosition(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceTopFlangePlatePosition"
    On Error GoTo ErrorHandler
  
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim iPort As SPSMemberAxisPortIndex
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim X0 As Double, Y0 As Double, Z0 As Double
    Dim oLine As IJLine
    Dim oAttrb As IJDAttributes
    Dim xVec As New DVector, yVec As New DVector, zVec As New DVector
    Dim clearance#
    Dim oPos As New DPosition, oPlateOrg As New DPosition
    Dim dist As Double
    Dim oMat As New DT4x4
    Dim oOcc As IJDOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim align As Long
    Dim bStartPoint As Boolean
    Dim oEndPos As New DPosition
    
        
    Set oOcc = pObject
    Set oSmartOccChild = pObject
    Set oAttrb = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped2Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    iPort = oSupped2Port.portIndex
    Set oSupped2Part = oSupped2Port.Part
    
    align = oAttrb.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Symmetry").Value
    clearance = oAttrb.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Clearance").Value
    
    CheckForUndefinedValueAndRaiseError oStructConn, align, "StructAlignment", 21

    oSupped2Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1
    If iPort = SPSMemberAxisStart Then
        oEndPos.Set X0, Y0, Z0
        bStartPoint = True
    ElseIf iPort = SPSMemberAxisEnd Then
        oEndPos.Set x1, y1, z1
        bStartPoint = False
    End If
    

    Set oLine = GetTangentLineAtCP(oSupped2Part, 8, bStartPoint)
    
    oLine.GetDirection x1, y1, z1
    xVec.Set x1, y1, z1
    
    If bStartPoint Then
        oLine.GetStartPoint X0, Y0, Z0
    Else
        oLine.GetEndPoint X0, Y0, Z0
        xVec.[Scale] -1 'reverse direction
    End If
    oPos.Set X0, Y0, Z0

    If align = 1 Then 'center
        dist = 0 'clearance / 2
    ElseIf align = 2 Then 'left, only supporte1 is cutback
        dist = -clearance / 2
    ElseIf align = 3 Then 'right , only supported2 is cutback
        dist = clearance / 2
    End If
    
    oPos.x = oPos.x + xVec.x * dist
    oPos.y = oPos.y + xVec.y * dist
    oPos.z = oPos.z + xVec.z * dist
    
    Set zVec = GetZVector(oSupped2Part, oEndPos)
    Set yVec = zVec.Cross(xVec)
    yVec.length = 1

    oPlateOrg.Set oPos.x, oPos.y, oPos.z
    
    oMat.LoadIdentity
    'X-axis of plate
    oMat.IndexValue(0) = xVec.x
    oMat.IndexValue(1) = xVec.y
    oMat.IndexValue(2) = xVec.z
    'Y-axis of plate
    oMat.IndexValue(4) = yVec.x
    oMat.IndexValue(5) = yVec.y
    oMat.IndexValue(6) = yVec.z
    'Z-axis of plate
    oMat.IndexValue(8) = zVec.x
    oMat.IndexValue(9) = zVec.y
    oMat.IndexValue(10) = zVec.z
    'Origin of plate
    oMat.IndexValue(12) = oPlateOrg.x
    oMat.IndexValue(13) = oPlateOrg.y
    oMat.IndexValue(14) = oPlateOrg.z
    
    'set the matrix on the Baseplate occurrence
    oOcc.Matrix = oMat
    oSmartOccChild.Update ' mark IJSmartOccurrence as modified so that the baseplate generation semantic fires

  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
Public Sub CMMigrateSpliceAsmConnTopFlangePlate(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateSpliceAsmConnTopFlangePlate"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub
'Bot Flange Plate SO custom methods............................................................................................
Public Sub CMConditionalSpliceAsmConnBotFlangePlate(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnBotFlangePlate"
  On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim oSectionAttrbs As IJDAttributes
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim spliceType As Long
    Dim secName1 As String, secName2 As String
    Dim rot1 As Double, rot2 As Double
    Dim rotTol As Double
    Dim secType As String
  
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    
    
    If colPorts.count < 2 Then
        Dim iJDObject As iJDObject
        Set iJDObject = oStructConn
        iJDObject.Remove
        Err.Raise E_FAIL
    End If

    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If

    If (oSupped1Port Is Nothing) Or (oSupped2Port Is Nothing) Then
        Err.Raise SPS_MACRO_WARNING
    End If
    
    Set oSupped1Part = oSupped1Port.Part
    Set oSupped2Part = oSupped2Port.Part
    
    secType = oSupped1Part.CrossSection.Definition.Type
    secName1 = oSupped1Part.CrossSection.SectionName
    secName2 = oSupped2Part.CrossSection.SectionName
    rot1 = oSupped1Part.Rotation.BetaAngle
    rot2 = oSupped2Part.Rotation.BetaAngle
    rotTol = Abs(rot1 - rot2)
    
    If (secName1 <> secName2) Or (rotTol > distTol) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSSpliceAsmConn")
    End If
    
    If secType = "WT" Or secType = "MT" Or secType = "ST" Then
        bIsNeeded = False
    End If
    
    spliceType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("SpliceWith").Value
    
    CheckForUndefinedValueAndRaiseError oSmartOccCAO, spliceType, "StructACSpliceTypes", 22

    If spliceType <> 1 And spliceType <> 2 Then
        bIsNeeded = False
    End If

  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

Public Sub CMConstructSpliceAsmConnBotFlangePlate(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnBotFlangePlate"
    On Error GoTo ErrorHandler
    Dim oPlate As IJStructCustomPlatePart
    Dim oFactory As IJStructCustomPlatePartFactory
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oRefColl   As IJDReferencesCollection
    Dim oSmartOccCAO      As IJSmartOccurrence
    Dim oSymbolOfCAO                As IJDSymbol
    Dim oMembObjects As IJDMemberObjects
    Dim oDesignParent As IJDesignParent
    Dim PlateThickness As Double
    Dim oAttrbs As IJDAttributes
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oMembObjects = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped2Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    
    Set oSuppedPart = oSupped2Port.Part
    
    Set oFactory = New StructCustomPlatePartFactory
    'Create the custom Plate.
    Set oPlate = oFactory.CreateCustomPlatePart(pResourceManager)
    Set oAttrbs = oSmartOccCAO.ItemObject
    On Error Resume Next
        Dim materialProxy As CollectionProxy
        'the interface below is added in V6. It may not exist in older catalog
        'New code and old catalog is unusual. The resume next above takes care of this though
        Set materialProxy = oAttrbs.CollectionOfAttributes("IJUASPSPlateMaterial")
    On Error GoTo ErrorHandler
    
    If Not materialProxy Is Nothing Then
        Dim strMaterial As String, strGrade As String
        Dim pIJAttribsCAO As IJDAttributes
        strMaterial = materialProxy.Item("SPSMaterial").Value
        strGrade = materialProxy.Item("SPSGrade").Value
        'set this as occurrence attribute
        Set pIJAttribsCAO = oSmartOccCAO
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSMaterial").Value = strMaterial
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSGrade").Value = strGrade
        
        SetPlateMaterial oPlate, strMaterial, strGrade
    Else
        'use hard coded material, as done prior to V6
        SetPlateMaterial oPlate, "Steel - Carbon", "A"
    End If
    'get default from the parent
    PlateThickness = oAttrbs.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Thickness1").Value
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, PlateThickness
    'set the found thickness on attrbs
'    oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value = PlateThickness
    
    Set oSmartOcc = oPlate
    ' Initialize the SmartOccurrence
    oSmartOcc.RootSelectorClass = "GenericRectPlatePart"
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(ReferencesCollection, pResourceManager)

    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    'create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSmartOcc, oRefColl
    Set pObj = oPlate 'return the created plate object to the custom assembly
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnBotFlangePlate(pMemberDesc As IJDMemberDescription)
    Const MT = "CMSetInputSpliceAsmConnBotFlangePlate"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oCustomPlate As StructCustomPlatePart
  
    Set oSmartOccCAO = pMemberDesc.CAO
    Set pIJAttribsCAO = oSmartOccCAO
    On Error Resume Next 'suppress the error that VB throws when QI fails
    Set oCustomPlate = pMemberDesc.object
    On Error GoTo ErrorHandler
  
    If Not oCustomPlate Is Nothing Then
        'copy plate category and type from parent to child
        oCustomPlate.CustomPlatePartCategory = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Category1").Value
        oCustomPlate.CustomPlatePartType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Type1").Value
    End If

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceBotFlangePlateSize(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceBotFlangePlateSize"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim sizeRule As Long
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSectionAttrbs As IJDAttributes
    Dim secDepth As Double, SecWidth As Double
    Dim depth As Double, Width As Double
    Dim oPlate As IJStructPlate
    Dim PlateThickness As Double
    Dim varThickness As Variant
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim gageDist As Double, distFromGage As Double
  
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped1Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If

    Set oSuppedPart = oSupped1Port.Part
    Set oSectionAttrbs = oSuppedPart.CrossSection.Definition
    
    secDepth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    SecWidth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    On Error Resume Next ' some sections don't support the interface IStructFlangedBoltGage
    gageDist = oSectionAttrbs.CollectionOfAttributes("IStructFlangedBoltGage").Item("gw").Value
    On Error GoTo ErrorHandler
    
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSSpliceAsmConn")
    End If
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1"), oAttrbs.CollectionOfAttributes("IJUASPSAsmConnPlatePart1")
    End If
      

    depth = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Length1").Value
    Width = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Width1").Value
    PlateThickness = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Thickness1").Value
    sizeRule = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("SizingRule").Value
    distFromGage = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("GageDistance1").Value

    CheckForUndefinedValueAndRaiseError oSmartOccCAO, sizeRule, "StructACSizingRule", 20

    Set oSmartOccChild = pPropertyDescriptions.object
    Set pIJAttribsChild = oSmartOccChild
    
    If Not pIJAttribsChild Is Nothing Then
        pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value = depth
        If sizeRule <> 2 Then ' Width is  calculated from gage
            If gageDist > 0 Then ' some sections don't have gage distance specified in the catalog. In that case we use the width
                pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = gageDist + 2 * distFromGage
            Else
                pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = Width
            End If
        Else 'user defined.
            pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = Width
       
        End If
        Set oPlate = oSmartOccChild
        
        On Error Resume Next
            Dim materialProxy As CollectionProxy
            'the interface below is added in V6. It may not exist in older catalog
            'New code and old catalog is unusual. The resume next above takes care of this though
            Set materialProxy = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial")
        On Error GoTo ErrorHandler
        If Not materialProxy Is Nothing Then
            If IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial")) Then
                Dim strMaterial As String, strGrade As String
                
                strMaterial = materialProxy.Item("SPSMaterial").Value
                strGrade = materialProxy.Item("SPSGrade").Value
                SetPlateMaterial oPlate, strMaterial, strGrade
            Else
                Dim oIJDMaterial As IJDMaterial
                Dim oPlateMaterial As IJStructMaterial
                Set oPlateMaterial = oPlate
                Set oIJDMaterial = oPlateMaterial.StructMaterial
                'copy material from  child to parent
                'this can only happen to a connection which was migrated from v5 to v6
                pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSMaterial").Value = oIJDMaterial.MaterialType
                pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSGrade").Value = oIJDMaterial.MaterialGrade
            End If
        End If
        
        SetPlateDimensions oPlate, PlateThickness
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart1").Item("Thickness1").Value = PlateThickness
    End If
    
  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
Public Sub CMComputeSpliceBotFlangePlateThickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceBotFlangePlateThickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceBotFlangePlatePosition(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceBotFlangePlatePosition"
    On Error GoTo ErrorHandler
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim iPort As SPSMemberAxisPortIndex
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim X0 As Double, Y0 As Double, Z0 As Double
    Dim oLine As IJLine
    Dim oAttrb As IJDAttributes
    Dim xVec As New DVector, yVec As New DVector, zVec As New DVector
    Dim clearance#
    Dim oPos As New DPosition, oPlateOrg As New DPosition
    Dim dist As Double
    Dim oMat As New DT4x4
    Dim oOcc As IJDOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim align As Long
    Dim bStartPoint As Boolean
    Dim oEndPos As New DPosition
    
        
    Set oOcc = pObject
    Set oSmartOccChild = pObject
    Set oAttrb = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If

    If oSupped2Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    
    iPort = oSupped2Port.portIndex
    Set oSupped2Part = oSupped2Port.Part
    
    align = oAttrb.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Symmetry").Value
    clearance = oAttrb.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Clearance").Value

    CheckForUndefinedValueAndRaiseError oStructConn, align, "StructAlignment", 21

    oSupped2Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1
    If iPort = SPSMemberAxisStart Then
        oEndPos.Set X0, Y0, Z0
        bStartPoint = True
    ElseIf iPort = SPSMemberAxisEnd Then
        oEndPos.Set x1, y1, z1
        bStartPoint = False
    End If
    
    Set oLine = GetTangentLineAtCP(oSupped2Part, 2, bStartPoint)
    oLine.GetDirection x1, y1, z1
    xVec.Set x1, y1, z1
    
    If bStartPoint Then
        oLine.GetStartPoint X0, Y0, Z0
    Else
        oLine.GetEndPoint X0, Y0, Z0
        xVec.[Scale] -1 'reverse direction
    End If
    oPos.Set X0, Y0, Z0
    

    If align = 1 Then 'center
        dist = 0 'clearance / 2
    ElseIf align = 2 Then 'left, only supporte1 is cutback
        dist = -clearance / 2
    ElseIf align = 3 Then 'right , only supported2 is cutback
        dist = clearance / 2
    End If
    
    oPos.x = oPos.x + xVec.x * dist
    oPos.y = oPos.y + xVec.y * dist
    oPos.z = oPos.z + xVec.z * dist
    
    Set zVec = GetZVector(oSupped2Part, oEndPos)
    ' reverse direction so that the vec is outward from cp2
    zVec.x = -zVec.x
    zVec.y = -zVec.y
    zVec.z = -zVec.z
    Set yVec = zVec.Cross(xVec)
    yVec.length = 1

    oPlateOrg.Set oPos.x, oPos.y, oPos.z
    
    oMat.LoadIdentity
    'X-axis of plate
    oMat.IndexValue(0) = xVec.x
    oMat.IndexValue(1) = xVec.y
    oMat.IndexValue(2) = xVec.z
    'Y-axis of plate
    oMat.IndexValue(4) = yVec.x
    oMat.IndexValue(5) = yVec.y
    oMat.IndexValue(6) = yVec.z
    'Z-axis of plate
    oMat.IndexValue(8) = zVec.x
    oMat.IndexValue(9) = zVec.y
    oMat.IndexValue(10) = zVec.z
    'Origin of plate
    oMat.IndexValue(12) = oPlateOrg.x
    oMat.IndexValue(13) = oPlateOrg.y
    oMat.IndexValue(14) = oPlateOrg.z
    
    'set the matrix on the Baseplate occurrence
    oOcc.Matrix = oMat
    oSmartOccChild.Update ' mark IJSmartOccurrence as modified so that the baseplate generation semantic fires

  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
Public Sub CMMigrateSpliceAsmConnBotFlangePlate(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateSpliceAsmConnBotFlangePlate"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub
'Top Flange Fill Plate1 SO custom methods............................................................................................
Public Sub CMConditionalSpliceAsmConnTopFlangeFillPlate1(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnTopFlangeFillPlate1"
  On Error GoTo ErrorHandler
    bIsNeeded = False
  Exit Sub
ErrorHandler:  HandleError MODULE, MT

End Sub

Public Sub CMConstructSpliceAsmConnTopFlangeFillPlate1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnTopFlangeFillPlate1"
    On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnTopFlangeFillPlate1(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnTopFlangeFillPlate1"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceTopFlangeFillPlate1Size(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceTopFlangeFillPlate1Size"
    On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceTopFlangeFillPlate1Thickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceTopFlangeFillPlate1Thickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceTopFlangeFillPlate1Position(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceTopFlangeFillPlate1Position"
    On Error GoTo ErrorHandler
  


  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMMigrateSpliceAsmConnTopFlangeFillPlate1(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateSpliceAsmConnTopFlangeFillPlate1"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub
'Top Flange Fill Plate2 SO custom methods............................................................................................
Public Sub CMConditionalSpliceAsmConnTopFlangeFillPlate2(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnTopFlangeFillPlate2"
  On Error GoTo ErrorHandler
    bIsNeeded = False
  Exit Sub
ErrorHandler:  HandleError MODULE, MT

End Sub

Public Sub CMConstructSpliceAsmConnTopFlangeFillPlate2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnTopFlangeFillPlate2"
    On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnTopFlangeFillPlate2(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnTopFlangeFillPlate2"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceTopFlangeFillPlate2Size(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceTopFlangeFillPlate2Size"
    On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceTopFlangeFillPlate2Thickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceTopFlangeFillPlate2Thickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceTopFlangeFillPlate2Position(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceTopFlangeFillPlate2Position"
    On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMMigrateSpliceAsmConnTopFlangeFillPlate2(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateSpliceAsmConnTopFlangeFillPlate2"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub
'End Plate1 SO custom methods............................................................................................
Public Sub CMConditionalSpliceAsmConnEndPlate1(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnEndPlate1"
  On Error GoTo ErrorHandler
    bIsNeeded = False
  Exit Sub
ErrorHandler:  HandleError MODULE, MT

End Sub

Public Sub CMConstructSpliceAsmConnEndPlate1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnEndPlate1"
    On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnEndPlate1(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnEndPlate1"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceEndPlate1Size(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceEndPlate1Size"
    On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceEndPlate1Thickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceEndPlate1Thickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceEndPlate1Position(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceEndPlate1Position"
    On Error GoTo ErrorHandler
  


  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMMigrateSpliceAsmConnEndPlate1(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const MT = "CMMigrateSpliceAsmConnEndPlate1"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
'End Plate2 SO custom methods............................................................................................
Public Sub CMConditionalSpliceAsmConnEndPlate2(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnEndPlate2"
  On Error GoTo ErrorHandler
    bIsNeeded = False
  Exit Sub
ErrorHandler:  HandleError MODULE, MT

End Sub

Public Sub CMConstructSpliceAsmConnEndPlate2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnEndPlate2"
    On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnEndPlate2(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnEndPlate2"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceEndPlate2Size(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceEndPlate2Size"
    On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceEndPlate2Thickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceEndPlate2Thickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceEndPlate2Position(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceEndPlate2Position"
    On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMMigrateSpliceAsmConnEndPlate2(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const MT = "CMMigrateSpliceAsmConnEndPlate2"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub



'Bot Flange Fill Plate1 SO custom methods............................................................................................
Public Sub CMConditionalSpliceAsmConnBotFlangeFillPlate1(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnBotFlangeFillPlate1"
  On Error GoTo ErrorHandler
    bIsNeeded = False
  Exit Sub
ErrorHandler:  HandleError MODULE, MT

End Sub

Public Sub CMConstructSpliceAsmConnBotFlangeFillPlate1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnBotFlangeFillPlate1"
    On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnBotFlangeFillPlate1(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnBotFlangeFillPlate1"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceBotFlangeFillPlate1Size(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceBotFlangeFillPlate1Size"
    On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceBotFlangeFillPlate1Thickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceBotFlangeFillPlate1Thickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceBotFlangeFillPlate1Position(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceBotFlangeFillPlate1Position"
    On Error GoTo ErrorHandler


  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMMigrateSpliceAsmConnBotFlangeFillPlate1(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateSpliceAsmConnBotFlangeFillPlate1"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub
'Bot Flange Fill Plate2 SO custom methods............................................................................................
Public Sub CMConditionalSpliceAsmConnBotFlangeFillPlate2(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnBotFlangeFillPlate2"
  On Error GoTo ErrorHandler
    bIsNeeded = False
  Exit Sub
ErrorHandler:  HandleError MODULE, MT

End Sub

Public Sub CMConstructSpliceAsmConnBotFlangeFillPlate2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnBotFlangeFillPlate2"
    On Error GoTo ErrorHandler
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnBotFlangeFillPlate2(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnBotFlangeFillPlate2"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceBotFlangeFillPlate2Size(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceBotFlangeFillPlate2Size"
    On Error GoTo ErrorHandler
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceBotFlangeFillPlate2Thickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceBotFlangeFillPlate2Thickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceBotFlangeFillPlate2Position(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceBotFlangeFillPlate2Position"
    On Error GoTo ErrorHandler


  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMMigrateSpliceAsmConnBotFlangeFillPlate2(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateSpliceAsmConnBotFlangeFillPlate2"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub

'Left Web Plate SO custom methods............................................................................................
Public Sub CMConditionalSpliceAsmConnLeftWebPlate(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnLeftWebPlate"
  On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim oSectionAttrbs As IJDAttributes
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim spliceType As Long
    Dim secName1 As String, secName2 As String
    Dim rot1 As Double, rot2 As Double
    Dim rotTol As Double
    Dim webPlatePos As Long
  
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    

    If colPorts.count < 2 Then
        Dim iJDObject As iJDObject
        Set iJDObject = oStructConn
        iJDObject.Remove
        Err.Raise E_FAIL
    End If

    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If (oSupped1Port Is Nothing) Or (oSupped2Port Is Nothing) Then
        Err.Raise SPS_MACRO_WARNING
    End If
    Set oSupped1Part = oSupped1Port.Part
    Set oSupped2Part = oSupped2Port.Part
    secName1 = oSupped1Part.CrossSection.SectionName
    secName2 = oSupped2Part.CrossSection.SectionName
    rot1 = oSupped1Part.Rotation.BetaAngle
    rot2 = oSupped2Part.Rotation.BetaAngle
    rotTol = Abs(rot1 - rot2)
    
    If (secName1 <> secName2) Or (rotTol > distTol) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSSpliceAsmConn")
    End If
    
    spliceType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("SpliceWith").Value
    webPlatePos = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("WebPlates").Value

    CheckForUndefinedValueAndRaiseError oSmartOccCAO, spliceType, "StructACSpliceTypes", 22
    CheckForUndefinedValueAndRaiseError oSmartOccCAO, webPlatePos, "StructACLeftRightBoth", 23

    If spliceType <> 1 And spliceType <> 3 Then
        bIsNeeded = False
    ElseIf webPlatePos <> 1 And webPlatePos <> 2 Then
        bIsNeeded = False
    End If

  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

Public Sub CMConstructSpliceAsmConnLeftWebPlate(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnLeftWebPlate"
    On Error GoTo ErrorHandler
    Dim oPlate As IJStructCustomPlatePart
    Dim oFactory As IJStructCustomPlatePartFactory
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oRefColl   As IJDReferencesCollection
    Dim oSmartOccCAO      As IJSmartOccurrence
    Dim oSymbolOfCAO                As IJDSymbol
    Dim oMembObjects As IJDMemberObjects
    Dim oDesignParent As IJDesignParent
    Dim PlateThickness As Double
    Dim oAttrbs As IJDAttributes
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oMembObjects = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped2Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If

    Set oSuppedPart = oSupped2Port.Part
    
    Set oFactory = New StructCustomPlatePartFactory
    'Create the custom Plate.
    Set oPlate = oFactory.CreateCustomPlatePart(pResourceManager)
    Set oAttrbs = oSmartOccCAO.ItemObject
    On Error Resume Next
        Dim materialProxy As CollectionProxy
        'the interface below is added in V6. It may not exist in older catalog
        'New code and old catalog is unusual. The resume next above takes care of this though
        Set materialProxy = oAttrbs.CollectionOfAttributes("IJUASPSPlateMaterial")
    On Error GoTo ErrorHandler
    
    If Not materialProxy Is Nothing Then
        Dim strMaterial As String, strGrade As String
        Dim pIJAttribsCAO As IJDAttributes
        strMaterial = materialProxy.Item("SPSMaterial").Value
        strGrade = materialProxy.Item("SPSGrade").Value
        'set this as occurrence attribute
        Set pIJAttribsCAO = oSmartOccCAO
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSMaterial").Value = strMaterial
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSGrade").Value = strGrade
        
        SetPlateMaterial oPlate, strMaterial, strGrade
    Else
        'use hard coded material, as done prior to V6
        SetPlateMaterial oPlate, "Steel - Carbon", "A"
    End If
    
    'get default from the parent
    PlateThickness = oAttrbs.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Thickness2").Value
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, PlateThickness
    'set the found thickness on attrbs
'    oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value = PlateThickness
    
    Set oSmartOcc = oPlate
    ' Initialize the SmartOccurrence
    oSmartOcc.RootSelectorClass = "GenericRectPlatePart"
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(ReferencesCollection, pResourceManager)

    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    'create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSmartOcc, oRefColl
    Set pObj = oPlate 'return the created plate object to the custom assembly
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnLeftWebPlate(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnLeftWebPlate"
  On Error GoTo ErrorHandler
  Dim pIJAttribsCAO As IJDAttributes
  Dim oSmartOccCAO As IJSmartOccurrence
  Dim oCustomPlate As StructCustomPlatePart
  
  Set oSmartOccCAO = pMemberDesc.CAO
  Set pIJAttribsCAO = oSmartOccCAO
  On Error Resume Next 'suppress the error that VB throws when QI fails
  Set oCustomPlate = pMemberDesc.object
  On Error GoTo ErrorHandler
  
  If Not oCustomPlate Is Nothing Then
    'copy plate category and type from parent to child
      oCustomPlate.CustomPlatePartCategory = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Category2").Value
      oCustomPlate.CustomPlatePartType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Type2").Value
  End If

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceLeftWebPlateSize(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceLeftWebPlateSize"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim sizeRule As Long
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSectionAttrbs As IJDAttributes
    Dim secDepth As Double, SecWidth As Double
    Dim depth As Double, Width As Double
    Dim oPlate As IJStructPlate
    Dim PlateThickness As Double
    Dim varThickness As Variant
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim gageDist As Double, distFromGage As Double
  
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If

    If oSupped1Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    
    Set oSuppedPart = oSupped1Port.Part
    Set oSectionAttrbs = oSuppedPart.CrossSection.Definition
    
    secDepth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    SecWidth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    On Error Resume Next ' some sections don't support the interface IStructFlangedBoltGage
    gageDist = oSectionAttrbs.CollectionOfAttributes("IStructFlangedBoltGage").Item("gw").Value
    On Error GoTo ErrorHandler
    
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSSpliceAsmConn")
    End If
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2"), oAttrbs.CollectionOfAttributes("IJUASPSAsmConnPlatePart2")
    End If
      

    depth = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Length2").Value
    Width = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Width2").Value
    PlateThickness = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Thickness2").Value
    sizeRule = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("SizingRule").Value
    distFromGage = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("GageDistance2").Value

    CheckForUndefinedValueAndRaiseError oSmartOccCAO, sizeRule, "StructACSizingRule", 20

    Set oSmartOccChild = pPropertyDescriptions.object
    Set pIJAttribsChild = oSmartOccChild
    
    If Not pIJAttribsChild Is Nothing Then
        pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value = depth
        If sizeRule <> 2 Then ' Width is  calculated from gage
            If gageDist > 0 Then ' some sections don't have gage distance specified in the catalog. In that case we use the width
                pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = (secDepth - 2 * gageDist) + 2 * distFromGage
            Else
                pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = Width
            End If
        Else 'user defined.
            pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = Width
        End If
        Set oPlate = oSmartOccChild
        On Error Resume Next
            Dim materialProxy As CollectionProxy
            'the interface below is added in V6. It may not exist in older catalog
            'New code and old catalog is unusual. The resume next above takes care of this though
            Set materialProxy = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial")
        On Error GoTo ErrorHandler
        If Not materialProxy Is Nothing Then
            If IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial")) Then
                Dim strMaterial As String, strGrade As String
                
                strMaterial = materialProxy.Item("SPSMaterial").Value
                strGrade = materialProxy.Item("SPSGrade").Value
                SetPlateMaterial oPlate, strMaterial, strGrade
            Else
                Dim oIJDMaterial As IJDMaterial
                Dim oPlateMaterial As IJStructMaterial
                Set oPlateMaterial = oPlate
                Set oIJDMaterial = oPlateMaterial.StructMaterial
                'copy material from  child to parent
                'this can only happen to a connection which was migrated from v5 to v6
                pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSMaterial").Value = oIJDMaterial.MaterialType
                pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSGrade").Value = oIJDMaterial.MaterialGrade
            End If
        End If
        
        SetPlateDimensions oPlate, PlateThickness
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Thickness2").Value = PlateThickness
    End If
  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
Public Sub CMComputeSpliceLeftWebPlateThickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceLeftWebPlateThickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub



Public Sub CMComputeSpliceLeftWebPlatePosition(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceLeftWebPlatePosition"
    On Error GoTo ErrorHandler
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim iPort As SPSMemberAxisPortIndex
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim X0 As Double, Y0 As Double, Z0 As Double
    Dim oLine As IJLine
    Dim oAttrb As IJDAttributes
    Dim xVec As New DVector, yVec As New DVector, zVec As New DVector
    Dim clearance#, tWeb#
    Dim oPos As New DPosition, oPlateOrg As New DPosition
    Dim dist As Double, sBtoB As Double
    Dim oMat As New DT4x4
    Dim oOcc As IJDOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim align As Long
    Dim offsetDist#
    Dim secType As String
    Dim bStartPoint As Boolean
    Dim oEndPos As New DPosition

    
 
    Set oOcc = pObject
    Set oSmartOccChild = pObject
    Set oAttrb = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If

    If oSupped2Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    iPort = oSupped2Port.portIndex
    Set oSupped2Part = oSupped2Port.Part
    
    align = oAttrb.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Symmetry").Value
    clearance = oAttrb.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Clearance").Value
    Set oAttrb = oSupped2Part.CrossSection.Definition

    CheckForUndefinedValueAndRaiseError oStructConn, align, "StructAlignment", 21

    oSupped2Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1
    If iPort = SPSMemberAxisStart Then
        oEndPos.Set X0, Y0, Z0
        bStartPoint = True
    ElseIf iPort = SPSMemberAxisEnd Then
        oEndPos.Set x1, y1, z1
        bStartPoint = False
    End If
    
    
    secType = oSupped2Part.CrossSection.Definition.Type
    
    Select Case secType
        Case "C", "MC", "L"
            Set oLine = GetTangentLineAtCP(oSupped2Part, 4, bStartPoint)
            If oSupped2Part.Rotation.Mirror = False Then
                offsetDist = 0
            Else
                tWeb = oAttrb.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
                offsetDist = tWeb
            End If
        Case "RS", "HSSR"

            If oSupped2Part.Rotation.Mirror = False Then
                Set oLine = GetTangentLineAtCP(oSupped2Part, 4, bStartPoint)
                offsetDist = 0
            Else
                Set oLine = GetTangentLineAtCP(oSupped2Part, 6, bStartPoint)
                offsetDist = 0
            End If
        Case "2L"
            Set oLine = GetTangentLineAtCP(oSupped2Part, 5, bStartPoint)
            tWeb = oAttrb.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
            sBtoB = oAttrb.CollectionOfAttributes("IJUA2L").Item("bb").Value
            offsetDist = tWeb + sBtoB / 2

        Case Else '"W", "S", "HP", "M","WT", "MT", "ST"
            Set oLine = GetTangentLineAtCP(oSupped2Part, 5, bStartPoint)
            tWeb = oAttrb.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
            offsetDist = tWeb / 2
    End Select
    
    oLine.GetDirection x1, y1, z1
    xVec.Set x1, y1, z1
    
    If bStartPoint Then
        oLine.GetStartPoint X0, Y0, Z0
    Else
        oLine.GetEndPoint X0, Y0, Z0
        xVec.[Scale] -1 'reverse direction
    End If
    oPos.Set X0, Y0, Z0
    

    If align = 1 Then 'center
        dist = 0 'clearance / 2
    ElseIf align = 2 Then 'left, only supporte1 is cutback
        dist = -clearance / 2
    ElseIf align = 3 Then 'right , only supported2 is cutback
        dist = clearance / 2
    End If
    
    oPos.x = oPos.x + xVec.x * dist
    oPos.y = oPos.y + xVec.y * dist
    oPos.z = oPos.z + xVec.z * dist
    
    Set zVec = GetYVector(oSupped2Part, oEndPos)
'    ' If the section is mirrored reverse direction so that the plate is on the same face
'    Commented  as left and right is not affected by miiror
'    If oSupped2Part.Rotation.Mirror = True Then
'        zVec.x = -zVec.x
'        zVec.y = -zVec.y
'        zVec.z = -zVec.z
'        tWeb = -tWeb
'    End If
    
    Set yVec = zVec.Cross(xVec)
    yVec.length = 1

    oPlateOrg.Set oPos.x, oPos.y, oPos.z
    
    
    oMat.LoadIdentity
    'X-axis of plate
    oMat.IndexValue(0) = xVec.x
    oMat.IndexValue(1) = xVec.y
    oMat.IndexValue(2) = xVec.z
    'Y-axis of plate
    oMat.IndexValue(4) = yVec.x
    oMat.IndexValue(5) = yVec.y
    oMat.IndexValue(6) = yVec.z
    'Z-axis of plate
    oMat.IndexValue(8) = zVec.x
    oMat.IndexValue(9) = zVec.y
    oMat.IndexValue(10) = zVec.z
    ' we need to offset the plate by half the web thickness
    oPos.Set 0, 0, offsetDist
    Set oPos = oMat.TransformPosition(oPos)
    
    'Origin of plate
    oMat.IndexValue(12) = oPlateOrg.x + oPos.x
    oMat.IndexValue(13) = oPlateOrg.y + oPos.y
    oMat.IndexValue(14) = oPlateOrg.z + oPos.z
    
    'set the matrix on the Baseplate occurrence
    oOcc.Matrix = oMat
    oSmartOccChild.Update ' mark IJSmartOccurrence as modified so that the baseplate generation semantic fires

  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

Public Sub CMMigrateSpliceAsmConnLeftWebPlate(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateSpliceAsmConnLeftWebPlate"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub
'Right Web Plate SO custom methods............................................................................................
Public Sub CMConditionalSpliceAsmConnRightWebPlate(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnRightWebPlate"
  On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim oSectionAttrbs As IJDAttributes
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim spliceType As Long
    Dim secName1 As String, secName2 As String
    Dim rot1 As Double, rot2 As Double
    Dim rotTol As Double
    Dim webPlatePos As Long
  
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    

    If colPorts.count < 2 Then
        Dim iJDObject As iJDObject
        Set iJDObject = oStructConn
        iJDObject.Remove
        Err.Raise E_FAIL
    End If
    
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If

    If (oSupped1Port Is Nothing) Or (oSupped2Port Is Nothing) Then
        Err.Raise SPS_MACRO_WARNING
    End If
    Set oSupped1Part = oSupped1Port.Part
    Set oSupped2Part = oSupped2Port.Part
    
    secName1 = oSupped1Part.CrossSection.SectionName
    secName2 = oSupped2Part.CrossSection.SectionName
    rot1 = oSupped1Part.Rotation.BetaAngle
    rot2 = oSupped2Part.Rotation.BetaAngle
    rotTol = Abs(rot1 - rot2)
    
    If (secName1 <> secName2) Or (rotTol > distTol) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSSpliceAsmConn")
    End If
    
    spliceType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("SpliceWith").Value
    webPlatePos = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("WebPlates").Value

    CheckForUndefinedValueAndRaiseError oSmartOccCAO, spliceType, "StructACSpliceTypes", 22

    If spliceType <> 1 And spliceType <> 3 Then
        bIsNeeded = False
    ElseIf webPlatePos <> 1 And webPlatePos <> 3 Then
        bIsNeeded = False
    End If

  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub

Public Sub CMConstructSpliceAsmConnRightWebPlate(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnRightWebPlate"
    On Error GoTo ErrorHandler
    Dim oPlate As IJStructCustomPlatePart
    Dim oFactory As IJStructCustomPlatePartFactory
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    Dim oRefColl   As IJDReferencesCollection
    Dim oSmartOccCAO      As IJSmartOccurrence
    Dim oSymbolOfCAO                As IJDSymbol
    Dim oMembObjects As IJDMemberObjects
    Dim oDesignParent As IJDesignParent
    Dim PlateThickness As Double
    Dim oAttrbs As IJDAttributes
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    
    'retrieve the inputs of the custom assembly occurrence
    Set oSmartOccCAO = pMemberDescription.CAO
    Set oMembObjects = pMemberDescription.CAO
    Set oDesignParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If

    If oSupped2Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    Set oSuppedPart = oSupped2Port.Part
    
    Set oFactory = New StructCustomPlatePartFactory
    'Create the custom Plate.
    Set oPlate = oFactory.CreateCustomPlatePart(pResourceManager)
    Set oAttrbs = oSmartOccCAO.ItemObject
    On Error Resume Next
        Dim materialProxy As CollectionProxy
        'the interface below is added in V6. It may not exist in older catalog
        'New code and old catalog is unusual. The resume next above takes care of this though
        Set materialProxy = oAttrbs.CollectionOfAttributes("IJUASPSPlateMaterial")
    On Error GoTo ErrorHandler
    
    If Not materialProxy Is Nothing Then
        Dim strMaterial As String, strGrade As String
        Dim pIJAttribsCAO As IJDAttributes
        strMaterial = materialProxy.Item("SPSMaterial").Value
        strGrade = materialProxy.Item("SPSGrade").Value
        'set this as occurrence attribute
        Set pIJAttribsCAO = oSmartOccCAO
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSMaterial").Value = strMaterial
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSGrade").Value = strGrade
        
        SetPlateMaterial oPlate, strMaterial, strGrade
    Else
        'use hard coded material, as done prior to V6
        SetPlateMaterial oPlate, "Steel - Carbon", "A"
    End If
    'get default from the parent
    PlateThickness = oAttrbs.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Thickness2").Value
    'set the default thickness on the child. Returns available thickness if asked thickness is not found
    SetPlateDimensions oPlate, PlateThickness
    'set the found thickness on attrbs
'    oAttrbs.CollectionOfAttributes("IJUASPSPlateThickness").Item("Thickness").Value = PlateThickness
    
    Set oSmartOcc = oPlate
    ' Initialize the SmartOccurrence
    oSmartOcc.RootSelectorClass = "GenericRectPlatePart"
    
    ' Create the reference collection
    Set oRefColl = oSymbolEntitiesFactory.CreateEntity(ReferencesCollection, pResourceManager)

    'Make the plate a child of the assembly connection
    oDesignParent.AddChild oPlate
    
    'create name for the plate
    GenerateNameForPlate oPlate
    
    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSmartOcc, oRefColl
    Set pObj = oPlate 'return the created plate object to the custom assembly
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnRightWebPlate(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnRightWebPlate"
  On Error GoTo ErrorHandler
  Dim pIJAttribsCAO As IJDAttributes
  Dim oSmartOccCAO As IJSmartOccurrence
  Dim oCustomPlate As StructCustomPlatePart
  
  Set oSmartOccCAO = pMemberDesc.CAO
  Set pIJAttribsCAO = oSmartOccCAO
  On Error Resume Next 'suppress the error that VB throws when QI fails
  Set oCustomPlate = pMemberDesc.object
  On Error GoTo ErrorHandler
  
  If Not oCustomPlate Is Nothing Then
    'copy plate category and type from parent to child
      oCustomPlate.CustomPlatePartCategory = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Category2").Value
      oCustomPlate.CustomPlatePartType = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Type2").Value
  End If

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceRightWebPlateSize(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceRightWebPlateSize"
    On Error GoTo ErrorHandler
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim sizeRule As Long
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSuppedPart As ISPSMemberPartPrismatic
    Dim oSectionAttrbs As IJDAttributes
    Dim secDepth As Double, SecWidth As Double
    Dim depth As Double, Width As Double
    Dim oPlate As IJStructPlate
    Dim PlateThickness As Double
    Dim varThickness As Variant
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim gageDist As Double, distFromGage As Double
  
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped1Port Is Nothing Then
        On Error GoTo ErrorHandler
    End If

    Set oSuppedPart = oSupped1Port.Part
    Set oSectionAttrbs = oSuppedPart.CrossSection.Definition
    
    secDepth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
    SecWidth = oSectionAttrbs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Width").Value
    On Error Resume Next ' some sections don't support the interface IStructFlangedBoltGage
    gageDist = oSectionAttrbs.CollectionOfAttributes("IStructFlangedBoltGage").Item("gw").Value
    On Error GoTo ErrorHandler
    
    
    Set pIJAttribsCAO = oSmartOccCAO
    Set oAttrbs = oSmartOccCAO.ItemObject
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSSpliceAsmConn")
    End If
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2")) Then
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2"), oAttrbs.CollectionOfAttributes("IJUASPSAsmConnPlatePart2")
    End If
      

    depth = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Length2").Value
    Width = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Width2").Value
    PlateThickness = pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Thickness2").Value
    sizeRule = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("SizingRule").Value
    distFromGage = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("GageDistance2").Value

    CheckForUndefinedValueAndRaiseError oSmartOccCAO, sizeRule, "StructACSizingRule", 20

    Set oSmartOccChild = pPropertyDescriptions.object
    Set pIJAttribsChild = oSmartOccChild
    If Not pIJAttribsChild Is Nothing Then
        pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Length").Value = depth
        If sizeRule <> 2 Then ' Width is  calculated from gage
            If gageDist > 0 Then ' some sections don't have gage distance specified in the catalog. In that case we use the width
                pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = (secDepth - 2 * gageDist) + 2 * distFromGage
            Else
                pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = Width
            End If
        Else 'user defined.
            pIJAttribsChild.CollectionOfAttributes("IJUASPSPlatePartDim").Item("Width").Value = Width
        
        End If
        Set oPlate = oSmartOccChild
        On Error Resume Next
            Dim materialProxy As CollectionProxy
            'the interface below is added in V6. It may not exist in older catalog
            'New code and old catalog is unusual. The resume next above takes care of this though
            Set materialProxy = pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial")
        On Error GoTo ErrorHandler
        If Not materialProxy Is Nothing Then
            If IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial")) Then
                Dim strMaterial As String, strGrade As String
                
                strMaterial = materialProxy.Item("SPSMaterial").Value
                strGrade = materialProxy.Item("SPSGrade").Value
                SetPlateMaterial oPlate, strMaterial, strGrade
            Else
                Dim oIJDMaterial As IJDMaterial
                Dim oPlateMaterial As IJStructMaterial
                Set oPlateMaterial = oPlate
                Set oIJDMaterial = oPlateMaterial.StructMaterial
                'copy material from  child to parent
                'this can only happen to a connection which was migrated from v5 to v6
                pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSMaterial").Value = oIJDMaterial.MaterialType
                pIJAttribsCAO.CollectionOfAttributes("IJUASPSPlateMaterial").Item("SPSGrade").Value = oIJDMaterial.MaterialGrade
            End If
        End If
        
        SetPlateDimensions oPlate, PlateThickness
        pIJAttribsCAO.CollectionOfAttributes("IJUASPSAsmConnPlatePart2").Item("Thickness2").Value = PlateThickness
    End If
  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
Public Sub CMComputeSpliceRightWebPlateThickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceRightWebPlateThickness"
    On Error GoTo ErrorHandler
    'This is a notification function and is called when the plate thickness is changed.
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub



Public Sub CMComputeSpliceRightWebPlatePosition(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceRightWebPlatePosition"
    On Error GoTo ErrorHandler
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped2Part As ISPSMemberPartPrismatic
    Dim iPort As SPSMemberAxisPortIndex
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim X0 As Double, Y0 As Double, Z0 As Double
    Dim oLine As IJLine
    Dim oAttrb As IJDAttributes
    Dim xVec As New DVector, yVec As New DVector, zVec As New DVector
    Dim clearance#, tWeb#
    Dim oPos As New DPosition, oPlateOrg As New DPosition
    Dim dist As Double, sBtoB As Double
    Dim oMat As New DT4x4
    Dim oOcc As IJDOccurrence
    Dim oSmartOccChild As IJSmartOccurrence
    Dim align As Long
    Dim offsetDist#
    Dim secType As String
    Dim bStartPoint As Boolean
    Dim oEndPos As New DPosition

    
 
    Set oOcc = pObject
    Set oSmartOccChild = pObject
    Set oAttrb = pPropertyDescriptions.CAO
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If

    If oSupped2Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    iPort = oSupped2Port.portIndex
    Set oSupped2Part = oSupped2Port.Part
    
    align = oAttrb.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Symmetry").Value
    clearance = oAttrb.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Clearance").Value
    Set oAttrb = oSupped2Part.CrossSection.Definition

    CheckForUndefinedValueAndRaiseError oStructConn, align, "StructAlignment", 21

    oSupped2Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1
    If iPort = SPSMemberAxisStart Then
        oEndPos.Set X0, Y0, Z0
        bStartPoint = True
    ElseIf iPort = SPSMemberAxisEnd Then
        oEndPos.Set x1, y1, z1
        bStartPoint = False
    End If

    secType = oSupped2Part.CrossSection.Definition.Type

    Select Case secType
        Case "C", "MC", "L"
            Set oLine = GetTangentLineAtCP(oSupped2Part, 4, bStartPoint)
            If oSupped2Part.Rotation.Mirror = False Then
                tWeb = oAttrb.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
                offsetDist = tWeb
            Else
                offsetDist = 0
            End If
        Case "RS", "HSSR"
            If oSupped2Part.Rotation.Mirror = False Then
                Set oLine = GetTangentLineAtCP(oSupped2Part, 6, bStartPoint)
                offsetDist = 0
            Else
                Set oLine = GetTangentLineAtCP(oSupped2Part, 4, bStartPoint)
                offsetDist = 0
            End If
        
        Case "2L"
            Set oLine = GetTangentLineAtCP(oSupped2Part, 5, bStartPoint)
            tWeb = oAttrb.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
            sBtoB = oAttrb.CollectionOfAttributes("IJUA2L").Item("bb").Value
            offsetDist = tWeb + sBtoB / 2

        Case Else '"W", "S", "HP", "M","WT", "MT", "ST"
            Set oLine = GetTangentLineAtCP(oSupped2Part, 5, bStartPoint)
            tWeb = oAttrb.CollectionOfAttributes("IStructFlangedSectionDimensions").Item("tw").Value
            offsetDist = tWeb / 2
    End Select
    
    oLine.GetDirection x1, y1, z1
    xVec.Set x1, y1, z1
    
    If bStartPoint Then
        oLine.GetStartPoint X0, Y0, Z0
    Else
        oLine.GetEndPoint X0, Y0, Z0
        xVec.[Scale] -1 'reverse direction
    End If
    oPos.Set X0, Y0, Z0
    

    If align = 1 Then 'center
        dist = 0 'clearance / 2
    ElseIf align = 2 Then 'left, only supporte1 is cutback
        dist = -clearance / 2
    ElseIf align = 3 Then 'right , only supported2 is cutback
        dist = clearance / 2
    End If
    
    oPos.x = oPos.x + xVec.x * dist
    oPos.y = oPos.y + xVec.y * dist
    oPos.z = oPos.z + xVec.z * dist
    
    Set zVec = GetYVector(oSupped2Part, oEndPos)
    'reverse the direction as the plate is on the right side
    zVec.x = -zVec.x
    zVec.y = -zVec.y
    zVec.z = -zVec.z
    
    Set yVec = zVec.Cross(xVec)
    yVec.length = 1

    oPlateOrg.Set oPos.x, oPos.y, oPos.z
    
    
    oMat.LoadIdentity
    'X-axis of plate
    oMat.IndexValue(0) = xVec.x
    oMat.IndexValue(1) = xVec.y
    oMat.IndexValue(2) = xVec.z
    'Y-axis of plate
    oMat.IndexValue(4) = yVec.x
    oMat.IndexValue(5) = yVec.y
    oMat.IndexValue(6) = yVec.z
    'Z-axis of plate
    oMat.IndexValue(8) = zVec.x
    oMat.IndexValue(9) = zVec.y
    oMat.IndexValue(10) = zVec.z
    ' we need to offset the plate by half the web thickness
    oPos.Set 0, 0, offsetDist
    Set oPos = oMat.TransformPosition(oPos)
    
    'Origin of plate
    oMat.IndexValue(12) = oPlateOrg.x + oPos.x
    oMat.IndexValue(13) = oPlateOrg.y + oPos.y
    oMat.IndexValue(14) = oPlateOrg.z + oPos.z
    
    'set the matrix on the Baseplate occurrence
    oOcc.Matrix = oMat
    oSmartOccChild.Update ' mark IJSmartOccurrence as modified so that the plate generation semantic fires

  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
Public Sub CMMigrateSpliceAsmConnRightWebPlate(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateSpliceAsmConnRightWebPlate"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub


'Cutback Plane custom methods......................................................................................................
Public Sub CMConditionalSpliceAsmConnCutbackPlane(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnCutbackPlane"
  On Error GoTo ErrorHandler
    bIsNeeded = True

  Exit Sub
ErrorHandler:  HandleError MODULE, MT

End Sub

Public Sub CMConstructSpliceAsmConnCutbackPlane(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnCutbackPlane"
    On Error GoTo ErrorHandler
    

    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic
    Dim oPlane As IJPlane
    Dim oGeom3DFactory As New GeometryFactory
    Dim iPort As SPSMemberAxisPortIndex
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim X0 As Double, Y0 As Double, Z0 As Double
    Dim oControlFlags As IJControlFlags
    
    Set oStructConn = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If

    If oSupped1Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    Set oSupped1Part = oSupped1Port.Part
    iPort = oSupped1Port.portIndex
    If iPort = SPSMemberAxisStart Then
        oSupped1Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1
    Else
        oSupped1Part.Axis.EndPoints x1, y1, z1, X0, Y0, Z0
    End If
    Set oPlane = oGeom3DFactory.Planes3d.CreateByPointNormal(pResourceManager, X0, Y0, Z0, x1 - X0, y1 - Y0, z1 - Z0)
     'Hide the plane
    Set oControlFlags = oPlane
    oControlFlags.ControlFlags(&H4) = &H4
    
    Set pObj = oPlane
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnCutbackPlane(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnCutbackPlane"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceAsmConnCutbackPlane(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceAsmConnCutbackPlane"
    On Error GoTo ErrorHandler

    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As ISPSSplitAxisPort
    Dim oSupped2Port As ISPSSplitAxisPort
    Dim oSupped1Part As ISPSMemberPartPrismatic
    Dim oPlane As IJPlane
    Dim oGeom3DFactory As New GeometryFactory
    Dim iPort As SPSMemberAxisPortIndex
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim X0 As Double, Y0 As Double, Z0 As Double
    Dim oMat As IJDT4x4
    Dim oVec As New DVector
    
    Set oStructConn = pPropertyDescriptions.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If

    If oSupped1Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    Set oSupped1Part = oSupped1Port.Part
    
    iPort = oSupped1Port.portIndex
    If iPort = SPSMemberAxisStart Then
        oSupped1Part.Axis.EndPoints X0, Y0, Z0, x1, y1, z1
        oSupped1Part.Rotation.GetTransformAtPosition X0, Y0, Z0, oMat, Nothing
        oVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
    Else
        oSupped1Part.Axis.EndPoints x1, y1, z1, X0, Y0, Z0
        oSupped1Part.Rotation.GetTransformAtPosition X0, Y0, Z0, oMat, Nothing
        oVec.Set oMat.IndexValue(0), oMat.IndexValue(1), oMat.IndexValue(2)
        oVec.[Scale] -1
    End If

    Set oPlane = pObject
    oPlane.SetRootPoint X0, Y0, Z0
    oPlane.SetNormal oVec.x, oVec.y, oVec.z
   
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub

Public Sub CMMigrateSpliceAsmConnCutbackPlane(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
  Const METHOD = "CMMigrateSpliceAsmConnCutbackPlane"
  On Error GoTo ErrorHandler
  
  'MsgBox METHOD
   
  Exit Sub
ErrorHandler:
  HandleError MODULE, METHOD
End Sub


'Cutback1 SO custom methods

Public Sub CMConditionalSpliceAsmConnCutback1(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnCutback1"
  On Error GoTo ErrorHandler
    bIsNeeded = True
  Exit Sub
ErrorHandler:  HandleError MODULE, MT

End Sub

Public Sub CMConstructSpliceAsmConnCutback1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnCutback1"
    On Error GoTo ErrorHandler
  
    Dim iUserType As IJDUserType
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim oSOParent As IJSmartOccurrence
    Dim oSOChild As IJSmartOccurrence
    Dim oDesignParent As IJDesignParent
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As IJPort
    Dim oSupped2Port As IJPort
    Dim oMembObjs As IJDMemberObjects
    Dim oFeatureFactory As New StructFeatureFactory
    
    'MsgBox MT
    
    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Set oMembObjs = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If
    
    If oSupped1Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
 
    Set oSOChild = oFeatureFactory.CreateStructFeature(pResourceManager)
    Set oDesignParent = oSOParent
    oDesignParent.AddChild oSOChild
    'create name
    GenerateNameForFeature oSOChild
    
    Set iUserType = oSOChild
    iUserType.UserType = "CPUASTRUCTPlanarCutbackFeatureOcc"
    oSOChild.RootSelection = "PlanarCutbackFeature"
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefCollChild = oSymFactory.CreateEntity(ReferencesCollection, pResourceManager)

    oRefCollChild.IJDEditJDArgument.SetEntity 1, oSupped1Port, CONST_ISPSSplitAxisEndPort, "RefColl"
    oRefCollChild.IJDEditJDArgument.SetEntity 2, oMembObjs.Item(11), IJPlane, "SPSSuppPlaneToRC_DEST" ' cutback plane

    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSOChild, oRefCollChild
    Set pObj = oSOChild ' return the newly created object
    
  
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnCutback1(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnCutback1"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceAsmConnCutback1(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceAsmConnCutback1"
    On Error GoTo ErrorHandler
    
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim clearance As Double
    Dim align As Long
  
    'MsgBox MT
    
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set pIJAttribsCAO = oSmartOccCAO
    
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn")) Then
    'This is probably executed when the parent and child are just constructed
        Set oAttrbs = oSmartOccCAO.ItemObject
        'copy defaults from parent
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSSpliceAsmConn")
    End If
    
    Set pIJAttribsChild = pPropertyDescriptions.object
    If Not pIJAttribsChild Is Nothing Then
        clearance = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Clearance").Value
        align = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Symmetry").Value

        CheckForUndefinedValueAndRaiseError oSmartOccCAO, align, "StructAlignment", 21

        If align = 1 Then 'center
            clearance = clearance / 2
        ElseIf align = 3 Then 'right , only supported2 is cutback
            clearance = 0
        End If
        pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback").Item("Clearance").Value = clearance
    End If
    
  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
'Cutback2 SO custom methods

Public Sub CMMigrateSpliceAsmConnCutback1(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateSpliceAsmConnCutback1"
    On Error GoTo ErrorHandler
    
    Dim pSmartOccurrence As IJSmartOccurrence
    Dim pReferencesCollection As IJDReferencesCollection
    Dim bIsInputMigrated As Boolean
    Dim oPoint As IJPoint
    Dim ii As Integer, eleCount As Integer
    Dim pObjectCollectionReplacing As IJDObjectCollection
    Dim bIsDeleted As Boolean
    
    Dim oOld As Object
    Dim oReplacing() As Object

    'MsgBox METHOD

    Set pSmartOccurrence = pMemberDesc.object
    Set pReferencesCollection = GetRefCollFromSmartOccurrence(pSmartOccurrence)
    
    GetPositionFromRefColl pReferencesCollection, oPoint

    eleCount = pReferencesCollection.IJDEditJDArgument.GetCount
    ReDim oReplacing(1 To eleCount)

    For ii = 1 To eleCount

        Set oOld = pReferencesCollection.IJDEditJDArgument.GetEntityByIndex(ii)
        
        Call pMigrateHelper.ObjectsReplacing(oOld, pObjectCollectionReplacing, bIsDeleted)
    
        If Not pObjectCollectionReplacing Is Nothing Then
            bIsInputMigrated = True
            SelectReplacingObject pObjectCollectionReplacing, oPoint, oReplacing(ii)
        Else
            Set oReplacing(ii) = oOld
        End If
        
        Set oOld = Nothing
        Set pObjectCollectionReplacing = Nothing
        
    Next ii

    If bIsInputMigrated Then
        Call pReferencesCollection.IJDEditJDArgument.RemoveAll
        pReferencesCollection.IJDEditJDArgument.SetEntity 1, oReplacing(1), CONST_ISPSSplitAxisEndPort, "RefColl"
        pReferencesCollection.IJDEditJDArgument.SetEntity 2, oReplacing(2), IJPlane, "SPSSuppPlaneToRC_DEST" ' cutback plane
    End If

    MigrateMemberObject pMemberDesc.object, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Sub CMConditionalSpliceAsmConnCutback2(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
  Const MT = "CMConditionalSpliceAsmConnCutback2"
  On Error GoTo ErrorHandler
    bIsNeeded = True
  Exit Sub
ErrorHandler:  HandleError MODULE, MT

End Sub

Public Sub CMConstructSpliceAsmConnCutback2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
  Const MT = "CMConstructSpliceAsmConnCutback2"
    On Error GoTo ErrorHandler
    Dim iUserType As IJDUserType
    Dim oRefCollChild   As IJDReferencesCollection
    Dim oSymFactory  As IJDSymbolEntitiesFactory
    Dim oSOParent As IJSmartOccurrence
    Dim oSOChild As IJSmartOccurrence
    Dim oDesignParent As IJDesignParent
    Dim oStructConn As IJAppConnection
    Dim colPorts As IJElements
    Dim oSupped1Port As IJPort
    Dim oSupped2Port As IJPort
    Dim oMembObjs As IJDMemberObjects
    Dim oFeatureFactory As New StructFeatureFactory
  
    Set oSOParent = pMemberDescription.CAO
    Set oStructConn = pMemberDescription.CAO
    Set oMembObjs = pMemberDescription.CAO
    Dim inputobj As IJElements
    oStructConn.enumPorts inputobj
    Dim ISportsVAlid As Variant
    ISportsVAlid = ISPSFACInputHelper_ValidateObjects(inputobj, colPorts)
    If Not colPorts Is Nothing Then
        GetSpliceMiterInputPorts colPorts, oSupped1Port, oSupped2Port
    Else
        Err.Raise E_FAIL
    End If

    If oSupped2Port Is Nothing Then
        Err.Raise SPS_MACRO_WARNING
    End If
    Set oSOChild = oFeatureFactory.CreateStructFeature(pResourceManager)
    Set oDesignParent = oSOParent
    oDesignParent.AddChild oSOChild
    'create name
    GenerateNameForFeature oSOChild

    Set iUserType = oSOChild
    iUserType.UserType = "CPUASTRUCTPlanarCutbackFeatureOcc"
    oSOChild.RootSelection = "PlanarCutbackFeature"
    Set oSymFactory = New DSymbolEntitiesFactory
    Set oRefCollChild = oSymFactory.CreateEntity(ReferencesCollection, pResourceManager)

    oRefCollChild.IJDEditJDArgument.SetEntity 1, oSupped2Port, CONST_ISPSSplitAxisEndPort, "RefColl"
    oRefCollChild.IJDEditJDArgument.SetEntity 2, oMembObjs.Item(11), IJPlane, "SPSSuppPlaneToRC_DEST" ' cutback plane

    ' Connect the SO to its model arguments
    ConnectSmartOccurrence oSOChild, oRefCollChild
    Set pObj = oSOChild ' return the newly created object
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMSetInputSpliceAsmConnCutback2(pMemberDesc As IJDMemberDescription)
  Const MT = "CMSetInputSpliceAsmConnCutback2"
  On Error GoTo ErrorHandler

  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub
Public Sub CMComputeSpliceAsmConnCutback2(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
 Const MT = "CMComputeSpliceAsmConnCutback2"
    On Error GoTo ErrorHandler
    
    Dim pIJAttribsCAO As IJDAttributes, pIJAttribsChild   As IJDAttributes
    Dim oSmartOccCAO As IJSmartOccurrence
    Dim oAttrbs As IJDAttributes
    Dim clearance As Double
    Dim align As Long
  
    Set oSmartOccCAO = pPropertyDescriptions.CAO
    Set pIJAttribsCAO = oSmartOccCAO
    
    If Not IsSOOverridden(pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn")) Then
    'This is probably executed when the parent and child are just constructed
        Set oAttrbs = oSmartOccCAO.ItemObject
        'copy defaults from parent
        CopyValuesToSOFromItem pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn"), oAttrbs.CollectionOfAttributes("IJUASPSSpliceAsmConn")
    End If
    
    Set pIJAttribsChild = pPropertyDescriptions.object
    If Not pIJAttribsChild Is Nothing Then
        clearance = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Clearance").Value
        align = pIJAttribsCAO.CollectionOfAttributes("IJUASPSSpliceAsmConn").Item("Symmetry").Value
        
        CheckForUndefinedValueAndRaiseError oSmartOccCAO, align, "StructAlignment", 21

        If align = 1 Then 'center
            clearance = clearance / 2
        ElseIf align = 2 Then 'left, only supporte1 is cutback
            clearance = 0
        End If
        pIJAttribsChild.CollectionOfAttributes("IJUASPSPlanarCutback").Item("Clearance").Value = clearance
    End If
    
    
  Exit Sub
ErrorHandler:
    HandleError MODULE, MT
    Err.Raise E_FAIL
End Sub
Public Sub CMMigrateSpliceAsmConnCutback2(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
    Const METHOD = "CMMigrateSpliceAsmConnCutback2"
    On Error GoTo ErrorHandler
    
    Dim pSmartOccurrence As IJSmartOccurrence
    Dim pReferencesCollection As IJDReferencesCollection
    Dim bIsInputMigrated As Boolean
    Dim oPoint As IJPoint
    Dim ii As Integer, eleCount As Integer
    Dim pObjectCollectionReplacing As IJDObjectCollection
    Dim bIsDeleted As Boolean
    
    Dim oOld As Object
    Dim oReplacing() As Object

    'MsgBox METHOD

    Set pSmartOccurrence = pMemberDesc.object
    Set pReferencesCollection = GetRefCollFromSmartOccurrence(pSmartOccurrence)
    
    GetPositionFromRefColl pReferencesCollection, oPoint

    eleCount = pReferencesCollection.IJDEditJDArgument.GetCount
    ReDim oReplacing(1 To eleCount)

    For ii = 1 To eleCount

        Set oOld = pReferencesCollection.IJDEditJDArgument.GetEntityByIndex(ii)
        
        Call pMigrateHelper.ObjectsReplacing(oOld, pObjectCollectionReplacing, bIsDeleted)
    
        If Not pObjectCollectionReplacing Is Nothing Then
            bIsInputMigrated = True
            SelectReplacingObject pObjectCollectionReplacing, oPoint, oReplacing(ii)
        Else
            Set oReplacing(ii) = oOld
        End If
        
        Set oOld = Nothing
        Set pObjectCollectionReplacing = Nothing
        
    Next ii
     
    If bIsInputMigrated Then
        Call pReferencesCollection.IJDEditJDArgument.RemoveAll
        pReferencesCollection.IJDEditJDArgument.SetEntity 1, oReplacing(1), CONST_ISPSSplitAxisEndPort, "RefColl"
        pReferencesCollection.IJDEditJDArgument.SetEntity 2, oReplacing(2), IJPlane, "SPSSuppPlaneToRC_DEST" ' cutback plane
    End If

    MigrateMemberObject pMemberDesc.object, pMigrateHelper

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub


'
' The following methods are generic for all the Custom assembly
'
'
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
  ' This method is in charge of the creation of the symbol definition object
  ' You can keep the current design unchanged for basic VB symbol definition.
   Const MT = "IJDUserSymbolServices_InstanciateDefinition"
  On Error GoTo ErrorHandler
 
  
  Dim pDefinition As IJDSymbolDefinition
  Dim pFact As IJCAFactory
  Set pFact = New CAFactory
  Set pDefinition = pFact.CreateCAD(ActiveConnection)
  
  ' Set definition progId and codebase
  pDefinition.ProgId = m_ItemProgId
  pDefinition.CodeBase = CodeBase
    
  ' Initialize the definition
  IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
  pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
  
  ' Persistence behavior
  pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
  pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
    
  'returned symbol definition
  Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
  Exit Function
ErrorHandler:  HandleError MODULE, MT
End Function
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  ' Name should be unique
  IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
 ' Obsolete method.
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
 ' Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
 IJDUserSymbolServices_EditOccurence = False
End Function





Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)

End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = NonParticipant
End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SPSMembers.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String

Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler
    Dim pAttrs As IJDAttributes
    Dim xSectionDepth As Double
    Dim pStructConn As IJAppConnection
    Dim pMemberPart As ISPSMemberPartPrismatic
    
    
    IJUserAttributeMgmt_OnAttributeChange = vbNullString
    
    

    ' check for the clearance to be atmost the depth of the member cross section
    If (pAttrToChange.AttrName = "Clearance") Then
        Set pStructConn = pIJDAttrs
        
        If Not pStructConn Is Nothing Then
            Dim depth As Double
            Dim SecWidth As Double
            Dim colPorts As IJElements
            Dim oCrossAttrs As IJDAttributes
            Dim pPort As IJPort
            
            pStructConn.enumPorts colPorts
            
            ' get the cross section depth from the member part
            If colPorts.count > 0 Then
                Set pPort = colPorts.Item(1)
                If Not pPort Is Nothing Then
                    Set pMemberPart = pPort.Connectable
                    If Not pMemberPart Is Nothing Then
                        Set oCrossAttrs = pMemberPart.CrossSection.Definition
                        If Not oCrossAttrs Is Nothing Then
                            depth = oCrossAttrs.CollectionOfAttributes("ISTRUCTCrossSectionDimensions").Item("Depth").Value
                        End If
                    End If
                End If
            End If
        End If
        If (varNewAttrValue > depth) Or (varNewAttrValue < 0) Then
            IJUserAttributeMgmt_OnAttributeChange = "Clearance is too large"
            Exit Function
        End If
    End If
    

    'make sure that if the material is edited, the new material has plate thicknesses
    'defined in catalog
      

    If (pAttrToChange.InterfaceName = "IJUASPSPlateMaterial") Then
    
        Dim strGrade As String, strMaterial As String
        Dim strOldGrade As String, strOldMaterial As String
        Dim oRefDataQuery As IJDStructServices
        Dim oMatlObj As IJDMaterial
        Dim oCatalogPOM As IJDPOM
        Dim matlThickCol As IJDCollection
        Dim pAttrColl As Collection
        Dim jj As Integer
        Dim pAttrDescr As IJAttributeDescriptor
        
        Set oRefDataQuery = New StructServices
        Set oCatalogPOM = GetCatalogResourceManager()
        Set pAttrColl = CollAllDisplayedValues

        For jj = 1 To pAttrColl.count
            Set pAttrDescr = pAttrColl.Item(jj)
            If pAttrDescr.InterfaceName = "IJUASPSPlateMaterial" Then

               If pAttrDescr.AttrName = "SPSGrade" Then

                   strGrade = pAttrDescr.AttrValue
               ElseIf pAttrDescr.AttrName = "SPSMaterial" Then
                   strMaterial = pAttrDescr.AttrValue
               End If
            End If
        Next
        
        If (strMaterial <> vbNullString) And (strGrade <> vbNullString) Then
            
            Set oMatlObj = oRefDataQuery.GetMaterialFromGradeAndType(oCatalogPOM, strMaterial, strGrade)
            If Not oMatlObj Is Nothing Then
    
                On Error Resume Next
                Set matlThickCol = oRefDataQuery.GetPlateDimensions(oCatalogPOM, strMaterial, strGrade)
                On Error GoTo ErrorHandler
                If Not matlThickCol Is Nothing Then
                    If matlThickCol.Size = 0 Then
                        IJUserAttributeMgmt_OnAttributeChange = "No plate thickness  available for the selected material and grade"
                    End If
                Else
                    IJUserAttributeMgmt_OnAttributeChange = "No plate thickness  available for the selected material and grade"
                End If
            Else
                IJUserAttributeMgmt_OnAttributeChange = "Selected material doesn't exist in catalog"
            End If
        End If

    End If
    
            

    Exit Function
ErrorHandler:      HandleError MODULE, METHOD

End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    
Const METHOD = "IJUserAttributeMgmt_OnPreCommit"

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SPSMembers.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function
Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SPSMembers.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String

End Function

Private Function IJUserAttributeMgmtParent_OnAttributeChange(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
'
End Function

Private Function IJUserAttributeMgmtParent_OnPreCommit(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmtParent_OnPreLoad(ByVal pParent As Object, ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    'get the parent
    Dim oParent As IJDMemberObjects
    Dim oAttribMgmntChild As IJUserAttributeMgmtParent
    Dim oSmartItem As IJSmartItem
    Dim ii As Long, jj As Long
    Dim oChild As Object
    
    Set oParent = pParent
    
    If Not oParent Is Nothing Then
        For ii = 1 To 10 'check if the object passed in is a plate member
            Set oChild = oParent.ItemByDispid(ii)
            If oChild Is pIJDAttrs Then
   

                Dim pAttrColl As Collection
                Set pAttrColl = CollAllDisplayedValues
           
                Dim pAttrDescr As IJAttributeDescriptor
                For jj = 1 To pAttrColl.count
                    Set pAttrDescr = pAttrColl.Item(jj)
                    If pAttrDescr.InterfaceName = "IJUASPSPlatePartDim" Then
                        'controlling  child's property
                        'make property read only in plate's GUI
                        pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                    
                    ElseIf pAttrDescr.InterfaceName = "IJDMaterial" Then
                        'controlling  child's property
                        'make property read only in plate's GUI
                        pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
 
 
                    ElseIf pAttrDescr.AttrName = "Thickness" Then
                        'disable edits of thickness from child's GUI
                        pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                    End If
                Next
                Exit For 'found a match so exit
            End If
        Next
    End If
End Function

Private Property Get ISPSFACInputHelper_UserAttributeMgmt() As SP3DStructInterfaces.IJUserAttributeMgmt

End Property

Private Property Get ISPSFACInputHelper_ValidateObjects(ByVal inputObjs As SP3DStructInterfaces.IJElements, relationObjs As SP3DStructInterfaces.IJElements) As SP3DStructInterfaces.SPSFACInputHelperStatus
 Const MT = "ISPSFACInputHelper_ValidateObjects"
    On Error GoTo ErrorHandler
    

    Dim oInputObj1 As Object, oInputObj2 As Object
    Dim portIdx As SPSMemberAxisPortIndex
    Dim oSupped1Part As ISPSMemberPartPrismatic, oSupped2Part As ISPSMemberPartPrismatic
    Dim oSupped1Port As ISPSSplitAxisPort, oSupped2Port As ISPSSplitAxisPort
    Dim oPoint1 As IJPoint, oPoint2 As IJPoint, oPoint3 As IJPoint
    Dim dist1 As Double, dist2 As Double
    
    Dim count As Integer
    Dim portCol As IJElements


    ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_Ok
    Set portCol = New JObjectCollection
    'filter out ports to portCol
    For count = 1 To inputObjs.count
        If TypeOf inputObjs.Item(count) Is IJPort Then
            portCol.Add inputObjs.Item(count)
        End If
    Next count

    '2 ports are expected
    If portCol.count = 2 Then
        Set oInputObj1 = portCol.Item(1)
        Set oInputObj2 = portCol.Item(2)
        If TypeOf oInputObj1 Is ISPSSplitAxisEndPort Then
            Set oSupped1Port = oInputObj1
            If TypeOf oInputObj2 Is ISPSSplitAxisPort Then
                Set oSupped2Port = oInputObj2
                If TypeOf oSupped2Port Is ISPSSplitAxisAlongPort Then
                    'get the end port
                    Set oSupped2Part = oSupped2Port.Part
                    portIdx = oSupped1Port.portIndex
                    Set oSupped1Part = oSupped1Port.Part
                    Set oPoint1 = oSupped1Part.PointAtEnd(portIdx)
                    Set oPoint2 = oSupped2Part.PointAtEnd(SPSMemberAxisStart)
                    Set oPoint3 = oSupped2Part.PointAtEnd(SPSMemberAxisEnd)
                    dist1 = oPoint1.DistFromPt(oPoint2)
                    dist2 = oPoint1.DistFromPt(oPoint3)
                    If dist1 < dist2 Then
                        Set oSupped2Port = oSupped2Part.AxisPort(SPSMemberAxisStart)
                    Else
                        Set oSupped2Port = oSupped2Part.AxisPort(SPSMemberAxisEnd)
                    End If
                End If
                If relationObjs Is Nothing Then
                    Set relationObjs = New JObjectCollection
                End If
                relationObjs.Clear
                relationObjs.Add oSupped1Port
                relationObjs.Add oSupped2Port
            Else
                ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
            End If
        Else
            ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_InvalidTypeOfObject
        End If
    Else
        ISPSFACInputHelper_ValidateObjects = SPSFACInputHelper_BadNumberOfObjects
    End If
    Exit Property
ErrorHandler:      HandleError MODULE, MT
End Property

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub
