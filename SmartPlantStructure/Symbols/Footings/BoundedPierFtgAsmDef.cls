VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoundedPierFtgAsmDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2006, Intergraph Corporation. All rights reserved.
'
'File
'    BoundedPierFtgAsmDef.cls
'
'Author
'       28th January 2003        AP
'
'Description
'
'Notes
'
'History:
'   09-May-06   JMS     DI#97751 - All lines are constructed through the
'                       geometry factory instead of being "new"ed
'   06-Jun-06   SPSDEV  Added CMMigrate method stubs
'   24-Jun-06   JMS     DI#60069 - Changes to allow editing of the weight and CG values
'                       changed call to SetWCG to call new interface to put weight and CG
'                       values since SetWCG is reservered for setting user defined values
'                       when the values here are the computed values
'  07-Jul-06    JMS     TR#101063 - Tolerate the nonexistence of IJWCGValueOrigin interface
'  08-Aug-06    AS      TR# 99969 Added split migration logic
'  06-Mar-07    RS      CR#41094 - Changes for placing a footing in space and dis/reconnect to members
'
'*******************************************************************
Option Explicit

Private Const MODULE = "BoundedPierFtgAsmDef"

Private Const INTERFACE_WCGValueOrigin As String = "IJWCGValueOrigin"
Private Const PROPERTY_DryWCGOrigin As String = "DryWCGOrigin"
Private Enum enumWeightCGDerivation
    WEIGHTCG_Computed = 2
    WEIGHTCG_UserDefined = 4
End Enum

Const CONST_ItemProgId As String = "SPSFootingMacros.BoundedPierFtgAsmDef"
Private Const strSourceFile As String = "BoundedPierFtgAsmDef.cls"
Private m_strErrorDescr As String
Private m_oLocalizer As IJLocalizer
Implements IJDUserSymbolServices
Implements ISPSFootingDefServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck
Implements ISPSFoundationInputHelper 'TR#71850
Implements ISPSTransformHelper

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
     IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
     IJDUserSymbolServices_GetDefinitionName = CONST_ItemProgId
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IMSSymbolEntities.IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler
     pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
     pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA

     ' Define the inputs -
     Dim pIH As IJDInputsHelper
     Set pIH = New InputHelper
     pIH.definition = pDefinition
     
     pIH.SetInput "SupportedMember"
       
     ' Aggregator Type
     Dim pAD As IJDAggregatorDescription
     Set pAD = pDefinition
     pAD.AggregatorClsid = "{439518F7-6759-4193-B380-DB95D12412A9}" 'CSPSFooting
     pAD.UserTypeClsid = "{043CDA28-9090-4a58-95D7-F25B37530C48}" 'CPUAPierFootingAsm
     pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAsm"
     pAD.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructAsm"
     pAD.SetCMSetInputs -1, -1
     pAD.SetCMRemoveInputs -1, -1
     pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAggregator"
     Set pAD = Nothing
     
     ' tr 74800
     Dim pCADefinition As IJCADefinition
     Set pCADefinition = pDefinition
     Let pCADefinition.CopyBackwardFlag = igCOPY_BACKWARD_TRIM
     Set pCADefinition = Nothing
     
     ' Aggregator property
     Dim pAPDs As IJDPropertyDescriptions
     Set pAPDs = pDefinition
     pAPDs.RemoveAll ' Remove all the previous property descriptions

     pAPDs.AddProperty "IJDAttributes", 1, IJDAttributes, "CMEvaluateCAO", imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IJWeightCG", 2, IJWeightCG, "CMEvaluateCAOWCG", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE
     pAPDs.AddProperty "IJStructElevationDatum", 3, IJStructElevationDatum, "CMEvaluateElevationDatum", imsCOOKIE_ID_USS_LIB
     
     Set pAPDs = Nothing
               
     ' Define the members
     Dim pMemberDescriptions As IJDMemberDescriptions
     Dim pMemberDescription As IJDMemberDescription
     Dim pPropertyDescriptions As IJDPropertyDescriptions
     Set pMemberDescriptions = pDefinition
     
     ' Remove all the previous member descriptions
     pMemberDescriptions.RemoveAll
      
     Set pMemberDescription = pMemberDescriptions.AddMember("Grout", 1, "CMConstructGrout", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalGrout"
     pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputGrout"
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructGrout"
     pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseGrout"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
      
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty GROUT_PROPERTIES, 1, IJDAttributes, "CMEvaluateGrout", imsCOOKIE_ID_USS_LIB
     pPropertyDescriptions.AddProperty GROUT_MATERIALPROP, 2, IJStructMaterial, "CMEvaluateGroutMaterial", imsCOOKIE_ID_USS_LIB
     pPropertyDescriptions.AddProperty GROUT_GEOMETRY, 3, IJGeometry, "CMEvaluateGroutGeometry", imsCOOKIE_ID_USS_LIB
          
     Set pMemberDescription = pMemberDescriptions.AddMember("Pier", 2, "CMConstructPier", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputPier"
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructPier"
     pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleasePier"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty PIER_PROPERTIES, 1, IJDAttributes, "CMEvaluatePier", imsCOOKIE_ID_USS_LIB
     pPropertyDescriptions.AddProperty PIER_MATERIAL_PROP, 2, IJStructMaterial, "CMEvaluatePierMaterial", imsCOOKIE_ID_USS_LIB
     pPropertyDescriptions.AddProperty PIER_GEOMETRY, 3, IJGeometry, "CMEvaluatePierGeometry", imsCOOKIE_ID_USS_LIB

     Set pMemberDescriptions = Nothing
     Set pMemberDescription = Nothing
     Set pPropertyDescriptions = Nothing
     
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal pResourceMgr As Object) As Object
' This method is in charge of the creation of the symbol definition object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler
     
     Dim pDefinition As IJDSymbolDefinition
     Dim pFact As IJCAFactory
     Set pFact = New CAFactory
     Set pDefinition = pFact.CreateCAD(pResourceMgr)
     
     ' Set definition progId and codebase
     pDefinition.ProgId = CONST_ItemProgId
     pDefinition.CodeBase = CodeBase
     
     ' Initialize the definition
     IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
     pDefinition.name = IJDUserSymbolServices_GetDefinitionName(defParams)
     
     ' Persistence behavior
     pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
     pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
     
     'returned symbol definition
     Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
Exit Function
ErrorHandler:      HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub
Public Sub CMConstructGrout(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructGrout"
On Error GoTo ErrorHandler
  
     Dim oGroutComp As ISPSFootingComponent
     Dim GroutComp As String
     GroutComp = GetCAODefAttribute(pMemberDescription, PIER_FOOTING_ASM, GROUT_COMPONENT)
     Call CreateComponent(GroutComp, pResourceManager, oGroutComp, pMemberDescription, False)
     
     Set pObj = oGroutComp
         
Exit Sub
ErrorHandler:       HandleError MODULE, METHOD
End Sub
Public Sub CMConditionalGrout(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
Const METHOD = "CMConditionalGrout"
On Error GoTo ErrorHandler

    'TR#67111 - Validate pier ht first before proceeding
    bIsNeeded = False
    If pMemberDescription Is Nothing Then
        Exit Sub
    ElseIf pMemberDescription.CAO Is Nothing Then
        Exit Sub
    End If
    
    'first check whether input is valid or not
    ValidatePierHt pMemberDescription


    'TR#67111  ''''''''''''''''''''''''''''''
    
    Dim OccAttribs As IJDAttributes, oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDescription.CAO
    Set OccAttribs = oSmartOcc
    Set oAttribs = oSmartOcc.ItemObject
    If Not IsSOOverridden(OccAttribs.CollectionOfAttributes(PIER_FOOTING_ASM)) Then
        CopyValuesToSOFromItem OccAttribs.CollectionOfAttributes(PIER_FOOTING_ASM), oAttribs.CollectionOfAttributes(PIER_FOOTING_ASM)
    End If

    Dim oAttr As IJDAttributes
    Set oAttr = oSmartOcc
    Dim Needed As Boolean

    If IsEmpty(oAttr.CollectionOfAttributes(PIER_FOOTING_ASM).Item(WITH_GROUT_PAD).Value) Then
        bIsNeeded = True
    Else
        Needed = oAttr.CollectionOfAttributes(PIER_FOOTING_ASM).Item(WITH_GROUT_PAD).Value
        bIsNeeded = Needed
    End If
     
Exit Sub
ErrorHandler:
  ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    End If
End Sub

Public Sub CMSetInputGrout(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputGrout"
On Error GoTo ErrorHandler
    
    Dim GroutComp As String
    GroutComp = GetCAODefAttribute(pMemberDesc, PIER_FOOTING_ASM, GROUT_COMPONENT)
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.object
    oSmartOcc.RootSelection = GroutComp
 
Exit Sub
ErrorHandler:       HandleError MODULE, METHOD
End Sub
Public Sub CMFinalConstructGrout(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructGrout"
On Error GoTo ErrorHandler

     Call AddSystemAndNameRule(pMemberDesc)
     
 
Exit Sub
ErrorHandler:       HandleError MODULE, METHOD
End Sub
Public Sub CMReleaseGrout(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMReleaseGrout"
On Error GoTo ErrorHandler


Exit Sub
ErrorHandler:       HandleError MODULE, METHOD
End Sub
Public Sub CMEvaluateGrout(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateGrout"
On Error GoTo ErrorHandler
    Dim oSymbol0fCAO As IMSSymbolEntities.IJDSymbol
    Dim oReferencesCollection0fCAO  As IMSSymbolEntities.IJDReferencesCollection
    Dim oMembSys As ISPSMemberSystem
    Dim oMembPart As ISPSMemberPartPrismatic
    
    Set oReferencesCollection0fCAO = GetRefCollection(pPropertyDescriptions.CAO)
    
    Dim x As Double, y As Double, z As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim CP As Long, dMembAng As Double

    Dim obj As Object

    If oReferencesCollection0fCAO.IJDEditJDArgument.GetCount > 0 Then
        Set obj = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1)
    End If
    
    If Not obj Is Nothing Then
        If TypeOf obj Is ISPSMemberSystem Then
            Set oMembSys = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1)
        End If
    End If

    Dim oFooting As ISPSFooting
    Set oFooting = pPropertyDescriptions.CAO
    
    Dim Matrix As IJDT4x4
    Dim oFootingMatrix As IJDT4x4
    Set Matrix = New DT4x4
    Matrix.LoadIdentity

    dMembAng = 0
    If oMembSys Is Nothing Then
        oFooting.GetPosition x, y, z
    Else
        GetFootingPositionFromMember oMembSys, x, y, z, dMembAng
        Set oMembPart = GetBottomMemberPart(oMembSys)
        oFooting.SetPosition x, y, z
    End If
     
    Dim pSymbol As IJDSymbol
    Set pSymbol = pObject
    Dim pOcc As IJDOccurrence
    On Error Resume Next
    Set pOcc = pSymbol
    If pOcc Is Nothing Then
         Exit Sub
    End If
    
    
    Matrix.IndexValue(12) = x
    Matrix.IndexValue(13) = y
    Matrix.IndexValue(14) = z

    Dim Vec As DVector
    Set Vec = New DVector
    Vec.Set 0, 0, 1
    Dim GroutRotationAngle As Double
    Dim GroutOrientation As Long
    Dim GroutSizingRule As Long
    Dim oGroutAttribs As IJDAttributes
    Set oGroutAttribs = pObject
    GroutRotationAngle = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_ROTATION_ANGLE).Value
    GroutOrientation = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_ORIENTATION).Value
    GroutSizingRule = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_SIZE_RULE).Value
    
    On Error GoTo ErrorHandler
    CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, GroutSizingRule, FOOTING_COMP_SIZE_RULE, 122
    CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, GroutOrientation, STRUCT_COORD_SYS_REF, 123
    On Error Resume Next
    
    If GroutOrientation = 1 Then '1 is Global
       Matrix.Rotate GroutRotationAngle, Vec
    ElseIf GroutOrientation = 2 Then '2 is Local to the SPSMemberPart
       GroutRotationAngle = GroutRotationAngle + dMembAng
       Matrix.Rotate GroutRotationAngle, Vec
    End If
           
    If oMembSys Is Nothing Then ' we have no member
        ' can get info from member to change grout so exit
        pOcc.Matrix = Matrix
        Exit Sub
    Else
        If GroutSizingRule = 1 Or GroutSizingRule = 2 Then
           Dim sRefStd As String
           Dim sSecName As String
           Dim Depth As Double
           Dim SecWidth As Double
           Dim GroutEdgeClearance As Double
           If Not oMembPart Is Nothing Then
               sRefStd = oMembPart.CrossSection.SectionStandard
               sSecName = oMembPart.CrossSection.SectionName
           End If
           Depth = GetCSAttribData(sSecName, sRefStd, CROSSSSECTION_DIMENSIONS, CROSSSECTION_DEPTH)
           SecWidth = GetCSAttribData(sSecName, sRefStd, CROSSSSECTION_DIMENSIONS, CROSSSECTION_WIDTH)
           GroutEdgeClearance = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_EDGE_CLEARANCE).Value
           
           'TR# 66868- Grout length & width calcualted considering orientation angle of section & grout
           Dim GroutShape As Long
           GroutShape = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_SHAPE).Value
           On Error GoTo ErrorHandler
           CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, GroutShape, PRISMATIC_FOOTING_SHAPES, 121
           On Error Resume Next
           
           Dim tmpAng As Double
           If GroutShape = 3 Then GroutRotationAngle = 0#
           tmpAng = Abs(GroutRotationAngle - dMembAng)
           If tmpAng > PI / 2 Then tmpAng = Abs(PI - tmpAng)
           
           Dim secDepthComp As Double
           Dim SecWidthComp As Double
           
           secDepthComp = Depth * Cos(tmpAng) + SecWidth * Sin(tmpAng)
           SecWidthComp = SecWidth * Cos(tmpAng) + Depth * Sin(tmpAng)
           
           If GroutShape = 2 Then
               oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value = _
                               SecWidthComp + (GroutEdgeClearance * 2)
               oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value = _
                                       secDepthComp + (GroutEdgeClearance * 2)
           ElseIf GroutShape = 3 Then
        '            If secDepthComp > SecWidthComp Then
        '                oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value = secDepthComp + (GroutEdgeClearance * 2)
        '                oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value = secDepthComp + (GroutEdgeClearance * 2)
        '            Else
        '                oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value = SecWidthComp + (GroutEdgeClearance * 2)
        '                oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value = SecWidthComp + (GroutEdgeClearance * 2)
        '            End If
               
               'TR#72794- if grout is circular then take diagonal + clearance=diameter. i.e. width & length
               oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value = Sqr(secDepthComp * secDepthComp + SecWidthComp * SecWidthComp) + (GroutEdgeClearance * 2)
               oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value = Sqr(secDepthComp * secDepthComp + SecWidthComp * SecWidthComp) + (GroutEdgeClearance * 2)
           End If
        End If
        
        pOcc.Matrix = Matrix
    End If

Exit Sub
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    End If

End Sub
Public Sub CMEvaluateGroutMaterial(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateGroutMaterial"
On Error GoTo ErrHandler
    Dim MemberObj As IJDMemberObjects
    Dim oAttrs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Dim iMaterial As IJDMaterial
    Set MemberObj = pPropertyDescriptions.CAO
    Set oSmartOcc = MemberObj.ItemByDispid(1)
    Set oAttrs = oSmartOcc
    Dim Material As String
    Dim oStructMatl As IJStructMaterial
    Dim Grade As String
    Material = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_MATERIAL).Value
    Grade = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_GRADE).Value
    If Not Material = vbNullString And Not Grade = vbNullString Then
        Set iMaterial = GetMaterialObject(Material, Grade)
        Set oStructMatl = oSmartOcc
        oStructMatl.StructMaterial = iMaterial
    End If
                    
Exit Sub
ErrHandler:      HandleError MODULE, METHOD
End Sub
Public Sub CMEvaluateGroutGeometry(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateGroutGeometry"
On Error GoTo ErrHandler
    
    Call CMEvaluateGrout(pPropertyDescriptions, pObject)

Exit Sub
ErrHandler:      HandleError MODULE, METHOD
End Sub


Public Sub CMConstructPier(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructPier"
On Error GoTo ErrorHandler
 
     Dim oPierComp As ISPSFootingComponent
     Dim PierComp As String
     PierComp = GetCAODefAttribute(pMemberDescription, PIER_FOOTING_ASM, PIER_COMPONENT)
     Call CreateComponent(PierComp, pResourceManager, oPierComp, pMemberDescription, True)
         
     Set pObj = oPierComp
     
Exit Sub
ErrorHandler:       HandleError MODULE, METHOD
End Sub
Public Sub CMSetInputPier(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputPier"
On Error GoTo ErrorHandler
 
    Dim PierComp As String
    PierComp = GetCAODefAttribute(pMemberDesc, PIER_FOOTING_ASM, PIER_COMPONENT)
    
    Dim oAssocRelation As IJDAssocRelation
    Dim oTOColloection As IJDTargetObjectCol
    Dim oAssocHelper As IJStructAssocCompute
    Dim lAssocFlags As Long
    
    Set oAssocRelation = pMemberDesc.CAO
    Set oTOColloection = oAssocRelation.CollectionRelations("IJDObject", "toToDoRecord")
    
    If oTOColloection.count > 0 Then
        Exit Sub
    End If

    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.object
    oSmartOcc.RootSelection = PierComp

    Dim NewRC As Boolean
    Dim obj As Object
    Dim pIRCAsm As IJDReferencesCollection
    Set pIRCAsm = GetRefCollection(pMemberDesc.CAO)
    Dim oReferencesCollection As IJDReferencesCollection
    
    NewRC = False

    If Not pIRCAsm Is Nothing Then
        Dim oPlane As IJPlane
        Set obj = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)
        Set oPlane = obj
        If oPlane Is Nothing Then
            Set obj = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(1)
        End If
       
        On Error Resume Next
        Set oReferencesCollection = GetRefCollection(oSmartOcc)
        On Error GoTo ErrorHandler
        
        If Not oReferencesCollection Is Nothing Then
            Dim j As Integer ' must remove old ref coll on ref coll if they exist (when moving from comb to single)
            For j = 1 To oReferencesCollection.IJDEditJDArgument.GetCount
                If TypeOf oReferencesCollection.IJDEditJDArgument.GetEntityByIndex(j) Is IJDReferencesCollection Then
                    Dim oObject As ijdObject
                    Set oObject = oReferencesCollection.IJDEditJDArgument.GetEntityByIndex(j)
                    If Not oObject Is Nothing Then
                        oObject.Remove
                    End If
                End If
            Next j
            oReferencesCollection.IJDEditJDArgument.RemoveAll
        Else
            'create one RC
            Dim oSymbolEntitiesFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
            Set oReferencesCollection = oSymbolEntitiesFactory.CreateEntity(referencesCollection, GetResourceMgr())
            NewRC = True
        End If
        
        If Not obj Is Nothing Then
            oReferencesCollection.IJDEditJDArgument.SetEntity 1, obj, IJPlane, "SPSSuppPlaneToRC_DEST"
            If NewRC Then
                ConnectSmartOccurrence oSmartOcc, oReferencesCollection
            End If
        End If
        
    End If
    
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub
Public Sub CMFinalConstructPier(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructPier"
On Error GoTo ErrorHandler
  
     Call AddSystemAndNameRule(pMemberDesc)
     
Exit Sub
ErrorHandler:       HandleError MODULE, METHOD
End Sub
Public Sub CMReleasePier(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMReleasePier"
On Error GoTo ErrorHandler


Exit Sub
ErrorHandler:       HandleError MODULE, METHOD
End Sub
Public Sub CMEvaluatePier(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluatePier"
On Error GoTo ErrorHandler

    Dim MemberObj As IJDMemberObjects
    Dim oGroutAttribs As IJDAttributes
    Dim oReferencesCollection0fCAO  As IMSSymbolEntities.IJDReferencesCollection
    Dim oMembSys As ISPSMemberSystem
    Dim oMembPart As ISPSMemberPartPrismatic
    Dim oSmartOcc As IJSmartOccurrence
    Set MemberObj = pPropertyDescriptions.CAO
    Dim GroutHt As Double, bIsNeeded As Boolean
    Dim oAttr As IJDAttributes
    Set oAttr = MemberObj
    bIsNeeded = oAttr.CollectionOfAttributes(PIER_FOOTING_ASM).Item(WITH_GROUT_PAD).Value
    Dim oObject1 As Object
    Dim oObject2 As Object
    
    If bIsNeeded Then
        Set oSmartOcc = MemberObj.ItemByDispid(1)
        Set oGroutAttribs = oSmartOcc
    '    NOTE: Need to check for existence of the GroutPad before using it height, if not get
    '    Bottom of column and mate to it for now...
        GroutHt = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_HEIGHT).Value
    End If


    Set oReferencesCollection0fCAO = GetRefCollection(pPropertyDescriptions.CAO)
    
    Dim bByPlaneOrDatum As Boolean
    bByPlaneOrDatum = False
     
    Dim RefCollectionCount As Long
    RefCollectionCount = oReferencesCollection0fCAO.IJDEditJDArgument.GetCount
     
    If RefCollectionCount > 0 Then
        
        Dim oPlane As IJPlane
        
        'the plane may be at index 1 or 2
        
        'it is at index 1 for a footing placed with point
        
        'it is at index 2 for a footing placed with member system.It is at index 2 even when the member system is deleted
        
        
        
        Set oObject1 = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1)
        If Not oObject1 Is Nothing Then
            If TypeOf oObject1 Is IJPlane Then
                Set oPlane = oObject1
            ElseIf TypeOf oObject1 Is ISPSMemberSystem Then ' if there is membersystem it is always at index 1
                Set oMembSys = oObject1
            End If
        End If
        
        'Following call doesn't error when there is nothing at index 2 (there is nothing at index 2 for footing placed by point)
        Set oObject2 = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(2)
        If Not oObject2 Is Nothing Then
            If TypeOf oObject2 Is IJPlane Then
                Set oPlane = oObject2
            End If
        End If
       
    End If

    Dim Normx As Double, Normy As Double, Normz As Double
    Dim Rtx As Double, Rty As Double, Rtz As Double
    Dim TotHt As Double
    
    If Not oPlane Is Nothing Then
        bByPlaneOrDatum = True
        oPlane.GetNormal Normx, Normy, Normz
        oPlane.GetRootPoint Rtx, Rty, Rtz
    Else
    
        Dim bUseDatum As Boolean
        Dim oCollProxy As IJDAttributesCol
        
        On Error Resume Next 'this interface was added in a later release so may not exist
        'in a migrated model
        Set oCollProxy = oAttr.CollectionOfAttributes(STRUCT_ELEVATION_DATUM)
        On Error GoTo ErrorHandler
            
        If Not oCollProxy Is Nothing Then
            bUseDatum = oCollProxy.Item(USE_ELEVATION_DATUM).Value
            If bUseDatum = True Then
                Rtz = oCollProxy.Item(BOTTOM_ELEVATION).Value
                Rtx = 0
                Rty = 0
                Normx = 0
                Normy = 0
                Normz = 1
                bByPlaneOrDatum = True
            End If
        End If
        If bUseDatum = False Then
            SPSToDoErrorNotify FootingToDoMsgCodelist, TDL_FTGMACROS_MISSING_SUPPORTING_PLANE, pPropertyDescriptions.CAO, Nothing
        End If
    
    End If
        

    
    
    Dim x As Double, y As Double, z As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim CP As Long, dMembAng As Double
    
    Dim oFooting As ISPSFooting
    Set oFooting = pPropertyDescriptions.CAO
    
    Dim Matrix As IJDT4x4
    Dim oFootingMatrix As IJDT4x4
    Set Matrix = New DT4x4
    Matrix.LoadIdentity

    If oMembSys Is Nothing Then
        oFooting.GetPosition x, y, z
    Else
        GetFootingPositionFromMember oMembSys, x, y, z, dMembAng
        Set oMembPart = GetBottomMemberPart(oMembSys)
        oFooting.SetPosition x, y, z
    End If

    
    Dim pSymbol As IJDSymbol
    Set pSymbol = pObject
    Dim pOcc As IJDOccurrence
    On Error Resume Next
    Set pOcc = pSymbol
    If pOcc Is Nothing Then
         Exit Sub
    End If
        
    Matrix.IndexValue(12) = x
    Matrix.IndexValue(13) = y
    Matrix.IndexValue(14) = z - GroutHt
    
    Dim oPierAttribs As IJDAttributes
    Set oPierAttribs = pObject
    
    Dim PierRotationAngle As Double
    Dim PierOrientation As Long
    Dim PierSizingRule As Long, PierShape As Long
    PierRotationAngle = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_ROTATION_ANGLE).Value
    PierOrientation = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_ORIENTATION).Value
    PierSizingRule = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_SIZING_RULE).Value
    PierShape = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_SHAPE).Value
    
    On Error GoTo ErrorHandler
    CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, PierShape, PRISMATIC_FOOTING_SHAPES, 124
    CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, PierSizingRule, FOOTING_COMP_SIZE_RULE, 125
    CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, PierOrientation, STRUCT_COORD_SYS_REF, 126
    On Error Resume Next
        
    Dim Vec As DVector
    Set Vec = New DVector
    Vec.Set 0, 0, 1
    
    If PierOrientation = 1 Then
        Matrix.Rotate PierRotationAngle, Vec
    ElseIf PierOrientation = 2 Then
        PierRotationAngle = PierRotationAngle + dMembAng
        Matrix.Rotate PierRotationAngle, Vec
    End If
    
    'TR# 66868-consider grout & pier rotation
    Dim GroutOrientation As Long
    Dim GroutShape As Long
    Dim GroutRotationAngle As Double
               
    GroutRotationAngle = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_ROTATION_ANGLE).Value
    GroutOrientation = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_ORIENTATION).Value
    GroutShape = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_SHAPE).Value
               
    CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, GroutShape, PRISMATIC_FOOTING_SHAPES, 121
    CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, GroutOrientation, STRUCT_COORD_SYS_REF, 123
               
    If GroutOrientation = 2 Then '2 is Local to the SPSMemberPart
        GroutRotationAngle = GroutRotationAngle + dMembAng
    End If
    If GroutShape = 3 Then GroutRotationAngle = 0# ' if grout is circular then not need to consider its angle
    
    Dim GroutWidth As Double, GroutLength As Double, PierEdgeClearance As Double
    
    If Not oMembSys Is Nothing Then ' we have a member
        If PierSizingRule = 1 Or PierSizingRule = 2 Then
           If Not bIsNeeded Then
               Dim sRefStd As String
               Dim sSecName As String
               Dim Depth As Double
               Dim SecWidth As Double
               Dim GroutEdgeClearance As Double
               If Not oMembPart Is Nothing Then
                   sRefStd = oMembPart.CrossSection.SectionStandard
                   sSecName = oMembPart.CrossSection.SectionName
               End If
               Depth = GetCSAttribData(sSecName, sRefStd, CROSSSSECTION_DIMENSIONS, CROSSSECTION_DEPTH)
               SecWidth = GetCSAttribData(sSecName, sRefStd, CROSSSSECTION_DIMENSIONS, CROSSSECTION_WIDTH)
               If Depth > SecWidth Then
                   GroutWidth = Depth
                   GroutLength = Depth
               Else
                   GroutWidth = SecWidth
                   GroutLength = SecWidth
               End If
               GroutRotationAngle = -dMembAng
            End If
        End If
    End If
           
    If bIsNeeded Then
        GroutWidth = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value
        GroutLength = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value
    End If
    
    If PierSizingRule = 1 Or PierSizingRule = 2 Then
        PierEdgeClearance = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_EDGE_CLEARANCE).Value
            
        'TR# 66868- Pier length & width calcualted considering orientation
        Dim groutwidthComp As Double, GroutLenComp As Double
        Dim tmpAng As Double
        If PierShape = 3 Then PierRotationAngle = 0#
        tmpAng = Abs(PierRotationAngle - GroutRotationAngle)
        If tmpAng > PI / 2 Then tmpAng = Abs(PI - tmpAng)
        
        groutwidthComp = GroutWidth * Cos(tmpAng) + GroutLength * Sin(tmpAng)
        GroutLenComp = GroutLength * Cos(tmpAng) + GroutWidth * Sin(tmpAng)
               
        If PierShape = 2 Then
            oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_WIDTH).Value = groutwidthComp + (PierEdgeClearance * 2)
            oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_LENGTH).Value = GroutLenComp + (PierEdgeClearance * 2)
        ElseIf PierShape = 3 Then
            'TR#72794- if pier is circular then take diagonal + clearance=diameter. i.e. width & length
            oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_WIDTH).Value = Sqr(GroutLenComp * GroutLenComp + groutwidthComp * groutwidthComp) + (PierEdgeClearance * 2)
            oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_LENGTH).Value = Sqr(GroutLenComp * GroutLenComp + groutwidthComp * groutwidthComp) + (PierEdgeClearance * 2)
        End If
    End If
    
    If bByPlaneOrDatum Then
       Dim oline As Line3d
       Dim GeomFactory As New IngrGeom3D.GeometryFactory
       Dim DummyFace As New Plane3d
       Dim temp As IJElements
       Set temp = New JObjectCollection ' elements
       Set oline = GeomFactory.Lines3d.CreateByPtVectLength(Nothing, x, y, z, 0, 0, -1, 1)
       oline.Infinite = True
       Dim oNewBottomSurf As IJSurface
       Dim code As Geom3dIntersectConstants
       Set DummyFace = GeomFactory.Planes3d.CreateByPointNormal(Nothing, Rtx, Rty, Rtz, Normx, Normy, Normz)
       Set oNewBottomSurf = DummyFace
    
       oNewBottomSurf.Intersect oline, temp, code
       
       Dim pt1 As Double, pt2 As Double, pt3 As Double
       If temp.count <> 0 Then
         Dim point As IJPoint
         Set point = New Point3d
         Set point = temp.Item(1)
         Dim dist As Double
         point.GetPoint pt1, pt2, pt3
       End If
       
       If pt3 > z Then
           On Error GoTo ErrorHandler
            SPSToDoErrorNotify FootingToDoMsgCodelist, TDL_FTGMACROS_BOTTOMPLANE_NOTVALID_CANNOTCOMPUTE, pPropertyDescriptions.CAO, Nothing
            Err.Raise E_FAIL
           
       End If
       
       If z > Rtz Then
           TotHt = Abs(z - pt3) - GroutHt
           oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_HEIGHT).Value = TotHt
       End If
    End If
    
    pOcc.Matrix = Matrix
          
Exit Sub
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    End If

End Sub
Public Sub CMEvaluatePierMaterial(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluatePierMaterial"
On Error GoTo ErrHandler
    Dim MemberObj As IJDMemberObjects
    Dim oAttrs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Dim iMaterial As IJDMaterial
    Set MemberObj = pPropertyDescriptions.CAO
    Set oSmartOcc = MemberObj.ItemByDispid(2)
    Set oAttrs = oSmartOcc
    Dim Material As String
    Dim oStructMatl As IJStructMaterial
    Dim Grade As String
    Material = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_MATERIAL).Value
    Grade = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_GRADE).Value
    If Not Material = vbNullString And Not Grade = vbNullString Then
        Set iMaterial = GetMaterialObject(Material, Grade)
        Set oStructMatl = oSmartOcc
        oStructMatl.StructMaterial = iMaterial
    End If

Exit Sub
ErrHandler:      HandleError MODULE, METHOD
End Sub
Public Sub CMEvaluatePierGeometry(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluatePierGeometry"
On Error GoTo ErrHandler
    
    Call CMEvaluatePier(pPropertyDescriptions, pObject)

Exit Sub
ErrHandler:      HandleError MODULE, METHOD
End Sub
Public Sub CMFinalConstructAsm(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAsm"
On Error GoTo ErrorHandler
    
  
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub


Public Sub CMConstructAsm(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMConstructAsm"
On Error GoTo ErrorHandler
    

  
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateCAO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const MT = "CMEvaluateCAO"
On Error GoTo ErrorHandler

    
    SetOccurrenceMatrix pObject

   
Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub
Public Sub CMEvaluateCAOWCG(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAOWCG"
On Error GoTo ErrHandler
    
    Exit Sub

ErrHandler:  HandleError MODULE, METHOD
End Sub
Private Sub CreateComponent(Component As String, ByVal pResourceManager As IUnknown, _
                ByRef oFootingComp As ISPSFootingComponent, ByVal pMembDescr As IJDMemberDescription, bSetInput As Boolean)
Const METHOD = "CreateComponent"
On Error GoTo ErrorHandler
     
     Dim oFtgFactory As SPSFootingFactory
     Set oFtgFactory = New SPSFootingFactory
     Set oFootingComp = oFtgFactory.CreateFootingComponent(pResourceManager)
     Dim oReferencesCollection As IMSSymbolEntities.IJDReferencesCollection
     Dim oSymbolEntitiesFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
     If bSetInput Then
         Dim oInputObj As Object
         Dim pIRCAsm As IJDReferencesCollection
         Set pIRCAsm = GetRefCollection(pMembDescr.CAO)
         If pIRCAsm.IJDEditJDArgument.GetCount > 1 Then
            Set oInputObj = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(2)
            Set oReferencesCollection = oSymbolEntitiesFactory.CreateEntity(referencesCollection, pResourceManager)
            oReferencesCollection.IJDEditJDArgument.SetEntity 1, oInputObj, IJPlane, "SPSSuppPlaneToRC_DEST"
         ElseIf pIRCAsm.IJDEditJDArgument.GetCount > 0 Then ' no member
            Set oInputObj = pIRCAsm.IJDEditJDArgument.GetEntityByIndex(1)
            Set oReferencesCollection = oSymbolEntitiesFactory.CreateEntity(referencesCollection, pResourceManager)
            oReferencesCollection.IJDEditJDArgument.SetEntity 1, oInputObj, IJPlane, "SPSSuppPlaneToRC_DEST"
         End If
     End If

     Dim oSmartOcc As IJSmartOccurrence
     Set oSmartOcc = oFootingComp
     Dim oItem As Object
     Dim strUserType As String
     Dim strSCName As String
     Dim oSmartItem As IJSmartItem
     Dim oSmartClass As IJSmartClass
     Dim oUserType As IJDUserType
     
     Set oItem = GetDefinition(Component)
    
     Set oSmartItem = oItem
     Set oSmartClass = oSmartItem.Parent
     strUserType = oSmartClass.SOUserType
     Set oUserType = oSmartOcc
     oUserType.UserType = strUserType
     strSCName = oSmartClass.SCName
     oSmartOcc.RootSelectorClass = strSCName
     oSmartOcc.RootSelection = oSmartItem.name
     
     If Not oReferencesCollection Is Nothing Then
        If oReferencesCollection.IJDEditJDArgument.GetCount > 0 Then
            ConnectSmartOccurrence oSmartOcc, oReferencesCollection
        End If
     End If
     
Exit Sub
ErrorHandler:     HandleError MODULE, METHOD
End Sub




 
Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler

    Dim i As Long
    Dim pAttrDescr As IJAttributeDescriptor
    Dim NonStateRO As Long

    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_FTGMACROS_ERROR, "ERROR")

    pAttrToChange.AttrValue = varNewAttrValue
    If (pAttrToChange.attrName = BOTTOM_ELEVATION) Then
        'for single footing, the footing origin is at the top center
        'of the footing
        Dim oFooting As SPSFooting
        Dim x#, y#, z#
        Set oFooting = pIJDAttrs
        oFooting.GetPosition x, y, z
        If varNewAttrValue > (z - TOL) Then
        
            IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_FTGMACROS_INVALID_Z_VALUE, "ERROR")
            Exit Function

        End If
    End If
    
    Set pAttrDescr = Nothing
    
    IJUserAttributeMgmt_OnAttributeChange = vbNullString
    
Exit Function
ErrorHandler:  HandleError MODULE, METHOD

End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String

End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrorHandler

    IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_FTGMACROS_ERROR, "ERROR")


    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim bUseElevDatum As Boolean

    
    Set pAttrColl = CollAllDisplayedValues
    
    For i = 1 To pAttrColl.count
        Set pAttrDescr = pAttrColl.Item(i)
        If (pAttrDescr.attrName = USE_ELEVATION_DATUM) Then
            pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly ' UseElevationDatum
            'is read only on the property page as there is no way to select a plane if 'UseElevationDatum' can be set to 'False'
            If pAttrDescr.AttrValue <> Empty Then ' in multi select mode, if the values are diffrent for diffrent footings, it is empty
                bUseElevDatum = pAttrDescr.AttrValue
            End If

        End If
    Next
    If (bUseElevDatum = False) Then
        'loop again to disable/enable Z value field based on how the footing is placed
        For i = 1 To pAttrColl.count
            Set pAttrDescr = pAttrColl.Item(i)
            If (pAttrDescr.attrName = BOTTOM_ELEVATION) Then
                pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly ' if the footing is placed
                'by Z value then it is made editable here.
            End If
        Next
    End If

Set pAttrColl = Nothing
Set pAttrDescr = Nothing


IJUserAttributeMgmt_OnPreLoad = vbNullString
Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructInterfaces.IJElements)
Const METHOD = "IJStructCustomFoulCheck_GetConnectedParts"
On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim pMemberSystem As ISPSMemberSystem
    Dim pDesignParts As IJElements
    Dim oSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oObject As Object
    
    Set oSmartOcc = pPartObject
    Set oRefColl = GetRefCollection(oSmartOcc)
    
    If Not oRefColl Is Nothing Then
        If Not oRefColl.IJDEditJDArgument Is Nothing Then
            If oRefColl.IJDEditJDArgument.GetCount > 0 Then
                Set oObject = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
                If Not oObject Is Nothing Then
                    If TypeOf oObject Is ISPSMemberSystem Then
                        Set pMemberSystem = oObject
                        If Not pMemberSystem Is Nothing Then
                            Set pDesignParts = pMemberSystem.DesignParts
                            For i = 1 To pDesignParts.count
                                pIJMonUnks.Add (pDesignParts.Item(i))
                            Next i
                        End If
                    End If
                End If
            End If
        End If
    End If
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD

End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = StandardGraphicEntity
End Sub

Private Sub ISPSFootingDefServices_GetInputs(ByVal FtgObject As Object, ByVal supported As SPSFootings.IJElements, ByVal supporting As SPSFootings.IJElements)
Const METHOD = "ISPSFootingDefServices_GetInputs"
On Error GoTo ErrorHandler
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = FtgObject
    Dim oRefColl As IJDReferencesCollection
    Set oRefColl = GetRefCollection(oSmartOcc)

    Dim i As Integer
    Dim cnt As Integer
    cnt = oRefColl.IJDEditJDArgument.GetCount
    If cnt > 1 Then
        supported.Add oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        supporting.Add oRefColl.IJDEditJDArgument.GetEntityByIndex(cnt)
    ElseIf cnt = 1 Then
        Dim oTemp As Object
        Set oTemp = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        If oTemp Is Nothing Then
            Set oTemp = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
        End If
        
        If TypeOf oTemp Is IJPlane Then
            supporting.Add oTemp
        Else
            supported.Add oTemp
        End If
    End If
    
Exit Sub
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Private Sub ISPSFootingDefServices_SetInputs(ByVal FtgObject As Object, ByVal FtgDefinitionObject As Object, ByVal supported As SPSFootings.IJElements, ByVal supporting As SPSFootings.IJElements)
'Works For Single Type Footing
'Create ReColl, may create more then one RefColl to support combined footing inputs
'where we group a column by more than one interface
'Could get column and look for a base-plate here too..
'Add suppd, suppg to RefColl by watched interface, either IJFullObject or Member specific?
Const METHOD = "ISPSFootingDefServices_SetInputs"
    Dim strUserType As String
    Dim strSCName As String
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oUserType As IJDUserType
    Dim oSmartOcc As IJSmartOccurrence
    Dim oFtgFactory As SPSFootingFactory
    Dim FtgSupporting As Boolean
    FtgSupporting = False
    Dim i As Integer
    Dim oReferencesCollection As IMSSymbolEntities.IJDReferencesCollection
         
    If supporting.count >= 1 Then
        FtgSupporting = True
    End If
    
    Set oSmartOcc = FtgObject
    
    ' See if there is already a reference collection and if not create one or clear the old one
    On Error Resume Next
    Set oReferencesCollection = GetRefCollection(oSmartOcc)
    If Not oReferencesCollection Is Nothing Then
        Dim j As Integer ' must remove old ref coll on ref coll if they exist (when moving from comb to single)
        For j = 1 To oReferencesCollection.IJDEditJDArgument.GetCount
            If TypeOf oReferencesCollection.IJDEditJDArgument.GetEntityByIndex(j) Is IJDReferencesCollection Then
                Dim oObject As ijdObject
                Set oObject = oReferencesCollection.IJDEditJDArgument.GetEntityByIndex(j)
                If Not oObject Is Nothing Then
                    oObject.Remove
                End If
            End If
        Next j
        oReferencesCollection.IJDEditJDArgument.RemoveAll
    Else
        Dim oSymbolEntitiesFactory As New IMSSymbolEntities.DSymbolEntitiesFactory
        Set oReferencesCollection = oSymbolEntitiesFactory.CreateEntity(referencesCollection, GetResourceMgr())
        ConnectSmartOccurrence oSmartOcc, oReferencesCollection
        Set oSymbolEntitiesFactory = Nothing
    End If
     
    Dim OldItem As IJSmartItem
    Dim strOldItemName As String
    On Error Resume Next
    Set OldItem = oSmartOcc.ItemObject
    Err.Clear
    
    If Not OldItem Is Nothing Then
        strOldItemName = OldItem.name
    End If
    
    Set oSmartItem = FtgDefinitionObject
    If strOldItemName <> oSmartItem.name Then
        Set oSmartClass = oSmartItem.Parent
        strUserType = oSmartClass.SOUserType
        Set oUserType = oSmartOcc
        oUserType.UserType = strUserType
        strSCName = oSmartClass.SCName
        oSmartOcc.RootSelectorClass = strSCName
        oSmartOcc.RootSelection = oSmartItem.name
    End If
    
    i = 1
    If supported.count > 0 Then ' add the member on ref collection
        If TypeOf supported.Item(1) Is ISPSMemberSystem Then ' only add member systems for single
            oReferencesCollection.IJDEditJDArgument.SetEntity 1, supported.Item(1), ISPSMemberSystemSuppingNotify1, "MembSysSuppingNotify1RC_DEST"
            i = i + 1
        End If
    End If
    
    
    If FtgSupporting Then
        oReferencesCollection.IJDEditJDArgument.SetEntity i, supporting.Item(1), IJPlane, "SPSSuppPlaneToRC_DEST"
    End If


    Set oReferencesCollection = Nothing
    Set oSymbolEntitiesFactory = Nothing
    Set oReferencesCollection = Nothing
    Set oSmartOcc = Nothing
    Set oUserType = Nothing
    Set oSmartItem = Nothing
    Set oSmartClass = Nothing
    
Exit Sub
ErrorHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    End If

End Sub


Private Function UserAttributeMgmt_Validate(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, sInterfaceName As String, sAttributeName As String, ByVal varAttributeValue As Variant) As String
Const METHOD = "UserAttributeMgmt_Validate"
On Error GoTo ErrorHandler

' first of all check if the symbol definition has CMCheck methods defined - TBD
    UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FTGMACROS_ERROR, "ERROR")

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
    
    Set oAttribute = pIJDAttrs.CollectionOfAttributes(sInterfaceName).Item(sAttributeName)

    If oAttribute.Value <> vbNullString Then
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
                    UserAttributeMgmt_Validate = ErrMessage
                    Exit Function
                Else
                End If
            End If
            On Error GoTo ErrorHandler
        End If
    End If
' get the list of interfaces implemented by the schema from IJDAttributes
' make sure that you are not looking into a system interface
' from the input interfaceName and propertyName, get the property type from catalog info
' select case on the property types, and in there, mention the valid attribute values for each propertyName

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
        UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FTGMACROS_SCHEMAERROR, "SchemaERROR")
        GoTo ErrorHandler
    End If
    Set oAttrCol = oAttributeMetaData.InterfaceAttributes(InterfaceID)
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        
        If oAttrObj.name = sAttributeName Then
            Select Case oAttrObj.Type
                Case DOUBLE_VALUE
                        If (varAttributeValue <= 0#) Then
                            UserAttributeMgmt_Validate = sAttributeName
                            Set oAttributeMetaData = Nothing
                            Exit Function
                        End If
            End Select
        End If
    Next
    
    UserAttributeMgmt_Validate = vbNullString
    Set oAttributeMetaData = Nothing
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function



Private Sub ValidatePierHt(pMemberDescription As IJDMemberDescription)
Const METHOD = "ValidatePierHt"
    
    Dim oReferencesCollection0fCAO  As IMSSymbolEntities.IJDReferencesCollection
    Dim oMembSys As ISPSMemberSystem
    Dim oMembPart As ISPSMemberPartPrismatic
    Dim GroutHt As Double
    Dim oAttr As IJDAttributes
    
    GroutHt = 0#

    Set oReferencesCollection0fCAO = GetRefCollection(pMemberDescription.CAO)
     
    Dim oFooting As ISPSFooting
    Set oFooting = pMemberDescription.CAO

    Dim x As Double, y As Double, z As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim oPlane As IJPlane
    Dim obj As Object
    Set obj = Nothing
    
    If oReferencesCollection0fCAO.IJDEditJDArgument.GetCount > 0 Then
        Set obj = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1)
    End If
    
    If Not obj Is Nothing Then
        If TypeOf obj Is ISPSMemberSystem Then
            Set oMembSys = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1)
        End If
    End If
    
    If oMembSys Is Nothing Then
        oFooting.GetPosition x, y, z
        If oReferencesCollection0fCAO.IJDEditJDArgument.GetCount > 0 Then
            Dim oTemp As Object
            Set oTemp = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1)
            If oTemp Is Nothing Then
                Set oTemp = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(2)
            End If
            If Not oTemp Is Nothing Then
                Set oPlane = oTemp
            End If
        End If
    Else  ' have a member
        Dim GlobalDelta As Double
        GetFootingPositionFromMember oMembSys, x, y, z, GlobalDelta
        Set oMembPart = GetBottomMemberPart(oMembSys)
        If oReferencesCollection0fCAO.IJDEditJDArgument.GetCount > 1 Then
            Set oPlane = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(2)
        End If
        oMembPart.Axis.EndPoints x, y, z, x1, y1, z1
        If z1 < z Then z = z1
     End If
     Dim Normx As Double, Normy As Double, Normz As Double
     Dim Rtx As Double, Rty As Double, Rtz As Double
     Dim bByPlaneOrDatum As Boolean
       
     If Not oPlane Is Nothing Then
        bByPlaneOrDatum = True
        oPlane.GetNormal Normx, Normy, Normz
        oPlane.GetRootPoint Rtx, Rty, Rtz
     Else
        Dim bUseDatum As Boolean
        Dim oCollProxy As IJDAttributesCol
        
        Set oAttr = pMemberDescription.CAO
        On Error Resume Next 'this interface was added in a later release so may not exist
        'in a migrated model
        Set oCollProxy = oAttr.CollectionOfAttributes(STRUCT_ELEVATION_DATUM)
        On Error GoTo ErrorHandler
            
        If Not oCollProxy Is Nothing Then
            bUseDatum = oCollProxy.Item(USE_ELEVATION_DATUM).Value
            If bUseDatum = True Then
                Rtz = oCollProxy.Item(BOTTOM_ELEVATION).Value
                Rtx = 0
                Rty = 0
                Normx = 0
                Normy = 0
                Normz = 1
                bByPlaneOrDatum = True
            End If
        End If
     End If
    
    
    If bByPlaneOrDatum Then
        Dim oline As Line3d
        Dim GeomFactory As New IngrGeom3D.GeometryFactory
        Dim DummyFace As New Plane3d
        Dim temp As IJElements

        Dim oNewBottomSurf As IJSurface
        Dim pt1 As Double, pt2 As Double, pt3 As Double
        Dim code As Geom3dIntersectConstants
        Dim point As IJPoint
        


        Set temp = New JObjectCollection ' elements
        Set oline = GeomFactory.Lines3d.CreateByPtVectLength(Nothing, x, y, z, 0, 0, -1, 1)
        oline.Infinite = True
                
        Set DummyFace = GeomFactory.Planes3d.CreateByPointNormal(Nothing, Rtx, Rty, Rtz, Normx, Normy, Normz)
        Set oNewBottomSurf = DummyFace

        oNewBottomSurf.Intersect oline, temp, code

        If temp.count <> 0 Then
          
          Set point = New Point3d
          Set point = temp.Item(1)
          point.GetPoint pt1, pt2, pt3
        End If
        
        If (z - pt3 - GroutHt) < 0# Then
            SPSToDoErrorNotify FootingToDoMsgCodelist, TDL_FTGMACROS_INVALID_PIERHEIGHT, pMemberDescription.CAO, Nothing
            Err.Raise E_FAIL
        End If
     End If
     
    Set oReferencesCollection0fCAO = Nothing
    Set oMembSys = Nothing
    Set oMembPart = Nothing
    Set oPlane = Nothing
    Set temp = Nothing
    Set DummyFace = Nothing
    Set point = Nothing
     
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

'TR#71850
Private Property Get ISPSFoundationInputHelper_ValidateObjects(ByVal inputSupported As Object, ByVal inputSupporting As Object, ByVal SupportedObjList As SP3DStructInterfaces.IJElements, ByVal ObjsinSelectSet As SP3DStructInterfaces.IJElements) As SP3DStructInterfaces.SPSFoundationInputHelperStatus
Const MT = "ISPSFoundationInputHelper_ValidateObjects"
On Error GoTo ErrorHandler
    
If inputSupporting Is Nothing Then
    ISPSFoundationInputHelper_ValidateObjects = CommonCheckMember(inputSupported, SupportedObjList)
Else
    ISPSFoundationInputHelper_ValidateObjects = CommonCheckBottomPlane(inputSupporting, SupportedObjList, ObjsinSelectSet)
End If
Exit Property
ErrorHandler:
    HandleError MODULE, MT
End Property

Private Sub Class_Initialize()
Set m_oLocalizer = New IMSLocalizer.Localizer
m_oLocalizer.Initialize App.Path & "\" & App.EXEName
End Sub

Private Sub Class_Terminate()
Set m_oLocalizer = Nothing
End Sub

'*************************************************************************
'Function
'CMMigrateAggregator
'
'Abstract
'Migrates thr footing to the correct surface if it is split.
'
'Arguments
'IJDMemberDescription interface of the member
'
'Return
'
'Exceptions
'
'***************************************************************************
Public Sub CMMigrateAggregator(oAggregatorDesc As IJDAggregatorDescription, oMigrateHelper As IJMigrateHelper)

  Const MT = "CMMigrateAggregator"
  On Error GoTo ErrorHandler
  
  MigrateFootingAggregator oAggregatorDesc, oMigrateHelper
    
  Exit Sub
ErrorHandler:  HandleError MODULE, MT
End Sub


Private Sub ISPSTransformHelper_Transform(ByVal Trans4x4 As SP3DStructInterfaces.IJDT4x4, ByVal ObjectToTransform As Object)
Const METHOD = "ISPSTransformHelper_Transform"
On Error GoTo ErrHandler
    Dim oFootingComponents As IJDMemberObjects
    Dim oPier As SPSFootingComponent
    Dim oGrout As SPSFootingComponent
    Dim bIgnoreOrientation As Boolean


    Dim strDebug As String
        
    Set oFootingComponents = ObjectToTransform
    
    'update grout's rotation
    strDebug = "updating grout's rotation"
    
    
    On Error Resume Next
    Set oGrout = oFootingComponents.ItemByDispid(1)
    On Error GoTo ErrHandler
    
    'if the footing is placed by point then update the rotaion. but if placed by member then update
    'only if orientation is global
    bIgnoreOrientation = Not IsFootingPlacedByMember(ObjectToTransform)
    
    If Not oGrout Is Nothing Then
        UpdateRotationAngle oGrout, FTG_GROUT_PAD, GROUT_ROTATION_ANGLE, GROUT_ORIENTATION, Trans4x4, bIgnoreOrientation
    End If

    'update pier's rotation
    strDebug = "updating pier's rotation"
    On Error Resume Next
    Set oPier = oFootingComponents.ItemByDispid(2)
    On Error GoTo ErrHandler
    
    If Not oPier Is Nothing Then
        UpdateRotationAngle oPier, FTG_PIER, PIER_ROTATION_ANGLE, PIER_ORIENTATION, Trans4x4, bIgnoreOrientation
    End If
    
    'Update the Datum Elevation so that it is consistent with applied transform (either applied interactively or through MDR)
    
    UpdateElevationDatumOnTransform ObjectToTransform, Trans4x4
    
    Exit Sub
ErrHandler:
Err.Raise ReportError(Err, strSourceFile, METHOD, strDebug).Number

End Sub


Public Sub CMEvaluateElevationDatum(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const MT = "CMEvaluateElevationDatum"
On Error GoTo ErrorHandler

    'if the footing is placed by plane then update datum value
    UpdateElevationDatumFromPlane pObject
    
Exit Sub
ErrorHandler: HandleError MODULE, MT
End Sub
