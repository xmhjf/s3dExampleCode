VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PierAndSlabFtgDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2006, Intergraph Corporation. All rights reserved.
'
'File
'    PierAndSlabFtgDef.cls
'
'Author
'       28th January 2003        AP
'
'Description
'
'Notes
'
'History:
'   24-Jun-2006 JMS DI#60069 - Changes to allow editing of the weight and CG values
'                   changed call to SetWCG to call new interface to put weight and CG
'                   values since SetWCG is reservered for setting user defined values
'                   when the values here are the computed values
'       07-Jul-2006 JMS TR#101063 - Tolerate the nonexistence of IJWCGValueOrigin interface
'  06-Mar-07   RS   CR#41094 - Changes for placing a footing in space and
'                   dis/reconnect to members
'  15-Apr-14    RRK    TR-CP-250717     Made change to CMEvaluateCAO method to solve Recorded Exception Minidump
'*******************************************************************
Option Explicit
Private Const MODULE = "PierAndSlabFtgDef"
Private Const strSourceFile = "PierAndSlabFtgDef.def"
Private Const INTERFACE_WCGValueOrigin As String = "IJWCGValueOrigin"
Private Const PROPERTY_DryWCGOrigin As String = "DryWCGOrigin"
Private Enum enumWeightCGDerivation
    WEIGHTCG_Computed = 2
    WEIGHTCG_UserDefined = 4
End Enum
Private Const CONST_ItemProgId As String = "SPSFootingMacros.PierAndSlabFtgDef"
Private Const MODELDATABASE = "Model"

Private m_strErrorDescr As String
Private m_oLocalizer As IJLocalizer
Dim bOnPreLoad As Boolean
Implements ISPSFootingDefServices
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements IJStructCustomFoulCheck
Implements ISPSFoundationInputHelper 'TR#71850
Implements ISPSTransformHelper
Implements ICustomOutputHelper



Private Function ICustomOutputHelper_GetComponents(ByVal assemblyObject As Object) As IJElements
Const METHOD = "ICustomOutputHelper_GetComponents"
On Error GoTo ErrHandler

    Dim colComponents As IJElements
    Dim colGroutAttribs As IJElements
    Dim colPierAttribs As IJElements
    Dim colSlabAttribs As IJElements
    
    Dim colGroutFaces As IJElements
    Dim colPierFaces As IJElements
    Dim colSlabFaces As IJElements
    Dim oFootingComp As IAssemblyComponent
    Dim oOutput As Object
    Dim oAttrs As IJDAttributes
    Dim strGrade As String, strMaterial As String
    Dim iSymbolOccMisc As iSymbolOccMisc
    Dim oOutPutProxy As IJDProxy
    Dim oRelation As IJDAssocRelation
    Dim oOutputs As IJDTargetObjectCol
    Dim oRelationship As IJDRelationship
    Dim lngCount As Long
    Dim lngIndex As Long
    Dim lPos1 As Long
    Dim lPos2 As Long
    Dim lPos3 As Long
    Dim lPos4 As Long
    Dim lPos5 As Long
    Dim colAttribs As CollectionProxy
    Dim oAttrib As AttributeProxy
    Dim oSymbol As IJDSymbol
    
    
    Set oSymbol = assemblyObject
    
    Set iSymbolOccMisc = assemblyObject
    Set oAttrs = assemblyObject
    
    Set colComponents = New JObjectCollection
    Set colGroutFaces = New JObjectCollection
    Set colPierFaces = New JObjectCollection
    Set colSlabFaces = New JObjectCollection
    
    Set oOutPutProxy = iSymbolOccMisc.BindToOCIfExists(SYMBOL_REP_NAME)
    Set oRelation = oOutPutProxy.Source

    Set oOutputs = oRelation.CollectionRelations(IJOUTPUTCOLLECTION, TO_OUTPUTS)
    lngCount = oOutputs.count
    For lngIndex = 1 To lngCount
                
        'the output below is not transformed by the symbol's occurrence matrix
        Set oOutput = oOutputs.Item(lngIndex)
        Set oRelationship = oOutputs.GetRelationshipToTarget(oOutput)
        lPos1 = InStr(1, oRelationship.name, "Grout", vbTextCompare)
        lPos2 = InStr(1, oRelationship.name, "Octagon", vbTextCompare)
        lPos3 = InStr(1, oRelationship.name, "Projection", vbTextCompare)
        lPos4 = InStr(1, oRelationship.name, "Pier", vbTextCompare)
        lPos5 = InStr(1, oRelationship.name, "Slab", vbTextCompare)
       
        Dim oOutPutRefProxy As Object
        
        'get the output that is transformed by the occurrence matrix
        Set oOutPutRefProxy = oSymbol.BindToOutput(SYMBOL_REP_NAME, oRelationship.name)
        
        'the looping could have been avoided by using the output name directly and calling
        'oSymbol.BindToOutput with the output name. But looping and getting relation name allowes
        'easy grouping components by pattern searching

        If ((lPos1 <> 0) Or (lPos2 <> 0) Or (lPos3 <> 0)) Then
            colGroutFaces.Add oOutPutRefProxy
        End If
        If ((lPos4 <> 0)) Then
            colPierFaces.Add oOutPutRefProxy
        End If
        If ((lPos5 <> 0)) Then
            colSlabFaces.Add oOutPutRefProxy
        End If
        Set oOutPutRefProxy = Nothing
        
    Next

    'create grout
    
    If colGroutFaces.count > 0 Then
        Set oFootingComp = New SPSAssemblyComponent
        'put name
        oFootingComp.name = "Grout"
        
        'put material and grade
        strMaterial = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_MATERIAL).Value
        strGrade = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_GRADE).Value
        
        oFootingComp.SetMaterialTypeAndGrade strMaterial, strGrade
        
        'put faces
        oFootingComp.Geometry = colGroutFaces
        
        'put attributes
        
        Set colGroutAttribs = New JObjectCollection
        
        For Each oAttrib In oAttrs.CollectionOfAttributes(FTG_GROUT_PAD)
            colGroutAttribs.Add oAttrib
        Next
        
        For Each oAttrib In oAttrs.CollectionOfAttributes(FTG_GROUT_PAD_DIM)
            colGroutAttribs.Add oAttrib
        Next
        oFootingComp.Attributes = colGroutAttribs
        colComponents.Add oFootingComp
    End If
    'create pier
    If colPierFaces.count > 0 Then
        Set oFootingComp = New SPSAssemblyComponent
        'put name
        oFootingComp.name = "Pier"
        
        'put material and grade
        strMaterial = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_MATERIAL).Value
        strGrade = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_GRADE).Value
        
        oFootingComp.SetMaterialTypeAndGrade strMaterial, strGrade
        
        'put faces
        oFootingComp.Geometry = colPierFaces
        
        'put attributes
        
        Set colPierAttribs = New JObjectCollection
        
        For Each oAttrib In oAttrs.CollectionOfAttributes(FTG_PIER)
            colPierAttribs.Add oAttrib
        Next
        
        For Each oAttrib In oAttrs.CollectionOfAttributes(FTG_PIER)
            colPierAttribs.Add oAttrib
        Next
        oFootingComp.Attributes = colPierAttribs
        colComponents.Add oFootingComp
    End If
    'create slab
    If colSlabFaces.count > 0 Then
        Set oFootingComp = New SPSAssemblyComponent
        'put name
        oFootingComp.name = "Slab"
        
        'put material and grade
        strMaterial = oAttrs.CollectionOfAttributes(FTG_SLAB).Item(SLAB_MATERIAL).Value
        strGrade = oAttrs.CollectionOfAttributes(FTG_SLAB).Item(SLAB_GRADE).Value
        
        oFootingComp.SetMaterialTypeAndGrade strMaterial, strGrade
        
        'put faces
        oFootingComp.Geometry = colSlabFaces
        
        'put attributes
        
        Set colSlabAttribs = New JObjectCollection
        
        For Each oAttrib In oAttrs.CollectionOfAttributes(FTG_SLAB)
            colSlabAttribs.Add oAttrib
        Next
        
        For Each oAttrib In oAttrs.CollectionOfAttributes(FTG_SLAB)
            colSlabAttribs.Add oAttrib
        Next
        oFootingComp.Attributes = colSlabAttribs
        
        colComponents.Add oFootingComp
    End If
    
    Set ICustomOutputHelper_GetComponents = colComponents
    Exit Function
ErrHandler:
    Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    
End Function

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
 'Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
 IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
  ' Name should be unique
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
  pAD.UserTypeClsid = "{27CEB5DD-51B1-49fc-A659-5A29E4210791}" 'CPUAPierAndSlabFooting
  pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructCA"
  pAD.SetCMSetInputs -1, -1
  pAD.SetCMRemoveInputs -1, -1
  Set pAD = Nothing
  
  ' Aggregator property
  Dim pAPDs As IJDPropertyDescriptions
  Set pAPDs = pDefinition
  pAPDs.RemoveAll ' Remove all the previous property descriptions
  'Adding properties that allow us to watch user interfaces as inputs for update.
  'By adding a CMEvaluate it allows us to treat it as input/output to set other values on the user interface
  'Used the same CMEvaluate for all so that we only call one complete update method
  ' NOTE: Not sure how many times this will be called for one or more interfaces modified?
  
  pAPDs.AddProperty "IJDAttributes", 1, IJDAttributes, "CMEvaluateCAO", imsCOOKIE_ID_USS_LIB
  pAPDs.AddProperty STRUCT_MATERIAL, 2, IJStructMaterial, "CMEvaluateCAOMaterial", imsCOOKIE_ID_USS_LIB
  pAPDs.AddProperty "IJWeightCG", 3, IJWeightCG, "CMEvaluateCAOWCG", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE

  Set pAPDs = Nothing
  
  Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
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
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub
Public Sub CMFinalConstructCA(pAggregatorDescription As IJDAggregatorDescription)
Const MT = "CMFinalConstructCA"
'Delegate this to the SetInputs since it will check for existence and create the RefCol if not there
'This is only called once in the lifetime of a CAD


  
Exit Sub
End Sub

Public Sub CMEvaluateCAO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAO"
On Error GoTo ErrHandler
    
    ' Get SO and Attributes pointer off the footing
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pObject
    Dim oAttribs As IJDAttributes
    Set oAttribs = pObject
    
    
 
    Dim oReferencesCollection0fCAO  As IMSSymbolEntities.IJDReferencesCollection
    Set oReferencesCollection0fCAO = GetRefCollection(oSmartOcc) '(pPropertyDescriptions.CAO)

    ' Get the x, y, z from the member or from the footing
    Dim oMembSys As ISPSMemberSystem
    
    If Not oReferencesCollection0fCAO Is Nothing Then
        If Not oReferencesCollection0fCAO.IJDEditJDArgument Is Nothing Then
            If Not oReferencesCollection0fCAO.IJDEditJDArgument.GetCount = 0 Then
                Set oMembSys = oReferencesCollection0fCAO.IJDEditJDArgument.GetEntityByIndex(1)
            End If
        End If
    End If
    
    Dim oFooting As ISPSFooting
    Set oFooting = pPropertyDescriptions.CAO
    Dim Matrix As IJDT4x4
    Dim oFootingMatrix As IJDT4x4
    Dim GlobalDelta As Double
        
    Set Matrix = New DT4x4
    Matrix.LoadIdentity

    Dim oMembPart As ISPSMemberPartPrismatic
    Dim x As Double, y As Double, z As Double
    Dim sRefStd As String
    Dim sSecName As String
    If oMembSys Is Nothing Then
        oFooting.GetPosition x, y, z
'        oFooting.GetMatrix oFootingMatrix
'        Matrix.LoadMatrix oFootingMatrix
'        GlobalDelta = GetPlanAngleFromMatrix(Matrix)
        GlobalDelta = 0
    Else  ' have a member
        GetFootingPositionFromMember oMembSys, x, y, z, GlobalDelta
        Set oMembPart = GetBottomMemberPart(oMembSys)
        oFooting.SetPosition x, y, z
     End If
        
       
    ' Set the location of the footing via the matrix
    ' Get the symbol
    Dim pSymbol As IJDSymbol
    Set pSymbol = pObject
    Dim pOcc As IJDOccurrence
    On Error Resume Next
    Set pOcc = pSymbol
    If pOcc Is Nothing Then
         Exit Sub
    End If

    ' Set the matrix
    Matrix.IndexValue(12) = x
    Matrix.IndexValue(13) = y
    Matrix.IndexValue(14) = z

    oAttribs.CollectionOfAttributes("IJUASPSPierAndSlabFooting").Item("GlobalDelta").Value = GlobalDelta
    
    If oMembSys Is Nothing Then ' have no member
        pOcc.Matrix = Matrix
        Exit Sub
    End If
        
        
    sRefStd = oMembPart.CrossSection.SectionStandard
    sSecName = oMembPart.CrossSection.SectionName
   
        
    ' Get Grout Attributes
    Dim GroutSizingRule As Long
    Dim Depth As Double
    Dim SecWidth As Double
    GroutSizingRule = oAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_SIZE_RULE).Value
    On Error GoTo ErrHandler
    CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, GroutSizingRule, FOOTING_COMP_SIZE_RULE, 122
    On Error Resume Next
    If GroutSizingRule = 1 Or GroutSizingRule = 2 Then
        Dim GroutEdgeClearance As Double
        Depth = GetCSAttribData(sSecName, sRefStd, CROSSSSECTION_DIMENSIONS, CROSSSECTION_DEPTH)
        SecWidth = GetCSAttribData(sSecName, sRefStd, CROSSSSECTION_DIMENSIONS, CROSSSECTION_WIDTH)
    
        GroutEdgeClearance = oAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_EDGE_CLEARANCE).Value
        Dim GroutShape As Long
        GroutShape = oAttribs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_SHAPE).Value
        On Error GoTo ErrHandler
        CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, GroutShape, PRISMATIC_FOOTING_SHAPES, 121
        On Error Resume Next
        
        Dim width As Double, length As Double
        If GroutShape = 2 Then                             ' rectangle
            width = GroutEdgeClearance * 2 + SecWidth
            length = GroutEdgeClearance * 2 + Depth
        ElseIf GroutShape = 3 Or GroutShape = 4 Then       ' circle or octagon
            If Depth > SecWidth Then
                width = GroutEdgeClearance * 2 + Depth
            Else
                width = GroutEdgeClearance * 2 + SecWidth
            End If
            length = width
        End If
        oAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value = width
        oAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value = length
     End If
     
     ' Get Pier Attributes
     Dim PierSizingRule As Long, PierShape As Long
     
     PierSizingRule = oAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_SIZING_RULE).Value
     On Error GoTo ErrHandler
     CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, PierSizingRule, FOOTING_COMP_SIZE_RULE, 125
     On Error Resume Next
     If PierSizingRule = 1 Or PierSizingRule = 2 Then
        Dim GroutWidth As Double, GroutLength As Double, PierEdgeClearance As Double
        GroutWidth = oAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value
        GroutLength = oAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value
        PierEdgeClearance = oAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_EDGE_CLEARANCE).Value
        PierShape = oAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_SHAPE).Value
        On Error GoTo ErrHandler
        CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, PierShape, PRISMATIC_FOOTING_SHAPES, 124
        On Error Resume Next
        If PierShape = 2 Then
            oAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_WIDTH).Value = GroutWidth + (PierEdgeClearance * 2)
            oAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_LENGTH).Value = GroutLength + (PierEdgeClearance * 2)
        ElseIf PierShape = 3 Then
            'TR#72794- if pier is circular then take diagonal +clearance=diametre. i.e. width & length
            oAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_WIDTH).Value = Sqr(GroutWidth * GroutWidth + GroutLength * GroutLength) + (PierEdgeClearance * 2)
            oAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_LENGTH).Value = Sqr(GroutWidth * GroutWidth + GroutLength * GroutLength) + (PierEdgeClearance * 2)
        End If
     End If
     
     ' Get slab attributes
     Dim SlabSizingRule As Long, SlabShape As Long
     SlabSizingRule = oAttribs.CollectionOfAttributes(FTG_SLAB).Item(SLAB_SIZING_RULE).Value
     
      If SlabSizingRule = 1 Or SlabSizingRule = 2 Then
        Dim PierWidth As Double, PierLength As Double, SlabEdgeClearance As Double
        PierWidth = oAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_WIDTH).Value
        PierLength = oAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_LENGTH).Value
        SlabEdgeClearance = oAttribs.CollectionOfAttributes(FTG_SLAB).Item(SLAB_EDGE_CLEARANCE).Value
        SlabShape = oAttribs.CollectionOfAttributes(FTG_SLAB).Item(SLAB_SHAPE).Value
        On Error GoTo ErrHandler
        CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, SlabShape, PRISMATIC_FOOTING_SHAPES, 127
        On Error Resume Next
        'TR#72794- if pier is circular then take diagonal +clearance=diametre. i.e. width & length
        If SlabShape = 2 Then
            oAttribs.CollectionOfAttributes("IJUASPSFtgSlabDim").Item(SLAB_WIDTH).Value = PierWidth + (SlabEdgeClearance * 2)
            oAttribs.CollectionOfAttributes("IJUASPSFtgSlabDim").Item(SLAB_LENGTH).Value = PierLength + (SlabEdgeClearance * 2)
        ElseIf SlabShape = 3 Then
            If PierShape = 2 Then
                oAttribs.CollectionOfAttributes("IJUASPSFtgSlabDim").Item(SLAB_WIDTH).Value = Sqr(PierWidth * PierWidth + PierLength * PierLength) + (SlabEdgeClearance * 2)
                oAttribs.CollectionOfAttributes("IJUASPSFtgSlabDim").Item(SLAB_LENGTH).Value = Sqr(PierWidth * PierWidth + PierLength * PierLength) + (SlabEdgeClearance * 2)
            Else
                oAttribs.CollectionOfAttributes("IJUASPSFtgSlabDim").Item(SLAB_WIDTH).Value = PierWidth + (SlabEdgeClearance * 2)
                oAttribs.CollectionOfAttributes("IJUASPSFtgSlabDim").Item(SLAB_LENGTH).Value = PierLength + (SlabEdgeClearance * 2)
            End If
        End If
     End If
    
     pOcc.Matrix = Matrix
Exit Sub
ErrHandler:
    ' For errors logged with E_FAIL, a todo list error will be generated so we should not
    '   be logging anything to the error log
    If Err.Number = E_FAIL Then
        Err.Raise E_FAIL
    Else
        Err.Raise ReportError(Err, strSourceFile, METHOD).Number
    End If

End Sub
Public Sub CMEvaluateCAOMaterial(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAOMaterial"
On Error GoTo ErrHandler
    
    Dim oAttrs As IJDAttributes
    Dim oDefAttr As IJDAttributes
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Dim iMaterial As IJDMaterial
    Dim Material As String
    Dim Grade As String
    
    Set oSmartOcc = pObject
    Set oAttrs = oSmartOcc
    
    
    Set oDefAttr = oSmartOcc.ItemObject
    If Not IsSOOverridden(oAttrs.CollectionOfAttributes(FTG_GROUT_PAD)) Then
        CopyValuesToSOFromItem oAttrs.CollectionOfAttributes(FTG_GROUT_PAD), oDefAttr.CollectionOfAttributes(FTG_GROUT_PAD)
    End If
    
    If Not IsSOOverridden(oAttrs.CollectionOfAttributes(FTG_PIER)) Then
        CopyValuesToSOFromItem oAttrs.CollectionOfAttributes(FTG_PIER), oDefAttr.CollectionOfAttributes(FTG_PIER)
    End If
    
    If Not IsSOOverridden(oAttrs.CollectionOfAttributes(FTG_SLAB)) Then
        CopyValuesToSOFromItem oAttrs.CollectionOfAttributes(FTG_SLAB), oDefAttr.CollectionOfAttributes(FTG_SLAB)
    End If
    
    'Grout Material
    Material = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_MATERIAL).Value
    Grade = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_GRADE).Value
    If Not Material = vbNullString And Not Grade = vbNullString Then
        Set iMaterial = GetMaterialObject(Material, Grade)
        Call AddRelationShip(oSmartOcc, iMaterial, "Grout")
    End If
  
    'Pier Material
    Material = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_MATERIAL).Value
    Grade = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_GRADE).Value
    If Not Material = vbNullString And Not Grade = vbNullString Then
        Set iMaterial = GetMaterialObject(Material, Grade)
        Call AddRelationShip(oSmartOcc, iMaterial, "Pier")
    End If
    
    'Slab Material
    Material = oAttrs.CollectionOfAttributes(FTG_SLAB).Item(SLAB_MATERIAL).Value
    Grade = oAttrs.CollectionOfAttributes(FTG_SLAB).Item(SLAB_GRADE).Value
    If Not Material = vbNullString And Not Grade = vbNullString Then
        Set iMaterial = GetMaterialObject(Material, Grade)
        Call AddRelationShip(oSmartOcc, iMaterial, "Slab")
    End If
    
Exit Sub
ErrHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateCAOWCG(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAOWCG"
On Error GoTo ErrHandler

    Dim MemberObj As IJDMemberObjects
    Dim oSmartOcc As IJSmartOccurrence
    Dim oAttrs As IJDAttributes
    Dim iWCG As ISPSComputedWeightCG
    Dim Grade As String, Material As String
    Dim density As Variant
    Dim Volume As Double
    Dim GroutWt As Double, PierWt As Double, SlabWt As Double
    Dim TotVol As Double, TotSurfArea As Double, dSurfArea As Double
    Dim AccumCG As DPosition
    Set AccumCG = New DPosition
    Dim iMaterial As IJDMaterial
    Dim length As Double, width As Double
    Dim Shape As Long
    
    Dim pSymbol As IJDSymbol
    Dim pOcc As IJDOccurrence
    Dim Matrix As IJDT4x4
    Set Matrix = New DT4x4
    Set pSymbol = pObject
    On Error Resume Next
    Set pOcc = pSymbol
    Set Matrix = pOcc.Matrix

    AccumCG.x = Matrix.IndexValue(12)
    AccumCG.y = Matrix.IndexValue(13)
    
    Set oSmartOcc = pObject
    Set oAttrs = oSmartOcc
    
    Dim oAttrCol As IJDAttributesCol
    Dim lWCGOrigin As Long

    lWCGOrigin = WEIGHTCG_Computed
    On Error Resume Next
    Set oAttrCol = oAttrs.CollectionOfAttributes(INTERFACE_WCGValueOrigin)
    On Error GoTo ErrHandler
    If Not oAttrCol Is Nothing Then
        lWCGOrigin = oAttrCol.Item(PROPERTY_DryWCGOrigin).Value
    End If
    
    If lWCGOrigin <> WEIGHTCG_UserDefined Then
        Dim oDefAttrs As IJDAttributes
        Set oDefAttrs = oSmartOcc.ItemObject
        
        If Not IsSOOverridden(oAttrs.CollectionOfAttributes(FTG_GROUT_PAD)) Then
            CopyValuesToSOFromItem oAttrs.CollectionOfAttributes(FTG_GROUT_PAD), oDefAttrs.CollectionOfAttributes(FTG_GROUT_PAD)
        End If
        If Not IsSOOverridden(oAttrs.CollectionOfAttributes(FTG_GROUT_PAD_DIM)) Then
            CopyValuesToSOFromItem oAttrs.CollectionOfAttributes(FTG_GROUT_PAD_DIM), oDefAttrs.CollectionOfAttributes(FTG_GROUT_PAD_DIM)
        End If
        
        If Not IsSOOverridden(oAttrs.CollectionOfAttributes(FTG_PIER)) Then
            CopyValuesToSOFromItem oAttrs.CollectionOfAttributes(FTG_PIER), oDefAttrs.CollectionOfAttributes(FTG_PIER)
        End If
    
        If Not IsSOOverridden(oAttrs.CollectionOfAttributes(FTG_PIER_DIM)) Then
            CopyValuesToSOFromItem oAttrs.CollectionOfAttributes(FTG_PIER_DIM), oDefAttrs.CollectionOfAttributes(FTG_PIER_DIM)
        End If
        
        If Not IsSOOverridden(oAttrs.CollectionOfAttributes(FTG_SLAB)) Then
            CopyValuesToSOFromItem oAttrs.CollectionOfAttributes(FTG_SLAB), oDefAttrs.CollectionOfAttributes(FTG_SLAB)
        End If
    
        If Not IsSOOverridden(oAttrs.CollectionOfAttributes("IJUASPSFtgSlabDim")) Then
            CopyValuesToSOFromItem oAttrs.CollectionOfAttributes("IJUASPSFtgSlabDim"), oDefAttrs.CollectionOfAttributes("IJUASPSFtgSlabDim")
        End If
        
        
        'Grout Material
        Material = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_MATERIAL).Value
        Grade = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_GRADE).Value
        Set iMaterial = GetMaterialObject(Material, Grade)
        If Not iMaterial Is Nothing Then
            density = iMaterial.density
        Else
            density = 2400 'approx density of concrete
        End If
        
        'determine what shape it is to calculate the volume
        Dim Isneeded As Boolean
        Isneeded = oAttrs.CollectionOfAttributes("IJUASPSPierAndSlabFooting").Item(WITH_GROUT_PAD).Value
        If Isneeded Then
            Dim GroutHeight As Double
            Shape = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD).Item(GROUT_SHAPE).Value
            CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, Shape, PRISMATIC_FOOTING_SHAPES, 121
            
            length = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_LENGTH).Value
            width = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_WIDTH).Value
            GroutHeight = oAttrs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_HEIGHT).Value
            
            If Shape = 2 Then
                'If its rectangular volume is
                Volume = length * width * GroutHeight
                dSurfArea = (2 * length * width) + (2 * length * GroutHeight) + (2 * width * GroutHeight)
            ElseIf Shape = 3 Then
                'If its Circular volume is
                Volume = (3.14159) * width ^ 2 * GroutHeight / 4
                dSurfArea = (2 * 3.14159 * width ^ 2 / 4) + (2 * 3.14159 * width / 2 * GroutHeight)
            End If
            GroutWt = (Volume * density)
            
            TotVol = TotVol + Volume
            TotSurfArea = TotSurfArea + dSurfArea
            AccumCG.z = AccumCG.z + (Matrix.IndexValue(14) - GroutHeight / 2 - AccumCG.z) * Volume / TotVol
        End If
            
        'Pier Material
        Material = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_MATERIAL).Value
        Grade = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_GRADE).Value
        Set iMaterial = GetMaterialObject(Material, Grade)
        If Not iMaterial Is Nothing Then
            density = iMaterial.density
        Else
            density = 2400 'approx density of concrete
        End If
    
        'determine what shape it is to calculate the volume
        Dim PierHeight As Double
        Shape = oAttrs.CollectionOfAttributes(FTG_PIER).Item(PIER_SHAPE).Value
        CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, Shape, PRISMATIC_FOOTING_SHAPES, 124
    
        length = oAttrs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_LENGTH).Value
        width = oAttrs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_WIDTH).Value
        PierHeight = oAttrs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_HEIGHT).Value
    
        If Shape = 2 Then
            'If its rectangular volume is
            Volume = length * width * PierHeight
            dSurfArea = (2 * length * width) + (2 * length * PierHeight) + (2 * width * PierHeight)
        ElseIf Shape = 3 Then
            'If its Circular volume is
            Volume = (3.14159) * width ^ 2 * PierHeight / 4
            dSurfArea = (2 * 3.14159 * width ^ 2 / 4) + (2 * 3.14159 * width / 2 * PierHeight)
        End If
        PierWt = (Volume * density)
        
        TotVol = TotVol + Volume
        TotSurfArea = TotSurfArea + dSurfArea
        AccumCG.z = AccumCG.z + (Matrix.IndexValue(14) - (GroutHeight + (PierHeight) / 2) - AccumCG.z) * Volume / TotVol
       
        
        'Slab Material
        Material = oAttrs.CollectionOfAttributes(FTG_SLAB).Item(SLAB_MATERIAL).Value
        Grade = oAttrs.CollectionOfAttributes(FTG_SLAB).Item(SLAB_GRADE).Value
        Set iMaterial = GetMaterialObject(Material, Grade)
        If Not iMaterial Is Nothing Then
            density = iMaterial.density
        Else
            density = 2400 'approx density of concrete
        End If
    
        'determine what shape it is to calculate the volume
        Dim SlabHeight As Double
        Shape = oAttrs.CollectionOfAttributes(FTG_SLAB).Item(SLAB_SHAPE).Value
        CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, Shape, PRISMATIC_FOOTING_SHAPES, 127
    
        length = oAttrs.CollectionOfAttributes("IJUASPSFtgSlabDim").Item(SLAB_LENGTH).Value
        width = oAttrs.CollectionOfAttributes("IJUASPSFtgSlabDim").Item(SLAB_WIDTH).Value
        SlabHeight = oAttrs.CollectionOfAttributes("IJUASPSFtgSlabDim").Item(SLAB_HEIGHT).Value
    
        If Shape = 2 Then
            'If its rectangular volume is
            Volume = length * width * SlabHeight
            dSurfArea = (2 * length * width) + (2 * length * GroutHeight) + (2 * width * SlabHeight)
        ElseIf Shape = 3 Then
            'If its Circular volume is
            Volume = (3.14159) * width ^ 2 * SlabHeight / 4
            dSurfArea = (2 * 3.14159 * width ^ 2 / 4) + (2 * 3.14159 * width / 2 * SlabHeight)
        End If
        SlabWt = (Volume * density)
        
        TotVol = TotVol + Volume
        TotSurfArea = TotSurfArea + dSurfArea
        AccumCG.z = AccumCG.z + (Matrix.IndexValue(14) - (GroutHeight + PierHeight + (SlabHeight) / 2) - AccumCG.z) * Volume / TotVol
       
        oAttrs.CollectionOfAttributes(IGENERIC_VOLUME).Item(ATTR_VOLUME).Value = TotVol
        oAttrs.CollectionOfAttributes(ISURFACE_AREA).Item(SURFACE_AREA).Value = TotSurfArea
        
        Set iWCG = oSmartOcc
          
        ' The following put property values was originally SetWCG on the IJWeightCG interface
        '   (which is reserved for setting user defined properties), hence changed to put values
        '   on a new interface
        iWCG.Weight = (GroutWt + PierWt + SlabWt)
        iWCG.CGx = AccumCG.x
        iWCG.CGy = AccumCG.y
        iWCG.CGz = AccumCG.z
    End If

Exit Sub
ErrHandler:
If Err.Description = "Undefined Value" Then
    Err.Raise E_FAIL
Else
    HandleError MODULE, METHOD
End If
End Sub
Public Sub CMEvaluateCAOVolume(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAOVolume"
On Error GoTo ErrHandler


Exit Sub
ErrHandler:  HandleError MODULE, METHOD
End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As SP3DStructInterfaces.IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrorHandler

    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_FTGMACROS_ERROR, "ERROR")
    
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
    
    ' when we change an attribute, we set the AttributeDescriptor_Changed flag
    ' This flag is supposed to be cleared on the client side after updating GOPC
    ' with the changes
    '
    ' We also set the AttributeDescriptor_ChangeAtCommit flag; this flag remains
    ' once set, to give us an idea of the attribute set that changed in this transaction
    
    pAttrToChange.AttrValue = varNewAttrValue
    If (pAttrToChange.attrName = PIER_SIZING_RULE) Then
        If (varNewAttrValue <> 3) Then 'User defined option for sizing rule
            'gray out the pier length and width on the GOPC
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = PIER_LENGTH) Or (pAttrDescr.attrName = PIER_WIDTH)) Then
                    If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                    Else
                        pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                    End If
                End If
            Next
        Else
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = PIER_LENGTH) Or (pAttrDescr.attrName = PIER_WIDTH)) Then
                    If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                        NonStateRO = Not (AttributeDescriptor_ReadOnly)
                        pAttrDescr.AttrState = pAttrDescr.AttrState And NonStateRO
                      End If
                End If
            Next
        End If
    ElseIf (pAttrToChange.attrName = SLAB_SIZING_RULE) Then
        If (varNewAttrValue <> 3) Then 'User defined option for sizing rule
            'gray out the pier length and width on the GOPC
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = SLAB_LENGTH) Or (pAttrDescr.attrName = SLAB_WIDTH)) Then
                    If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                    Else
                        pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                    End If
                End If
            Next
        Else
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = SLAB_LENGTH) Or (pAttrDescr.attrName = SLAB_WIDTH)) Then
                    If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                        NonStateRO = Not (AttributeDescriptor_ReadOnly)
                        pAttrDescr.AttrState = pAttrDescr.AttrState And NonStateRO
                      End If
                End If
            Next
        End If
    ElseIf (pAttrToChange.attrName = GROUT_SIZE_RULE) Then
        If (varNewAttrValue <> 3) Then 'User defined option for sizing rule
            'gray out the pier length and width on the GOPC
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = GROUT_LENGTH) Or (pAttrDescr.attrName = GROUT_WIDTH)) Then
                    If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                    Else
                        pAttrDescr.AttrState = pAttrDescr.AttrState Or AttributeDescriptor_ReadOnly
                    End If
                End If
            Next
        Else
            Set pColl = CollAllDisplayedValues
            For i = 1 To pColl.count
                Set pAttrDescr = pColl.Item(i)
                If ((pAttrDescr.attrName = GROUT_LENGTH) Or (pAttrDescr.attrName = GROUT_WIDTH)) Then
                    If (pAttrDescr.AttrState And AttributeDescriptor_ReadOnly) Then
                        NonStateRO = Not (AttributeDescriptor_ReadOnly)
                        pAttrDescr.AttrState = pAttrDescr.AttrState And NonStateRO
                      End If
                End If
            Next
        End If
    End If
    IJUserAttributeMgmt_OnAttributeChange = vbNullString
    
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreCommit"
On Error GoTo ErrorHandler

    IJUserAttributeMgmt_OnPreCommit = m_oLocalizer.GetString(IDS_FTGMACROS_ERROR, "ERROR")

    IJUserAttributeMgmt_OnPreCommit = vbNullString
    
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As SP3DStructInterfaces.IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrorHandler
    IJUserAttributeMgmt_OnPreLoad = m_oLocalizer.GetString(IDS_FTGMACROS_ERROR, "ERROR")
    bOnPreLoad = True ' optimization to avoid value validation in OnAttrChange
    
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    Dim attrName As String
    Dim ErrStr As String
    
    Set pAttrColl = CollAllDisplayedValues
    For i = 1 To pAttrColl.count
        Set pAttrDescr = pAttrColl.Item(i)
            ErrStr = IJUserAttributeMgmt_OnAttributeChange(pIJDAttrs, CollAllDisplayedValues, pAttrDescr, pAttrDescr.AttrValue)
            If Len(ErrStr) > 0 Then
                bOnPreLoad = False
                Exit Function
            End If
    Next
    
    bOnPreLoad = False

    IJUserAttributeMgmt_OnPreLoad = vbNullString
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
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

Private Sub ISPSFootingDefServices_GetInputs(ByVal FtgObject As Object, _
        ByVal supported As IJElements, ByVal supporting As IJElements)
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
        supported.Add oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    End If
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
'Works For Single Type Footing
'Create ReColl, may create more then one RefColl to support combined footing inputs
'where we group a column by more than one interface
'Could get column and look for a base-plate here too..
'Add suppd, suppg to RefColl by watched interface, either IJFullObject or Member specific?
Private Sub ISPSFootingDefServices_SetInputs(ByVal FtgObject As Object, ByVal FtgDefinitionObject As Object, ByVal supported As SPSFootings.IJElements, ByVal supporting As SPSFootings.IJElements)
    Const METHOD = "ISPSFootingDefServices_SetInputs"
    
    Dim strUserType As String
    Dim strSCName As String
    Dim oSmartItem As IJSmartItem
    Dim oSmartClass As IJSmartClass
    Dim oUserType As IJDUserType
    Dim oSmartOcc As IJSmartOccurrence
    Dim oFtgFactory As SPSFootingFactory
    Dim i As Integer
    Dim oReferencesCollection As IMSSymbolEntities.IJDReferencesCollection

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
    
    If supported.count > 0 Then
        If TypeOf supported.Item(1) Is ISPSMemberSystem Then ' only add member systems for single
            For i = 1 To supported.count
                oReferencesCollection.IJDEditJDArgument.SetEntity i, supported.Item(i), ISPSMemberSystemSuppingNotify1, "MembSysSuppingNotify1RC_DEST"
            Next i
        End If
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
'                    UserAttributeMgmt_Validate = "Symbol CMCheck Failed"
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
        If oAttrObj.name = PIER_ROTATION_ANGLE Or oAttrObj.name = SLAB_ROTATION_ANGLE Or oAttrObj.name = GROUT_ROTATION_ANGLE Or oAttrObj.name = PIER_SIZE_INC Or oAttrObj.name = SLAB_SIZE_INC Then
        Else
            If oAttrObj.name = sAttributeName Then
                Select Case oAttrObj.Type
                    Case DOUBLE_VALUE
                            If (varAttributeValue <= 0#) Then
                                UserAttributeMgmt_Validate = m_oLocalizer.GetString(IDS_FTGMACROS_NEGATIVE_ATTRIBVAL, "Negative Attribute Value")
                                Set oAttributeMetaData = Nothing
                                Exit Function
                            End If
                End Select
            End If
        End If
    Next
    
    UserAttributeMgmt_Validate = vbNullString
    Set oAttributeMetaData = Nothing
Exit Function
ErrorHandler:  HandleError MODULE, METHOD

End Function


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

Private Sub ISPSTransformHelper_Transform(ByVal Trans4x4 As SP3DStructInterfaces.IJDT4x4, ByVal ObjectToTransform As Object)
Const METHOD = "ISPSTransformHelper_Transform"
On Error GoTo ErrHandler
Dim strDebug As String
Dim bIgnoreOrientation As Boolean

    'if the footing is placed by point then update the rotaion. but if placed by member then update
    'only if orientation is global
    bIgnoreOrientation = Not IsFootingPlacedByMember(ObjectToTransform)

    strDebug = "update grout's angle"
    UpdateComponentRotationAngles ObjectToTransform, FTG_GROUT_PAD, GROUT_ROTATION_ANGLE, GROUT_ORIENTATION, Trans4x4, bIgnoreOrientation
    strDebug = "update Piers's angle"
    UpdateComponentRotationAngles ObjectToTransform, FTG_PIER, PIER_ROTATION_ANGLE, PIER_ORIENTATION, Trans4x4, bIgnoreOrientation
    strDebug = "update slab's angle"
    UpdateComponentRotationAngles ObjectToTransform, FTG_SLAB, SLAB_ROTATION_ANGLE, SLAB_ORIENTATION, Trans4x4, bIgnoreOrientation
    
    Exit Sub
ErrHandler:
Err.Raise ReportError(Err, strSourceFile, METHOD, strDebug).Number

End Sub

