VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MergedPierAsmDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2006, Intergraph Corporation. All rights reserved.
'
'File
'    MergedPierAsmDef.cls
'
'Author
'       Nov 2004        MD
'
'Description
'
'Notes
'
'History:
'   09-May-06   JMS     DI#97751 - All lines are constructed through the
'                       geometry factory instead of being "new"ed
'   24-Jun-2006 JMS     DI#60069 - Changes to allow editing of the weight and CG values
'                       changed call to SetWCG to call new interface to put weight and CG
'                       values since SetWCG is reservered for setting user defined values
'                       when the values here are the computed values
'  07-Jul-2006 JMS      TR#101063 - Tolerate the nonexistence of IJWCGValueOrigin interface
'  06-Mar-07   RS & SS  CR#41094 - Changes for placing a footing in space and
'                       dis/reconnect to members
'
'*******************************************************************
Option Explicit

Private Const MODULE = "MergedPierAsmDef"

Private Const INTERFACE_WCGValueOrigin As String = "IJWCGValueOrigin"
Private Const PROPERTY_DryWCGOrigin As String = "DryWCGOrigin"
Private Enum enumWeightCGDerivation
    WEIGHTCG_Computed = 2
    WEIGHTCG_UserDefined = 4
End Enum

Private Const CONST_ItemProgId As String = "SPSFootingMacros.MergedPierAsmDef"
Private Const strSourceFile As String = "MergedPierAsmDef.cls"
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
     pAD.UserTypeClsid = "{C702A190-C0B7-49E9-AF90-CBE8F38F9C5D}" 'CPUAMergedPierAsm
     pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAsm"
     pAD.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructAsm"
     pAD.SetCMSetInputs -1, -1
     pAD.SetCMRemoveInputs -1, -1
     pAD.SetCMAdaptClone imsCOOKIE_ID_USS_LIB, "CMAdaptCloneAsm"
     pAD.SetCMIdentifyClone imsCOOKIE_ID_USS_LIB, "CMIdentifyCloneAsm"
     pAD.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateAsm"
     
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
      
     Set pMemberDescription = pMemberDescriptions.AddMember("Combined Footing Grout", 1, "CMConstructGrout", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalGrout"
     pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputGrout"
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructGrout"
     pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseGrout"
     pMemberDescription.SetCMCount imsCOOKIE_ID_USS_LIB, "CMGroutCount"
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateGrout"
     pMemberDescription.SetCMIdentifyClone imsCOOKIE_ID_USS_LIB, "CMIdentifyCloneGrout"
     pMemberDescription.SetCMAdaptClone imsCOOKIE_ID_USS_LIB, "CMAdaptCloneGrout"
'
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty GROUT_PROPERTIES, 1, IJDAttributes, "CMEvaluateGrout", imsCOOKIE_ID_USS_LIB
     pPropertyDescriptions.AddProperty GROUT_MATERIALPROP, 2, IJStructMaterial, "CMEvaluateGroutMaterial", imsCOOKIE_ID_USS_LIB
     pPropertyDescriptions.AddProperty GROUT_GEOMETRY, 3, IJGeometry, "CMEvaluateGroutGeometry", imsCOOKIE_ID_USS_LIB
          
     Set pMemberDescription = pMemberDescriptions.AddMember("Combined Footing Pier", 2, "CMConstructPier", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputPier"
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructPier"
     pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleasePier"
     pMemberDescription.SetCMAdaptClone imsCOOKIE_ID_USS_LIB, "CMAdaptClonePier"
     pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalPier"
     pMemberDescription.SetCMCount imsCOOKIE_ID_USS_LIB, "CMPierCount"
     pMemberDescription.SetCMIdentifyClone imsCOOKIE_ID_USS_LIB, "CMIdentifyClonePier"
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigratePier"

     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty PIER_PROPERTIES, 1, IJDAttributes, "CMEvaluatePier", imsCOOKIE_ID_USS_LIB
     pPropertyDescriptions.AddProperty PIER_MATERIAL_PROP, 2, IJStructMaterial, "CMEvaluatePierMaterial", imsCOOKIE_ID_USS_LIB
     pPropertyDescriptions.AddProperty PIER_GEOMETRY, 3, IJGeometry, "CMEvaluatePierGeometry", imsCOOKIE_ID_USS_LIB
     
     Set pMemberDescription = pMemberDescriptions.AddMember("Control Point", 3, "CMConstructControlPoint", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputControlPoint"
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructControlPoint"
     pMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseControlPoint"
     pMemberDescription.SetCMAdaptClone imsCOOKIE_ID_USS_LIB, "CMAdaptCloneControlPoint"
     pMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalControlPoint"
     pMemberDescription.SetCMCount imsCOOKIE_ID_USS_LIB, "CMControlPointCount"
     pMemberDescription.SetCMIdentifyClone imsCOOKIE_ID_USS_LIB, "CMIdentifyCloneControlPoint"
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateControlPoint"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "Control Point Properties", 1, IJDAttributes, "CMComputeControlPoints", imsCOOKIE_ID_USS_LIB
     pPropertyDescriptions.AddProperty "Control Point Geometry Properties", 2, IJGeometry, "CMEvaluateControlPointsGeometry", imsCOOKIE_ID_USS_LIB
     
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
  
     Set pFact = Nothing
     Set pDefinition = Nothing
     
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

    bIsNeeded = False
    If pMemberDescription Is Nothing Then
        Exit Sub
    ElseIf pMemberDescription.CAO Is Nothing Then
        Exit Sub
    End If
    
    'first check whether input is valid or not
    Dim strError As String
    Dim bosFlag As Boolean
    bosFlag = True ' if set to true then validate function will check BOS levels of all members
    strError = ValidateInputMembersCombFtg(pMemberDescription, bosFlag)

    If strError <> vbNullString Then
        Err.Raise E_FAIL, "CMConditionalGrout", strError
    End If

    ''''''''''''''''''''''''''''''''''''
    
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
     
    Set oSmartOcc = Nothing
    Set oAttribs = Nothing
    Set OccAttribs = Nothing
    Set oAttr = Nothing
    
    Exit Sub
    
ErrorHandler:
Err.Raise ReportError(Err, strSourceFile, METHOD, strError).Number
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

     Call EvaluateGrout(pPropertyDescriptions, pObject)
    
     Exit Sub
     
ErrorHandler:     HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGroutMaterial(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateGroutMaterial"
On Error GoTo ErrHandler

    Dim MemberObj As IJDMemberObjects
    Dim oAttrs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Dim iMaterial As IJDMaterial
    Set MemberObj = pPropertyDescriptions.CAO
    Set oSmartOcc = MemberObj.ItemByDispid(1, pPropertyDescriptions.index)
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
                    
     Set MemberObj = Nothing
     Set oSmartOcc = Nothing
     Set oAttrs = Nothing
     Set oStructMatl = Nothing
     
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

    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pMemberDesc.object
    oSmartOcc.RootSelection = PierComp
    
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
     
     Dim oReferencesCollection0fCAO  As IMSSymbolEntities.IJDReferencesCollection
     Dim oRefColl  As IMSSymbolEntities.IJDReferencesCollection
     Dim oRefColl2  As IMSSymbolEntities.IJDReferencesCollection
     Dim GroutHt As Double
     Dim PierEdgeClearance As Double
     Dim xmid As Double, ymid As Double, z As Double
     Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
     Dim pSymbol As IJDSymbol
     Dim pOcc As IJDOccurrence
     Dim supported As IJElements
     Dim supporting As IJElements
     
     Set pSymbol = pObject
     On Error Resume Next
     Set pOcc = pSymbol
     If pOcc Is Nothing Then
          Exit Sub
     End If
     On Error GoTo ErrorHandler

     GroutHt = GetGroutHt(pPropertyDescriptions)
    
     Set oRefColl = GetRefCollection(pPropertyDescriptions.CAO)
     Set oReferencesCollection0fCAO = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
     Set oRefColl2 = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
     Set oRefColl = Nothing
     
     Set supported = GetUpdatedRefColl(oReferencesCollection0fCAO)
     Set supporting = GetUpdatedRefColl(oRefColl2)
          
     Dim bByPlaneOrDatum As Boolean
     bByPlaneOrDatum = False
     
     Dim oPlane As IJPlane
     If supporting.count > 0 Then
        Set oPlane = supporting.Item(1)
     End If
     
     Dim Normx As Double, Normy As Double, Normz As Double
     Dim Rtx As Double, Rty As Double, Rtz As Double
     Dim TotHt As Double
     Dim oAttr As IJDAttributes

    If Not oPlane Is Nothing Then
        oPlane.GetNormal Normx, Normy, Normz
        oPlane.GetRootPoint Rtx, Rty, Rtz
        Set oPlane = Nothing
        bByPlaneOrDatum = True
    Else
        'Err.Clear - commenting as this is unnecessary
        Dim bUseDatum As Boolean
        Dim oCollProxy As IJDAttributesCol
        Set oAttr = pPropertyDescriptions.CAO
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

     Set oRefColl2 = Nothing
     
     Dim PierRotationAngle As Double
     Dim PierSizingRule As Long, PierShape As Long
     Dim oPierAttribs As IJDAttributes
     Dim MemberObj As IJDMemberObjects
     
     Set MemberObj = pPropertyDescriptions.CAO
     Dim oSmartOcc As IJSmartOccurrence
     Set oSmartOcc = MemberObj.ItemByDispid(2, pPropertyDescriptions.index)
     Set oPierAttribs = oSmartOcc  'Set oPierAttribs = pObject

     PierRotationAngle = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_ROTATION_ANGLE).Value
     PierSizingRule = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_SIZING_RULE).Value
     PierShape = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_SHAPE).Value
     
     Call GetColumnsRangeAboutRotatedZ(supported, PierRotationAngle, xmin, ymin, xmax, ymax, z, xmid, ymid)
     
     Dim Matrix As IJDT4x4
     Set Matrix = New DT4x4
         
     Matrix.LoadIdentity
     Matrix.IndexValue(12) = xmid
     Matrix.IndexValue(13) = ymid
     Matrix.IndexValue(14) = z - GroutHt

     Dim Vec As DVector
     Set Vec = New DVector
     Vec.Set 0, 0, 1
     
     CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, PierShape, PRISMATIC_FOOTING_SHAPES, 124
     CheckForUndefinedValueAndRaiseError pPropertyDescriptions.CAO, PierSizingRule, FOOTING_COMP_SIZE_RULE, 125

     Matrix.Rotate PierRotationAngle, Vec

     If PierSizingRule = 1 Or PierSizingRule = 2 Then
        PierEdgeClearance = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_EDGE_CLEARANCE).Value
        
        'get appropriate grout width length componenets to be considered for deciding pier length & width
        Dim GroutWidth As Double, GroutLength As Double
        GetGroutWidthLengthComponents pPropertyDescriptions, xmin, xmax, ymin, ymax, PierRotationAngle, GroutWidth, GroutLength
        
        Dim WidthComp As Double, LenComp As Double
        Dim tmpAng As Double
        tmpAng = Abs(PierRotationAngle)
        If tmpAng > PI / 2 Then tmpAng = Abs(PI - tmpAng)
        
        WidthComp = ((ymax - ymin) + GroutWidth)
        LenComp = ((xmax - xmin) + GroutLength)
        
        ' Pier size = range size+ clearance+  appropriate grout width/length
        oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_WIDTH).Value = WidthComp + (PierEdgeClearance * 2)
        oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_LENGTH).Value = LenComp + (PierEdgeClearance * 2)
     End If

     If bByPlaneOrDatum Then
        Dim oline As Line3d
        Dim GeomFactory As New IngrGeom3D.GeometryFactory
        Dim DummyFace As New Plane3d
        Dim temp As IJElements
        Set temp = New JObjectCollection ' elements
        Set oline = GeomFactory.Lines3d.CreateByPtVectLength(Nothing, xmid, ymid, z, 0, 0, -1, 1)
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
        
        On Error GoTo ErrorHandler
        Dim strError As String
        If pt3 > z Then
            SPSToDoErrorNotify FootingToDoMsgCodelist, TDL_FTGMACROS_BOTTOMPLANE_NOTVALID_CANNOTCOMPUTE, pPropertyDescriptions.CAO, Nothing
            Err.Raise E_FAIL
        Else
            If z > Rtz Then
                TotHt = Abs(z - pt3) - GroutHt
                oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_HEIGHT).Value = TotHt
            Else
                SPSToDoErrorNotify FootingToDoMsgCodelist, TDL_FTGMACROS_BOTTOMPLANE_NOTVALID_CANNOTCOMPUTE, pPropertyDescriptions.CAO, Nothing
                Err.Raise E_FAIL
                
            End If
        End If
        
     End If
     
     pOcc.Matrix = Matrix
    
     Set oReferencesCollection0fCAO = Nothing
     Set pSymbol = Nothing
     Set Vec = Nothing
     Set oPierAttribs = Nothing
     Set pOcc = Nothing
     Set temp = Nothing
     Set DummyFace = Nothing
     Set oNewBottomSurf = Nothing
     Set point = Nothing
     Set MemberObj = Nothing
     Set oSmartOcc = Nothing
     Set supporting = Nothing
     Set supported = Nothing
     
    On Error GoTo ErrorHandler
    If strError <> vbNullString Then
        Err.Raise E_FAIL, "CMEvaluatePier", strError
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

Public Sub CMEvaluatePierMaterial(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluatePierMaterial"
On Error GoTo ErrHandler
    
    Dim MemberObj As IJDMemberObjects
    Dim oAttrs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    Dim iMaterial As IJDMaterial
    Set MemberObj = pPropertyDescriptions.CAO
    Set oSmartOcc = MemberObj.ItemByDispid(2, pPropertyDescriptions.index)
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

    Set MemberObj = Nothing
    Set oSmartOcc = Nothing
    Set oAttrs = Nothing
    Set oStructMatl = Nothing
    Set iMaterial = Nothing
    
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
    
ErrorHandler: HandleError MODULE, METHOD
End Sub


Public Sub CMConstructAsm(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMConstructAsm"
On Error GoTo ErrorHandler
    
    Exit Sub
    
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateCAO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAO"
On Error GoTo ErrHandler
 
    
    SetOccurrenceMatrix pObject
    
    'Need to make all the points equal on Z plane, this can be wrong when mirror copy happens
    Dim supported As SPSFootings.IJElements
    Dim supporting As SPSFootings.IJElements
    Set supported = New JObjectCollection
    Set supporting = New JObjectCollection
    Call CombineFootings_GetInputs(pObject, supported, supporting)
    ChangePointsZToSameZAsFirstPoint supported
    
   
    Exit Sub
    
ErrHandler:  HandleError MODULE, METHOD
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
            oReferencesCollection.IJDEditJDArgument.SetEntity 1, oInputObj
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
     
     Set oReferencesCollection = Nothing
     Set pIRCAsm = Nothing
     Set oInputObj = Nothing
     Set oSmartOcc = Nothing
     Set oItem = Nothing
     Set oUserType = Nothing
     
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
        If ValidateBottomZValueForCombinedFooting(pIJDAttrs, CDbl(varNewAttrValue)) = False Then
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

    Dim i As Integer, j As Integer
    Dim pMemberSystem As ISPSMemberSystem
    Dim pDesignParts As IJElements
    Dim oSmartOcc As IJSmartOccurrence
    Dim oRefColl As IJDReferencesCollection
    Dim oRefColl1 As IJDReferencesCollection
    Dim supported As IJElements
    
    Set oSmartOcc = pPartObject
    Set oRefColl = GetRefCollection(oSmartOcc)
    If oRefColl.IJDEditJDArgument.GetCount > 0 Then
        Set oRefColl1 = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
        Set supported = GetUpdatedRefColl(oRefColl1)
        
        For i = 1 To supported.count
            If Not supported.Item(i) Is Nothing Then
                If TypeOf supported.Item(i) Is ISPSMemberSystem Then
                    Set pMemberSystem = supported.Item(i)
                    Set pDesignParts = pMemberSystem.DesignParts
                    For j = 1 To pDesignParts.count
                        pIJMonUnks.Add (pDesignParts.Item(j))
                    Next j
                    Set pMemberSystem = Nothing
                    Set pDesignParts = Nothing
                End If
            End If
        Next i
    End If
    Set oSmartOcc = Nothing
    Set oRefColl = Nothing
    Set oRefColl1 = Nothing
    Set pMemberSystem = Nothing
    Set pIJMonUnks = Nothing
    Set supported = Nothing
    
    Exit Sub
    
ErrorHandler:     HandleError MODULE, METHOD
End Sub

Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = StandardGraphicEntity
End Sub

Private Sub ISPSFootingDefServices_GetInputs(ByVal FtgObject As Object, ByVal supported As SPSFootings.IJElements, ByVal supporting As SPSFootings.IJElements)
Const METHOD = "ISPSFootingDefServices_GetInputs"
On Error Resume Next
   
   Call CombineFootings_GetInputs(FtgObject, supported, supporting)
   
    Exit Sub
    
ErrorHandler:      HandleError MODULE, METHOD
End Sub

Private Sub ISPSFootingDefServices_SetInputs(ByVal FtgObject As Object, ByVal FtgDefinitionObject As Object, ByVal supported As SPSFootings.IJElements, ByVal supporting As SPSFootings.IJElements)
Const METHOD = "ISPSFootingDefServices_SetInputs"
On Error GoTo ErrorHandler

    If supported.count < 1 Then
        SPSToDoErrorNotify FootingToDoMsgCodelist, TDL_FTGMACROS_INSUFFICIENT_INPUTS, FtgObject, Nothing
        Err.Raise E_FAIL
    Else
        ChangePointsZToSameZAsFirstPoint supported
        Call CombineFootings_SetInputs(FtgObject, FtgDefinitionObject, supported, supporting)
    End If

    Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
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
'
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
    Set dInputs = Nothing
    
    Exit Function
    
ErrorHandler:  HandleError MODULE, METHOD
End Function

Public Sub CMGroutCount(ByVal pMemberDescription As IJDMemberDescription, ByRef lCount As Long)
Const METHOD = "CMGroutCount"
On Error GoTo ErrorHandler

    lCount = GetGroutCount(pMemberDescription)

    Exit Sub
    
ErrorHandler:       HandleError MODULE, METHOD
End Sub

Public Sub CMMigrateGrout(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateGrout"
On Error GoTo ErrorHandler

    Exit Sub

ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMIdentifyCloneGrout(ByVal pMemberDescription As IJDMemberDescription, ByRef pObject As Object)
Const METHOD = "CMIdentifyCloneGrout"
On Error GoTo ErrorHandler
  
    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMAdaptCloneGrout(ByVal pAD As IJDMemberDescription)
Const METHOD = "CMAdaptCloneGrout"
On Error GoTo ErrorHandler
  
    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMAdaptCloneAsm(ByVal pAD As IJDAggregatorDescription)
Const METHOD = "CMAdaptCloneAsm"
On Error GoTo ErrorHandler
  
    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMIdentifyCloneAsm(ByVal pAD As IJDAggregatorDescription, ByRef pObject As Object)
Const METHOD = "CMIdentifyCloneAsm"
On Error GoTo ErrorHandler
  
    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMMigrateAsm(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateAsm"
On Error GoTo ErrorHandler

    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMMigratePier(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigratePier"
On Error GoTo ErrorHandler

    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalPier(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
Const METHOD = "CMConditionalPier"
On Error GoTo ErrorHandler

    'first check whether evaluated pier height is going to be valid or not.
    Dim strError As String

    strError = ValidatePierHeight(pMemberDescription)

    If strError <> vbNullString Then
        Err.Raise E_FAIL, "CMConditionalPier", strError
    End If
    
    Exit Sub
ErrorHandler:
Err.Raise ReportError(Err, strSourceFile, METHOD, strError).Number
End Sub

Public Sub CMIdentifyClonePier(ByVal pMemberDescription As IJDMemberDescription, ByRef pObject As Object)
Const METHOD = "CMIdentifyClonePier"
On Error GoTo ErrorHandler
  
    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMPierCount(ByVal pMemberDescription As IJDMemberDescription, ByRef lCount As Long)
Const METHOD = "CMPierCount"
On Error GoTo ErrorHandler

    lCount = -1

    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMAdaptClonePier(ByVal pAD As IJDMemberDescription)
Const METHOD = "CMAdaptClonePier"
On Error GoTo ErrorHandler
  
    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMMigrateControlPoint(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateControlPoint"
On Error GoTo ErrorHandler
   
  Exit Sub
  
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMIdentifyCloneControlPoint(ByVal pMemberDescription As IJDMemberDescription, ByRef pObject As Object)
Const METHOD = "CMIdentifyCloneControlPoint"
On Error GoTo ErrorHandler
  
    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMControlPointCount(ByVal pMemberDescription As IJDMemberDescription, ByRef lCount As Long)
Const METHOD = "CMControlPointCount"
On Error GoTo ErrorHandler

    lCount = 8 + GetGroutCount(pMemberDescription)

    Exit Sub
    
ErrorHandler:       HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalControlPoint(ByVal pMemberDescription As IJDMemberDescription, ByRef bIsNeeded As Boolean)
Const METHOD = "CMConditionalControlPoint"
On Error GoTo ErrorHandler

    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMAdaptCloneControlPoint(ByVal pAD As IJDMemberDescription)
Const METHOD = "CMAdaptCloneControlPoint"
On Error GoTo ErrorHandler
  
    Exit Sub
    
ErrorHandler:   HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseControlPoint(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMReleaseControlPoint"
On Error GoTo ErrorHandler

    Exit Sub
    
ErrorHandler:       HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructControlPoint(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructControlPoint"
On Error GoTo ErrorHandler
  
     Call CreateContolPointRelation(pMemberDesc)

     Call GenerateNameForFooting(pMemberDesc.object)
     
    Exit Sub
    
ErrorHandler:       HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputControlPoint(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMSetInputControlPoint"
On Error GoTo ErrorHandler
     
    Exit Sub
    
ErrorHandler: HandleError MODULE, METHOD
End Sub

Public Sub CMConstructControlPoint(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructControlPoint"
On Error GoTo ErrorHandler

    Dim m_oControlPoint As IJControlPoint
     
    Set m_oControlPoint = CreateControlPoint(pResourceManager, False)
    
    m_oControlPoint.SubType = cpFoundation
    If pMemberDescription.index > 8 Then
       m_oControlPoint.Type = cpControlPoint
    Else
       m_oControlPoint.Type = cpKeyPoint
    End If
    Call SetControlFlags(CTL_FLAG_NO_DISPLAY_IN_NON_GRAPHIC_VIEW, m_oControlPoint)
   
    Set pObj = m_oControlPoint
   
    Set m_oControlPoint = Nothing

    Exit Sub
    
ErrorHandler:       HandleError MODULE, METHOD
End Sub

Public Sub CMComputeControlPoints(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeControlPoints"
On Error GoTo ErrorHandler

    Dim oPierAttribs As IJDAttributes
    Dim MemberObj As IJDMemberObjects
    Dim oSmartOcc As IJSmartOccurrence
    Dim PierWidth As Double, PierLength As Double, PierHeight As Double, PierRotationAngle As Double

    Dim m_oControlPoint As IJControlPoint
    Dim ControlPt As IJPoint
    Dim x As Double, y As Double, z As Double
    Dim xmid As Double, ymid As Double
    Dim oReferencesCollection0fCAO  As IMSSymbolEntities.IJDReferencesCollection
    Dim oRefColl  As IMSSymbolEntities.IJDReferencesCollection
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double
    Dim dMembAng As Double
    Dim sSecName As String, sRefStd As String
    Dim supported As IJElements

    Set m_oControlPoint = pObject
    Set MemberObj = pPropertyDescriptions.CAO
    Set oSmartOcc = MemberObj.ItemByDispid(2)
    Set oPierAttribs = oSmartOcc
    PierWidth = oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_WIDTH).Value
    PierLength = oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_LENGTH).Value
    PierHeight = oPierAttribs.CollectionOfAttributes(FTG_PIER_DIM).Item(PIER_HEIGHT).Value
    PierRotationAngle = oPierAttribs.CollectionOfAttributes(FTG_PIER).Item(PIER_ROTATION_ANGLE).Value
    
    Set oRefColl = GetRefCollection(pPropertyDescriptions.CAO)
    Set oReferencesCollection0fCAO = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
    Set supported = GetUpdatedRefColl(oReferencesCollection0fCAO)
       
    Call GetColumnsRangeAboutRotatedZ(supported, PierRotationAngle, xmin, ymin, xmax, ymax, z, xmid, ymid)
    'first 8 points represent the corners of the pier. Four at the top and four at the bottom. Rest of the points represnt column
    'locations
    If pPropertyDescriptions.index > 8 Then
        GetMemberBottomCenterProperties supported, x, y, z, dMembAng, sSecName, sRefStd, pPropertyDescriptions.index - 8
    Else
        Select Case pPropertyDescriptions.index
            Case 1, 2, 3, 4
                z = z - GetGroutHt(pPropertyDescriptions)
            Case 5, 6, 7, 8
                z = z - GetGroutHt(pPropertyDescriptions) - PierHeight
        End Select
        
        'the pier may be rotated relative to the global. So need to compute corner points based on pier length/width and rotation
        
        Dim lngQuadrant As Long
        
       
        
        '8 points fall into 4 quadrants
        lngQuadrant = pPropertyDescriptions.index Mod 4
        
        If lngQuadrant = 0 Then lngQuadrant = 4
        
        Dim xLoc As Double, yLoc As Double ' along the pier's local coordinate system
        Select Case lngQuadrant
        
        Case 1
             xLoc = PierLength / 2
             yLoc = PierWidth / 2

        Case 2
             xLoc = -PierLength / 2
             yLoc = PierWidth / 2
        Case 3
             xLoc = -PierLength / 2
             yLoc = -PierWidth / 2
        Case 4
             xLoc = PierLength / 2
             yLoc = -PierWidth / 2
        End Select
        x = xmid + (xLoc * Cos(PierRotationAngle) - yLoc * Sin(PierRotationAngle)) ' rotate to global
        y = ymid + (xLoc * Sin(PierRotationAngle) + yLoc * Cos(PierRotationAngle)) 'rotate to global
        
    End If

    Set ControlPt = New Point3d
    Set ControlPt = m_oControlPoint
     
    ControlPt.SetPoint x, y, z
    
    Set oRefColl = Nothing
    Set oReferencesCollection0fCAO = Nothing
    Set MemberObj = Nothing
    Set m_oControlPoint = Nothing
    Set oSmartOcc = Nothing
    Set oPierAttribs = Nothing
    Set ControlPt = Nothing
    Set supported = Nothing
     
    Exit Sub
    
ErrorHandler:       HandleError MODULE, METHOD
End Sub

Private Function ValidatePierHeight(ByVal pMemberDescription As IJDMemberDescription) As String
Const METHOD = "ValidatePierHeight"
On Error GoTo ErrorHandler

     Dim oReferencesCollection0fCAO  As IMSSymbolEntities.IJDReferencesCollection
     Dim oRefColl  As IMSSymbolEntities.IJDReferencesCollection
     Dim oRefColl2  As IMSSymbolEntities.IJDReferencesCollection
     Dim GroutHt As Double
     Dim x As Double, y As Double, z As Double
     Dim MemberObj As IJDMemberObjects
     Dim oAttr As IJDAttributes
     Dim bIsNeeded As Boolean
     Dim oSmartOcc As IJSmartOccurrence
     Dim oGroutAttribs As IJDAttributes
     
     ValidatePierHeight = vbNullString

     ' get grout height first
     Set MemberObj = pMemberDescription.CAO
     Set oAttr = MemberObj
     bIsNeeded = oAttr.CollectionOfAttributes(PIER_FOOTING_ASM).Item(WITH_GROUT_PAD).Value

     GroutHt = 0#
     If bIsNeeded Then 'check for existence of the GroutPad
        Set oSmartOcc = MemberObj.ItemByDispid(1, 1) ' get first grout height. It will be same for all in case of mergedpier & combined slab asm
        Set oGroutAttribs = oSmartOcc
        GroutHt = oGroutAttribs.CollectionOfAttributes(FTG_GROUT_PAD_DIM).Item(GROUT_HEIGHT).Value
     End If
     
     Set MemberObj = Nothing
     Set oGroutAttribs = Nothing
     Set oAttr = Nothing
     Set oSmartOcc = Nothing

     Dim supported As IJElements
     Dim supporting As IJElements
     
     Set oRefColl = GetRefCollection(pMemberDescription.CAO)
     Set oReferencesCollection0fCAO = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
     Set oRefColl2 = oRefColl.IJDEditJDArgument.GetEntityByIndex(2)
     Set oRefColl = Nothing
     Set supported = GetUpdatedRefColl(oReferencesCollection0fCAO)
     Set supporting = GetUpdatedRefColl(oRefColl2)
          
     Dim IsPlane As Boolean
     IsPlane = False

     Dim oPlane As IJPlane
     If supporting.count > 0 Then
        Set oPlane = supporting.Item(1)
     End If
    
    Dim Normx As Double, Normy As Double, Normz As Double
    Dim Rtx As Double, Rty As Double, Rtz As Double
    Dim TotHt As Double
    Dim bByPlaneOrDatum As Boolean
    bByPlaneOrDatum = False


    If Not oPlane Is Nothing Then
        oPlane.GetNormal Normx, Normy, Normz
        oPlane.GetRootPoint Rtx, Rty, Rtz
        Set oPlane = Nothing
        bByPlaneOrDatum = True
    Else
        'Err.Clear - commenting as this is unnecessary
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


     Set oRefColl2 = Nothing
     
     Call GetColumnsRange(supported, x, y, x, y, z)
       
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
        
        If (z - pt3 - GroutHt) < 0# Then
            SPSToDoErrorNotify FootingToDoMsgCodelist, TDL_FTGMACROS_PIERHEIGHT_INVALID_CHECK, pMemberDescription.CAO, Nothing
            Err.Raise E_FAIL
        End If
        
     End If

     Set oReferencesCollection0fCAO = Nothing
     Set temp = Nothing
     Set DummyFace = Nothing
     Set oNewBottomSurf = Nothing
     Set point = Nothing
     Set supported = Nothing
     Set supporting = Nothing
     
     Exit Function
     
ErrorHandler:       HandleError MODULE, METHOD
End Function

Public Sub CMEvaluateControlPointsGeometry(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateControlPointsGeometry"
On Error GoTo ErrorHandler
 
    Call CMComputeControlPoints(pPropertyDescriptions, pObject)
 
    Exit Sub
    
ErrorHandler:       HandleError MODULE, METHOD
End Sub

Private Property Get ISPSFoundationInputHelper_ValidateObjects(ByVal inputSupported As Object, ByVal inputSupporting As Object, ByVal SupportedObjList As SP3DStructInterfaces.IJElements, ByVal ObjsinSelectSet As SP3DStructInterfaces.IJElements) As SP3DStructInterfaces.SPSFoundationInputHelperStatus
Const MT = "ISPSFoundationInputHelper_ValidateObjects"
On Error GoTo ErrorHandler
    
    If inputSupporting Is Nothing Then
        ISPSFoundationInputHelper_ValidateObjects = CommonCombinedCheckMember(inputSupported, SupportedObjList)
    Else
        ISPSFoundationInputHelper_ValidateObjects = CommonCombinedCheckBottomPlane(inputSupporting, SupportedObjList, ObjsinSelectSet)
    End If
    Exit Property
    
ErrorHandler:   HandleError MODULE, MT
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
    Dim oFootingComponents As IJDMemberObjects
    Dim oPier As SPSFootingComponent
    Dim oGrout As SPSFootingComponent
    Dim oRefCollColumns  As IJDReferencesCollection
    Dim oRefColl  As IJDReferencesCollection
    Dim colColumns As IJElements
    Dim lngGrout As Long
    Dim idx As Long
    Dim strDebug As String
    Dim bIgnoreOrientation As Boolean
        
    Set oFootingComponents = ObjectToTransform
    
    'update grout's rotation
    strDebug = "updating grout's rotation"
    Set oRefColl = GetRefCollection(ObjectToTransform)
    Set oRefCollColumns = oRefColl.IJDEditJDArgument.GetEntityByIndex(1)
     
    Set colColumns = GetUpdatedRefColl(oRefCollColumns)
     
    lngGrout = colColumns.count
    
    'if the footing is placed by point then update the rotaion. but if placed by member then update
    'only if orientation is global
    bIgnoreOrientation = Not IsFootingPlacedByMember(ObjectToTransform)
    
    
    For idx = 1 To lngGrout
        On Error Resume Next
        Set oGrout = oFootingComponents.ItemByDispid(1, idx)
        On Error GoTo ErrHandler
        
        If Not oGrout Is Nothing Then
            UpdateRotationAngle oGrout, FTG_GROUT_PAD, GROUT_ROTATION_ANGLE, GROUT_ORIENTATION, Trans4x4, bIgnoreOrientation
        End If
    Next idx
    
    'update pier's rotation
    strDebug = "updating pier's rotation"
    On Error Resume Next
    Set oPier = oFootingComponents.ItemByDispid(2)
    On Error GoTo ErrHandler
    
    If Not oPier Is Nothing Then
        UpdateRotationAngle oPier, FTG_PIER, PIER_ROTATION_ANGLE, PIER_ORIENTATION, Trans4x4, True
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

