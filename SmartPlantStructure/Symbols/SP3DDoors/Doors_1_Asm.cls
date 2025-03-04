VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Doors_1_Asm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE = "Doors_1_Asm:"  'Used for error messages

Implements IJDUserSymbolServices
Implements IJFoulCheckForUSS


Public m_bDoors_1_Asm_InitializeInProgress As Boolean

Private m_oEquipCADHelper As IJEquipCADHelper

 
Public Function GetDefinitionName(ByVal definitionParameters As Variant) As String
  ' Name should be unique
    Let GetDefinitionName = "SP3DDoorsAsm.Doors_1_Asm"
End Function

Private Sub Class_Initialize()
    Set m_oEquipCADHelper = New CADServices
    m_oEquipCADHelper.ProjectName = "SP3DDoorsAsm"
    m_oEquipCADHelper.ClassName = "Doors_1_Asm"
    m_oEquipCADHelper.OccurrenceRootClass = orcEquipment
End Sub

Private Sub Class_Terminate()
    Set m_oEquipCADHelper = Nothing
End Sub

'
' IJDUserSymbolServices implementation
'
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    Dim iSymbolDefinition As IJDSymbolDefinition
    Set iSymbolDefinition = m_oEquipCADHelper.InstanciateDefinition(CodeBase, defParams, ActiveConnection)
   
    ' en mode static, la definition est rechargee a chaque fois de la .dll
    ' le CopyBackwardFlag marque la CADefinition modified depuis la SmartOccurrence CES !
    'iSymbolDefinition.MetaDataOption = igSYMBOL_STATIC_METADATA
    
        IJDUserSymbolServices_InitializeSymbolDefinition iSymbolDefinition
    Set IJDUserSymbolServices_InstanciateDefinition = iSymbolDefinition
End Function
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(iSymbolDefinition As IJDSymbolDefinition)
    Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
    Let sSourceFile = "Doors_1_Asm.bas": Let sMethod = "InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
 
    If m_bDoors_1_Asm_InitializeInProgress = True Then
        sError = "SimpleDoor_1_Asm::already in progress"
        Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
    End If
    Let m_bDoors_1_Asm_InitializeInProgress = True
    
      
    Dim pCADefinition As IJCADefinition
    Set pCADefinition = iSymbolDefinition
    Let pCADefinition.CopyBackwardFlag = igCOPY_BACKWARD_TRIM
    
    sError = "Define the aggregator"
    Call DefineAggregator(iSymbolDefinition)
    
    'No members
    sError = "Define the members"
    Dim pMemberDescriptions As IJDMemberDescriptions
    Set pMemberDescriptions = iSymbolDefinition
        pMemberDescriptions.RemoveAll
    
    Let m_bDoors_1_Asm_InitializeInProgress = False
Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    Let IJDUserSymbolServices_GetDefinitionName = GetDefinitionName(definitionParameters)
End Function
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputColl As Object, ByRef arrayOfInputs())
    Call InvokeRepresentation(sblOcc, repName, outputColl, arrayOfInputs)
End Sub
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
 ' Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
    Let IJDUserSymbolServices_EditOccurence = EditOccurence(pSymbolOccurence, transactionMgr)
End Function
'
' custom methods for aggregator
'
Public Sub CMSetInputs4SO(pAD As IJDAggregatorDescription)
    Call myCMSetInputs4SO(pAD)
End Sub
Public Sub CMMigrate4SO(pAD As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)
    'Call myCMMigrate4SO(pAD, pMigrateHelper)
End Sub
Public Sub CMIdentifyClone4SO(ByVal pAD As IJDAggregatorDescription, ByRef pObject As Object)
    'pObject should return the clone of the CAO
End Sub


Public Sub CMAdaptClone4SO(ByVal pAD As IJDAggregatorDescription)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Doors_1_Asm.bas": Let sMethod = "CMAdaptClone4SO"
On Error GoTo ErrorHandler

'Comment code since cutoutContour is now trimmed from initial set for Doors
'(CStructCutoutToolCES::TrimInitialSet)

    'As equip mating cnstr is not copied , we got cutoutContour copied by proxy without cutoutAE
    'this is not good, we just want to copy the door without cutout, so remove it here
'    Dim oCloneSmartOcc As IJSmartOccurrence
'    Set oCloneSmartOcc = pAD.CAO

'    'Only remove if mating constraint not cloned
'    Dim oClonePortDeck As IJPort 'plate support
'    Dim dEqpOffset As Double
'
'    Set oClonePortDeck = GetDoorMatingSurfaces(oCloneSmartOcc, dEqpOffset)
'
'    If oClonePortDeck Is Nothing Then
'        Dim oSblCutoutOutput As Object
'        Dim iSymbol As IJDSymbol
'        Set iSymbol = oCloneSmartOcc
'        Set oSblCutoutOutput = iSymbol.BindToOutput("CutoutContour", "CutoutCurve")
'        If oSblCutoutOutput Is Nothing Then
'           Err.Raise ReportError(Err, MODULE, METHOD, "oSblCutoutOutput is nothing").Number
'        End If
'
'        'get connected cutoutcontour if any
'        Dim oStructCutoutContour As IJDStructCutoutTool
'        Set oStructCutoutContour = CutoutCurve_GetCutoutContour(oSblCutoutOutput)
'        If Not oStructCutoutContour Is Nothing Then
'            'delete cutoutContour
'            Dim iCutoutObj As IJDObject
'            Set iCutoutObj = oStructCutoutContour
'            iCutoutObj.Remove
'        End If
'
'        'Also remove refcoll arguments (do not delete otherwise equip will be deleted)
'        Dim iRefColl As IJDReferencesCollection
'        Set iRefColl = GetRefCollFromSmartOccurrence(oCloneSmartOcc)
'        If Not iRefColl Is Nothing Then
'            iRefColl.IJDEditJDArgument.RemoveAll
'        End If
'    End If
    
 Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub


Public Sub CMEvaluateCutout(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "Doors_1_Asm.bas": Let sMethod = "CMEvaluateCutout"

Dim sMissingSupportingObjectError As String: Let sMissingSupportingObjectError = "Door/Window is no longer associated to a supporting object. Property failed to compute. Undo the deletion of the supporting object, or mate equipment to a supporting object."
On Error GoTo ErrorHandler

Dim iOppositeDeckPort As IJPort
Dim iprevConnectable As IJConnectable
Dim iEquip As IJEquipment
Dim iSmartOcc As IJSmartOccurrence
Dim iSymbol As IJDSymbol
Dim oSblCutoutOutput As Object
Dim iConnectable As IJConnectable
Dim iRefColl   As IJDReferencesCollection
Dim oSupportType    As SupportType
Dim oprevSupportType As SupportType
Dim sprevOperationProgId As String
Dim sOperationProgId As String
Dim nbRef As Long

'Get Symbol "CutoutContour" output
Set iSmartOcc = pPropertyDescriptions.CAO
Set iEquip = iSmartOcc
Set iSymbol = iSmartOcc
Set oSblCutoutOutput = iSymbol.BindToOutput("CutoutContour", "CutoutCurve")
If oSblCutoutOutput Is Nothing Then
    sError = "CMEvaluateCutout::oSblCutoutOutput is nothing"
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If

'GetCutoutContour from oSblCutoutOutput if any
Dim oprevStructCutoutContour As StructCutoutContour
Set oprevStructCutoutContour = GetDoorCutoutContour(oSblCutoutOutput)
If Not oprevStructCutoutContour Is Nothing Then
    'Retrieve Connectable from the equipment (Nothing if no cutout)
    Set iprevConnectable = GetConnectableFromCutoutContour(oprevStructCutoutContour)
    If iprevConnectable Is Nothing Then
        SPSToDoErrorNotifyEx "SP3DDoorsAsm", IDS_SP3DDOORSSAM_MISSING_SUPPORTING_OBJECT, sMissingSupportingObjectError, iSmartOcc, Nothing
        Err.Raise E_FAIL
    End If
End If

Dim iPortDeck As IJPort 'plate support
Dim bDefaultMatingCnstrExist As Boolean
Dim bReferenceExist As Boolean
Dim dEqpOffset As Double

bDefaultMatingCnstrExist = False
Set iPortDeck = GetDoorMatingSurfaces(iSmartOcc, dEqpOffset)

If Not iPortDeck Is Nothing Then
    bDefaultMatingCnstrExist = True
End If

'Get Connectable type and set corresponding progId
If Not (iprevConnectable Is Nothing) Then
    If TypeOf iprevConnectable Is ISPSSlabEntity Then
       sprevOperationProgId = "StructGeneric.StructCutoutOperationAE.1"
       oprevSupportType = SLAB
    Else 'ISPSWallPart
       sprevOperationProgId = "SP3DStructGeneric.StructCutoutOperation"
       oprevSupportType = WALL
    End If
End If

'mating constraint has been removed, remove previous cutout if any
If bDefaultMatingCnstrExist = False And Not (iprevConnectable Is Nothing) Then
    RemoveCutoutFromContour iprevConnectable, oSblCutoutOutput, sprevOperationProgId
    Exit Sub
End If

'default mating constraint exists, create or update cutout
If bDefaultMatingCnstrExist = True Then

    Dim iObject As IJDObject
    Dim iPOMUnk As IUnknown
    Set iObject = oSblCutoutOutput
    Set iPOMUnk = iObject.ResourceManager
    Dim iOperatorElts As IJElements
    Dim iStructOperationPattern As IJStructOperationPattern
    Dim iStructCutoutOperationAE As StructCutoutOperationAE
    Dim ostructcutoutcontourfactory As StructCutoutContourFactory
    Dim oStructCutoutContour As StructCutoutContour
    Dim bSymbolContourIsCutout As Boolean
    
    'Gets the object from which the port is coming
    Set iConnectable = iPortDeck.Connectable
    If iConnectable Is Nothing Then
        SPSToDoErrorNotifyEx "SP3DDoorsAsm", IDS_SP3DDOORSSAM_MISSING_SUPPORTING_OBJECT, sMissingSupportingObjectError, iSmartOcc, Nothing
        Err.Raise E_FAIL
    End If
    
    'Get Connectable type and set corresponding progId
    If TypeOf iConnectable Is ISPSSlabEntity Then
       oSupportType = SLAB
       sOperationProgId = "StructGeneric.StructCutoutOperationAE.1"
    Else 'ISPSWallPart
       oSupportType = WALL
       sOperationProgId = "SP3DStructGeneric.StructCutoutOperation"
    End If
    
    'if a cutout exists and its support is different
    'this means that SO is replacing its support, need to remove previous cutoutcontour and reference
    ' and go on
    If Not (iprevConnectable Is Nothing) Then
        If Not (iprevConnectable Is iConnectable) Then
            RemoveCutoutFromContour iprevConnectable, oSblCutoutOutput, sprevOperationProgId
            'reset iprevConnectable
            Set iprevConnectable = Nothing
        End If
    End If
    
    'Set Cutting Limit to create ModelBody (created from the planar contour in the direction of its normal)
    'MinCutOutDistance/MaxCutOutDistance represent the start/end of the cutting tool from the profile curve
    'MaxCutOutDistance is the distance between the Equipment and the support + width of support.
    'MinCutOutDistance is taken as -MaxCutOutDistance to work when support is mirrored

    
     Dim dMinCutOutDistance As Double
     Dim dPrevMinCutOutDistance As Double
     Dim dMaxCutOutDistance As Double
     Dim dprevMaxCutOutDistance As Double
     Dim bHasMinCutOutDistanceChanged As Boolean
     Dim bHasMaxCutOutDistanceChanged As Boolean
  
     bHasMinCutOutDistanceChanged = True
     bHasMaxCutOutDistanceChanged = True
     
     'dist +/- dextendedLength
     Dim dextendedLength As Double
     Dim dthickness As Double
     dextendedLength = 0.2
     dthickness = GetSupportThickness(iConnectable)
     
     'EqpOffset can be negative
     dMaxCutOutDistance = Abs(Abs(dEqpOffset) + dthickness + dextendedLength)
     dMinCutOutDistance = -dMaxCutOutDistance
    
     bSymbolContourIsCutout = False
     Set iStructOperationPattern = iConnectable
     
  
     iStructOperationPattern.GetOperationPattern sOperationProgId, iOperatorElts, iStructCutoutOperationAE
     If (Not iOperatorElts Is Nothing) Then
         For Each oStructCutoutContour In iOperatorElts
             ' If the symbol contour is not yet in the collection add it
             If oStructCutoutContour.InputContour Is oSblCutoutOutput Then
                 bSymbolContourIsCutout = True
                 'Get previous cuttingDepth
                 dprevMaxCutOutDistance = oStructCutoutContour.MaxCutOutDistance
                 dPrevMinCutOutDistance = oStructCutoutContour.MinCutOutDistance
                 If Abs(dprevMaxCutOutDistance - dMaxCutOutDistance) < 0.001 Then
                     bHasMaxCutOutDistanceChanged = False
                 End If
                 If Abs(dPrevMinCutOutDistance - dMinCutOutDistance) < 0.001 Then
                     bHasMinCutOutDistanceChanged = False
                 End If
                 Exit For
             End If
         Next
     Else
               Set iOperatorElts = New JObjectCollection
     End If
     
     If bSymbolContourIsCutout = False Then

         'Create the StructCutoutContour and add it as operator
         Set ostructcutoutcontourfactory = New StructCutoutContourFactory
         Set oStructCutoutContour = ostructcutoutcontourfactory.CreateStructCutoutContour(iPOMUnk, oSblCutoutOutput)
         
         'Do not display in treeview
         Dim iControlFlags As IJControlFlags
         Set iControlFlags = oStructCutoutContour
         iControlFlags.ControlFlags(CTL_FLAG_SYSTEM_MASK) = CTL_FLAG_NO_DISPLAY Or CTL_FLAG_NO_HILITE Or CTL_FLAG_NO_LOCATE Or CTL_FLAG_NO_DISPLAY_IN_NON_GRAPHIC_VIEW
         Set iControlFlags = Nothing
         
         oStructCutoutContour.MaxCutOutDistance = dMaxCutOutDistance
         oStructCutoutContour.MinCutOutDistance = dMinCutOutDistance
         
         ' Add cutout Contour as operator
         iOperatorElts.Add oStructCutoutContour
         
         iStructOperationPattern.SetOperationPattern sOperationProgId, iOperatorElts, iStructCutoutOperationAE
     Else
        'bSymbolContourIsCutout = true i.e iprevConnectable should exist, modify cutout
        If iprevConnectable Is Nothing Then
            SPSToDoErrorNotifyEx "SP3DDoorsAsm", IDS_SP3DDOORSSAM_MISSING_SUPPORTING_OBJECT, sMissingSupportingObjectError, iSmartOcc, Nothing
            Err.Raise E_FAIL
        End If
         If bHasMinCutOutDistanceChanged = True Or bHasMaxCutOutDistanceChanged = True Then
             oStructCutoutContour.MaxCutOutDistance = dMaxCutOutDistance
             oStructCutoutContour.MinCutOutDistance = dMinCutOutDistance
         End If
     End If


     iOperatorElts.Clear
     Set iOperatorElts = Nothing
     Set ostructcutoutcontourfactory = Nothing
     Set oStructCutoutContour = Nothing
     Set iStructCutoutOperationAE = Nothing
     Set iStructOperationPattern = Nothing
     Set iConnectable = Nothing
     Set iprevConnectable = Nothing
End If

 Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub


Public Sub myCMSetInputs4SO(ByVal pAD As IJDAggregatorDescription)
    Const sMethodName As String = "myCMSetInputs4SO": Dim sContext As String: On Error GoTo ErrorHandler
    On Error GoTo ErrorHandler
    
    m_oEquipCADHelper.SetSmartItemAsInputToSymbol pAD
    
    Exit Sub
ErrorHandler:
    Call LogAndRaiseError(Err, sContext, "AssemblyDefinition::" + sMethodName)
End Sub

Sub DefineAggregator(iSymbolDefinition As IJDSymbolDefinition, Optional bIsCustomized As Boolean = True)
    Dim iAggregatorDescription As IJDAggregatorDescription
    Set iAggregatorDescription = iSymbolDefinition
    
    ' define properties
    If True Then
        Let iAggregatorDescription.UserTypeClsid = "{643D1162-E891-11D4-9827-00104BD1CC25}"          ' SOSlot
        Let iAggregatorDescription.AggregatorClsid = m_oEquipCADHelper.OccurrenceRootClassGUID()
    End If
    
    ' define custom methods
    If bIsCustomized Then
        Call iAggregatorDescription.SetCMSetInputs(imsCOOKIE_ID_USS_LIB, "CMSetInputs4SO")
        Call iAggregatorDescription.SetCMMigrate(imsCOOKIE_ID_USS_LIB, "CMMigrate4SO")
        Call iAggregatorDescription.SetCMIdentifyClone(imsCOOKIE_ID_USS_LIB, "CMIdentifyClone4SO")
        Call iAggregatorDescription.SetCMAdaptClone(imsCOOKIE_ID_USS_LIB, "CMAdaptClone4SO")
    End If
         
    ' define custom properties
    If bIsCustomized Then
         Dim pPropertyDescriptions As IJDPropertyDescriptions
         Set pPropertyDescriptions = iSymbolDefinition
         
         Call pPropertyDescriptions.RemoveAll
         Call pPropertyDescriptions.AddProperty("EvalCutout", 1, "IJSOCompute", "CMEvaluateCutout", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE)
    End If
End Sub

Public Sub InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputColl As Object, ByRef arrayOfInputs())
    ' Obsolete method.
End Sub
Public Function EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
    ' Obsolete method.
    ' Instead you can record your custom command within the definition (see IJDCommandDescription interface)
    Let EditOccurence = False
End Function
'*************************************************************************
'Function
'IJFoulCheckForUSS
'   GetConnectedParts
'
'Abstract
'Adds objects that are related with the SmartOccurrence ( Door/Window support ) to the list of connected objects
'
'Arguments
'pPartObject is the SmartOccurrence object
'pIJMonUnks is the list of monikers
'
'Return
'
'Exceptions
'
'***************************************************************************

Private Sub IJFoulCheckForUSS_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As GSCADFoulCheck.IJElements)
    Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
    Let sSourceFile = "Doors_1_Asm.bas": Let sMethod = "IJFoulCheckForUSS_GetConnectedParts"
    On Error GoTo ErrorHandler
    
    Dim iSmartOcc As IJSmartOccurrence
    Dim oConnectable As Object
    Dim iPortDeck As IJPort
    Dim dEqpOffset As Double
    
    Set iSmartOcc = pPartObject
    Set iPortDeck = GetDoorMatingSurfaces(iSmartOcc, dEqpOffset)

    If Not iPortDeck Is Nothing Then
        Set oConnectable = iPortDeck.Connectable
        If Not oConnectable Is Nothing Then
            pIJMonUnks.Add oConnectable
        End If
    End If
     
    Set iSmartOcc = Nothing
    Set oConnectable = Nothing
    Set iPortDeck = Nothing
   
Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub

Private Sub IJFoulCheckForUSS_GetFoulInterfaceType(pFoulInterfaceType As GSCADFoulCheck.FoulInterfaceType)
    pFoulInterfaceType = StandardGraphicEntity
End Sub
