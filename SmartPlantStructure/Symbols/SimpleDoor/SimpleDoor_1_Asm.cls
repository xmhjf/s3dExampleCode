VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleDoor_1_Asm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Implements IJDUserSymbolServices
Implements IJFoulCheckForUSS



Public m_bDoors_1_Asm_InitializeInProgress As Boolean

Private m_oEquipCADHelper As IJEquipCADHelper


Public Function GetDefinitionName(ByVal definitionParameters As Variant) As String
  ' Name should be unique
    Let GetDefinitionName = "SimpleDoorAsm.SimpleDoor_1_Asm"
End Function

Private Sub Class_Initialize()
    Set m_oEquipCADHelper = New CADServices
    m_oEquipCADHelper.ProjectName = "SimpleDoorAsm"
    m_oEquipCADHelper.ClassName = "SimpleDoor_1_Asm"
    m_oEquipCADHelper.OccurrenceRootClass = orcEquipment
End Sub

Private Sub Class_Terminate()
    Set m_oEquipCADHelper = Nothing
End Sub

'
' IJDUserSymbolServices implementation
'
Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    Dim pSymbolDefinition As IJDSymbolDefinition
    Set pSymbolDefinition = m_oEquipCADHelper.InstanciateDefinition(CodeBase, defParams, ActiveConnection)
   
    ' en mode static, la definition est rechargee a chaque fois de la .dll
    ' le CopyBackwardFlag marque la CADefinition modified depuis la SmartOccurrence CES !
    'pSymbolDefinition.MetaDataOption = igSYMBOL_STATIC_METADATA
    
        IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition
End Function
Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDefinition As IJDSymbolDefinition)
    Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
    Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    
    If m_bDoors_1_Asm_InitializeInProgress = True Then
        sError = "SimpleDoor_1_Asm::already in progress"
        Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
    End If
    Let m_bDoors_1_Asm_InitializeInProgress = True
    
       
    Dim pCADefinition As IJCADefinition
    Set pCADefinition = pSymbolDefinition
    Let pCADefinition.CopyBackwardFlag = igCOPY_BACKWARD_TRIM
    
    sError = "Define the aggregator"
    Call DefineAggregator(pSymbolDefinition)
    
    'No members
    sError = "Define the members"
    Dim pMemberDescriptions As IJDMemberDescriptions
    Set pMemberDescriptions = pSymbolDefinition
        pMemberDescriptions.RemoveAll
    
    Let m_bDoors_1_Asm_InitializeInProgress = False
Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub
Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    Let IJDUserSymbolServices_GetDefinitionName = GetDefinitionName(definitionParameters)
End Function
Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    Call InvokeRepresentation(sblOcc, repName, outputcoll, arrayOfInputs)
End Sub
Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
 ' Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
    Let IJDUserSymbolServices_EditOccurence = EditOccurence(pSymbolOccurence, TransactionMgr)
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
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "CMAdaptClone4SO"
On Error GoTo ErrorHandler

    Exit Sub
    
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub

Public Sub CMEvalCutoutThickness(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "CMEvalCutoutThickness"

Dim sMissingSupportingObjectError As String: Let sMissingSupportingObjectError = "Door/Window is no longer associated to a supporting object. Property failed to compute. Undo the deletion of the supporting object, or mate equipment to a supporting object."

On Error GoTo ErrorHandler

'Get door mating constraint
Dim oSmartOcc As IJSmartOccurrence
Dim oPortDeck As IJPort 'plate support
Dim activeEntity As IJAssemblyConstraintAE

Set oSmartOcc = pPropertyDescriptions.CAO

'just retrieve door mating constraint if any
Set oPortDeck = GetDoorMatingSurfaces2(oSmartOcc, activeEntity)

If Not oPortDeck Is Nothing Then
    'Get connectable thickness to set as symbol input
    Dim dthickness As Double
    Dim dMatinfOffset As Double
    Dim oConnectable As IJConnectable
    Set oConnectable = oPortDeck.Connectable
    If oConnectable Is Nothing Then
        SPSToDoErrorNotifyEx "SP3DDoorsAsm", IDS_SP3DDOORSSAM_MISSING_SUPPORTING_OBJECT, sMissingSupportingObjectError, oSmartOcc, Nothing
        Err.Raise E_FAIL
    End If
    
    dthickness = GetSupportThickness(oConnectable)
    
    
    'retrieve symbol parameter FPR and DMF
    Dim pFPRAttribute As IJDAttribute
    Dim pDMFAttribute As IJDAttribute
    Dim pTopFrameDepthAttribute As IJDAttribute
    Dim pLowerFrameDepthAttribute As IJDAttribute
    Dim pLeftFrameDepthAttribute As IJDAttribute
    Dim pRightFrameDepthAttribute As IJDAttribute
    Dim lFPR As Long
    Dim dDMF As Double
    Dim dTopFrameDepth As Double
    Dim dLowerFrameDepth As Double
    Dim dLeftFrameDepth As Double
    Dim dRightFrameDepth As Double
    Dim dMaxFrameDepth As Double
    
    Set pFPRAttribute = GetAttribute(pObject, "IJUASDoor", "DoorFramePosition")
    Set pDMFAttribute = GetAttribute(pObject, "IJUASDoor", "DMF")
    
    lFPR = pFPRAttribute.Value
    dDMF = pDMFAttribute.Value
    
    Set pTopFrameDepthAttribute = GetAttribute(pObject, "IJUASDoor", "TopFrameDepth")
    Set pLowerFrameDepthAttribute = GetAttribute(pObject, "IJUASDoor", "LowerFrameDepth")
    Set pLeftFrameDepthAttribute = GetAttribute(pObject, "IJUASDoor", "LeftFrameDepth")
    Set pRightFrameDepthAttribute = GetAttribute(pObject, "IJUASDoor", "RightFrameDepth")
    
    dTopFrameDepth = pTopFrameDepthAttribute.Value
    dLowerFrameDepth = pLowerFrameDepthAttribute.Value
    dLeftFrameDepth = pLeftFrameDepthAttribute.Value
    dRightFrameDepth = pRightFrameDepthAttribute.Value
    
    dMaxFrameDepth = dTopFrameDepth
    If dLowerFrameDepth > dMaxFrameDepth Then dMaxFrameDepth = dLowerFrameDepth
    If dLeftFrameDepth > dMaxFrameDepth Then dMaxFrameDepth = dLeftFrameDepth
    If dRightFrameDepth > dMaxFrameDepth Then dMaxFrameDepth = dRightFrameDepth
    
    
    'FPR=0: centered, FPR=1: flush, FPR=2 DMF
    'Set computed DMF for FPR = 0 or 1
    If lFPR = 0 Then
        If dMaxFrameDepth > 0# Then  ' need to get center of frame exactly on center of the wall
            dMatinfOffset = (dthickness / 2) + (dMaxFrameDepth / 2)
            pDMFAttribute.Value = (dthickness / 2) - (dMaxFrameDepth / 2)
        Else
            dMatinfOffset = dthickness / 2
            pDMFAttribute.Value = dMatinfOffset
        End If
        
    ElseIf lFPR = 1 Then
        dMatinfOffset = dMaxFrameDepth
        pDMFAttribute.Value = 0
    Else
        dMatinfOffset = dDMF + dMaxFrameDepth
    End If
    
    ' Update mating offset only if value changed
    If Abs(activeEntity.offset + dMatinfOffset) > 0.001 Then
        'Set the new position of Frames by setting the mating offset
        activeEntity.offset = -dMatinfOffset
    End If
    
End If


Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub

Public Sub CMEvaluateCutout(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Dim sSourceFile As String: Dim sMethod As String: Dim sError As String
Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "CMEvaluateCutout"
On Error GoTo ErrorHandler

Dim oOppositeDeckPort As IJPort
Dim oprevDeckPort As IJPort
Dim oprevOppositeDeckPort As IJPort
Dim oprevConnectable As IJConnectable
Dim oEquip As IJEquipment
Dim oSmartOcc As IJSmartOccurrence
Dim pSymbol As IJDSymbol
Dim oSblCutoutOutput As Object
Dim oConnectable As IJConnectable

Dim oSupportType    As SupportType
Dim oprevSupportType As SupportType
Dim sprevOperationProgId As String
Dim sOperationProgId As String
Dim nbRef As Long

'Get Symbol "CutoutContour" output
Set oSmartOcc = pPropertyDescriptions.CAO
Set oEquip = oSmartOcc
Set pSymbol = oSmartOcc
Set oSblCutoutOutput = pSymbol.BindToOutput("CutoutContour", "CutoutCurve")
If oSblCutoutOutput Is Nothing Then
    sError = "CMEvaluateCutout::oSblCutoutOutput is nothing"
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End If

'GetCutoutContour from oSblCutoutOutput if any
Dim oprevStructCutoutContour As StructCutoutContour
Set oprevStructCutoutContour = GetDoorCutoutContour(oSblCutoutOutput)
If Not oprevStructCutoutContour Is Nothing Then
    'Retrieve Connectable from the equipment (Nothing if no cutout)
    Set oprevConnectable = GetConnectableFromCutoutContour(oprevStructCutoutContour)
    If oprevConnectable Is Nothing Then
        sError = "CMEvaluateCutout::cannot retrieve connectable of the Cutout"
        Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
    End If
End If

Dim oPortDeck As IJPort 'plate support
Dim bDefaultMatingCnstrExist As Boolean
Dim bReferenceExist As Boolean
Dim EqpOffset As Double

bDefaultMatingCnstrExist = False
Set oPortDeck = GetDoorMatingSurfaces(oSmartOcc, EqpOffset)

If Not oPortDeck Is Nothing Then
    bDefaultMatingCnstrExist = True
End If

'Get Connectable type and set corresponding progId
If Not (oprevConnectable Is Nothing) Then
    If TypeOf oprevConnectable Is ISPSSlabEntity Then
       sprevOperationProgId = "StructGeneric.StructCutoutOperationAE.1"
       oprevSupportType = SLAB
    Else 'ISPSWallPart
       sprevOperationProgId = "SP3DStructGeneric.StructCutoutOperation"
       oprevSupportType = WALL
    End If
End If

'mating constraint has been removed, remove previous cutout if any
If bDefaultMatingCnstrExist = False And Not (oprevConnectable Is Nothing) Then
    RemoveCutoutFromContour oprevConnectable, oSblCutoutOutput, sprevOperationProgId
    Exit Sub
End If

'default mating constraint exists, create or update cutout
If bDefaultMatingCnstrExist = True Then

    Dim pObj As IJDObject
    Dim oResourceManager As IUnknown
    Set pObj = oSblCutoutOutput
    Set oResourceManager = pObj.ResourceManager
    Dim oCollectionOfOperators As IJElements
    Dim OperationPattern As IJStructOperationPattern
    Dim oStructCutoutOperationAE As StructCutoutOperationAE
    Dim ostructcutoutcontourfactory As StructCutoutContourFactory
    Dim oStructCutoutContour As StructCutoutContour
    Dim bSymbolContourIsCutout As Boolean
    Dim oSymbolEntitiesFactory  As New DSymbolEntitiesFactory
    
    'Gets the object from which the port is coming
    Set oConnectable = oPortDeck.Connectable
    If oConnectable Is Nothing Then
        sError = "CMEvaluateCutout::iConnectable is nothing"
        Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
    End If
    
    'Get Connectable type and set corresponding progId
    If TypeOf oConnectable Is ISPSSlabEntity Then
       oSupportType = SLAB
       sOperationProgId = "StructGeneric.StructCutoutOperationAE.1"
    Else 'ISPSWallPart
       oSupportType = WALL
       sOperationProgId = "SP3DStructGeneric.StructCutoutOperation"
    End If
    
    'if a cutout exists and its support is different
    'this means that SO is replacing its support, need to remove previous cutoutcontour and reference
    ' and go on
    If Not (oprevConnectable Is Nothing) Then
        If Not (oprevConnectable Is oConnectable) Then
            RemoveCutoutFromContour oprevConnectable, oSblCutoutOutput, sprevOperationProgId
            'reset oprevConnectable
            Set oprevConnectable = Nothing
        End If
        GetOppositeFaceFromSupport oPortDeck, oOppositeDeckPort, oSupportType
        If oOppositeDeckPort Is Nothing Then
            sError = "CMEvaluateCutout::oOppositeDeckPort is nothing"
            Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
        End If
    End If
    
    'Set Cutting Limit to create ModelBody (created from the planar contour in the direction of its normal)
    'MinCutOutDistance/MaxCutOutDistance represent the start/end of the cutting tool from the profile curve
    'MaxCutOutDistance is the distance between the Equipment and the support + width of support.
    'MinCutOutDistance is taken as -MaxCutOutDistance to work when support is mirrored

    
     Dim MinCutOutDistance As Double
     Dim prevMinCutOutDistance As Double
     Dim MaxCutOutDistance As Double
     Dim prevMaxCutOutDistance As Double
     Dim HasMinCutOutDistanceChanged As Boolean
     Dim HasMaxCutOutDistanceChanged As Boolean
  
     HasMinCutOutDistanceChanged = True
     HasMaxCutOutDistanceChanged = True
     
     'dist +/- extendedLength
     Dim extendedLength As Double
     Dim dthickness As Double
     extendedLength = 0.2
     dthickness = GetSupportThickness(oConnectable)
    
     'EqpOffset can be negative
     MaxCutOutDistance = Abs(Abs(EqpOffset) + dthickness + extendedLength)
     MinCutOutDistance = -MaxCutOutDistance
     
     bSymbolContourIsCutout = False
     Set OperationPattern = oConnectable
     
  
     OperationPattern.GetOperationPattern sOperationProgId, oCollectionOfOperators, oStructCutoutOperationAE
     If (Not oCollectionOfOperators Is Nothing) Then
         For Each oStructCutoutContour In oCollectionOfOperators
             ' If the symbol contour is not yet in the collection add it
             If oStructCutoutContour.InputContour Is oSblCutoutOutput Then
                 bSymbolContourIsCutout = True
                 'Get previous cuttingDepth
                 prevMaxCutOutDistance = oStructCutoutContour.MaxCutOutDistance
                 prevMinCutOutDistance = oStructCutoutContour.MinCutOutDistance
                 If Abs(prevMaxCutOutDistance - MaxCutOutDistance) < 0.001 Then
                     HasMaxCutOutDistanceChanged = False
                 End If
                 If Abs(prevMinCutOutDistance - MinCutOutDistance) < 0.001 Then
                     HasMinCutOutDistanceChanged = False
                 End If
                 Exit For
             End If
         Next
     Else
         Set oCollectionOfOperators = New JObjectCollection
     End If
     
     If bSymbolContourIsCutout = False Then

         'Create the StructCutoutContour and add it as operator
         Set ostructcutoutcontourfactory = New StructCutoutContourFactory
         Set oStructCutoutContour = ostructcutoutcontourfactory.CreateStructCutoutContour(oResourceManager, oSblCutoutOutput)
         
         'Do not display in treeview
         Dim oControlFlags As IJControlFlags
         Set oControlFlags = oStructCutoutContour
         oControlFlags.ControlFlags(CTL_FLAG_SYSTEM_MASK) = CTL_FLAG_NO_DISPLAY Or CTL_FLAG_NO_HILITE Or CTL_FLAG_NO_LOCATE Or CTL_FLAG_NO_DISPLAY_IN_NON_GRAPHIC_VIEW
         Set oControlFlags = Nothing
         
         'Retrieve opposite port to set as input of SO
         If oOppositeDeckPort Is Nothing Then
            GetOppositeFaceFromSupport oPortDeck, oOppositeDeckPort, oSupportType
            If oOppositeDeckPort Is Nothing Then
                sError = "CMEvaluateCutout::oOppositeDeckPort is nothing"
                Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
            End If
         End If
         
         oStructCutoutContour.MaxCutOutDistance = MaxCutOutDistance
         oStructCutoutContour.MinCutOutDistance = MinCutOutDistance
         
         ' Add cutout Contour as operator
         oCollectionOfOperators.Add oStructCutoutContour
         
         OperationPattern.SetOperationPattern sOperationProgId, oCollectionOfOperators, oStructCutoutOperationAE
     Else
        'bSymbolContourIsCutout = true i.e oprevConnectable should exist, modify cutout
        If oprevConnectable Is Nothing Then
            sError = "CMEvaluateCutout::oprevConnectable is nothing while there is a contour"
            Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
        End If
         
         If HasMinCutOutDistanceChanged = True Or HasMaxCutOutDistanceChanged = True Then
             oStructCutoutContour.MaxCutOutDistance = MaxCutOutDistance
             oStructCutoutContour.MinCutOutDistance = MinCutOutDistance
         End If
     End If


     oCollectionOfOperators.Clear
     Set oCollectionOfOperators = Nothing
     Set ostructcutoutcontourfactory = Nothing
     Set oStructCutoutContour = Nothing
     Set oStructCutoutOperationAE = Nothing
     Set OperationPattern = Nothing
     Set oConnectable = Nothing
     Set oprevConnectable = Nothing
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

Sub DefineAggregator(pSymbolDefinition As IJDSymbolDefinition, Optional bIsCustomized As Boolean = True)
    Dim pAggregatorDescription As IJDAggregatorDescription
    Set pAggregatorDescription = pSymbolDefinition
    
    ' define properties
    If True Then
        Let pAggregatorDescription.UserTypeClsid = "{643D1162-E891-11D4-9827-00104BD1CC25}"          ' SOSlot
        Let pAggregatorDescription.AggregatorClsid = m_oEquipCADHelper.OccurrenceRootClassGUID()
    End If
    
    ' define custom methods
    If bIsCustomized Then
        Call pAggregatorDescription.SetCMSetInputs(imsCOOKIE_ID_USS_LIB, "CMSetInputs4SO")
        Call pAggregatorDescription.SetCMMigrate(imsCOOKIE_ID_USS_LIB, "CMMigrate4SO")
        Call pAggregatorDescription.SetCMIdentifyClone(imsCOOKIE_ID_USS_LIB, "CMIdentifyClone4SO")
        Call pAggregatorDescription.SetCMAdaptClone(imsCOOKIE_ID_USS_LIB, "CMAdaptClone4SO")
    End If
         
    ' define custom properties
    If bIsCustomized Then
         Dim pPropertyDescriptions As IJDPropertyDescriptions
         Set pPropertyDescriptions = pSymbolDefinition
         
         Call pPropertyDescriptions.RemoveAll
         Call pPropertyDescriptions.AddProperty("EvalCutoutThickness", 1, "IJUASDoor", "CMEvalCutoutThickness", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE)
         Call pPropertyDescriptions.AddProperty("EvalCutout", 2, "IJSOCompute", "CMEvaluateCutout", imsCOOKIE_ID_USS_LIB, igPROCESS_PD_AFTER_SYMBOL_UPDATE)
    End If
End Sub

Public Sub InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
    ' Obsolete method.
End Sub
Public Function EditOccurence(ByRef pSymbolOccurence As Object, ByVal TransactionMgr As Object) As Boolean
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
    Let sSourceFile = "SimpleDoor_1_Asm.bas": Let sMethod = "IJFoulCheckForUSS_GetConnectedParts"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oConnectable As Object
    Dim oPortDeck As IJPort
    Dim EqpOffset As Double
    
    Set oSmartOcc = pPartObject
    Set oPortDeck = GetDoorMatingSurfaces(oSmartOcc, EqpOffset)

    If Not oPortDeck Is Nothing Then
        Set oConnectable = oPortDeck.Connectable
        If Not oConnectable Is Nothing Then
            pIJMonUnks.Add oConnectable
        End If
    End If
     
    Set oSmartOcc = Nothing
    Set oConnectable = Nothing
    Set oPortDeck = Nothing
   
Exit Sub
ErrorHandler:
    Err.Raise ReportError(Err, sSourceFile, sMethod, sError).Number
End Sub

Private Sub IJFoulCheckForUSS_GetFoulInterfaceType(pFoulInterfaceType As GSCADFoulCheck.FoulInterfaceType)
    pFoulInterfaceType = StandardGraphicEntity
End Sub
