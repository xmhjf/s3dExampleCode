VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateEdgeToMemberFaceDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'
'   Item definition: basic template to define an item
'

Const m_ItemProgid As String = CUSTOMERID + "AssyConRul.PlateEdgeToMemberFaceDef"
Const m_ItemName As String = CUSTOMERID + "AssyConRul.PlateEdgeToMemberFaceDef"
Const m_FamilyProgid As String = ""

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\PlateEdgeToMemberFaceDef.cls"

Implements IJDUserSymbolServices

Public Sub ItemInputs(pIH As IJDInputsHelper)
    On Error GoTo ErrorHandler
    
    Dim strError As String

    strError = "Setting bounded and bounding objects as inputs."
    pIH.SetInput INPUT_BOUNDED_OR_PENETRATED_OBJECT
    pIH.SetInput INPUT_BOUNDING_OR_PENETRATING_OBJECT
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemInputs", strError).Number
End Sub

Public Sub ItemAggregator(pAD As IJDAggregatorDescription)
    On Error GoTo ErrorHandler
    ' TODO - Add aggregator description below
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemAggregator").Number
End Sub

Public Sub ItemMembers(pMDs As IJDMemberDescriptions)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim pMemDesc As IJDMemberDescription
    
    strError = "Constructing physical connection."
    Set pMemDesc = pMDs.AddMember("PlateEdgeToMemberFace1", 1, "CMConstructPC", imsCOOKIE_ID_USS_LIB)
    pMemDesc.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConstructPCCon"
    Set pMemDesc = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ItemMembers", strError).Number
End Sub

' ** Start CM **
' *******************************************************************************************
' If needed Add Custom Method HERE
' *******************************************************************************************
Public Sub CMConstructPC(ByVal pMemberDescription As IJDMemberDescription, _
                                        ByVal pResourceManager As IUnknown, _
                                        ByRef pObject As Object)
    On Error GoTo ErrorHandler
    
    Dim strError As String
    Dim clsAssyConnCM As New AssyConnDefCM
    Dim strWeld As String
    PCTypePlateEdgeToMemberFace pMemberDescription, pResourceManager, strWeld
    
    strError = "Constructing physical connection."
    
    'Get Wrapper Class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    Set pAssyConn.object = pMemberDescription.CAO
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
            
    Set oPort1 = pAssyConn.Port1
    Set oPort2 = pAssyConn.Port2
    
    Dim oAPSPlt As New StructDetailObjects.PlatePart
    Set oAPSPlt.object = oPort1.Connectable
    
    Dim oPlateSystem As IJPlateSystem
    Dim oHelper As IJPlatePartAttributes
    
    Set oHelper = New PlatePartUtils
    
    
    Dim bISMacro As Boolean
    Dim oGeometricConstruction As IJGeometricConstruction
    Dim sGCType As String
    
    Set oPlateSystem = oHelper.GetPartRootPlateSystem(oPort1.Connectable)
    
    If Not oPlateSystem Is Nothing Then
    
        bISMacro = oPlateSystem.IsMacroPlate(oGeometricConstruction)
        
        If bISMacro Then
            sGCType = oGeometricConstruction.Type
        End If
        
    End If
    
    If InStr(1, sGCType, "NosePlate", vbTextCompare) = 0 Then
        Set pObject = clsAssyConnCM.CMConstruct_PhysConnFromACPorts(pMemberDescription, pResourceManager, strWeld)
    Else
        Set oPort1 = oAPSPlt.baseport(BPT_Lateral)  'Use complete lateral port for PC creation
        
        ' Get the Assembly connection, since it is the parent of the PC
        Dim pSystemParent As IJSystemChild ' this is the slot's ca
        Set pSystemParent = pMemberDescription.CAO
           
        ' Create physical connection
        Dim oPhysicalConnection As New PhysicalConn
        Dim oSLPort1 As IJPort
        Dim oSLPort2 As IJPort
        
        Dim oStructDetailObjectHelper As New StructDetailObjects.Helper
        
        Set oSLPort1 = oStructDetailObjectHelper.GetEquivalentLastPort(oPort1)
        Set oSLPort2 = oStructDetailObjectHelper.GetEquivalentLastPort(oPort2)
        Set oStructDetailObjectHelper = Nothing
        Set oPort1 = Nothing
        Set oPort2 = Nothing
    
        Call oPhysicalConnection.Create(pResourceManager, oSLPort2, oSLPort1, _
                                        strWeld, pSystemParent, ConnectionStandard)
                                        
        Set pObject = oPhysicalConnection.object
    End If
    Set clsAssyConnCM = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPC", strError).Number
End Sub
Public Function PCTypePlateEdgeToMemberFace(pMemberDescription, pResourceManager, strWeld As String)
    On Error GoTo ErrorHandler
    Dim strError As String
    Dim oMember As Object
   
    Dim AssyConn As New StructDetailObjects.AssemblyConn
    
    Set AssyConn.object = pMemberDescription.CAO
    
    Dim oPlatePort As IJPort
    
    Set oPlatePort = AssyConn.Port1
    
    Dim MemberPart As New StructDetailObjects.MemberPart
   
    Set oMember = AssyConn.ConnectedObject2
    Set MemberPart.object = oMember
    Dim TFlgRPort As IJPort
    Dim TFlgLPort As IJPort
    Dim BFlgRPort As IJPort
    Dim BFlgLPort As IJPort
    Dim oUtility As GSCADShipGeomOps.SGOModelBodyUtilities
    Set oUtility = New SGOModelBodyUtilities
    Dim HasIntrstGeometry As Boolean
    Dim bTFL As Boolean
    Dim bBFL As Boolean
    Dim bTFR As Boolean
    Dim bBFR As Boolean
    If TypeOf MemberPart.object Is IJConnectable Then
        CrossSection_Flanges MemberPart.object, bTFL, bBFL, bTFR, bBFR
        If bTFR Then
            'getting the TopFlgRight port
            Set TFlgRPort = GetLateralSubPortBeforeTrim(MemberPart.object, JXSEC_TOP_FLANGE_RIGHT)
            'Checking if there is an interseccting geometry between Memberport and platePort
            HasIntrstGeometry = oUtility.HasIntersectingGeometry(oPlatePort.Geometry, TFlgRPort.Geometry)
        End If
        If bTFL And Not HasIntrstGeometry Then
            'getting the TopFlgLeft port
            Set TFlgLPort = GetLateralSubPortBeforeTrim(MemberPart.object, JXSEC_TOP_FLANGE_LEFT)
            HasIntrstGeometry = oUtility.HasIntersectingGeometry(oPlatePort.Geometry, TFlgLPort.Geometry)
        End If
        If bBFR And Not HasIntrstGeometry Then
            Set BFlgRPort = GetLateralSubPortBeforeTrim(MemberPart.object, JXSEC_BOTTOM_FLANGE_RIGHT)
            'Checking if there is an interseccting geometry between Memberport and platePort
            HasIntrstGeometry = oUtility.HasIntersectingGeometry(oPlatePort.Geometry, BFlgRPort.Geometry)
        End If
        If bBFL And Not HasIntrstGeometry Then
            Set BFlgLPort = GetLateralSubPortBeforeTrim(MemberPart.object, JXSEC_BOTTOM_FLANGE_LEFT)
            HasIntrstGeometry = oUtility.HasIntersectingGeometry(oPlatePort.Geometry, BFlgLPort.Geometry)
        End If
        
        If HasIntrstGeometry Then
            'its an edge to edge connection, setting it to Butt weld
            strWeld = "ButtWeld"
        Else
            ' its not an edge to edge connection, setting it to Tee weld
            strWeld = "TeeWeld"
        End If
    End If
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "PCTypePlateEdgeToMemberFace", strError).Number
End Function

' ** End CM **
' ********************************************************************************************
'         !!!!! Start Private Code !!!!!
'                 - Following Code Should not be edited
'                 - It exposes the Selector as a regular symbol definition
' ********************************************************************************************

Public Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    ' Name should be unique
    IJDUserSymbolServices_GetDefinitionName = m_ItemName
End Function

Public Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IJDSymbolDefinition)

    ' Remove all existing defined Input and Output (Representations)
    ' before defining the current Inputs and Outputs
    pDefinition.IJDInputs.RemoveAllInput
    pDefinition.IJDRepresentations.RemoveAllRepresentation

    pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
    pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
      
    ' define the inputs
    Dim pIH As IJDInputsHelper
    Set pIH = New InputHelper
    pIH.definition = pDefinition
    pIH.InitAs m_FamilyProgid
    ItemInputs pIH
    
    Dim pPDs As IJDPropertyDescriptions
    
    ' define the aggregator
    Dim pAD As IJDAggregatorDescription
    Set pAD = pDefinition
    Set pPDs = pAD
    pPDs.RemoveAll ' Remove all the previous property descriptions
    ItemAggregator pAD
     
    ' define the members
    Dim pMDs As IJDMemberDescriptions
    Set pMDs = pDefinition
    pMDs.RemoveAll ' Remove all the previous Member descriptions
    ItemMembers pMDs
End Sub

Public Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal ActiveConnection As Object) As Object
    Dim pDefinition As IJDSymbolDefinition
    Dim pCAFactory As New CAFactory
    
    Set pDefinition = pCAFactory.CreateCAD(ActiveConnection)
    
    ' Set definition progId and codebase
    pDefinition.ProgId = m_ItemProgid
    pDefinition.CodeBase = CodeBase
    pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
      
    ' Initialize the definition
    IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
    
    Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
End Function

Public Sub IJDUserSymbolServices_InvokeRepresentation(ByVal sblOcc As Object, ByVal repName As String, ByVal outputcoll As Object, ByRef arrayOfInputs())
End Sub

Public Function IJDUserSymbolServices_EditOccurence(ByRef pSymbolOccurence As Object, ByVal transactionMgr As Object) As Boolean
End Function


Public Sub CMConstructPCCon(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    bIsNeeded = True
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructPCCon").Number
End Sub