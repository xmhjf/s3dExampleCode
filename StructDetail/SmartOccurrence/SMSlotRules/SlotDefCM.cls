VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SlotDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "SlotRules\SlotDefCM.cls"

Private sERROR As String
Private sMETHOD As String


'***********************************************************************
' METHOD:  CAConstruct_PhysConn
'
' DESCRIPTION:  Constructs the Physical Connection
'               Currently this method assumes that:
'                     Penetrating object is a profile
'                     Penetrated Object is a plate
'***********************************************************************

Public Function CAConstruct_PhysConn(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown, _
                                    strStartClass As String, _
                                    ByVal ePenetratingSubPort As JXSEC_CODE, _
                                    ByVal ePenetratedSubPort As JXSEC_CODE) As Object

    sMETHOD = "Slot_FinalConstruct"
    
    ' Get Class
    Dim pSlot As New StructDetailObjects.Slot ' create new slot wrapper class
    sERROR = "Setting Slot Inputs"
    Set pSlot.object = pMemberDescription.CAO
    
    ' Get Penetrating Part's Port, figure out if it's a plate or profile
    sERROR = "Getting Penetrating object's port"
    Dim oPenetratingPort As IJPort
    If TypeOf pSlot.Penetrating Is IJProfile Then
        Dim oPenetratingProfile As New StructDetailObjects.ProfilePart
        Set oPenetratingProfile.object = pSlot.Penetrating
        Set oPenetratingPort = oPenetratingProfile.SubPort(ePenetratingSubPort)
    ElseIf TypeOf pSlot.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oBasePort As IJPort
        Dim oMappedPorts As JCmnShp_CollectionAlias
        Set oMappedPorts = New Collection
        
        oSlotMappingRule.GetEmulatedPorts pSlot.Penetrating, pSlot.Penetrated, oBasePort, oMappedPorts
        
        Dim pSDOHelper As New StructDetailObjects.Helper
        Set oPenetratingPort = pSDOHelper.GetEquivalentLastPort(oMappedPorts.Item(CStr(ePenetratingSubPort)))
    Else
        'Unsupported Penetrating Object
        Exit Function
    End If
   
    ' Get Penetrated Part's Port, figure out if it's a plate or profile
    sERROR = "Getting Penetrated object's port"
    Dim oPenetratedPort As IJPort
    If TypeOf pSlot.Penetrated Is IJPlate Then
        Dim oPenetratedPlate As New StructDetailObjects.PlatePart
        Set oPenetratedPlate.object = pSlot.Penetrated
        Set oPenetratedPort = oPenetratedPlate.CutoutSubPort(pSlot.object, ePenetratedSubPort)
    ElseIf TypeOf pSlot.Penetrated Is IJProfile Then
        Dim oPenetratedProfile As New StructDetailObjects.ProfilePart
        Set oPenetratedProfile.object = pSlot.Penetrated
        Set oPenetratedPort = oPenetratedProfile.CutoutSubPort(pSlot.object, ePenetratedSubPort)
    ElseIf TypeOf pSlot.Penetrated Is ISPSMemberPartPrismatic Then
        Dim oMbrPartWrapper As New StructDetailObjectsEx.SDOMemberPart
        Set oMbrPartWrapper.object = pSlot.Penetrated
        Set oPenetratedPort = oMbrPartWrapper.CutoutPort(pSlot.object, ePenetratedSubPort)
    Else
        'Unsupported Penetrated Object
        Exit Function
    End If
        
    ' Get the slot, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sERROR = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
 
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sERROR = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                               oPenetratedPort, _
                               oPenetratingPort, _
                               strStartClass, _
                               pSystemParent, ConnectionStandard
                               
    sERROR = "Setting Physical Connection to private variable"
    Set CAConstruct_PhysConn = oPhysicalConnection.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
        
End Function


'***********************************************************************
' METHOD:  Slot_FinalConstruct
'
' DESCRIPTION:  Does the final construction of the slot.  Adds to the graph
'***********************************************************************
Public Sub Slot_FinalConstruct(ByVal pAggregatorDescription As IJDAggregatorDescription)

  On Error GoTo ErrorHandler
  sMETHOD = "Slot_FinalConstruct"
  
  ' Create Slot and add cut
  Dim oSlot As New StructDetailObjects.Slot
  Set oSlot.object = pAggregatorDescription.CAO
  
  oSlot.AddSlotToCut
  
  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number

End Sub

Public Sub CMConstructFET1(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    
    On Error GoTo ErrorHandler
    Set pObject = CMConstructFETForFeature(pMemberDescription, pResourceManager)
    
    Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMConstructFET1").Number
End Sub

Public Sub CMGetEdgeTreatment(ByRef pMD As IJDMemberDescription, _
                          ByRef bIsNeeded As Boolean)
 On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FreeEdgeTreatment) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    
    Dim sApplyTreatment As String
    sApplyTreatment = GetAnswer_ApplyTreatment(pMD, SMARTTYPE_SLOT)
    
    If LCase(sApplyTreatment) = LCase("No") Then
        bIsNeeded = False
    ElseIf LCase(sApplyTreatment) = LCase("Yes") Then
        bIsNeeded = True
    End If
    
       ' Putting this update call here as a convenience
    ' Should be put in a more logical location
    UpdateDependentCorners pMD.CAO
    
    
    Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMGetEdgeTreatment").Number
End Sub
