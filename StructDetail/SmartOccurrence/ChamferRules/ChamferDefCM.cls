VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChamferDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "S:\StructDetail\Data\SmartOccurrence\ChamferRules\ChamferDefCM.cls"
Private sERROR As String
Private sMETHOD As String


'***********************************************************************
' METHOD:  CAConstruct_PhysConn
'
' DESCRIPTION:  Constructs the Physical Connection
'***********************************************************************

Public Function CAConstruct_PhysConn(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    Dim strError As String
    On Error GoTo ErrorHandler

    sMETHOD = "CAConstruct_PhysConn"
    
    ' Get Class
    Dim pChamfer As New StructDetailObjects.Chamfer ' create new chamfer wrapper class
    sERROR = "Setting Chamfer Inputs"
    Set pChamfer.object = pMemberDescription.CAO ' Intialize the wrapper with slot
    
    ' Get chamfered Part's Port
    Dim oChamferedPort As IJPort ' This is the chamfered port
    Dim pSDOHelper As New StructDetailObjects.Helper
    sERROR = "Getting chamfered object"
    Set oChamferedPort = pSDOHelper.GetEquivalentLastPort(pChamfer.Chamfered) ' Get the chamfered part from wrapper
    
    ' Get the driving Part's Port
    Dim oDrivingPort As IJPort
    sERROR = "Getting the driving parts port"
    Set oDrivingPort = pSDOHelper.GetEquivalentLastPort(pChamfer.DrivesChamfer)
    
    ' Get the parent of the PC, which is the AssyConn with the Chamfer
    Dim pSystemParent As IJSystemChild
    sERROR = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
    
    Dim pAssyConn As IJSystem
    Set pAssyConn = pSystemParent.GetParent
       
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sERROR = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                                            oChamferedPort, _
                                            oDrivingPort, _
                                            strStartClass, _
                                            pAssyConn, ConnectionStandard

    sERROR = "Setting Physical Connection to private variable"
    Set CAConstruct_PhysConn = oPhysicalConnection.object

    Set oDrivingPort = Nothing
    Set oChamferedPort = Nothing
    Set oPhysicalConnection = Nothing
    Set pSystemParent = Nothing
    Set pChamfer = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CAConstruct_PhysConn", sERROR).Number
        
End Function


'***********************************************************************
' METHOD:  Chamfer_FinalConstruct
'
' DESCRIPTION:  Does the final construction of the chamfer.  Adds to the graph
'***********************************************************************
Public Sub Chamfer_FinalConstruct(ByVal pAggregatorDescription As IJDAggregatorDescription)

  On Error GoTo ErrorHandler
  sMETHOD = "Chamfer_FinalConstruct"
  
  ' Create Slot and add cut
  Dim oChamfer As New StructDetailObjects.Chamfer
  Set oChamfer.object = pAggregatorDescription.CAO
  
  oChamfer.AddChamferToCut
  
  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "Chamfer_FinalConstruct").Number

End Sub


