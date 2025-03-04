VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Repair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
'
' File: Repair.cls
'
' Abstract:
'   Repair rule to update a slot that was posted by SMCheckMfctySlotAngle.Check
'

Option Explicit

Private Const m_sClassName As String = "Repair"
Private Const m_sRuleProgID = m_sProjectName + m_sClassName
Private Const m_sModule = m_sProjectPath + m_sClassName + ".cls"

Implements IJCheckMfctyRepair

' Method: IJCheckMfctyRepair_Repair
'
' Abstract:
'   Implementation of the IJCheckMfctyRepair interface
'
' Description: Recomput slot
'
' Inputs:
'     vObject:  The object to be repaired.
'******************************************************************************
Private Sub IJCheckMfctyRepair_Repair(ByRef vObject As Variant)
   Const sMethod As String = "IJCheckMfctyRepair_Repair"
   
   On Error GoTo ErrorHandler
   Dim oObject As Object
   
   Set oObject = vObject
   If oObject Is Nothing Then
      Exit Sub
   End If
   
   If Not TypeOf oObject Is IJStructFeature Then
      Exit Sub
   End If
   
   Dim oStructFeature As IJStructFeature
   Dim eSFType As StructFeatureTypes
   
   Set oStructFeature = oObject
   eSFType = oStructFeature.get_StructFeatureType
   
   If eSFType <> SF_Slot Then
      Exit Sub
   End If
   
   Dim oStructAssocTools As New SP3DStructGenericTools.StructAssocTools
   Dim oDObject As IJDObject
    
   ' The CheckMfcty checking mechanism has set the status of the object
   ' to Rejected.  The Revision Manager cannot recompute if the object
   ' is not set to working.
   ' Therefore the status of the object must be changed.
   Set oDObject = oStructFeature
   oDObject.ApprovalStatus = Working
   Set oStructFeature = Nothing
   
   oStructAssocTools.UpdateObject oDObject, IID_IJFullObject
   Set oDObject = Nothing
   Set oStructAssocTools = Nothing
    
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, m_sModule, sMethod, "Error").Number

End Sub
