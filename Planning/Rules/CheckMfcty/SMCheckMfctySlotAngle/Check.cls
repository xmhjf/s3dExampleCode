VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Copyright (C) 2009 Intergraph Corporation.  All Rights Reserved.
'
' File: Check.cls
'
' Abstract:
'   Check Manufacturability rule that checks if assembly orientation stored
'   in slot occurrence is still the same as the orientation of the assembly
'   that penetrated part is in.
'
Option Explicit

Private Const m_sClassName As String = "Check"
Private Const m_sRuleProgID = m_sProjectName + m_sClassName
Private Const m_sModule = m_sProjectPath + m_sClassName + ".cls"

Implements IJCheckMfcty
' Method: IJCheckMfcty_ExecuteCheck
'
' Description:
'    This method is called by the Check Manufacturability mechanism
'  to check if assembly orientation stored in slot occurrence is the
'  same as the orientation of the assembly that penetrated part is in.
'  If not, the slot will be posted in the list view so that it can be
'  updated when user choses to Repair it.
'
' Inputs:
'     oItemMonikers :   The monikers for items in the working set.
'     bStopOnError  :   Operator option to stop after first error has been encountered.
'     vOptionCodes  :   Not used.
'     oCallingObj   :   The object that made the call to this method.
'                       This is used to pass back information to the Check
'                       Manufacturability command.
'
Private Sub IJCheckMfcty_ExecuteCheck( _
        ByVal oItemMonikers As GSCADGlobalCheck.IJElements, _
        ByVal bStopOnError As Boolean, _
        ByRef vOptionCodes() As Variant, _
        ByVal oCallingObject As Object)
   Const sMethod = "IJCheckMfcty_ExecuteCheck"
   On Error GoTo ErrorHandler
   
   If oItemMonikers Is Nothing Then
      Exit Sub
   End If
   
   Dim vMoniker As Variant
   Dim oElement As Object
   Dim oPOM As IJDPOM
   Dim oCallBack As IJCheckMfctyCallback
   Dim nItemProcessed As Long
   Dim nItemCount As Long
   
   nItemProcessed = 0
   nItemCount = oItemMonikers.Count
   
   Set oCallBack = oCallingObject
   Set oPOM = GetPOM("Model")
   
   For Each vMoniker In oItemMonikers
      Set oElement = oPOM.GetObject(vMoniker)
            
      If TypeOf oElement Is IJStructFeature Then
         Dim oStructFeature As IJStructFeature
         Dim eFeatureType As StructFeatureTypes
                  
         Set oStructFeature = oElement
         eFeatureType = oStructFeature.get_StructFeatureType
         If eFeatureType = SF_Slot Then
            Dim bInError As Boolean
            Dim sMsgToBePosted As String
            
            CheckSlot oStructFeature, _
                      bInError, _
                      sMsgToBePosted
            
            If bInError = True Then
               '
               ' The arguments for IJCheckMfctyCallback.OnCheckError are:
               '   oStructFeature         Slot in consistency
               '   m_sRuleProgID          ProgID of this rule
               '   Severity of error      (101 error or 102 warning)
               '   105                    Rule type of struct detailing check manufacturability
               '                          (M:\ShipCatalogData\BulkLoad\DataFiles\ShipCheckMfctyRules.xls)
               '   sMsgToBePosted         Message to be posted
               '   "COM Repair Object"    Type of repair to be made
               '   "SMCheckMfctySlotAngle.RepairSlot"   Progid of the COM object that will make the repair
               '
               oCallBack.OnCheckError oStructFeature, _
                                      m_sRuleProgID, _
                                      102, _
                                      105, _
                                      sMsgToBePosted, _
                                      "COM Repair Object", _
                                      "SMCheckMfctySlotAngle.Repair"

               If bStopOnError = True Then
                  Exit For
               End If
            End If
         End If
         Set oStructFeature = Nothing
      End If
      Set oElement = Nothing
      
      nItemProcessed = nItemProcessed + 1
      ' Update the progress bar.
      ' First argument is percentage(0-100) of processed items
      oCallBack.OnProgress (nItemProcessed * 100) / nItemCount, 0
   Next
   Set oPOM = Nothing
   Set oCallBack = Nothing
   
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, m_sModule, sMethod, "Error").Number

End Sub

Private Sub CheckSlot(ByVal oSlot As Object, _
                      ByRef bInError As Boolean, _
                      ByRef sMsgToBePosted As String)
   Const sMethod As String = "CheckSlot"
   On Error GoTo ErrorHandler
      
   Dim sUserInterfaceName As String
   Dim eValueType As EResultType
   Dim sTableName As String
   Dim nCodeListValue As Long
   Dim sShortDesc As String
   Dim sLongDesc As String
   Dim dOcOrientX As Double
   Dim dOcOrientY As Double
   Dim dOcOrientZ As Double
   Dim oSDHelper As New StructDetailObjects.Helper
   
   bInError = False
   sMsgToBePosted = "Unknown"

   ' If occurrence slot angle is overridden, no need to check
   
   ' Get assemby orientation stored on occurrence
   sUserInterfaceName = "IJUASlotAssyOrientation"
   oSDHelper.GetCustomAttributeParamValue _
                           oSlot, _
                           sUserInterfaceName, _
                           "AssyOrientationX", _
                           dOcOrientX, _
                           eValueType, _
                           sTableName, _
                           nCodeListValue, _
                           sShortDesc, _
                           sLongDesc
   
   oSDHelper.GetCustomAttributeParamValue _
                           oSlot, _
                           sUserInterfaceName, _
                           "AssyOrientationY", _
                           dOcOrientY, _
                           eValueType, _
                           sTableName, _
                           nCodeListValue, _
                           sShortDesc, _
                           sLongDesc
   
   oSDHelper.GetCustomAttributeParamValue _
                           oSlot, _
                           sUserInterfaceName, _
                           "AssyOrientationZ", _
                           dOcOrientZ, _
                           eValueType, _
                           sTableName, _
                           nCodeListValue, _
                           sShortDesc, _
                           sLongDesc
   Set oSDHelper = Nothing
   
   ' Get assembly orientation
   Dim oAssemblyChild As IJAssemblyChild
   Dim oAssemblyOrientation As IJAssemblyOrientation
   
   Dim dAssyOrientX As Double
   Dim dAssyOrientY As Double
   Dim dAssyOrientZ As Double
   
   Dim oSDSlotWrapper As New StructDetailObjects.Slot
   Dim oPenetrated As Object
   
   Set oSDSlotWrapper.Object = oSlot
   Set oPenetrated = oSDSlotWrapper.Penetrated
   
   If TypeOf oPenetrated Is IJAssemblyChild Then
      Dim oAssembly As Object
      Dim oViewMatrix As IJDT4x4
      
      Set oAssemblyChild = oPenetrated
      If Not oAssemblyChild Is Nothing Then
         Set oAssembly = oAssemblyChild.Parent
         Set oAssemblyChild = Nothing
         Set oAssemblyOrientation = oAssembly
         Set oAssembly = Nothing
         
         If Not oAssemblyOrientation Is Nothing Then
            ' The part was already in an assembly
            Set oViewMatrix = oAssemblyOrientation.ViewMatrix
            Set oAssemblyOrientation = Nothing
            
            dAssyOrientX = oViewMatrix.IndexValue(8)
            dAssyOrientY = oViewMatrix.IndexValue(9)
            dAssyOrientZ = oViewMatrix.IndexValue(10)
            Set oViewMatrix = Nothing
            
            If Abs(dOcOrientX - dAssyOrientX) > 0.0001 Or _
               Abs(dOcOrientY - dAssyOrientY) > 0.0001 Or _
               Abs(dOcOrientZ - dAssyOrientZ) > 0.0001 Then
               bInError = True
               
               sMsgToBePosted = "Click 'Repair' to fix the inconsistency between:" & vbNewLine & _
                                "Slot occurrence assembly orientation ( " & _
                                Format(dOcOrientX, "0.000") & ", " & _
                                Format(dOcOrientY, "0.000") & ", " & _
                                Format(dOcOrientZ, "0.000") & " )" & vbNewLine & _
                                "and Assembly orientation ( " & _
                                Format(dAssyOrientX, "0.000") & ", " & _
                                Format(dAssyOrientY, "0.000") & ", " & _
                                Format(dAssyOrientZ, "0.000") & " )"
            End If
         End If
      End If
   End If
   Set oPenetrated = Nothing
   
   Exit Sub
   
ErrorHandler:
   Err.Raise LogError(Err, m_sModule, sMethod, "Error").Number

End Sub

