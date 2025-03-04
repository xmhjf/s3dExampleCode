VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MfgTabMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2010, Intergraph Corporation.  All rights reserved.
'
'  File: MfgTabMigration.cls
'
'  Abstract: The file contains an implementation of the default split migration rule
'            for the MfgTab in StructManufacturing.
'
'  History:
'     Suma Mallena  16/11/10    Initial creation.
'
'***************************************************************************
Option Explicit
Implements IJMfgSplitMigration
Private Const MODULE = "MfgTabMigration"

Private Function IJMfgSplitMigration_MigrateObject(ByVal pOldObj As Object, ByVal pReplacingObjColl As GSCADStructMfgGlobals.IJElements, ByVal pMfgObj As Object, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
'Arguments:
' pOldObj ==>  Plate "P1" which is being deleted.
' pReplacingObjColl ==> Collection of replacing new plates ( P2 and P3 ).
' pbCreateMfg ==>   Boolean flag indicating if software should automatically create MfgTabs for the new plates.
' Output of this routine will be one of the plate in pReplacingObjColl collection.
' NOTE:  The rule is not limited to two plates.  "n" number of plates could be passed to the rule after a split.  The rule would still only return one plate meeting the criteria.
'
' pbCreateMfg   Output of MigrateObject()   Action
' ===========   =========================   ======
' False         Nothing           MfgTab ( M1 ) will be deleted. NO new MfgTabs will be created.
' True          Nothing           MfgTab ( M1 ) will be deleted. new MfgTabs will be created for all the new plates ( for P2 and P3 ).
' False         Object P2         MfgTab ( M1 ) will be re-connected from P1 to P2. new MfgTabs will NOT be created for the remaining new plates ( for P3 ).
' True          Object P2         MfgTab ( M1 ) will be re-connected from P1 to P2. new MfgTabs will be created for the remaining new plates ( for P3 ).
' ===========   =========================   ======
    Const METHOD = "IJMfgSplitMigration_MigrateObject"
    On Error GoTo ErrorHandler
    
    Dim oObject As Object
    Dim bOverlapByDrivenPort As Boolean
    Dim bOverlapByDrivingPort As Boolean
    Dim oRuleHlpr As IJMfgRuleHelpersSupport
    Set oRuleHlpr = New MfgRuleHelpersSupport
    Dim oPortColl As IJElements
    Dim oPort As IJPort
    
    Dim oStrMfgTab As IJStrMfgTab
    Dim oPlatePart As Object
    Dim oDrivenPorts As IJElements
    Dim oDrivingPorts As IJElements
    Dim oPartItem As IJDPart
    Dim oDrivenPort As IJPort
    Dim oDrivingPort As Object
    Dim oTabPoint As IJPoint

   If TypeOf pOldObj Is IJPlatePart Then 'This is For Plate part
        Dim dMaxEdgeLength          As Double
        Dim oEdgePortObj            As IJPort
        
        Set oStrMfgTab = pMfgObj
        oStrMfgTab.GetTabInputs oPlatePart, oDrivenPorts, oDrivingPorts, oPartItem, oTabPoint
        
        Dim oIJDAttr As IJDAttribute
        Set oIJDAttr = GetAttribute(oPartItem, "IJUASMPlateTabType", "TabType")
        
        Dim lTabType As Long
        lTabType = oIJDAttr.Value
        
        If (Not oDrivenPorts Is Nothing) And (Not oDrivingPorts Is Nothing) Then
        
            If pReplacingObjColl.Contains(pOldObj) Then 'Partial (Restart) split case
                
                Set oPortColl = oDrivenPorts
            
                If oPortColl.Count = 1 Then
                    Set oPort = oPortColl.Item(1)
                    Set IJMfgSplitMigration_MigrateObject = oPort.Connectable
                End If
              
            Else 'direct split case
                Dim oPortLapCol As IJElements
                Set oPortLapCol = New JObjectCollection
                
                If lTabType = 4 Then 'tab along edge '((eSelMethod = SMTabCorner) Or (eSelMethod = SMTabMid) Or (eSelMethod = SMTabReferenceObject)) Then
                    For Each oObject In pReplacingObjColl
                        Dim oTempPlatePart As IJPlatePart
                        Set oTempPlatePart = oObject
                        
                        Dim oGeomHelper As New MfgGeomHelper
                        
                        Dim oPlateSurface As IJSurfaceBody
                        Set oPlateSurface = oGeomHelper.GetSurfaceFromPlatePart(oTempPlatePart, True)
                        
                        Dim oModelBody As IJDModelBody
                        Set oModelBody = oPlateSurface
                        
                        Dim oTabPosition As IJDPosition
                        Set oTabPosition = New DPosition
                        
                        Dim dx As Double, dy As Double, dz As Double
                        oTabPoint.GetPoint dx, dy, dz
                        
                        oTabPosition.Set dx, dy, dz
                        
                        Dim oDummyPos As IJDPosition
                        Dim dDistFromPlate As Double
                        oModelBody.GetMinimumDistanceFromPosition oTabPosition, oDummyPos, dDistFromPlate
                        
                        If dDistFromPlate < 0.5 Then
                            oPortLapCol.Add oObject
                            Exit For
                        End If
                    Next
                Else
                    Set oDrivenPort = oDrivenPorts.Item(1)
                    Set oDrivingPort = oDrivingPorts.Item(1)
                    
                    For Each oObject In pReplacingObjColl
                        bOverlapByDrivenPort = oRuleHlpr.RangesOverlap(oDrivenPort, oObject)
                        bOverlapByDrivingPort = oRuleHlpr.RangesOverlap(oDrivingPort, oObject)
                        If (bOverlapByDrivenPort = True) And (bOverlapByDrivingPort = True) Then
                            oPortLapCol.Add oObject
                        End If
                        Set oObject = Nothing
                    Next
                End If
                
                If Not oPortLapCol Is Nothing Then
                    If oPortLapCol.Count = 1 Then
                        Set IJMfgSplitMigration_MigrateObject = oPortLapCol.Item(1)
                    End If
                End If
            End If
        End If
    End If
        
    Set oObject = Nothing
    Set oRuleHlpr = Nothing
    ' no settings to clone
    pOptionalArg = False
        
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 9002, , "RULES")
End Function

Private Function IJMfgSplitMigration_ReverseMigrate(ByVal pReplacedPartColl As GSCADStructMfgGlobals.IJElements, ByVal pReplacingPart As Object, ByVal pMfgObjsColl As GSCADStructMfgGlobals.IJElements, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
Const METHOD = "IJMfgSplitMigration_ReverseMigrate"
On Error GoTo ErrorHandler
    
    pbCreateMfg = False ' set this to true only if migration is not needed but want to create a new MfgTab
    pOptionalArg = False ' this makes send to clone the settings only when the above the flag is true
    
   ' Initialization.
    Dim indMfgTab           As Long
    Dim oStrMfgTab          As IJStrMfgTab

    ' Loop for each element in the manufactruing objects collection.
    For indMfgTab = 1 To pMfgObjsColl.Count

        Set oStrMfgTab = pMfgObjsColl.Item(indMfgTab)
        Set IJMfgSplitMigration_ReverseMigrate = oStrMfgTab

    Next indMfgTab

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 9002, , "RULES")
End Function

Private Function GetAttribute(pObject As Object, oAttributeIID As Variant, strAttributeName As String) As IJDAttribute
Const METHOD = "GetAttribute"
On Error GoTo ErrorHandler
    
    Dim oAttributes     As IJDAttributes
    Set oAttributes = pObject
    
    Dim oAttributesCol  As IJDAttributesCol
    Set oAttributesCol = oAttributes.CollectionOfAttributes(oAttributeIID)
    Set oAttributes = Nothing

    Dim i               As Integer
    Dim oAttribute      As IJDAttribute
    
    For i = 1 To oAttributesCol.Count
        Set oAttribute = oAttributesCol.Item(i)
        If oAttribute.AttributeInfo.Name = strAttributeName Then
            Set GetAttribute = oAttribute
            Exit For
        End If
    Next i
    
    Set oAttributesCol = Nothing
    Set oAttribute = Nothing

    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 9002, , "RULES")
End Function
