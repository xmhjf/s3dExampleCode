VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarginMigration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements IJMfgSplitMigration
Private Const MODULE = "MarginMigration"

Private Function IJMfgSplitMigration_MigrateObject(ByVal pOldObj As Object, ByVal pReplacingObjColl As GSCADStructMfgGlobals.IJElements, ByVal pMfgObj As Object, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
'Arguments:
' pOldObj ==>  Plate "P1" which is being deleted.
' pReplacingObjColl ==> Collection of replacing new plates ( P2 and P3 ).
' pbCreateMfg ==>   Boolean flag indicating if software should automatically create Margins for the new plates.
' Output of this routine will be one of the plate in pReplacingObjColl collection.
' NOTE:  The rule is not limited to two plates.  "n" number of plates could be passed to the rule after a split.  The rule would still only return one plate meeting the criteria.
'
' pbCreateMfg   Output of MigrateObject()   Action
' ===========   =========================   ======
' False         Nothing           Margin ( M1 ) will be deleted. NO new Margins will be created.
' True          Nothing           Margin ( M1 ) will be deleted. new Margins will be created for all the new plates ( for P2 and P3 ).
' False         Object P2         Margin ( M1 ) will be re-connected from P1 to P2. new Margins will NOT be created for the remaining new plates ( for P3 ).
' True          Object P2         Margin ( M1 ) will be re-connected from P1 to P2. new Margins will be created for the remaining new plates ( for P3 ).
' ===========   =========================   ======
    Const METHOD = "IJMfgSplitMigration_MigrateObject"
    On Error GoTo ErrorHandler
    
    Dim oObject As Object
    Dim bOverlap As Boolean
    Dim oRuleHlpr As IJMfgRuleHelpersSupport
    Set oRuleHlpr = New MfgRuleHelpersSupport
    Dim oMfgDefinition As IJMfgDefinition
    Dim oMarginPort As IJPort
    Dim oPortColl As IJElements
    Dim oPort As IJPort
    Dim oEdgeLength As Double
    Dim oEntityHelper As IJMfgEntityHelper
    
    If TypeOf pOldObj Is IJStructProfilePart Then
        If pReplacingObjColl.Contains(pOldObj) Then 'Partial split case
            Set oPortColl = GetMarginPortCollection(pMfgObj)
            
            If oPortColl.Count = 0 Then
                Set IJMfgSplitMigration_MigrateObject = Nothing
                Exit Function
            End If
            
            If oPortColl.Count = 1 Then
                Set oPort = oPortColl.Item(1)
                Set IJMfgSplitMigration_MigrateObject = oPort.Connectable
            Else
                'not a possible case
            End If
        Else 'Direct split case
        
            Set oMfgDefinition = pMfgObj
            Set oMarginPort = oMfgDefinition.GetPort
              
            If oMarginPort Is Nothing Then
                Set IJMfgSplitMigration_MigrateObject = Nothing
                Exit Function
            End If
            
            For Each oObject In pReplacingObjColl
                bOverlap = oRuleHlpr.RangesOverlap(oMarginPort, oObject)
                If (bOverlap = True) Then
                    Set IJMfgSplitMigration_MigrateObject = oObject
                End If
            Next
        End If
    Else 'This is For Plate part
        Dim dMaxEdgeLength          As Double
        Dim oEdgePortObj            As IJPort
        
        If pReplacingObjColl.Contains(pOldObj) Then 'Partial split case
            Set oPortColl = GetMarginPortCollection(pMfgObj)
            
            If oPortColl.Count = 0 Then
                Set IJMfgSplitMigration_MigrateObject = Nothing
                Exit Function
            End If
            
            If oPortColl.Count = 1 Then
                Set oPort = oPortColl.Item(1)
                Set IJMfgSplitMigration_MigrateObject = oPort.Connectable
            Else
                pbCreateMfg = True
                
                'Partial migration case
                'Get lengthiest among the ports
                Set oEntityHelper = New MfgEntityHelper
                dMaxEdgeLength = -1# 'maximum EdgeLength
                
                Dim oLengthiestPort As IJPort
                For Each oPort In oPortColl
                    Set oEdgePortObj = oEntityHelper.GetEdgePortGivenFacePort(oPort, CTX_BASE)
                    If (Not oEdgePortObj Is Nothing) Then
                        oEdgeLength = EdgePortlength(oEdgePortObj)
                        
                        If oEdgeLength > dMaxEdgeLength Then
                            ' If so, store this length as maximum and the part index
                            dMaxEdgeLength = oEdgeLength
                            Set oLengthiestPort = oEdgePortObj
                        End If
                    End If
                Next
                'return the Connectable of lengthiest
                Set IJMfgSplitMigration_MigrateObject = oLengthiestPort.Connectable
            End If
        Else 'direct split case
            Dim oFabMargin As IJDFabMargin
            Set oFabMargin = pMfgObj
            
            Dim oOldPortWB As Object
            Dim oPortLapCol As IJElements
            Set oPortLapCol = New JObjectCollection
            
            Set oMfgDefinition = pMfgObj
            Set oMarginPort = oMfgDefinition.GetPort
            
            If oFabMargin.MarginGeometryMode = FullFabricationMargin Then
                Set oOldPortWB = oMarginPort
            Else
                Dim oRegMarginWireBody As IJWireBody
                Set oRegMarginWireBody = GetWireBodyOfRegMargin(oFabMargin)
                Set oOldPortWB = oRegMarginWireBody
            End If
            
            For Each oObject In pReplacingObjColl
                bOverlap = oRuleHlpr.RangesOverlap(oOldPortWB, oObject)
                If (bOverlap = True) Then
                    oPortLapCol.Add oObject
                End If
                Set oObject = Nothing
            Next
             
            Dim lMaxEdgeLengthIndex     As Long
            ' Initialization.
            dMaxEdgeLength = -1# 'maximum EdgeLength
            lMaxEdgeLengthIndex = 0 'index of the max edgelength part in pReplacingObjColl
             
            If Not oPortLapCol Is Nothing Then
                If oPortLapCol.Count = 1 Then
                    Set IJMfgSplitMigration_MigrateObject = oPortLapCol.Item(1)
                    Exit Function
                End If
                
                pbCreateMfg = True
    
                Dim indPlate As Long
                For indPlate = 1 To oPortLapCol.Count
                     Dim oPlatePartSupport As IJPlatePartSupport
                     Set oPlatePartSupport = New PlatePartSupport
                     Dim oPartSupport As IJPartSupport
                     Set oPartSupport = oPlatePartSupport
                     Set oPartSupport.Part = oPortLapCol.Item(indPlate)
                      
                    If TypeOf oPartSupport.Part Is IJPlatePart Then
                      
                        Dim oConnectable As IJConnectable
                        Set oConnectable = oPartSupport.Part
                          
                        Dim oEnumPortsColl As IJElements
                        Set oEnumPortsColl = New JObjectCollection
                          
                        oConnectable.enumPorts oEnumPortsColl, PortFace
                        Dim oFacePortObj As IJPort
                        Dim indPort As Long
                                                                                                                                           
                        If Not oEnumPortsColl Is Nothing Then
                            For indPort = 1 To oEnumPortsColl.Count
                                Set oFacePortObj = oEnumPortsColl.Item(indPort)
                                Dim bDirectionOverLap As Boolean
                                     
                                bDirectionOverLap = oRuleHlpr.ComparePortAttributes(oFacePortObj, oMarginPort)
                                If (bDirectionOverLap = True) Then
                                    Set oEntityHelper = New MfgEntityHelper
                                    Set oEdgePortObj = oEntityHelper.GetEdgePortGivenFacePort(oFacePortObj, CTX_BASE)
                                    If (Not oEdgePortObj Is Nothing) Then
                                        oEdgeLength = EdgePortlength(oEdgePortObj)
                                        
                                        If oEdgeLength > dMaxEdgeLength Then
                                            ' If so, store this length as maximum and the part index
                                            dMaxEdgeLength = oEdgeLength
                                            lMaxEdgeLengthIndex = indPlate
                                        End If
                                    End If
                                End If
                            Next indPort
                        End If
                    End If
                Next indPlate
                'Return the part which is having the maximum edge length.
                Set IJMfgSplitMigration_MigrateObject = oPortLapCol.Item(lMaxEdgeLengthIndex)
            End If
        End If
    End If
        Set oMarginPort = Nothing
        Set oObject = Nothing
        Set oRuleHlpr = Nothing
        Set oConnectable = Nothing
        Set oFacePortObj = Nothing
        Set oEdgePortObj = Nothing
        Set oEntityHelper = Nothing

        ' no settings to clone
        pOptionalArg = False
        
        Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 6009, , "RULES")
End Function
Private Function EdgePortlength(oPort As IJPort) As Double
Const METHOD = "EdgePortlength"

   On Error GoTo ErrorHandler
   
   EdgePortlength = 0#
   Dim oWireBody As IJWireBody
   Dim oGeometry As Object
   Set oGeometry = oPort.Geometry
   If TypeOf oGeometry Is IJWireBody Then
      Set oWireBody = oGeometry
      EdgePortlength = LengthWireBody(oWireBody)
   End If
   
CleanUp:
   Set oWireBody = Nothing
   Set oGeometry = Nothing

   Exit Function
   
ErrorHandler:
    Err.Raise Err.Number
End Function
Private Function LengthWireBody(ByRef oWB As IJWireBody) As Double
Const METHOD = "LengthWireBody"

   Dim anMfgMGHelper As IJMfgMGHelper
   Dim oCS As IJComplexString
   
   Set anMfgMGHelper = New GSCADMathGeom.MfgMGHelper
   
   anMfgMGHelper.WireBodyToComplexString oWB, oCS
   LengthWireBody = LengthComplexString(oCS)

CleanUp:
   Set anMfgMGHelper = Nothing
   Set oCS = Nothing
   
   Exit Function
   
ErrorHandler:
    Err.Raise Err.Number
End Function
Private Function LengthComplexString(ByRef oCS As IJComplexString) As Double
Const METHOD = "LengthComplexString"

   On Error GoTo ErrorHandler
   
   LengthComplexString = 0#
   Dim aCurve As IJCurve
   Dim ppDisp As Object
   Dim lIndex As Long
   For lIndex = 1 To oCS.CurveCount
      Set aCurve = Nothing
      Set ppDisp = Nothing
      oCS.GetCurve lIndex, ppDisp
      If TypeOf ppDisp Is IJCurve Then
         Set aCurve = ppDisp
         LengthComplexString = LengthComplexString + aCurve.Length
      End If
   Next

CleanUp:
   Set aCurve = Nothing
   Set ppDisp = Nothing
   
   Exit Function
   
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function IJMfgSplitMigration_ReverseMigrate(ByVal pReplacedPartColl As GSCADStructMfgGlobals.IJElements, ByVal pReplacingPart As Object, ByVal pMfgObjsColl As GSCADStructMfgGlobals.IJElements, pbCreateMfg As Boolean, Optional pOptionalArg As Boolean) As Object
    Const METHOD = "IJMfgSplitMigration_ReverseMigrate"
    On Error GoTo ErrorHandler
    
    'Return the margin with largest port length
        
    pbCreateMfg = False ' set this to true only if migration is not needed but want to create a new margin
    pOptionalArg = False ' this makes send to clone the settings only when the above the flag is true
    
    ' Initialization.
    Dim indMargin           As Long
    Dim dMaxLength     As Double
    Dim lMaxIndex    As Long
    
    dMaxLength = -1# 'maximum surface area
    lMaxIndex = 0 'index of the biggest plate part in pReplacingObjColl

    'Create entity support helper to get the base edge port from the face port
    Dim oEntityHelper   As IJMfgEntityHelper
    Set oEntityHelper = New MfgEntityHelper
        
    Dim oEdgePortObj    As IJPort
    
    ' Loop for each element in the manufactruing objects collection.
    For indMargin = 1 To pMfgObjsColl.Count
        
        Dim oFabMargin As IJMfgDefinition
        Set oFabMargin = pMfgObjsColl.Item(indMargin)
        
        Dim oFacePortObj   As IJPort
        Set oFacePortObj = oFabMargin.GetPort
        
        If Not oFacePortObj Is Nothing Then
            'Does the port has similar port on new part
            Dim bSimilarPort As Boolean
            If TypeOf oFacePortObj.Connectable Is IJPlatePart Then
                Set oEdgePortObj = oEntityHelper.GetEdgePortGivenFacePort(oFacePortObj, CTX_BASE)
                If (Not oEdgePortObj Is Nothing) Then
                
                    Dim oEdgeLength         As Double
                    oEdgeLength = EdgePortlength(oEdgePortObj)
                    
                    If oEdgeLength > dMaxLength Then
                        ' If so, assume this as the Margin with maximum length
                        dMaxLength = oEdgeLength
                        Set IJMfgSplitMigration_ReverseMigrate = oFabMargin
                    End If
                    
                    Set oEdgePortObj = Nothing
                    Set oFacePortObj = Nothing
                    Set oFabMargin = Nothing
                End If
            Else 'Type is IJPorfilePart
                Set IJMfgSplitMigration_ReverseMigrate = oFabMargin
            End If
        End If
    Next indMargin
    
    Set oEntityHelper = Nothing
       
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 6009, , "RULES")
End Function
'This method with the idea that Margin may have more than one ports( in partial split case)
Private Function GetMarginPortCollection(oMargin As Object) As IJElements
Const METHOD = "GetMarginPortCollection"
   On Error GoTo ErrorHandler
        
        Set GetMarginPortCollection = New JObjectCollection
        Dim strRoleName As String
        strRoleName = "StrDetailPort_DEST"
                
        Dim oIJDAssocRelation As IJDAssocRelation
        Dim oIJDTargetObjectColl As IJDTargetObjectCol
        
        Set oIJDAssocRelation = oMargin
        Set oIJDTargetObjectColl = oIJDAssocRelation.CollectionRelations _
                                ("IJMfgDefinition", strRoleName)
        
        Dim iPortIndex As Long
        For iPortIndex = 1 To oIJDTargetObjectColl.Count
            Dim oPort As Object
            Set oPort = oIJDTargetObjectColl.Item(iPortIndex)
            GetMarginPortCollection.Add oPort
        Next
                
CleanUp:
    Exit Function
   
ErrorHandler:
    Err.Raise Err.Number
End Function

Private Function GetWireBodyOfRegMargin(oFabMargin As IJDFabMargin) As IJWireBody
    
    Dim oPart As Object
    Dim oTempPort As IJPort
    Dim oPort As IJPort

    Dim CreationType As MarginCreation
    Dim eMode As EnumMarginMode
    Dim lMarginType As Long
    Dim lDirection As Long
    Dim lGeometryChange As Long
    Dim dStretchOffset As Double
    Dim dStartValue As Double
    Dim dEndvalue As Double
    Dim enumIntersectionMethodStart As EnumMarginSelectionMethod
    Dim enumIntersectionMethodEnd As EnumMarginSelectionMethod
    
    Dim oStartPos As IJDPosition
    Dim oEndPos As IJDPosition
    
    Dim oEntityHelper As MfgEntityHelper
    Set oEntityHelper = New MfgEntityHelper
    
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    oFabMargin.GetFabMarginInputs oPart, oTempPort, CreationType, eMode, lMarginType, _
                            lDirection, lGeometryChange, dStretchOffset, dStartValue, dEndvalue, _
                            enumIntersectionMethodStart, enumIntersectionMethodEnd, Nothing, _
                            Nothing, oStartPos, oEndPos
                            
    Set oPort = oEntityHelper.GetEdgePortGivenFacePort(oTempPort, CTX_BASE)
    
    Dim oRelatedPort As IJPort
    Set oRelatedPort = oEntityHelper.GetRelatedPort(oPort)
    
    Dim oWireBody As IJWireBody
    Set oWireBody = oRelatedPort.Geometry
    
    Dim oPortCs As IJComplexString
    
    oMfgMGHelper.WireBodyToComplexString oWireBody, oPortCs
    oMfgMGHelper.TrimCurveByPoints oPortCs, oStartPos, oEndPos
    
    Dim oTrimmedWireBody As IJWireBody
    oMfgMGHelper.ComplexStringToWireBody oPortCs, oTrimmedWireBody
    
    Set GetWireBodyOfRegMargin = oTrimmedWireBody
    
CleanUp:
    Set oPart = Nothing
    Set oPort = Nothing
    Set oTempPort = Nothing
    Set oStartPos = Nothing
    Set oEndPos = Nothing
    Set oPortCs = Nothing
    Set oWireBody = Nothing
    
End Function
