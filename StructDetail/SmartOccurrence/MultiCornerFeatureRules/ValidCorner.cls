VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ValidCorners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const m_sClassName As String = "ValidCorners"
Const m_sProjectPath = "MultiCornerFeatureRule"
Const MODULE = m_sProjectPath + m_sClassName + ".cls"

'*********************************************************************************************
'  Copyright (C) 2016, Intergraph Corporation.  All rights reserved.
'
'  Project     : MultiCornerFeatureRule
'  File        : ValidCorners.cls
'
'  Description :
'
'  Author      :
'
'  History     :
'               - Created
'   05-Feb-2016  -Modified -GHM- TR-287173 IsCornerValid() is updated with proper validation null checks
'                               and enumPorts is called on IJConnectable instead of IJStructConnectable
'*********************************************************************************************

Public Function IsCornerValid(ByVal oFirstPort As IJPort, ByVal oSecondPort As IJPort) As Boolean

    Const METHOD = "::IsCornerValid"
    On Error GoTo ErrorHandler
    
    If oFirstPort Is Nothing Or oSecondPort Is Nothing Then
        Exit Function
    End If
    
    Dim sMsg As String

    Dim ValidCorner As Boolean
    Dim oAsmParent As IJAssembly
    Dim oPart As Object
    
    Set oPart = oFirstPort.Connectable

    ValidCorner = True
    
    If TypeOf oPart Is IJPlate Then
            Dim oPlate As New StructDetailObjects.PlatePart
            Set oPlate.object = oPart
    End If
    
    Dim oFirstStructPort As IJStructPort
    Dim oSecondStructPort As IJStructPort
    Set oFirstStructPort = oFirstPort
    Set oSecondStructPort = oSecondPort
    Dim pObjThatCreatedFirstPort As IUnknown
    Dim pObjThatCreatedSecondPort As IUnknown
    Dim pOperation As IJStructOperation
    Dim oSDUtils As New StructDetailHelper
   
   ' Check if the plate tightness is set to true if it is then it is
    ' not a valid corner
    Dim BOOL As Boolean
    If oPlate.Tightness = NonTight Then
        ValidCorner = True
    Else
        ValidCorner = False
    End If
    Dim oStructConnectable As IJStructConnectable
    Dim oCollection As IJElements
    Dim oConnectableStructPort As IJStructPort
    
    ' Creating StructFilter, built in ShipStructure\Client, as an object to remove
    ' ShipStructure dependency
    Dim oStructFilter As Object
    Set oStructFilter = CreateObject("STFilter.StructFilter")
    
    Dim ValidFirstEdge As Boolean
    Dim ValidSecondEdge As Boolean
                
    ' Get the object that created that port and see if it is seam and set the
    ' validity based on the type of the seam
    oSDUtils.FindOperatorForOperationInGraphByID oPart, _
                                               oFirstStructPort.OperationID, _
                                               oFirstStructPort.OperatorID, _
                                               pOperation, _
                                               pObjThatCreatedFirstPort
                                               
    
    Dim pFirstseamType As IJDSeamType
    Dim FirstStructSeam As StructSeamType
    FirstStructSeam = -1
    If Not pObjThatCreatedFirstPort Is Nothing Then
        If TypeOf pObjThatCreatedFirstPort Is IJSeam Then
             Set pFirstseamType = pObjThatCreatedFirstPort
             FirstStructSeam = pFirstseamType.SeamType
        End If
    End If
    
    oSDUtils.FindOperatorForOperationInGraphByID oPart, _
                                               oSecondStructPort.OperationID, _
                                               oSecondStructPort.OperatorID, _
                                               pOperation, _
                                               pObjThatCreatedSecondPort
                                               
    
    Dim pSecondseamType As IJDSeamType
    Dim SecondStructSeam As StructSeamType
    SecondStructSeam = -1
    If Not pObjThatCreatedSecondPort Is Nothing Then
        If TypeOf pObjThatCreatedSecondPort Is IJSeam Then
             Set pSecondseamType = pObjThatCreatedSecondPort
             SecondStructSeam = pSecondseamType.SeamType
        End If
    End If
    
    ' From the plate part get to the plate system
         
    Dim oPlateSystem As IJPlateSystem
    Dim oParentSystem As Object
    
    Set oParentSystem = oPlate.ParentSystem
    
    If Not oParentSystem Is Nothing Then
        If TypeOf oParentSystem Is IJPlateSystem Then
            Set oPlateSystem = oParentSystem
        End If
    End If
    
    
    If Not oPlateSystem Is Nothing Then
        Dim oSystemChild As IJSystemChild
        Set oSystemChild = oPlateSystem
        While TypeOf oSystemChild.GetParent Is IJPlateSystem
            Set oSystemChild = oSystemChild.GetParent
        Wend
             
        Set oStructConnectable = oSystemChild
       
       
       ' Get all the global edge ports of the main plate system in to a collection
       oStructConnectable.enumConnectableTransientPorts oCollection, vbNullString, False, PortEdge, JS_TOPOLOGY_FILTER_NONE, True
    
       Dim FirstEdgeER As Boolean
       Dim SecondEdgeER As Boolean
       Dim oSystem As IJSystem
            Dim oChildPlates As IJDTargetObjectCol
        Set oSystem = oSystemChild
        Set oChildPlates = oSystem.GetChildren()
        Dim Child As Object
        Dim ind As Integer
        Dim ERCollection As New collection
        Dim oERSys As IJERSystem
    
        Dim oPartInfo As PartInfo
        Dim x1 As Double
        Dim y1 As Double
        Dim z1 As Double
        Dim x2 As Double
        Dim y2 As Double
        Dim z2 As Double
        Dim x3 As Double
        Dim y3 As Double
        Dim z3 As Double
        Dim x4 As Double
        Dim y4 As Double
        Dim z4 As Double
        Dim x5 As Double
        Dim y5 As Double
        Dim z5 As Double
        Dim x6 As Double
        Dim y6 As Double
        Dim z6 As Double
        Dim tol As Double
        tol = 0.01
        Set oPartInfo = New PartInfo
                               
        Dim objConnectable As IJStructPort
        Dim oERSystem As IJERSystem
        Dim ERConnectable As IJConnectable
        ' Traverse through all the children of the plate and see if there are any
        ' Edge Reinforcements in it, if so add them to a collection
        
        Dim ERPortColl As IJElements
        
        For ind = 1 To oChildPlates.Count
            On Error Resume Next
            Set Child = oChildPlates.Item(ind)
            Err.Clear
            On Error GoTo ErrorHandler
            
            If Not Child Is Nothing Then
                If TypeOf Child Is IJERSystem Then
                    ERCollection.Add Child
                End If
            End If
            Set Child = Nothing
        Next
        
        ' for each port in the port collection of the plate system see if any of its
        ' ports are same as the ones that are returned by the corner feature based on
        ' the Operation ID and operator ID. This was done because the ports that are
        ' coming back from the corner feature were face ports and we need the corresponding
        ' edge ports to see if they are free edges and if there are any edge reinforcements
        ' on that edge.
        Dim oTransPort As IJTransientPort
        Dim oPeristentPort As IJStructPort
        For Each oConnectableStructPort In oCollection
            
            If TypeOf oConnectableStructPort Is IJTransientPort Then
                
                Set oTransPort = oConnectableStructPort
                oTransPort.GetPersistentPort oPeristentPort, True
                
                Set oConnectableStructPort = Nothing
                Set oConnectableStructPort = oPeristentPort
                
                Set oTransPort = Nothing
                Set oPeristentPort = Nothing
            End If
            
            If (oConnectableStructPort.OperationID = oFirstStructPort.OperationID And oConnectableStructPort.OperatorID = oFirstStructPort.OperatorID) Then
                Dim FirstFreeEdgePort As IJPort
                Set FirstFreeEdgePort = oConnectableStructPort
                oStructFilter.Clear
                ' Filter based on the FreeEdgePortOnRootObject Filter.
                oStructFilter.AddCriteria "[STFilterFunctions.StructFilterFunctions,IsFreeEdgePortOnRootObject]"
                If oStructFilter.PassCriteria(FirstFreeEdgePort) > 0 Then
                    ValidFirstEdge = True
                End If
                ' Go through all the Edge reinforcements in the collection and get
                ' its ports in a collection, see if the port returned from the corner
                ' feature is with in a tolerance and if so determine that there is
                ' an edge reinforcement on that edge
                For Each oERSys In ERCollection
                    If ValidFirstEdge Then
                        Set ERConnectable = oERSys
                        ERConnectable.enumPorts ERPortColl
                        For Each objConnectable In ERPortColl
                            If (objConnectable.OperationID = oFirstStructPort.OperationID And objConnectable.OperatorID = oFirstStructPort.OperatorID) Then
                                oPartInfo.GetVertexRange oERSys, x1, y1, z1, x2, y2, z2
                                oPartInfo.GetVertexRange oSecondStructPort, x3, y3, z3, x4, y4, z4
                                If ((x1 > x4 - tol And x1 < x4 + tol And y1 > y4 - tol And y1 < y4 + tol And z1 > z4 - tol And z1 < z4 + tol) Or (x2 > x3 - tol And x2 < x3 + tol And y2 > y3 - tol And y2 < y3 + tol And z2 > z3 - tol And z2 < z3 + tol) Or _
                                    (x2 > x4 - tol And x2 < x4 + tol And y2 > y4 - tol And y2 < y4 + tol And z2 > z4 - tol And z2 < z4 + tol) Or (x1 > x3 - tol And x1 < x3 + tol And y1 > y3 - tol And y1 < y3 + tol And z1 > z3 - tol And z1 < z3 + tol)) Then
                                     FirstEdgeER = True
                                    Exit For
                                End If
                            End If
                            Set objConnectable = Nothing
                        Next
                     End If
                     Set ERPortColl = Nothing
                     Set objConnectable = Nothing
                     Set oERSystem = Nothing
                     Set ERConnectable = Nothing
                       
                Next
            End If
        Next
        For Each oConnectableStructPort In oCollection
            If TypeOf oConnectableStructPort Is IJTransientPort Then
                
                Set oTransPort = oConnectableStructPort
                oTransPort.GetPersistentPort oPeristentPort, True
                
                Set oConnectableStructPort = Nothing
                Set oConnectableStructPort = oPeristentPort
                
                Set oTransPort = Nothing
                Set oPeristentPort = Nothing
                
            End If
            If (oConnectableStructPort.OperationID = oSecondStructPort.OperationID And oConnectableStructPort.OperatorID = oSecondStructPort.OperatorID) Then
                Dim SecondFreeEdgePort As IJPort
                Set SecondFreeEdgePort = oConnectableStructPort
                oStructFilter.Clear
                oStructFilter.AddCriteria "[STFilterFunctions.StructFilterFunctions,IsFreeEdgePortOnRootObject]"
                If oStructFilter.PassCriteria(SecondFreeEdgePort) > 0 Then
                    ValidSecondEdge = True
                 End If
               
                For Each oERSys In ERCollection
                    If ValidSecondEdge Then
                        Set ERConnectable = oERSys
                        ERConnectable.enumPorts ERPortColl
                        For Each objConnectable In ERPortColl
                            If (objConnectable.OperationID = oSecondStructPort.OperationID And objConnectable.OperatorID = oSecondStructPort.OperatorID) Then
                                oPartInfo.GetVertexRange oERSys, x1, y1, z1, x2, y2, z2
                                oPartInfo.GetVertexRange oFirstStructPort, x3, y3, z3, x4, y4, z4
                                                     
                                If ((x1 > x4 - tol And x1 < x4 + tol And y1 > y4 - tol And y1 < y4 + tol And z1 > z4 - tol And z1 < z4 + tol) Or (x2 > x3 - tol And x2 < x3 + tol And y2 > y3 - tol And y2 < y3 + tol And z2 > z3 - tol And z2 < z3 + tol) Or _
                                    (x2 > x4 - tol And x2 < x4 + tol And y2 > y4 - tol And y2 < y4 + tol And z2 > z4 - tol And z2 < z4 + tol) Or (x1 > x3 - tol And x1 < x3 + tol And y1 > y3 - tol And y1 < y3 + tol And z1 > z3 - tol And z1 < z3 + tol)) Then
                                 
                                    SecondEdgeER = True
                                    Exit For
                                End If
                            End If
                            Set objConnectable = Nothing
                        Next
                        Set oERSys = Nothing
                    End If
                    
                    Set ERPortColl = Nothing
                    Set objConnectable = Nothing
                    Set oERSystem = Nothing
                    Set ERConnectable = Nothing
                Next
            End If
            Set SecondFreeEdgePort = Nothing
        Next
       
       
        'based on the spec decide if it is valid corner or not
        
          If (FirstStructSeam = sstStrakingSeam) Or _
             (FirstStructSeam = sstDesignSeam) Or _
             (SecondStructSeam = sstStrakingSeam) Or _
             (SecondStructSeam = sstDesignSeam) Or _
             ((FirstStructSeam = sstPlanningSeam) And (SecondStructSeam = sstPlanningSeam)) Or _
             ((ValidFirstEdge And Not FirstEdgeER) Or (ValidSecondEdge And Not SecondEdgeER)) Then
        
                ValidCorner = False
        End If
        Set oChildPlates = Nothing
        Set oSystem = Nothing
        
        Set oPartInfo = Nothing
        Set ERCollection = Nothing
    Else
        'If parent plate system is nothing
        
        On Error GoTo ErrorHandler
        
        Dim oAssocRelation As IJDAssocRelation
        Dim index As Integer
        Dim RelCol As IJDRelationshipCol
        Dim collection As Object
        Dim PortStr As String
        'Changing the string IID as part of DI-CP-217568
        'PortStr = "{5CF7C404-546D-11D2-B328-080036024603}"
        PortStr = "{96eb9676-6530-11d1-977f-080036754203}"
       ' IID_IJGeometry = "{96eb9676-6530-11d1-977f-080036754203}"
        FirstEdgeER = False
        SecondEdgeER = False
        
        Dim oConnectable As IJConnectable
        Set oConnectable = oPart
        
        'Get Edge Ports
        oConnectable.enumPorts oCollection, PortEdge
        
        'Loop through all the Edge ports and find if any ER is associated with the given First Port
        If Not oCollection Is Nothing Then
            For Each oConnectableStructPort In oCollection
                If (oConnectableStructPort.OperationID = oFirstStructPort.OperationID And oConnectableStructPort.OperatorID = oFirstStructPort.OperatorID) Then
                    If TypeOf oConnectableStructPort Is IJDAssocRelation Then
                        Set oAssocRelation = oConnectableStructPort
                        Set collection = oAssocRelation.CollectionRelations(PortStr, "ERDefiningPort_DEST")
                        If Not collection Is Nothing Then
                            Set RelCol = collection
                            If RelCol.Count > 0 Then
                                FirstEdgeER = True
                            End If
                        End If
                    End If
                    Set oConnectableStructPort = Nothing
                    Set oAssocRelation = Nothing
                    Set RelCol = Nothing
                End If
            Next
        End If
        
        Set oCollection = Nothing
        Set oConnectable = oSecondPort.Connectable
        
        'Get Edge Ports
        oConnectable.enumPorts oCollection, PortEdge
        
        'Loop through all the Edge ports and find if any ER is associated with the given First Port
        If Not oCollection Is Nothing Then
            For Each oConnectableStructPort In oCollection
                If (oConnectableStructPort.OperationID = oSecondStructPort.OperationID And oConnectableStructPort.OperatorID = oSecondStructPort.OperatorID) Then
                    If TypeOf oConnectableStructPort Is IJDAssocRelation Then
                        Set oAssocRelation = oConnectableStructPort
                        Set collection = oAssocRelation.CollectionRelations(PortStr, "ERDefiningPort_DEST")
                        If Not collection Is Nothing Then
                            Set RelCol = collection
                             If RelCol.Count > 0 Then
                                SecondEdgeER = True
                             End If
                        End If
                    End If
                    Set oConnectableStructPort = Nothing
                    Set oAssocRelation = Nothing
                     Set RelCol = Nothing
                End If
            Next
        End If
        
        ValidFirstEdge = False
        ValidSecondEdge = False
        
        If Not pObjThatCreatedSecondPort Is Nothing Then
            If Not TypeOf pObjThatCreatedSecondPort Is IJConnectable Then
                ValidSecondEdge = True
            End If
        End If
        
        If Not pObjThatCreatedFirstPort Is Nothing Then
            If Not TypeOf pObjThatCreatedFirstPort Is IJConnectable Then
                ValidFirstEdge = True
            End If
        End If
        
        If ((ValidFirstEdge And Not FirstEdgeER) Or (ValidSecondEdge And Not SecondEdgeER)) Then
          ValidCorner = False
        End If
        
    End If
    
    Set oStructFilter = Nothing
    Set pObjThatCreatedFirstPort = Nothing
    Set pObjThatCreatedSecondPort = Nothing
    Set pOperation = Nothing
    Set oFirstPort = Nothing
    Set oSecondPort = Nothing
 
    IsCornerValid = ValidCorner
    
    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD, sMsg
    
End Function

'*************************************************************************
'Function
'HandleError
'
'Abstract
' called by other subs and fuctions during error. This method logs the error
' and returns success
'
'input
'Module(file) name, method name
'
'Return
'success
'
'Exceptions
'
'***************************************************************************
Private Sub HandleError(sModule As String, sMETHOD As String, Optional sExtraInfo As String = "")
    Dim oEditErrors As IJEditErrors
    
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.AddFromErr Err, sExtraInfo, sMETHOD, sModule
    End If
    Set oEditErrors = Nothing
End Sub

