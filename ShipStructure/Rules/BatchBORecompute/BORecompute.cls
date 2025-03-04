VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BORecompute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements PrjMgtMiddleServicesWrapper.IBatchMiddleCommand

Private m_oMonikers As IEnumMoniker
Private m_oMiddleCtx As IJMiddleContext
Private Const IID_IJStructGeometry As String = "{6034AD40-FA0B-11d1-B2FD-080036024603}"
Private Const IID_IJGeometry As String = "{96EB9676-6530-11d1-977F-080036754203}"
Private Const IID_IJStructOperation = "{1E275841-C22A-11D1-B2E5-080036024603}"
Private Const IID_IJCompartEntity = "{D7973FA7-6ECA-429D-A316-226315C3E5B6}"
Private Const IID_IJCompartGeometry = "{04681CD7-C958-4D57-9B4E-AA064647AAD8}"
Private Const IID_IJCompartSplitAE = "{4C0D36D9-7334-11D2-A494-080036B9C303}"
Private m_oBatchProcessHelper   As S3DBatchProcessHelper
Private m_oBatchProcessLogHelper As S3DBatchProcessLogHelper
Private m_oObjectsToRecompute As IJElements

Private Enum StructEntityTypes
    ENTITY_TYPE_UNKNOWN = 0
    ENTITY_TYPE_DESIGN_SEAM = 1
    ENTITY_TYPE_PROFILE_SYSTEM = 2
    ENTITY_TYPE_REFCURVE = 3
End Enum
'Implement IBatchCommand's Initialize Method
Private Function IBatchMiddleCommand_Initialize(ByVal oResourceMgr As IJDPOM, ByVal JobInfoWrapper As PrjMgtMiddleServicesWrapper.JobInfoWrapper) As Boolean

    IBatchMiddleCommand_Initialize = False
    
    If m_oMiddleCtx Is Nothing Then
        Set m_oMiddleCtx = New GSCADMiddleContext
    End If
    
    Dim strProgramData As String
    strProgramData = JobInfoWrapper.CommandArgsInXMLFormat
    
    'Instantiate object Collection.
    Set m_oObjectsToRecompute = New JObjectCollection
    
    'Instantiate S3D Batch Process Log Helper.
    Set m_oBatchProcessLogHelper = New S3DBatchProcessLogHelper
    m_oBatchProcessLogHelper.LogFileName = "BORecompute"
    
   'Instantiate S3D Batch Process Helper
    Set m_oBatchProcessHelper = New S3DBatchProcessHelper
    
    'To get the objects selected in the batch command
    Dim filters     As IJElements
    Dim assemblies  As IJElements
    Dim objects     As IJElements
    m_oBatchProcessHelper.GetInputs strProgramData, filters, assemblies, objects
    
    'Get Objects in All Filters and add to elements collection.
    If Not filters Is Nothing Then
        If filters.Count > 0 Then
            m_oBatchProcessLogHelper.Log "Batch Input Type : Filters"
            m_oBatchProcessLogHelper.Log "Number of Filters : " & filters.Count
        End If
        
        Dim i As Integer
        For i = 1 To filters.Count
            m_oObjectsToRecompute.AddElements m_oBatchProcessHelper.GetFilterObjects(filters.Item(i))
        Next i
    End If
    
    'Get Objects in All Assemblies add to elements collection.
    If Not assemblies Is Nothing Then
        If assemblies.Count > 0 Then
            m_oBatchProcessLogHelper.Log "Batch Input Type : Assemblies"
            m_oBatchProcessLogHelper.Log "Number of Assemblies : " & assemblies.Count
        End If
    
        Dim j As Integer
        For j = 1 To assemblies.Count
            m_oObjectsToRecompute.AddElements m_oBatchProcessHelper.GetAssemblyObjects(assemblies.Item(j), "", True)
        Next j
    End If
    
   'Add the Objects to Collection.
    If Not objects Is Nothing Then
        If objects.Count > 0 Then
            m_oBatchProcessLogHelper.Log "Batch Input Type : Objects"
            m_oBatchProcessLogHelper.Log "Number of Objects : " & objects.Count
        End If
    
        Dim k As Integer
        For k = 1 To objects.Count
            If oResourceMgr.SupportsInterface(objects.Item(i), "IJAssemblyBase") Then
                m_oObjectsToRecompute.AddElements m_oBatchProcessHelper.GetAssemblyObjects(objects.Item(k), "", True)
            ElseIf oResourceMgr.SupportsInterface(objects.Item(i), "IJSystem") Then
                m_oObjectsToRecompute.AddElements m_oBatchProcessHelper.GetSystemObjects(objects.Item(k))
            Else
                m_oObjectsToRecompute.Add objects.Item(k)
            End If
        Next k
    End If
    
    IBatchMiddleCommand_Initialize = True
    
End Function
'Implement IBatchCommand's Execute Method
Private Function IBatchMiddleCommand_Execute(strLogFiles As String, strErrorLogFiles As String) As Boolean
    
    IBatchMiddleCommand_Execute = False
    
    m_oBatchProcessLogHelper.Log "Recomputing Objects Started......."
    
    'Recompute object under filter
    Dim i As Integer
    For i = 1 To m_oObjectsToRecompute.Count
        
        Dim oPOM As IJDPOM
        Set oPOM = m_oMiddleCtx.GetResourceManager("model")
        
        Dim oObject As IJDObject
        Set oObject = oPOM.GetObject(m_oObjectsToRecompute.Item(i))
        
        On Error Resume Next
       
        Dim oObjGeometry As Object
        Set oObjGeometry = GetRootGeometry(oObject)
        
        ' For entities like Seams, Profiles etc., get the underlying landing curve
        Dim objLandingCurveGeometry As Object
        Set objLandingCurveGeometry = GetRootLandingCurve(oObject)
        
        ' Mark the selected object dirty on all of its interfaces
        UpdateInterfaces m_oObjectsToRecompute.Item(i), oPOM
        
        ' If we could get a geometry, mark it dirty too
        If Not oObjGeometry Is Nothing Then
            UpdateInterfaces m_oObjectsToRecompute.Item(i), oPOM
        End If
        
        ' Mark the interfaces of Landing Curve geometry as dirty so as to trigger landing curve semantics
        If Not objLandingCurveGeometry Is Nothing Then
            UpdateInterfaces m_oObjectsToRecompute.Item(i), oPOM
        End If
        
        Dim namedItem As IJNamedItem
        Set namedItem = oObject
        m_oBatchProcessLogHelper.Log "Recomputed : " & namedItem.Name
    Next
    
    m_oBatchProcessLogHelper.Log "Recomputing Objects Ended......."
   
    Dim oContext      As IJContext
    Set oContext = GetJContext()
    
    Dim oAccessMiddle As IJDAccessMiddle
    Set oAccessMiddle = oContext.GetService("ConnectMiddle")
    oAccessMiddle.CommitTransaction "Commit"
    
    strLogFiles = m_oBatchProcessLogHelper.LogFilePath
    
    m_oBatchProcessLogHelper.CloseLog
    
    IBatchMiddleCommand_Execute = True
    
End Function

Public Sub UpdateInterfaces(varMoniker As Variant, oPOM As IJDPOM)
    
    Dim oInfos As IJDInfosCol
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim oHelper As AssocHelper
    
    Const E_ACCESSDENIED As Long = -2147467259
    
    Set oHelper = New AssocHelper
    
    ' Get all the supported interfaces
    Set oInfos = GetSupportedInterfaces(varMoniker, oPOM)
    
    ' Disable error handling, since UpdateObject may fail if Access is denied
    ' (for example, due to permission group limitations)
    On Error Resume Next
    
    ' Mark each interface as "updated"
    If Not oInfos Is Nothing Then
        For Each oInterfaceInfo In oInfos
            ' Mfg process and Marking SOs are connected to detailed part using IJStructMfgProduction interface
            ' when we update detailed part using BO recompute we do not want to automatically update Mfg objects.
            ' So if we stop updating the detailed part on IJStructMfgProduction interface it would not trigger
            ' the SO semantics and hence it would not update the Mfg objects, they would become out of date and user
            ' has to use the MSM command to update them.
            If oInterfaceInfo.Name <> "IJStructMfgProduction" Then
                oHelper.UpdateObject oPOM.GetObject(varMoniker), oInterfaceInfo.Type
                
                ' If there's an access problem, then stop trying to update
                If Err.Number = E_ACCESSDENIED Then
                    Exit For
                End If
            End If
        Next
    
        ' Kludge -- make sure interface IJStructGeometry is marked as updated.
        ' Note: this will NOT fail if the object doesn't support the interface.
        Const IID_IJStructGeometry As String = "{6034AD40-FA0B-11D1-B2FD-080036024603}"
        oHelper.UpdateObject oPOM.GetObject(varMoniker), IID_IJStructGeometry
        
        Set oInterfaceInfo = Nothing
    End If
    
    Set oInfos = Nothing
    Set oHelper = Nothing

End Sub

Public Function GetSupportedInterfaces(varMoniker As Variant, oPOM As IJDPOM) As IJDInfosCol
    
    Dim oReposQuery As IJDRepositoryQuery
    Set oReposQuery = oPOM
    
    Dim strClassType As String
    strClassType = oPOM.ClsidOf(varMoniker)
    
    Dim oClassInfo As IJDClassInfo
    For Each oClassInfo In oReposQuery.ClassCollection
        If oClassInfo.ClassType = strClassType Then
            Set GetSupportedInterfaces = oClassInfo.ImplementedInterfaces()
            Exit For
        End If
    Next
   
    Set oReposQuery = Nothing
    Set oClassInfo = Nothing
End Function

Private Function GetRootGeometry(oStructEntity As Object) As Object
    On Error Resume Next
    Dim oIsStructEntity As IJStructGeometry
    Dim oAssocRelation  As IJDAssocRelation
    Dim oRelation As IJDRelationship
    Dim oGeom As Object
    'From the entity, get the geometry
    Set oAssocRelation = oStructEntity
    If Not oAssocRelation Is Nothing Then
        Dim oRelCol As IJDRelationshipCol
        Set oRelCol = oAssocRelation.CollectionRelations(IID_IJGeometry, "StructToGeometry_DEST")
        Set oRelation = oRelCol.Item(1)
        
        If Not oRelation Is Nothing Then
            Set oGeom = oRelation.Origin()
            Dim oOperation As IJStructOperation
            Dim cnt As Long
            Dim oPlane As IJPlane
            Dim thisType As StructOperation
            
            Do While (IsRootGeometry(oGeom) = False) And oIsStructEntity Is Nothing
                cnt = cnt + 1
                Set oAssocRelation = oGeom
                Set oRelCol = oAssocRelation.CollectionRelations(IID_IJStructGeometry, "StructOperation_RSLT1_DEST")
                Set oRelation = oRelCol.Item(1)
                Set oOperation = oRelation.Origin()
                'oOperation.GetOperationType thisType
                'MsgBox "Op Type: " & thisType
                Set oAssocRelation = oOperation
                Set oRelCol = oAssocRelation.CollectionRelations(IID_IJStructOperation, "StructOperation_OPRND_ORIG")
                Set oRelation = oRelCol.Item(1)
                Set oGeom = oRelation.Destination()
                
                Set oIsStructEntity = oGeom
                
                If cnt > 20 Then Exit Do
            Loop
        Else
            Set oRelCol = oAssocRelation.CollectionRelations(IID_IJCompartEntity, "Geometry")
            If oRelCol.Count > 0 Then
                Set oRelation = oRelCol.Item(1)
                Set oGeom = oRelation.Destination()
                
                'Check for Multiple volumes command
                Set oAssocRelation = oGeom
                If Not oAssocRelation Is Nothing Then
                    Dim oAE As Object
                    Set oRelCol = oAssocRelation.CollectionRelations(IID_IJCompartGeometry, "CompartAE")
                    Set oRelation = oRelCol.Item(1)
                    
                    Set oAE = oRelation.Origin()
                    Set oAssocRelation = oAE
                    If Not oAssocRelation Is Nothing Then
                        Set oRelCol = oAssocRelation.CollectionRelations(IID_IJCompartSplitAE, "VolumeInput")
                        Set oRelation = oRelCol.Item(1)
                        
                        If Not oRelation Is Nothing Then
                            Dim oUnBoundedGeom As Object
                            Set oUnBoundedGeom = oRelation.Destination()
                            
                            If Not oUnBoundedGeom Is Nothing Then
                                Set oGeom = oUnBoundedGeom
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Set GetRootGeometry = oGeom
    End If
    
    
End Function

' Get the underlying landing curve of entity that is being recomputed
Private Function GetRootLandingCurve(oStructEntity As Object) As Object
    
    On Error GoTo PROC_ERROR

    Dim objLandCurve As IJLandCurve

    Dim lngEntityType As StructEntityTypes
    lngEntityType = GetEntityType(oStructEntity)
    
    If lngEntityType = ENTITY_TYPE_DESIGN_SEAM Then
        ' Use Seam Utils to get the landing curve
        
        Dim objSeam As IJSeam
        Set objSeam = oStructEntity
        
        Dim objSeamUtils As IJSeamAttributes
        Set objSeamUtils = New GSCADCreateModifyUtilities.SeamUtils
        
        objSeamUtils.GetLandingCurveFromSeam objSeam, objLandCurve
        
    ElseIf lngEntityType = ENTITY_TYPE_PROFILE_SYSTEM Then
        ' Use Profile Utils to get the landing curve
        
        Dim objStiffener As IJStiffener
        Set objStiffener = oStructEntity
        
        Dim objProfileUtils As IJProfileAttributes
        Set objProfileUtils = New GSCADCreateModifyUtilities.ProfileUtils
        
        objProfileUtils.GetRootLandingCurveFromProfile objStiffener, objLandCurve
        
    ElseIf lngEntityType = ENTITY_TYPE_REFCURVE Then
        ' Use Ref Curve Utils to get the landing curve
        
        Dim objRefCurve As IJRefCurveOnSurface
        Set objRefCurve = oStructEntity
        
        Dim objRefCurveUtils As IJRefCurveAttributes
        Set objRefCurveUtils = New GSCADCreateModifyUtilities.RefCurveUtils
        
        objRefCurveUtils.GetLandingCurveFromRefCurve objRefCurve, objLandCurve
        
    Else
        ' Add other entities as required.
    End If

PROC_EXIT:
    Set GetRootLandingCurve = objLandCurve
    
    Exit Function

PROC_ERROR:
    ' For time being, do not error out

    GoTo PROC_EXIT

End Function

Public Function IsRootGeometry(oGeomObj As Object) As Boolean
    On Error Resume Next
    IsRootGeometry = False
    Dim oAssocRelation As IJDAssocRelation
    Set oAssocRelation = oGeomObj
    Dim oRelCol As IJDRelationshipCol
    
    Set oRelCol = oAssocRelation.CollectionRelations(IID_IJStructGeometry, "GeometryGeneration_RSLT1_DEST")
   
    If Not oRelCol Is Nothing Then
        IsRootGeometry = oRelCol.Count() > 0
    End If

End Function

' Figures out the whether the given entity is a seam, profile etc.
Private Function GetEntityType( _
        oStructEntity As Object) _
        As StructEntityTypes
    
    Dim lngEntityType As StructEntityTypes
    lngEntityType = ENTITY_TYPE_UNKNOWN
    
    Dim objStructFilter As StructFilter
    Set objStructFilter = New StructFilter
    
    Dim objEntityList As IJElements
    Set objEntityList = New GSCADAppConnections.AppConnHlp
    objEntityList.Add oStructEntity
    
    Do
        objStructFilter.Clear
        
        ' Use a filter function to check whether the entity is a Design Seam
        objStructFilter.AddCriteria "[STFilterFunctions.StructFilterFunctions,IsADesignSeam]"
        If objStructFilter.PassCriteria(objEntityList) = objEntityList.Count Then
            ' Entity is a Design Seam
        
            lngEntityType = StructEntityTypes.ENTITY_TYPE_DESIGN_SEAM
        
            Exit Do
        End If
    
        If TypeOf oStructEntity Is IJStiffenerSystem Then
            ' Entity is a Profile System
            
            lngEntityType = StructEntityTypes.ENTITY_TYPE_PROFILE_SYSTEM
            
            Exit Do
        End If
        
        If TypeOf oStructEntity Is IJRefCurveOnSurface Then
            ' Entity is a Reference Curve (could be Plate Knuckle as well)
            
            lngEntityType = StructEntityTypes.ENTITY_TYPE_REFCURVE
            
            Exit Do
        End If
        
        ' Could handle other cases later
        
    Loop Until True
    
    Set objStructFilter = Nothing
    Set objEntityList = Nothing
    
    GetEntityType = lngEntityType
End Function


