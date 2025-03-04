VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMergeSplitRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************************
'Copyright (C) 2005-2006, Intergraph Corporation. All rights reserved.
'
'Abstract:
' Rule to Check the Overlapping of the Compartments
'Description: This class implements the IJDCompartMergeSplitRule
'PROGID:CmpartMergeSplitRule.CMergeSplitRule
'
'Notes:
'
'History
'       Raghuveer       29th Jun 06                 Creation.
'*********************************************************************

Option Explicit

Private Const MODULE = "CCompartMergeSplitRule"
Private Const PARTCLASS = "PartClass"
Private Const TKWorkingSet = "WorkingSet"

Implements IJCompartMergeSplitRule

Private Function GetPartClassType(oSpacePart As IJDPart) As String
Const METHOD = "GetPartClassType"
On Error GoTo ErrorHandler

    Dim oPartClass                  As IJDPartClass
    Dim oPartRelationHelper         As IMSRelation.DRelationHelper
    Dim oPartCollection             As IMSRelation.DCollectionHelper
    Dim strClass                    As String
    
    If oSpacePart Is Nothing Then
        GetPartClassType = vbNullString
    Else
        If oSpacePart.PartDescription <> vbNullString Then
        
            'Get the Relation Helper
            Set oPartRelationHelper = oSpacePart
                
            'get the partclass
            Set oPartCollection = oPartRelationHelper.CollectionRelations(IID_IJDPart, PARTCLASS)
                       
            Set oPartClass = oPartCollection.Item(1)
        End If
        
        
        If Not oPartClass Is Nothing Then
            'Compare the part class type and localize it
            Select Case oPartClass.PartClassType
                Case "ShipZoneClass"
                    strClass = "ShipZoneClass"
                Case "VoidSpaceClass"
                    strClass = "VoidSpaceClass"
                Case "CompartmentClass"
                    strClass = "CompartmentClass"
                Case Else
                    strClass = ""
            End Select
            
            GetPartClassType = strClass
            
        End If
        
    End If

    Set oPartClass = Nothing
    Set oPartCollection = Nothing
    Set oPartRelationHelper = Nothing

Exit Function
ErrorHandler:
     Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Public Property Get IID_IJDPart() As Variant
    IID_IJDPart = InitGuid(&H7D, &H1B, &HB1, &H9A, &H98, &H92, &HD1, &H11, &HBD, &HDC, &H0, &H60, &H97, &H3D, &H48, &H5)
End Property

Public Property Get IID_IJDPartClass() As Variant
    IID_IJDPartClass = InitGuid(&H7F, &H6F, &H3E, &H7B, &HFB, &H93, &HD1, &H11, &HBD, &HDD, &H0, &H60, &H97, &H3D, &H48, &H5)
End Property

'This InitGuid is different in ordering of bytes from the function InitGuid
'   of CommonApp Utilities.
Private Function InitGuid(a As Byte, b As Byte, c As Byte, d As Byte, _
                e As Byte, f As Byte, g As Byte, h As Byte, _
                i As Byte, j As Byte, k As Byte, l As Byte, _
                m As Byte, n As Byte, o As Byte, p As Byte) As Variant

    Dim Guid(0 To 15) As Byte
    
    Guid(0) = a
    Guid(1) = b
    Guid(2) = c
    Guid(3) = d
    Guid(4) = e
    Guid(5) = f
    Guid(6) = g
    Guid(7) = h
    Guid(8) = i
    Guid(9) = j
    Guid(10) = k
    Guid(11) = l
    Guid(12) = m
    Guid(13) = n
    Guid(14) = o
    Guid(15) = p
    
    InitGuid = Guid
End Function

'Iterates through the catalog and Get the available Compart part
Private Function GetAvailbleCompartPart(ByVal strCompartClass As String) As Object
Const METHOD = "GetAvailbleCompartPart"
On Error GoTo ErrorHandler

    Dim RelationHelper              As IMSRelation.DRelationHelper
    Dim CollectionHelper            As IMSRelation.DCollectionHelper
    Dim oPartClass                  As IJDPartClass
    Dim Index                       As Long
    Dim oPart                       As IJDPart
    Dim i                           As Long
    Dim PartClassColl               As GSCADRefDataGlobal.IJDCollection
    Dim oObj                        As Object
   
    Set GetAvailbleCompartPart = Nothing
   
    'Get the partclass collection based on the Compart class type
    Dim oCommand            As JMiddleCommand
    Dim oAllObjects         As IEnumMoniker
    Dim oPOM                As IJDPOM
    Dim strconn             As String
   
    Set oCommand = New JMiddleCommand

    Set oPOM = GetActiveConnection.GetResourceManager(GetActiveConnectionName("catalog"))
    strconn = oPOM.ConnectionType
    
    ' Get all the objects matches with input name from the model through query
    With oCommand
        .Prepared = False
        .QueryLanguage = LANGUAGE_SQL
        .ActiveConnection = oPOM.DatabaseID
        
        If StrComp(strconn, "ORACLE", vbTextCompare) = 0 Then
            .CommandText = "Select oid from REFDATPartClass where PartClassType = '" & strCompartClass & "'"
        Else
            .CommandText = "Select oid from dbo.REFDATPartClass where PartClassType = '" & strCompartClass & "'"
        End If
        
        Set oAllObjects = .SelectObjects()
    End With
    
    Set oCommand = Nothing
    
    Dim oObjVariant     As Variant
    
    Set PartClassColl = New GSCADRefDataGlobal.Collection
    
    On Error Resume Next
    ' Get the first object in the collection.
    If Not oAllObjects Is Nothing Then
        For Each oObjVariant In oAllObjects
            PartClassColl.Add oPOM.GetObject(oObjVariant)
        Next
    End If
    Err.Clear
    
    On Error GoTo ErrorHandler
    
    Set oAllObjects = Nothing
    Set oObjVariant = Nothing
        
    For i = 1 To PartClassColl.Size
        Set oPartClass = PartClassColl.Item(i)
        Set RelationHelper = oPartClass
        Set CollectionHelper = RelationHelper.CollectionRelations _
                                (IID_IJDPartClass, "Part")
        For Index = 1 To CollectionHelper.Count
            Set oPart = CollectionHelper.Item(Index)
            Set GetAvailbleCompartPart = oPart
            Exit For
        Next Index
    Next i
    
    Set PartClassColl = Nothing
    Set oPart = Nothing
    Set RelationHelper = Nothing
    Set CollectionHelper = Nothing
 
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_MERGESPLITRULE_VOIDSPACEPARTNOTFOUND)
End Function

Private Sub SetNewPart(oInputCompartEntity As IJCompartEntity, oVoidSpacePart As IJDPart)
Const METHOD = "SetNewPart"
On Error GoTo ErrorHandler

    Dim oCopyHelper     As IJCompartRelationHelper 'IJDRelationsCopyHelper
    Dim i               As Long
    Dim oSpaceEntity    As IJCompartEntity
    Dim oSpaceChild     As IJSpaceChild
    Dim oSpaceFolder    As IJSpaceParent
    Dim oParent         As Object
    Dim oIJDObject      As IJDObject
    Dim strRuleName     As String
    Dim strSpaceName    As String
    Dim oNamedItem      As IJNamedItem
    Dim oOldPart        As IJDPart
    Dim oCompartOld     As IJCompartEntity
    Dim strOldPartNumber    As String

    Set oCopyHelper = New CompartRelationHelper
        
        Set oCompartOld = oInputCompartEntity
        Set oOldPart = oCompartOld.CompartDefinition
        
        If Not oOldPart Is Nothing Then
            strOldPartNumber = oOldPart.PartNumber
        End If
        
        If strOldPartNumber <> oVoidSpacePart.PartNumber Then
        
            strRuleName = GetName(oCompartOld)
            Set oNamedItem = oCompartOld
            strSpaceName = oNamedItem.Name
                    
            Set oSpaceChild = oCompartOld
            Set oParent = oSpaceChild.GetParent
            
            If oParent Is Nothing Then Exit Sub
        
            'Create a new object of new space type
            Set oSpaceEntity = CreateSpaceObject(oVoidSpacePart, oParent)
            
            'Copy the relation (and SpaceGeometry)
            oCopyHelper.CopyRelationsAndProperties oCompartOld, oSpaceEntity
                    
            SetName strSpaceName, oSpaceEntity
          
            'Delete the old object
            Set oIJDObject = oCompartOld
            'Remove the Space/SpaceChild Relation before deleting
            Set oSpaceChild = oCompartOld
            Set oSpaceFolder = oSpaceChild.GetParent
            oSpaceFolder.RemoveChild oCompartOld
            oIJDObject.Remove
         
        End If

wrapup:
    Set oCopyHelper = Nothing
    Set oParent = Nothing
    Set oSpaceEntity = Nothing
    Set oSpaceChild = Nothing
    Set oSpaceFolder = Nothing
    Set oIJDObject = Nothing
    Set oNamedItem = Nothing

Exit Sub
ErrorHandler:
     Err.Raise CompartLogError(Err, MODULE, METHOD, , , CMPART_CUSTOMERRORS_MERGESPLITRULE_FAILED_UPDATEDEFINITION)
End Sub

Private Sub SetName(strName As String, oSpace As Object)
On Error GoTo ErrorHandler
Const METHOD = "SetName"
    
    Dim oNamedItem              As IJNamedItem
    Set oNamedItem = oSpace
    oNamedItem.Name = strName
    Set oNamedItem = Nothing
    
Exit Sub
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Sub

Private Function GetName(oSpace As Object) As String
Const METHOD = "GetName"
On Error GoTo ErrorHandler

    Dim oNamedItem As IJNamedItem
    Set oNamedItem = oSpace
    GetName = oNamedItem.Name
    Set oNamedItem = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

'Create Space Entity Object
Private Function CreateSpaceObject(oPartNew As IJDPart, oParentSpace As Object) As Object
Const METHOD = "CreateSpaceObject"
On Error GoTo ErrorHandler
    
    Dim oProxy              As Object
    Dim oCompartFactory     As IJCompartFactoryEntity
    Dim oModelPOM           As IJDPOM

    Set oModelPOM = GetActiveConnection.GetResourceManager(GetActiveConnectionName("Model"))
            
    'Get the Proxy of the part
    Set oProxy = oModelPOM.GetProxy(oPartNew, True)
    
    Set oCompartFactory = New CompartFactoryEntity
    
    If oParentSpace Is Nothing Then
        Exit Function
    End If
            
    Set CreateSpaceObject = oCompartFactory.CreateEntity("VoidSpaceEntity.VoidSpace.1", oParentSpace, oProxy)
     
    Set oProxy = Nothing
    Set oCompartFactory = Nothing
   
    Exit Function
ErrorHandler:
   Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Private Function IJCompartMergeSplitRule_HandleType(ByVal oInputCompartmentObj As Object, ByVal pTargetType As Object) As Boolean
Const METHOD = "IJMergeSplitRule_HandleType"
On Error GoTo ErrorHandler
    
    Dim oObj                         As Object
    Dim lCount                      As Long
    Dim lIndex                      As Long
    Dim oInputCompartEntity         As IJCompartEntity
    Dim oInputCompartPart           As IJDPart
    Dim strInputPartClass           As String
    Dim strTargetPartClass          As String
    Dim oPart                       As IJDPart
    Dim oVoidSpacePart              As IJDPart
    Dim oNamedItem                  As IJNamedItem
            
    Set oInputCompartEntity = oInputCompartmentObj
    Set oNamedItem = oInputCompartEntity
    Set oPart = oInputCompartEntity.CompartDefinition
    
    strInputPartClass = GetPartClassType(oPart)
    strTargetPartClass = GetPartClassType(pTargetType)

    If (strInputPartClass = strTargetPartClass) And (strInputPartClass = "CompartmentClass") Then
        
        IJCompartMergeSplitRule_HandleType = True
        'Change the Input Compartment's type To Void Space
        
        On Error Resume Next
        Set oVoidSpacePart = GetAvailbleCompartPart("VoidSpaceClass")
        On Error GoTo ErrorHandler
                                                                                                                        
        If Not oVoidSpacePart Is Nothing Then
            On Error Resume Next
            SetNewPart oInputCompartEntity, oVoidSpacePart
            On Error GoTo ErrorHandler
        End If
                        
    End If
                
    Set oPart = Nothing
    Set oInputCompartEntity = Nothing
    Set oInputCompartPart = Nothing
    
Exit Function
ErrorHandler:
    Err.Raise CompartLogError(Err, MODULE, METHOD)
End Function

Private Function GetActiveConnection() As IJDAccessMiddle
    Const METHOD = "GetActiveConnection"
    On Error GoTo ErrorHandler
    
    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    oCmnAppGenericUtil.GetActiveConnection GetActiveConnection

    Set oCmnAppGenericUtil = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function


Private Function GetActiveConnectionName(strDB As String) As String

    Dim jContext As IJContext
    Dim oDBTypeConfiguration As IJDBTypeConfiguration
    
    'Get the middle context
    Set jContext = GetJContext()
    
    'Get IJDBTypeConfiguration from the Context.
    Set oDBTypeConfiguration = jContext.GetService("DBTypeConfiguration")
    
    'Get the Model DataBase ID given the database type
    GetActiveConnectionName = oDBTypeConfiguration.get_DataBaseFromDBType(strDB)
    
    Set jContext = Nothing
    Set oDBTypeConfiguration = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function
