VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CompPartClassRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: CompartRules
'
'  Abstract: The file contains an implementation of the naming rules
'
'  History:
'           Raghuveer Thandur           Creation        15-Aug-2006
''          Raghuveer                   30-Oct-2006              TR-CP·11655 - Compartments under the same node can have the same name.
'***************************************************************************


Option Explicit

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259
Private Const PARTCLASS = "PartClass"
  
Implements IJNameRule

Private Const MODULE = "CompPartClassRule: "
Private Const strCountFormat = "0000"

Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IJElements, ByVal pActiveEntity As Object)
Const METHOD = "IJNameRule_ComputeName"
On Error GoTo ErrorHandler

    Dim oPartClass                  As IJDPartClass
    Dim oPartRelationHelper         As IMSRelation.DRelationHelper
    Dim oPartCollection             As IMSRelation.DCollectionHelper
    Dim strClass                    As String
    Dim oSpacePart                  As IJDPart
    Dim oCompartEntity              As IJCompartEntity
    Dim strName                     As String
    Dim nCount                      As Long
    Dim oNamedItem                  As IJNamedItem
        
    Dim JContext As IJContext
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    Dim oModelResourceMgr As IUnknown
    Dim oNameCounter As IJNameCounter
    Dim strLocation As String
        
    Set oCompartEntity = pEntity
    Set oNamedItem = pEntity
    
    If oNamedItem.Name <> vbNullString Then GoTo cleanup
    
    'Get the connection to the model database
    Set JContext = GetJContext()
    
    Set oDBTypeConfig = JContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = JContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
    
    Set oSpacePart = oCompartEntity.CompartDefinition

    'Get the Relation Helper
    Set oPartRelationHelper = oSpacePart
        
    'get the partclass
    Set oPartCollection = oPartRelationHelper.CollectionRelations(IID_IJDPart, PARTCLASS)
    
    Set oPartClass = oPartCollection.Item(1)
    
    strName = oPartClass.PartClassType
                                
    strLocation = vbNullString
    
    'GetCountEx:Returns the number of occurrence of a string in addtion to the LocationID
    nCount = oNameCounter.GetCountEx(oModelResourceMgr, strName, strLocation)
    
    'Add LocationID, if available
    If strLocation <> vbNullString Then
        strName = strName & "-" & strLocation & "-" & Format(nCount, strCountFormat)
    Else
        strName = strName & "-" & Format(nCount, strCountFormat)
    End If
    
    oNamedItem.Name = strName
    
            
cleanup:
    Set oPartClass = Nothing
    Set oPartCollection = Nothing
    Set oPartRelationHelper = Nothing
        
Exit Sub
ErrorHandler:
    m_oErrors.Add Err.Number, "CompDescRule::IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub

'****************************************************************************************************
'Author - HRM
'Description
'   All the Naming Parents that need to participate in an objects naming are added here to the
'   IJElements collection. Dummy function which does nothing
'****************************************************************************************************
Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IJElements
Const METHOD = "IJNameRule_GetNamingParents"
On Error GoTo ErrorHandler

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

Exit Function
ErrorHandler:
    m_oErrors.Add Err.Number, "CompDescRule::IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function
Public Property Get IID_IJDPart() As Variant
    IID_IJDPart = InitGuid(&H7D, &H1B, &HB1, &H9A, &H98, &H92, &HD1, &H11, &HBD, &HDC, &H0, &H60, &H97, &H3D, &H48, &H5)
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
