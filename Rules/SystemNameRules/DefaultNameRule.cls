VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 1999, Intergraph Corporation.  All rights reserved.
'
'  Project: SystemsAndSpecs SystemNameRules
'  Class:   DefaultNameRule
'
'  Abstract: The file contains a sample implementation of a naming rule for systems.
'
'  Author: Bathina Balakrishna Reddy
'
' 03/21/00      Rob Lemley                  TR CP10966: NameRules.CommonNameRule does not support
'                                           different GSCAD object types
' 01 May 2000   Bathina Balakrishna Reddy   Removed client tier references and moved to middle tier.
' 17 Oct 2000   jpf                         Changed to use part name instead of type string in name
'
'       Modified to use the IJProdModelItem.TypeString (with blanks
'       removed) as the object's "basename" for naming.
'
'       Fixed to remove extra blank in name introduced by the
'       VB Str() function.  Use Format() function and create
'       zero-padded fixed-width counter field in name.
'
' 1 June 2001   Jayanth                     Modified to make use of IJNamedItem instead of
'                                           IJProdModelItem
' 06/20/02      TM                          Removed trying to use part name because systems
'                                           don't have part names.
' 10/21/03      TM                          Modified ComputeName method to not include names
'                                           of parents.  Modified GetNamingParents to return
'                                           NULL.
'
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "DefaultNameRule: "
Private Const DEFAULTBASE = "Defaultname"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name
Private Const PARTROLE = "part"
Private IID_IJPartOcc As Variant

Dim m_oErrors As IJEditErrors

Private Const E_FAIL = -2147467259



Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
'  Creates a name for the object passed in. The name is based on the parents
'  name and object name.It is something like this: "Base Name" + "Object Name" + Index.
'  "Base Name" is derived from the Naming Parents and the Index is unique for the "Base Name"
'  It is assumed that all Naming Parents and the Object implement IJPRodModelItem.
'  The Naming Parents are added in AddNamingParents() of the same interface.
'  Both these methods are called from naming rule semantic.
'
' Arguments:
'  oObject - Input.  Child object that needs to have the naming rule naming.
'  oParents - Input.  Naming parents collection.
'  oActiveEntity - Input.  Naming rules active entity on which the NamingParentsString is stored.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal oParents As IJElements, ByVal oActiveEntity As Object)

    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo label
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    
    'Get the middle context
    Set jContext = GetJContext()
    
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)

    Dim oNameCounter As IJNameCounter
    Dim oNamedItem As IJNamedItem

    Dim strName As String
    Dim strNameBasis As String
    
    Dim nCount As Long
    Dim strSystemName As String
    Dim strLocation As String

    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService

    Set oNamedItem = oObject
    
    'Start the system name with its type string.
    strSystemName = oNamedItem.TypeString
    
    'Remove blanks:
    strSystemName = Join(Split(strSystemName), "")

    'No parents - use just the strSystemName as the base
    strNameBasis = oActiveEntity.NamingParentsString
    
    'Check if the new parent name string constructed and the old parent name string
    '  existing are the same.  If they are, we do not need to generate a new name
    If strSystemName <> strNameBasis Then
        oActiveEntity.NamingParentsString = strSystemName

        'Increment and append a counter.
        strLocation = vbNullString
        nCount = oNameCounter.GetCountEx(oModelResourceMgr, strSystemName, strLocation)
        
        If Not strLocation = vbNullString Then
            strName = strSystemName + "-" + strLocation + "-" + Format(nCount, strCountFormat)
        Else
            strName = strSystemName + "-" + Format(nCount, strCountFormat)
        End If
        
        oNamedItem.Name = strName
    End If
   
    'Clean up
    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    
    Set oNameCounter = Nothing
    Set oNamedItem = Nothing
    
Exit Sub

label:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Description:
' Initial design:
' All the Naming Parents that need to participate in an objects naming are added here to the
' IJElements collection. The parents added here are used in computing the name of the object in
' ComputeName() of the same interface. Both these methods are called from naming rule semantic.
'
' V4 update:
' The naming parent relationship is not defined to cross database boundries.  For now,
' there will not be any naming parent.
'
' Arguments:
'  oEntity - Input.  Child object that needs to have the naming rule naming.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo label

    Set IJNameRule_GetNamingParents = Nothing
    
    Exit Function

label:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, "IJNameRule_GetNamingParents", Err.Description
    Err.Raise E_FAIL
End Function

