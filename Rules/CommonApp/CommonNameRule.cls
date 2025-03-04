VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommonNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 1999,2003 Intergraph Corporation.  All rights reserved.
'
'  Project: NameRules
'
'  Abstract: The file contains a sample implementation of a naming rule
'       If the object is a part occurrence, then then the name rule is
'           "part name" + "Unique counter"
'       else if the object is not a part occurrence
'           "Class User Name" + "Unique counter"

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
' 1 June 2001   JY                     Modified to make use of IJNamedItem instead of
'                                           IJProdModelItem
'       SS Oct/07/2003
'           TR CP49909: fixed the usage of LocationID
' Oct/18/2003   JPF         Removed using the system parent name as prefix for named object
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Implements IJNameRule
                                                        
Const vbInvalidArg = &H80070057

Private Const Module = "CommonNameRule: "
Private Const DEFAULTBASE = "Defaultname"
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name

Dim m_oErrors As IJEditErrors

Private Sub Class_Initialize()

    'DI-CP-164764    Prevent DLL unloading for high usage DLLs
    If m_ForceDllInProcess = 0 Then
        m_ForceDllInProcess = ForceDllInProcess(AddressOf ForceDllInProcess)
    End If

    Set m_oErrors = New IMSErrorLog.JServerErrors

End Sub

Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub

'********************************************************************
' Description:
' Creates a name for the object passed in.  If the object is a part occ, then
' then name is the part name + a unique counter
' if the object is not a part occ, then the name is the user name for the class
' plus a unique counter
' Both these methods are called from naming rule semantic.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)

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
    Dim strPartname As String
    Dim strLocation As String

    If oObject Is Nothing Then
        Err.Raise vbInvalidArg, Module, METHOD
    End If

    Set oNameCounter = New GSCADNameGenerator.NameGeneratorService

    Set oNamedItem = oObject
    
    ' Get name of part if this object is a part occurrence object
    strPartname = GetPartName(oObject)
    If strPartname = "" Then
        ' if we cannot get the part name, then use the type string
        strPartname = oNamedItem.TypeString
    End If

    ' Use part Name for basename, remove blanks:
    strPartname = Join(Split(strPartname), "")

    ' no parents - use just the strPartName as the base
    strNameBasis = oActiveEntity.NamingParentsString

    'Check if New parent name string constructed and old parent name string existing are same
    'if they are the same, we do not need to generate a new name
    If strPartname <> strNameBasis Then
        oActiveEntity.NamingParentsString = strPartname
            
        strLocation = vbNullString
        nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocation)
        
        If Not (strLocation = vbNullString) Then
            strName = strPartname + "-" + strLocation + "-" + Format(nCount, strCountFormat)
        Else
            strName = strPartname + "-" + Format(nCount, strCountFormat)
        End If
        
        oNamedItem.Name = strName
    End If
   
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
'Description
' For the Generic rule, the naming parent would be the part object if the named object
' is a part occurrence object.  However, the naming parent relationship is not
' defined to cross database boundaries.  For now there will not be any naming parent
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

