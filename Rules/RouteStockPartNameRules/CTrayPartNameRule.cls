VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrayPartNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'********************************************************************
' Copyright (C) 2008 - 2009, Intergraph Corporation.  All Rights Reserved.
'
' File:  CTrayPartNameRule.cls
'
' Author: Bharath
'
' Abstract:
'    The File contains an Implementation of a NamingRule for CableTrayStockParts.
'    The Name of a CableTray stock part is set as follows: RunName + "-" + "TRAYStock" + "-" + SequenceNumber
' History:
'      18-Nov-2004   Bharath                 Created
'*************************************************************************************************************


Option Explicit

Implements IJNameRule
Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259
Private Const MODULE = "CPComponentNameRule: "
Private Const MODELDATABASE = "Model"
Private Const strCountFormat = "0000"       ' define fixed-width number field for name
Private Sub Class_Initialize()
    Set m_oErrors = New IMSErrorLog.JServerErrors
    If m_oErrors Is Nothing Then
               Set m_oErrors = CreateObject("IMSErrorLog.Errors")
               'Or use the following line instead for the Singleton version
               'Set m_oErrors = CreateObject("IMSErrorLog.ErrorsSingleton")
    End If
End Sub
Private Sub Class_Terminate()
    Set m_oErrors = Nothing
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''IJNameRule''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''/////////////////////////////////////////////////////////////////////////////
'' METHOD:         IJNameRule_ComputeName
' Description:
'  Creates a name for the object passed in. The Naming Parents are added in AddNamingParents()
'  of the same interface. Both these methods are called from NamingRule Semantic.
'
' Arguments:
'  oObject       - Input.  Child object that needs to have the NamingRule naming.
'  oParents      - Input.  Naming parents collection.
'  oActiveEntity - Input.  Naming rules active entity on which the NamingParentsString is stored.
'/////////////////////////////////////////////////////////////////////////////


Private Sub IJNameRule_ComputeName(ByVal oEntity As Object, ByVal oParents As IJElements, ByVal oActiveEntity As Object)
    
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrHndler

    Dim oModelResourceMgr As IUnknown
    Dim oNameCounter As IJNameCounter
    Dim oNamedItem As IJNamedItem
    Dim oParent As IJNamedItem
    
    Dim strName As String
    Dim strRunName As String
    Dim strSequenceNo As String
    
    Dim lcount As Long
    Dim strLocation As String
    Dim strTray As String

    If oEntity Is Nothing Then Exit Sub
    
    strTray = "TRAY"
    
   'To get the SystemParentName
    Set oParent = oParents.Item(1)
    strRunName = oParent.Name
    Set oParent = Nothing
    'To get the SequenceNumber
   
    If strRunName <> oActiveEntity.NamingParentsString Then
        'Get the middle context
        GetMiddleContext oModelResourceMgr
        'Adding the RunName as the Naming ParentString
        oActiveEntity.NamingParentsString = strRunName
        'Increment and append a counter.
        Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
        strLocation = vbNullString
        lcount = oNameCounter.GetCountEx(oModelResourceMgr, strRunName, strLocation)
        
        strSequenceNo = Format(lcount, strCountFormat)
        'Name is Formed in the Required Sequence
        strName = strRunName + "-" + strTray + "-" + strSequenceNo
         
         Set oNamedItem = oEntity
        'Setting Name on the Object
        oNamedItem.Name = strName
    End If
    
    If m_oErrors Is Nothing Then
               Set m_oErrors = CreateObject("IMSErrorLog.Errors")
               'Or use the following line instead for the Singleton version
               'Set m_oErrors = CreateObject("IMSErrorLog.ErrorsSingleton")
    End If
     
    'Clean up
    Set oModelResourceMgr = Nothing
    Set oNameCounter = Nothing
    Set oNamedItem = Nothing
Exit Sub
   
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Sub



'///////////////////////////////////////////////////////////////////////////////////////
' METHOD:         IJNameRule_GetNamingParents
' Description:
'  All the Naming Parents that need to Participate in an Objects Naming are added here to the
' IJElements collection. The Parents added here are used in Computing the Name of the Object in
' ComputeName() of the same Interface. Both these methods are called from Naming Rule Semantic.
'
' Arguments:
'  oEntity        - Input.  Child object that needs to have the NamingRule naming.
'
'/////////////////////////////////////////////////////////////////////////////


Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements

    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrHndler
    Dim oParent As IJDesignParent
    Dim oChild As IJDesignChild
    
   'Getting the Parent
    Set oChild = oEntity
    Set oParent = oChild.GetParent
    
    On Error GoTo ErrHndler
    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection
    
    If Not (oParent Is Nothing) Then
    Call IJNameRule_GetNamingParents.Add(oParent)
    End If
    
    If m_oErrors Is Nothing Then
               Set m_oErrors = CreateObject("IMSErrorLog.Errors")
               'Or use the following line instead for the Singleton version
               'Set m_oErrors = CreateObject("IMSErrorLog.ErrorsSingleton")
    End If
    
    'Clean Up
    Set oParent = Nothing
    Set oChild = Nothing
    Exit Function
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Private Sub GetMiddleContext(oModelResourceMgr As IUnknown)
    Dim jContext              As IJContext
    Dim oDBTypeConfig         As IJDBTypeConfiguration
    Dim oConnectMiddle        As IJDAccessMiddle
    Dim strModelDBID          As String
    
    Const METHOD = "GetMiddleContext"
    On Error GoTo ErrHndler
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
    
    'Clean up
    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
Exit Sub
ErrHndler:
    ' log the error in middle tier and propagate the error code to the caller
    m_oErrors.Add Err.Number, MODULE & METHOD, Err.Description
    Err.Raise E_FAIL
End Sub


