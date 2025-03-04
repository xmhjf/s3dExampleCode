VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BlockNameRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
' Copyright (C) 2002, Intergraph Corporation.  All rights reserved.
'
' Project:
'   NameRules
'
' Abstract:
'   The file contains the VB-namerule for Blocks
'
' Author:
'   Ole Mortensen oss/olmo
'
' History:
'   jun 27 2002     oss/olmo    Created
'   sep 24 2002     oss/olmo    Modified. If IJNamedItem.Name="" then calculate name,
'                               else do not recalculate name because the name is either the
'                               calculated or a userdefined.
'   Apr 21 2009     Kishore     DM-CP-163399  Performance problem when creating a new assembly
'******************************************************************

Option Explicit

Implements IJNameRule
                                                        
Private Const Module = "CPlnNameRules.BlockNameRule:: "
Private Const MODELDATABASE = "Model"
'Private Const strCountFormat = "000000"       ' define fixed-width number field for name
Private Const ERRORPROGID As String = "IMSErrorLog.ServerErrors"
Private m_oErrors As IJEditErrors       ' To collect and propagate the errors.
Private m_oError As IJEditError         ' The error to raise.
Private Sub Class_Initialize()
   Set m_oErrors = CreateObject(ERRORPROGID)
End Sub

Private Sub Class_Terminate()
   Set m_oErrors = Nothing
   Set m_oError = Nothing
End Sub

'********************************************************************
' Overall description:
' Creates a name for the block object passed in. Creation only takes place if the
' currnet name is empty (""). BlockName is based on parents name + index.
' E.g. Block with parentname B0.2.1 gets name B0.2.1.X where X is unique.
' No checks are performes as to wheather the blockname allready exists.
'
' The ComputeName method is called by semantic.
'
'
' It is assumed that all Naming Parents and the Object implement IJNamedItem.
' The Naming Parents are added in AddNamingParents() of the same interface.
' Both these methods are called from naming rule semantic.
'********************************************************************
Private Sub IJNameRule_ComputeName(ByVal oObject As Object, ByVal elements As IJElements, ByVal oActiveEntity As Object)
    Const METHOD = "IJNameRule_ComputeName"
    On Error GoTo ErrorHandler
    
    Dim jContext As IJContext
    Dim oModelResourceMgr As IUnknown
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
    Dim oNameCounter As IJNameCounter
    Dim oChildName As IJNamedItem
    Dim nCount As Long
    Dim strPartname As String
    Dim oParentName As IJNamedItem
    Dim oBlockParent As GSCADBlockItem.IJBlock
    Dim oAssyChild As IJAssemblyChild
    Dim strLocationID As String
    Dim strSeparator As String
    
    strSeparator = "-"

    'Get the middle context
    Set jContext = GetJContext()

    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")

    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType(MODELDATABASE)
    Set oModelResourceMgr = oConnectMiddle.GetResourceManager(strModelDBID)
       
    If oObject Is Nothing Then
        Err.Raise E_ACCESSDENIED, Module, METHOD
    End If
    
    'check if block allready has a name
    On Error Resume Next
    Set oChildName = oObject
    Err.Clear
    On Error GoTo ErrorHandler
    
    
    If oChildName.Name = "" Then
        ' get parent of this block
        Set oAssyChild = oObject
        If Not oAssyChild Is Nothing Then
        
            On Error Resume Next
            Set oBlockParent = oAssyChild.Parent
            Err.Clear
            
            On Error GoTo ErrorHandler
                    
            ' get max sub block index, N, which exist below the parent e.g. Bx.y.z.N
            Dim subblockindex As Long
            
            If oBlockParent Is Nothing Then
                 'get name from parent
                 
                subblockindex = GetBlockIndex(oAssyChild.Parent)
               
                Dim oAsmChild      As GSCADAsmHlpers.IJAssemblyChild
                Dim oAssy           As IJAssembly
                
                Set oAssy = oAssyChild.Parent
                
                Do While Not TypeOf oAssy Is IJBlock
                    'passed was not a block, get its parent
                    Set oAsmChild = oAssy
                    Set oAssy = Nothing
                    Set oAssy = oAsmChild.Parent
                Loop
                
                 'get name from parent
                Set oParentName = oAssy
                                        
            Else
                subblockindex = oBlockParent.MaxSubBlockIndex(strSeparator)
                 'get name from parent
                Set oParentName = oBlockParent
            End If
        End If
        
       
        'get IJNamedItem from this and set new name
        Dim sParentName As String
        
        If Not oParentName Is Nothing Then
            sParentName = oParentName.Name
        End If
        
        strPartname = oChildName.TypeString
        Set oNameCounter = New GSCADNameGenerator.NameGeneratorService
        nCount = oNameCounter.GetCountEx(oModelResourceMgr, strPartname, strLocationID)
            If Not sParentName = vbNullString Then
                If Not strLocationID = vbNullString Then
                    oChildName.Name = sParentName & "." & strLocationID & strSeparator & (subblockindex + 1)
                Else
                    'oChildName.Name = sParentName & "." & (subblockindex + 1)
                    oChildName.Name = sParentName & strSeparator & (subblockindex + 1)
                End If
            Else
                oChildName.Name = "Unspecified"
            End If
    End If 'if name=""
    
CleanUp:
    Set oBlockParent = Nothing
    Set oAssyChild = Nothing

    Set jContext = Nothing
    Set oModelResourceMgr = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oNameCounter = Nothing
    Set oChildName = Nothing
    Set oParentName = Nothing
    Exit Sub
    
ErrorHandler:
   m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
    GoTo CleanUp
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Description
' All the Naming Parents that need to participate in an objects naming are added here
' to the IJElements collection. The parent added here is root of assembly hierarchy
' and is used in computing the name of the object in ComputeName() of the same interface. Both these methods are called from
' naming rule semantic.
'
' For Blocks do NOT add a naming parent as no semantic should be involved.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJNameRule_GetNamingParents(ByVal oEntity As Object) As IJElements
    Const METHOD = "IJNameRule_GetNamingParents"
    On Error GoTo ErrorHandler

    Set IJNameRule_GetNamingParents = New IMSCoreCollections.JObjectCollection

Exit Function
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
End Function

Private Function GetBlockIndex(oAssembly As IJAssembly) As Long
Const METHOD = "GetBlockIndex"
On Error GoTo ErrorHandler

    Dim oChildren       As IJDTargetObjectCol
    Dim oName           As IJNamedItem
    Dim strName         As String
    
    Dim i               As Long
    Dim ltempIndex      As Long
    Dim lMaxIndex       As Long
    
    lMaxIndex = 0
    
    Set oChildren = oAssembly.GetChildren
    
    For i = 1 To oChildren.Count
        Set oName = oChildren.Item(i)
        ltempIndex = GetIndex(oName.Name)
        If ltempIndex > lMaxIndex Then
            lMaxIndex = ltempIndex
        End If
    Next
    
    GetBlockIndex = lMaxIndex
    
Exit Function
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
End Function

Public Function GetIndex(strName As String) As Long
Const METHOD = "GetIndex"
On Error GoTo ErrorHandler

    Dim strArray() As String
    Dim iUBound As Integer
    On Error Resume Next
    
    strArray = Split(strName, "-")
    iUBound = UBound(strArray)

    GetIndex = CLng((strArray(iUBound)))
    On Error GoTo ErrorHandler

    Exit Function
ErrorHandler:
    m_oError = m_oErrors.AddFromErr(Err, Module & " - " & METHOD)
    m_oError.Raise
End Function
