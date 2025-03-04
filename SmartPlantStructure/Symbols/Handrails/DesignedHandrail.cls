VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesignedHandrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************************************
'  Copyright (C) 2010, Intergraph Corporation.  All rights reserved.
'
'  Project: \Symbols\Handrails\HandrailMacros.vbp
'  File:  \Symbols\Handrail\DesignedHandrail.cls
'
'  Description:  Handrail as system owner only.
'
'  Author: Structure Team

'  History:
'  01-29-10     MOH     CR 52186  add ability to "convert" handrail to member systems.
'  04-26-11     GG      TR 187117  Converted Handrails add invalid entry on TDL after copy or MDR
'******************************************************************

    Private Const MODULE = "SPSHandrailMacros.DesignedHandrail"
    Const defProgId As String = "SPSHandrailMacros.DesignedHandrail"

    Private m_oErrors As IJEditErrors
    Private Const E_FAIL = -2147467259

    Implements IJDUserSymbolServices
    Implements ISPSHandrailSymbolServices

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    IJDUserSymbolServices_GetDefinitionName = defProgId
End Function

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal definitionParameters As Variant, ByVal pResourceMgr As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler

    'Create a Symbol Definition Object.
    Dim pSymbolFactory As New DSymbolEntitiesFactory
    Dim pSymbolDefinition As IJDSymbolDefinition

    Set pSymbolDefinition = pSymbolFactory.CreateEntity(definition, pResourceMgr)
    IJDUserSymbolServices_InitializeSymbolDefinition pSymbolDefinition

    'Now Let's do some nomenclature stuff
    pSymbolDefinition.ProgId = defProgId
    pSymbolDefinition.CodeBase = CodeBase
    pSymbolDefinition.name = pSymbolDefinition.ProgId
    pSymbolDefinition.CacheOption = igSYMBOL_GEOM_FREE

    'returned symbol defintion
    Set IJDUserSymbolServices_InstanciateDefinition = pSymbolDefinition
    Set pSymbolFactory = Nothing
    Set pSymbolDefinition = Nothing

    Exit Function

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise (E_FAIL)
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pSymbolDef As IMSSymbolEntities.IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler

    ' Clean previous data
    pSymbolDef.IJDInputs.RemoveAllInput
    pSymbolDef.IJDRepresentations.RemoveAllRepresentation

    Dim libDesc As New DLibraryDescription
    Dim libCookie As Long

    libDesc.name = "mySelfAsLib"
    libDesc.Type = imsLIBRARY_IS_ACTIVEX
    libDesc.Properties = imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES
    libDesc.Source = defProgId
    pSymbolDef.IJDUserMethods.SetLibrary libDesc
    ' Get the lib/method cookie
    libCookie = libDesc.Cookie

    Dim oSymbolCache As New CustomCache
    oSymbolCache.SetupCustomCache pSymbolDef
    
    Dim rep As IJDRepresentation
    Set rep = New DRepresentation

    rep.RepresentationId = SimplePhysical 'define an aspect0 (Simple_physical)
    rep.name = "Physical"
    rep.Description = "Physical representation"

    Dim pUM As IMSSymbolEntities.IJDUserMethods
    Set pUM = pSymbolDef

    Dim mthCookie As Long
    mthCookie = pUM.GetMethodCookie("Physical", libCookie)
    rep.IJDRepresentationStdCustomMethod.SetCMEvaluate libCookie, mthCookie
    rep.IJDOutputs.Property = rep.IJDOutputs.Property + igCOLLECTION_VARIABLE

    pSymbolDef.IJDRepresentations.Add rep
    rep.Reset

    pSymbolDef.GeomOption = igSYMBOL_GEOM_FREE
    pSymbolDef.CacheOption = igSYMBOL_CACHE_OPTION_NOT_SHARED

    Exit Sub

ErrorHandler:
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Sub

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

Private Sub ISPSHandrailSymbolServices_EvaluateGeometry(ByVal HandrailBO As SPSHandrails.ISPSHandrail, ByVal PartOccChanged_IN As Boolean, ByVal Sketch3dCurve_IN As Object, ByVal AttributeCol_IN As SPSHandrails.IJDAttributesCol, ByVal AttributesChanged_IN As Boolean, ByVal OccAttributeCol_INOUT As SPSHandrails.IJDAttributesCol, ByVal OccAttributesChanged_IN As Boolean)

End Sub

Private Sub ISPSHandrailSymbolServices_EvaluateLocalWCG(ByVal HandrailBO As SPSHandrails.ISPSHandrail, ByVal OccAttributeCol As SPSHandrails.IJDInfosCol, weight As Double, CogX As Double, CogY As Double, CogZ As Double)

End Sub

Public Sub Physical(ByVal pIRepSCM As IJDRepresentationStdCustomMethod)

End Sub

Private Sub ISPSHandrailSymbolServices_GetHandrailLength(ByVal HandrailBO As SPSHandrails.ISPSHandrail, length As Double)

End Sub

Private Sub ISPSHandrailSymbolServices_GetTopRailRadius(ByVal pIJDPart As SPSHandrails.IJDPart, radius As Double)

End Sub

