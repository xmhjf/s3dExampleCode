VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Importedhandrail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*******************************************************************
'  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
'
'  Project: \Symbols\Handrails\HandrailMacros.vbp
'  File:  \Symbols\Handrail\ImportedHandrail.cls
'
'  Description:  Imported Handrail with no graphics
'
'  Author: Structure Team

'  History:
'  05-Aug-08   WR      TR-CP-131048 - Added new method to ISPSHandrailSymbolServices to get the top rail radius
'******************************************************************

Private Const MODULE = "SPSHandrailMacros.ImportedHandrail"
Const defProgId As String = "SPSHandrailMacros.ImportedHandrail"

Private m_oErrors As IJEditErrors
Private Const E_FAIL = -2147467259

Implements IJDUserSymbolServices
Implements ISPSHandrailSymbolServices
Implements IJStructCustomFoulCheck

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
    
    Set pSymbolDefinition = pSymbolFactory.CreateEntity(Definition, pResourceMgr)
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
    Err.Raise E_FAIL
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

Public Sub Physical(pIRepSCM As IJDRepresentationStdCustomMethod)

End Sub


Private Sub ISPSHandrailSymbolServices_GetHandrailLength(ByVal HandrailBO As SPSHandrails.ISPSHandrail, length As Double)

End Sub

Private Sub ISPSHandrailSymbolServices_GetTopRailRadius(ByVal pIJDPart As SPSHandrails.IJDPart, radius As Double)

End Sub
'*****************************************************************************
'
'Function
'
'   <IJStructCustomFoulCheck_GetConnectedParts>
'
'Abstract
'   This method GetConnectedParts is on IJStructCustomFoulCheck interface.
'   The user can write custom code in here that tells the handrail business object
'   implemented in the middle tier about legal clashes for interference checking.
'   The objects populating the IJElements list will be ignored as the candidates for
'   interference with the handrail being computed
'
'Arguments
'
'   pPartPobject    -   the handrail object being computed (input)
'   pIJMonUnks      -   list of output collection of valid interfering objects (input/output)
'
'Return
'   None
'Exceptions
'   None
'******************************************************************************

Private Sub IJStructCustomFoulCheck_GetConnectedParts(ByVal pPartObject As Object, ByVal pIJMonUnks As SP3DStructGeneric.IJElements)
    'do not remove this method
End Sub
'*****************************************************************************
'Function
'
'   <IJStructCustomFoulCheck_GetFoulInterfaceType>
'
'Abstract
'   This method GetFoulInterfaceType is on IJStructCustomFoulCheck interface.
'   This tells the handrail business object implemented in the middle tier
'   if the handrail macro is going to participate in interference checking or not
'
'   pFoulInterfaceType = NonParticipant indicates that the handrail is not going
'                                       to participate in interference checking
'
'   pFoulInterfaceType = StandardGraphicEntity indicates otherwise
'
'Arguments
'   pFoulInterfaceType  -   foul interface type indicate the participation of the
'                           object in interference process (input/output)
'
'Return
'   None
'Exceptions
'   None
'******************************************************************************
Private Sub IJStructCustomFoulCheck_GetFoulInterfaceType(pFoulInterfaceType As SP3DStructGeneric.FoulInterfaceType)
    pFoulInterfaceType = NonParticipant ' as the imported object doesn't have a geometry, no need to participate in foul check
End Sub

