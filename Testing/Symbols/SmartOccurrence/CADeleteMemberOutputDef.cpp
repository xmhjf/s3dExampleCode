// CADeleteMemberOutputDef.cpp : Implementation of CCADeleteMemberOutputDef

#include "stdafx.h"
#include "CADeleteMemberOutputDef.h"
#include "CustomAssembly.h"
#include "Attributes.h"
#include "IMSDObject.h"
#include "Geom3d.h"

// CCADeleteMemberOutputDef

STDMETHODIMP CCADeleteMemberOutputDef::InstanciateDefinition(BSTR CodeBase, VARIANT definitionParameters, LPDISPATCH pResourceMgr, LPDISPATCH * ppSymbolDefDisp)
{
	const CComBSTR cProgId(L"ATPSO.CADeleteMemberOutputDef");
	HRESULT hr = S_OK;
	
	try
	{
		if (ppSymbolDefDisp == NULL)
			hr_throw( E_INVALIDARG, IID_IJDUserSymbolServices, __FILE__, L"Invalid symbol definition argument passed to InstanciateDefinition", "", 0 , "");
		*ppSymbolDefDisp = NULL;

		// This method is in charge of the creation of the symbol definition object
		CComPtr<IJCAFactory>	pFact;

		hr = pFact.CoCreateInstance( CLSID_CAFactory, NULL, CLSCTX_ALL );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pFact.CoCreateInstance( IID_CAFactory failed", "", 0 , "");
	    
		CComPtr<IDispatch>	 pDefinitionDisp;
		hr = pFact->get_CreateCAD(pResourceMgr, &pDefinitionDisp);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pFact->CreateCAD failed", "", 0 , "");
		CComPtr<IJDSymbolDefinition>	pDefinition;
		hr = pDefinitionDisp->QueryInterface(IID_IJDSymbolDefinition, (void**)&pDefinition);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pDefinitionDisp->QueryInterface(IID_IJSymbolDefinition failed", "", 0 , "");
	    
		// Set definition progId and codebase
		hr = pDefinition->put_ProgID( cProgId );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pDefinition->put_ProgID failed", "", 0 , "");

		hr = pDefinition->put_CodeBase( CodeBase );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pDefinition->put_CodeBase failed", "", 0 , "");
	    
		// Initialize the definition
		hr = InitializeCustomAssemblyDefinition( pDefinition );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"InitializeSymbolDefinition failed", "", 0 , "");
	    
		CComBSTR	 sDefinitionName;
		hr = GetDefinitionName(definitionParameters, &sDefinitionName);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"GetDefinitionName failed", "", 0 , "");
	    
		hr = pDefinition->put_Name( sDefinitionName );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pDefinition->put_Name failed", "", 0 , "");
    
		// Return definition
		*ppSymbolDefDisp = pDefinitionDisp.Detach();
	}
	catch ( CCMNException& ce)
	{
		hr = ce.GetHR();
	}
	catch(...) 
	{	
		hr = E_FAIL;
	}

	return hr;
}
	
STDMETHODIMP CCADeleteMemberOutputDef::InitializeSymbolDefinition(IJDSymbolDefinition * * ppSymbolDefDisp)
{
	HRESULT hr = InitializeCustomAssemblyDefinition( *ppSymbolDefDisp );

	return hr;
}

HRESULT CCADeleteMemberOutputDef::InitializeCustomAssemblyDefinition(IJDSymbolDefinition* pSymbolDefDisp)
{
	const CComBSTR CLSID_AssemblyMembers1Relationship = "{45E4020F-F8D8-47A1-9B00-C9570C1E0B17}";

	const CComBSTR  INTERFACE_IJDAttributes = "{B25FD387-CFEB-11D1-850B-080036DE8E03}";
	const CComBSTR  INTERFACE_IJGeometry = "{96eb9676-6530-11d1-977f-080036754203}";
	const CComBSTR  INTERFACE_Sphere = "IJUATestDotNetSphere";
	const CComBSTR  INTERFACE_ModifyOutput = "IJUAModifyOutput";
	const CComBSTR  INTERFACE_IJSmartOccurrence = "IJSmartOccurrence";

	const CComBSTR  PROPERTY_SphereDiameter(L"Diameter");
	const CComBSTR  PROPERTY_SphereOriginX(L"OriginX");
	const CComBSTR  PROPERTY_SphereOriginY(L"OriginY");
	const CComBSTR  PROPERTY_SphereOriginZ(L"OriginZ");

	HRESULT hr = S_OK;
	try
	{
		if ( pSymbolDefDisp == NULL )
			hr_throw( E_INVALIDARG, IID_IJDUserSymbolServices, __FILE__, L"Invalid symbol definition argument passed to InitializeCustomAssemblyDefinition", "", 0 , "");

		// Persistence behavior
		hr = pSymbolDefDisp->put_SupportOnlyOption( igSYMBOL_NOT_SUPPORT_ONLY );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pDefinition->put_SupportOnlyOption failed", "", 0 , "");
		hr = pSymbolDefDisp->put_MetaDataOption( igSYMBOL_DYNAMIC_METADATA );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pDefinition->put_MetaDataOption failed", "", 0 , "");
    
		// Define the inputs -
		CComPtr<IJDInputs> oInputs;
		hr = pSymbolDefDisp->QueryInterface( IID_IJDInputs, (void**) &oInputs);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->QueryInterface( IID_IJDInputs failed", "", 0 , "");

		CComPtr<IJDInput> oInput;
		hr = oInput.CoCreateInstance(CLSID_DInput, NULL, CLSCTX_INPROC_SERVER);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput.CoCreateInstance(CLSID_DInput failed", "", 0 , "");
    
		hr = oInput->put_Name(CComBSTR(L"SupportingObject"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Name failed", "", 0 , "");

		hr = oInput->put_Index( 1 );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"put_Index( 1 ) failed", "", 0 , "");
    
		oInputs->Add( oInput, CComVariant(1) );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInputs->Add( oInput, CComVariant(1) ) failed", "", 0 , "");
    
		hr = oInput->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->Reset() failed", "", 0 , "");
    
		oInput.Release();
		oInputs.Release();
    
		// Aggregator Type
		CComPtr<IJDAggregatorDescription> pAD;
		hr = pSymbolDefDisp->QueryInterface( IID_IJDAggregatorDescription, (void**)&pAD);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->QueryInterface( IID_IJDAggregatorDescription failed", "", 0 , "");
		hr = pAD->put_AggregatorClsid( CComBSTR(L"{A2A655C1-E2F5-11D4-9825-00104BD1CC25}") ); // CSmartOccurrence
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pAD->put_AggregatorClsid failed", "", 0 , "");
		hr = pAD->SetCMFinalConstruct( CComVariant(imsCOOKIE_ID_USS_LIB), CComVariant(CComBSTR(L"CMFinalConstructAsm")) );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pAD->SetCMFinalConstruct failed", "", 0 , "");
		hr = pAD->SetCMConstruct( CComVariant(imsCOOKIE_ID_USS_LIB), CComVariant(CComBSTR(L"CMConstructAsm")) );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pAD->SetCMConstruct failed", "", 0 , "");
		hr = pAD->SetCMSetInputs( CComVariant(-1), CComVariant(-1) );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pAD->SetCMSetInputs failed", "", 0 , "");
		hr = pAD->SetCMRemoveInputs( CComVariant(-1), CComVariant(-1) );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pAD->SetCMRemoveInputs failed", "", 0 , "");
    
		pAD.Release();
    
		CComPtr<IJCADefinition> pCADefinition;
		hr = pSymbolDefDisp->QueryInterface( IID_IJCADefinition, (void**)&pCADefinition);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->QueryInterface( IID_IJCADefinition failed", "", 0 , "");
	    hr = pCADefinition->put_CopyBackwardFlag(igCOPY_BACKWARD_TRIM);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pCADefinition->put_CopyBackwardFlag failed", "", 0 , "");
		pCADefinition.Release();
    
		// Aggregator property
		CComPtr<IJDPropertyDescriptions> pAPDs;
		hr = pSymbolDefDisp->QueryInterface( IID_IJDPropertyDescriptions, (void**)&pAPDs);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->QueryInterface( IID_IJDPropertyDescriptions failed", "", 0 , "");
		hr = pAPDs->RemoveAll(); // Remove all the previous property descriptions
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pAPDs->RemoveAll failed", "", 0 , "");
    
		CComPtr<IJDPropertyDescription>	pD;
		hr = pAPDs->AddProperty(CComBSTR(L"AttributeMods"), 1, INTERFACE_IJDAttributes, CComVariant(CComBSTR(L"CMEvaluateCAO")), CComVariant(imsCOOKIE_ID_USS_LIB), igPROCESS_PD_AFTER_SYMBOL_UPDATE, &pD);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->QueryInterface( IID_IJDPropertyDescriptions failed", "", 0 , "");
        
		pD.Release();
		pAPDs.Release();
    
		CComPtr<IJDMemberDescriptions> pMemberDescriptions;
		hr = pSymbolDefDisp->QueryInterface( IID_IJDMemberDescriptions, (void**)&pMemberDescriptions);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->QueryInterface( IID_IJDMemberDescriptions failed", "", 0 , "");
		// Remove all the previous member descriptions
		hr = pMemberDescriptions->RemoveAll();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pMemberDescriptions->RemoveAll failed", "", 0 , "");
    
		CComPtr<IJDMemberDescription> pMemberDescription;
		hr = pMemberDescriptions->AddMember(CComBSTR(L"Sphere"), 1, CComVariant(CComBSTR(L"CMConstructSphere")), CComVariant(imsCOOKIE_ID_USS_LIB), &pMemberDescription);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pMemberDescriptions->AddMember failed", "", 0 , "");
		hr = pMemberDescription->put_RelationshipClsid( CLSID_AssemblyMembers1Relationship );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pMemberDescription->put_RelationshipClsid failed", "", 0 , "");
    
		pMemberDescriptions.Release();
		pMemberDescription.Release();
	}
	catch ( CCMNException& ce)
	{
		hr = ce.GetHR();
	}
	catch(...) 
	{	
		hr = E_FAIL;
	}

	return hr;
}
STDMETHODIMP CCADeleteMemberOutputDef::InvokeRepresentation(LPDISPATCH pSymbolOccurrence, BSTR pRepName, LPDISPATCH pOutputColl, SAFEARRAY * * arrayOfInputs)
{
	return S_OK;
}
	
STDMETHODIMP CCADeleteMemberOutputDef::EditOccurence(LPDISPATCH * pSymbolOccurrence, LPDISPATCH pTransactionMgr, VARIANT_BOOL * pbHasOwnEditForm)
{
	return S_OK;
}
	
STDMETHODIMP CCADeleteMemberOutputDef::GetDefinitionName(VARIANT definitionParameters, BSTR * pDefName)
{
	const CComBSTR cName(L"ATPSO.CADeleteMemberOutputDef");

	*pDefName = cName.Copy();
	return S_OK;
}

HRESULT CCADeleteMemberOutputDef::GetSymbolVersion( BSTR *pstrVersion )
{
    if (  !pstrVersion )              
        return E_INVALIDARG;

	*pstrVersion = SysAllocString(L"1.0.0.0");
	if ( ! *pstrVersion )
		return E_OUTOFMEMORY;

	return S_OK;
}

STDMETHODIMP CCADeleteMemberOutputDef::CMFinalConstructAsm( IJDAggregatorDescription** pAggregatorDescription)
{
	HRESULT hr = S_OK;

	return hr;
}

STDMETHODIMP CCADeleteMemberOutputDef::CMConstructAsm( IJDAggregatorDescription** pAggregatorDescription)
{
	HRESULT hr = S_OK;

	return hr;
}

STDMETHODIMP CCADeleteMemberOutputDef::CMEvaluateCAO( IJDPropertyDescription** pPropertyDescriptions, IDispatch** pObject)
{
	const CComBSTR INTERFACE_ModifyOutput(L"IJUAModifyOutput");

	HRESULT hr = S_OK;

	try
	{
		if (pPropertyDescriptions == NULL || *pPropertyDescriptions == NULL)
			hr_throw( E_INVALIDARG, IID_ICADeleteMemberOutputDef, __FILE__, L"Invalid property description argument passed to CMEvaluateCAO", "", 0 , "");
		if (pObject == NULL || *pObject == NULL)
			hr_throw( E_INVALIDARG, IID_ICADeleteMemberOutputDef, __FILE__, L"Invalid object argument passed to CMEvaluateCAO", "", 0 , "");

		CComPtr<IDispatch> oOccurrenceDisp;
		hr = (*pPropertyDescriptions)->get_CAO( &oOccurrenceDisp );
		hr_onFail_throw( hr, IID_ICADeleteMemberOutputDef, __FILE__, L"(*pPropertyDescriptions)->get_CAO failed", "", 0 , "");
		CComPtr<IJSmartOccurrence> oSmartOcc;
		hr = oOccurrenceDisp->QueryInterface( IID_IJSmartOccurrence, (void**) &oSmartOcc);
		hr_onFail_throw( hr, IID_ICADeleteMemberOutputDef, __FILE__, L"oOccurrenceDisp->QueryInterface( IID_IJSmartOccurrenc failed", "", 0 , "");
		CComPtr<IJDAttributes> oAttrs; 
		hr = oOccurrenceDisp->QueryInterface( IID_IJDAttributes, (void**) &oAttrs);
		hr_onFail_throw( hr, IID_ICADeleteMemberOutputDef, __FILE__, L"oOccurrenceDisp->QueryInterface( IID_IJDAttributes failed", "", 0 , "");
		CComPtr<IJDAttributesCol> oAttrCol;
		hr = oAttrs->get_CollectionOfAttributes(CComVariant(INTERFACE_ModifyOutput), &oAttrCol );
		hr_onFail_throw( hr, IID_ICADeleteMemberOutputDef, __FILE__, L"(*oAttrs->get_CollectionOfAttributes failed", "", 0 , "");
		hr = S_OK;
		if ( oAttrCol != NULL )
		{
			CComPtr<IJDAttribute> oAttr;
			hr = oAttrCol->get_Item(CComVariant(CComBSTR(L"DeleteOutput")), &oAttr );
			hr = S_OK;
			if ( oAttr != NULL )
			{
				CComVariant varValue;
				oAttr->get_Value( &varValue );
				long lValue = 0;
				if ( varValue.vt == VT_I4 ) lValue = (long)varValue.lVal;
				CComPtr<IJDMemberObjects> oMembers;
				if ( lValue == 2 ) // Delete the output (by first notifying the SO that we are doing this
				{
					CComPtr<IDispatch> oDefinitionDisp;
					hr = (*pPropertyDescriptions)->get_Definition( &oDefinitionDisp );
                    CComPtr<IJDMemberDescriptions> oMemberDescrs; 
					hr = oDefinitionDisp->QueryInterface( IID_IJDMemberDescriptions, (void**)&oMemberDescrs);
					hr_onFail_throw( hr, IID_ICADeleteMemberOutputDef, __FILE__, L"(*oAttrs->get_CollectionOfAttributes failed", "", 0 , "");
                    CComPtr<IJDMemberDescription> oMemberDescr; 
					hr = oMemberDescrs->get_ItemByDispid(1, &oMemberDescr);
					hr = S_OK;
					if ( oMemberDescr != NULL )
					{
						hr = oSmartOcc->QueryInterface( IID_IJDMemberObjects, (void**)&oMembers );
						hr_onFail_throw( hr, IID_ICADeleteMemberOutputDef, __FILE__, L"hr = oSmartOcc->QueryInterface( IID_IJDMemberObjects failed", "", 0 , "");
						hr = oMembers->RemoveDuringGame( oMemberDescr, igMEMBER_REMOVED_UNNEEDED );
						hr_onFail_throw( hr, IID_ICADeleteMemberOutputDef, __FILE__, L"oMembers->RemoveDuringGame failed", "", 0 , "");
					}
				}
				else if ( lValue == 3 ) // Delete the output so a ToDo record is generated
				{
					hr = oSmartOcc->QueryInterface( IID_IJDMemberObjects, (void**)&oMembers );
					hr_onFail_throw( hr, IID_ICADeleteMemberOutputDef, __FILE__, L"hr = oSmartOcc->QueryInterface( IID_IJDMemberObjects failed", "", 0 , "");
					CComPtr<IDispatch> oMemberDescrDisp;
					hr = oMembers->get_ItemByDispid(1, -1, &oMemberDescrDisp);
					if ( oMemberDescrDisp != NULL )
					{
						CComPtr<IJDObject> oAssemblyOutput;
						hr = oMemberDescrDisp->QueryInterface(IID_IJDObject, (void**)&oAssemblyOutput);
						if ( oAssemblyOutput != NULL)
						{
							hr = oAssemblyOutput->Remove();
						}
					}
					hr = S_OK;
				}
			}
		}

	}
	catch ( CCMNException& ce)
	{
		hr = ce.GetHR();
	}
	catch(...) 
	{	
		hr = E_FAIL;
	}

	return hr;
}

STDMETHODIMP CCADeleteMemberOutputDef::CMConstructSphere( IJDMemberDescription* pMemberDescription, IUnknown* pResourceManager, IDispatch** pObj)
{
	const CComBSTR INTERFACE_Sphere(L"IJUATestDotNetSphere");
	const CComBSTR  PROPERTY_SphereDiameter(L"Diameter");
	const CComBSTR  PROPERTY_SphereOriginX(L"OriginX");
	const CComBSTR  PROPERTY_SphereOriginY(L"OriginY");
	const CComBSTR  PROPERTY_SphereOriginZ(L"OriginZ");

	HRESULT hr = S_OK;
    double lDiameter = 0.0;
    lDiameter = GetCAOAttributeDouble(pMemberDescription, INTERFACE_Sphere, PROPERTY_SphereDiameter);
    double lOriginX = 0.0;
    lOriginX = GetCAOAttributeDouble(pMemberDescription, INTERFACE_Sphere, PROPERTY_SphereOriginX);
    double lOriginY = 0.0;
    lOriginY = GetCAOAttributeDouble(pMemberDescription, INTERFACE_Sphere, PROPERTY_SphereOriginY);
    double lOriginZ = 0.0;
    lOriginZ = GetCAOAttributeDouble(pMemberDescription, INTERFACE_Sphere, PROPERTY_SphereOriginZ);

    CComPtr<IJGeometryFactory>	oGeometryFactory;
	hr = oGeometryFactory.CoCreateInstance( CLSID_GeometryFactory, NULL, CLSCTX_INPROC_SERVER);
	if ( SUCCEEDED(hr) )
	{
		CComPtr<ISpheres3d>	oSpheres;
		hr = oGeometryFactory->get_Spheres3d( &oSpheres );
		if ( oSpheres )
		{
			CComPtr<IJSphere>	oSphere;
			hr = oSpheres->CreateByCenterRadius(pResourceManager, lOriginX, lOriginY, lOriginZ, lDiameter / 2.0, VARIANT_TRUE, &oSphere);
			if ( oSphere )
			{
				hr = oSphere->QueryInterface( IID_IDispatch, (void**) pObj );
			}
		}
	}

	return hr;
}

STDMETHODIMP CCADeleteMemberOutputDef::CMSetInputSphere( IJDMemberDescription** pMemberDesc)
{
	return S_OK;
}

STDMETHODIMP CCADeleteMemberOutputDef::CMFinalConstructSphere( IJDMemberDescription** pMemberDesc)
{
	return S_OK;
}

STDMETHODIMP CCADeleteMemberOutputDef::CMReleaseSphere( IJDMemberDescription** pMemberDesc)
{
	return S_OK;
}

STDMETHODIMP CCADeleteMemberOutputDef::CMEvaluateSphereProperties( IJDPropertyDescription** pPropertyDescriptions, IDispatch** pObject)
{
	return S_OK;
}

STDMETHODIMP CCADeleteMemberOutputDef::CMEvaluateSphereGeometry( IJDPropertyDescription** pPropertyDescriptions, IDispatch** pObject)
{
	return S_OK;
}

double CCADeleteMemberOutputDef::GetCAOAttributeDouble(IJDMemberDescription* pMemberDescription, BSTR InterfaceName, BSTR AttributeName)
{
	HRESULT hr = S_OK;
	double dValue = 1.0E-12;

	if ( pMemberDescription == NULL ) return dValue;

	CComPtr<IDispatch> oOccurrenceDisp;
	hr = pMemberDescription->get_CAO( &oOccurrenceDisp );
	if ( oOccurrenceDisp != NULL )
	{
		CComPtr<IJDAttributes> oAttrs;
		hr = oOccurrenceDisp->QueryInterface( IID_IJDAttributes, (void**)&oAttrs );
		if ( oAttrs != NULL )
		{
			CComPtr<IJDAttributesCol> oAttrsCol;
			hr = oAttrs->get_CollectionOfAttributes( CComVariant(InterfaceName), &oAttrsCol );
			if ( oAttrsCol )
			{
				CComPtr<IJDAttribute> oAttr;
				hr = oAttrsCol->get_Item(CComVariant(AttributeName), &oAttr );
				if ( oAttr != NULL )
				{
					CComVariant	varValue;
					hr = oAttr->get_Value( &varValue );
					if ( varValue.vt == VT_R8 ) dValue = varValue.dblVal;
				}
			}
		}
	}

	return dValue;
}
