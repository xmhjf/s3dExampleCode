// CAEvaluateAfterSymbolDef.cpp : Implementation of CCAEvaluateAfterSymbolDef

#include "stdafx.h"
#include "CAEvaluateAfterSymbolDef.h"
#include "CustomAssembly.h"
#include "Attributes.h"
#include "IMSDObject.h"
#include "Geom3d.h"

// CCAEvaluateAfterSymbolDef

STDMETHODIMP CCAEvaluateAfterSymbolDef::InstanciateDefinition(BSTR CodeBase, VARIANT definitionParameters, LPDISPATCH pResourceMgr, LPDISPATCH * ppSymbolDefDisp)
{
	const CComBSTR cProgId(L"ATPSO.CAEvaluateAfterSymbolDef");
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
	
STDMETHODIMP CCAEvaluateAfterSymbolDef::InitializeSymbolDefinition(IJDSymbolDefinition * * ppSymbolDefDisp)
{
	HRESULT hr = InitializeCustomAssemblyDefinition(*ppSymbolDefDisp);

	return hr;
}

HRESULT CCAEvaluateAfterSymbolDef::InitializeCustomAssemblyDefinition(IJDSymbolDefinition *pSymbolDefDisp)
{
	const CComBSTR CLSID_AssemblyMembers1Relationship = "{45E4020F-F8D8-47A1-9B00-C9570C1E0B17}";

	const CComBSTR  INTERFACE_IJDAttributes = "{B25FD387-CFEB-11D1-850B-080036DE8E03}";
	const CComBSTR  INTERFACE_IJGeometry = "{96eb9676-6530-11d1-977f-080036754203}";
	const CComBSTR  INTERFACE_Sphere = "IJUATestDotNetSphere";
	const CComBSTR  INTERFACE_Torus = "IJUATestDotNetTorus";
	const CComBSTR  INTERFACE_ModifyOutput = "IJUAModifyOutput";
	const CComBSTR  INTERFACE_IJSmartOccurrence = "IJSmartOccurrence";

	const CComBSTR  PROPERTY_SphereDiameter(L"Diameter");
	const CComBSTR  PROPERTY_SphereOriginX(L"OriginX");
	const CComBSTR  PROPERTY_SphereOriginY(L"OriginY");
	const CComBSTR  PROPERTY_SphereOriginZ(L"OriginZ");

	HRESULT hr = S_OK;
	try
	{
		if (pSymbolDefDisp == NULL)
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
    
		hr = oInputs->Add( oInput, CComVariant(1) );
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
    
		CComPtr<IJDPropertyDescription>	pPD;
		hr = pAPDs->AddProperty(INTERFACE_Sphere, 1, INTERFACE_Sphere, CComVariant(CComBSTR(L"CMEvaluateCAOBefore")), CComVariant(imsCOOKIE_ID_USS_LIB), igPROCESS_PD_BEFORE_SYMBOL_UPDATE, &pPD);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pAPDs->AddProperty(INTERFACE_Sphere failed", "", 0 , "");
		pPD.Release();
		hr = pAPDs->AddProperty(INTERFACE_Torus, 2, INTERFACE_Torus, CComVariant(CComBSTR(L"CMEvaluateCAOAfter")), CComVariant(imsCOOKIE_ID_USS_LIB), igPROCESS_PD_AFTER_SYMBOL_UPDATE, &pPD);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pAPDs->AddProperty(INTERFACE_Torus failed", "", 0 , "");
        
		pPD.Release();
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
		hr = pMemberDescription->SetCMSetInputs( CComVariant(imsCOOKIE_ID_USS_LIB), CComVariant(CComBSTR(L"CMSetInputSphere")) );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pMemberDescription->SetCMSetInputs failed", "", 0 , "");
		hr = pMemberDescription->SetCMFinalConstruct( CComVariant(imsCOOKIE_ID_USS_LIB), CComVariant(CComBSTR(L"CMFinalConstructSphere")) );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pMemberDescription->SetCMFinalConstruct failed", "", 0 , "");
		hr = pMemberDescription->SetCMRelease( CComVariant(imsCOOKIE_ID_USS_LIB), CComVariant(CComBSTR(L"CMReleaseSphere")) );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pMemberDescription->SetCMRelease failed", "", 0 , "");
		hr = pMemberDescription->put_RelationshipClsid( CLSID_AssemblyMembers1Relationship );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pMemberDescription->put_RelationshipClsid failed", "", 0 , "");
    
		hr = pMemberDescription->QueryInterface( IID_IJDPropertyDescriptions, (void**) &pAPDs );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pMemberDescription->QueryInterface( IID_IJDPropertyDescription failed", "", 0 , "");
		hr = pAPDs->AddProperty( CComBSTR(L"SphereProperties"), 1, INTERFACE_IJDAttributes, CComVariant(CComBSTR(L"CMEvaluateSphereProperties")), CComVariant(imsCOOKIE_ID_USS_LIB), igPROCESS_PD_AFTER_SYMBOL_UPDATE, &pPD );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"hr = pAPDs->AddProperty( CComBSTR(SphereProperties failed", "", 0 , "");
		pPD.Release();
		hr = pAPDs->AddProperty( CComBSTR(L"SphereGeometry"), 2, INTERFACE_IJGeometry, CComVariant(CComBSTR(L"CMEvaluateSphereGeometry")), CComVariant(imsCOOKIE_ID_USS_LIB), igPROCESS_PD_AFTER_SYMBOL_UPDATE, &pPD );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pAPDs->AddProperty( CComBSTR(SphereGeometry failed", "", 0 , "");
		
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

STDMETHODIMP CCAEvaluateAfterSymbolDef::InvokeRepresentation(LPDISPATCH pSymbolOccurrence, BSTR pRepName, LPDISPATCH pOutputColl, SAFEARRAY * * arrayOfInputs)
{
	return S_OK;
}
	
STDMETHODIMP CCAEvaluateAfterSymbolDef::EditOccurence(LPDISPATCH * pSymbolOccurrence, LPDISPATCH pTransactionMgr, VARIANT_BOOL * pbHasOwnEditForm)
{
	return S_OK;
}
	
STDMETHODIMP CCAEvaluateAfterSymbolDef::GetDefinitionName(VARIANT definitionParameters, BSTR * pDefName)
{
	const CComBSTR cName(L"ATPSO.CAEvaluateAfterSymbolDef");

	*pDefName = cName.Copy();
	return S_OK;
}

HRESULT CCAEvaluateAfterSymbolDef::GetSymbolVersion( BSTR *pstrVersion )
{
    if (  !pstrVersion )              
        return E_INVALIDARG;

	*pstrVersion = SysAllocString(L"1.0.0.0");
	if ( ! *pstrVersion )
		return E_OUTOFMEMORY;

	return S_OK;
}

STDMETHODIMP CCAEvaluateAfterSymbolDef::CMFinalConstructAsm( IJDAggregatorDescription** pAggregatorDescription)
{
	return S_OK;
}

STDMETHODIMP CCAEvaluateAfterSymbolDef::CMConstructAsm( IJDAggregatorDescription** pAggregatorDescription)
{
	HRESULT hr = S_OK;

	return hr;
}

STDMETHODIMP CCAEvaluateAfterSymbolDef::CMEvaluateCAOBefore( IJDPropertyDescription** pPropertyDescriptions, IDispatch** pObject)
{
	const CComBSTR INTERFACE_Sphere(L"IJUATestDotNetSphere");
	const CComBSTR  PROPERTY_SphereDiameter(L"Diameter");
	const CComBSTR  PROPERTY_SphereOriginX(L"OriginX");
	const CComBSTR  PROPERTY_SphereOriginY(L"OriginY");
	const CComBSTR  PROPERTY_SphereOriginZ(L"OriginZ");

	HRESULT hr = S_OK;

	try
	{
		if (pPropertyDescriptions == NULL || *pPropertyDescriptions == NULL)
			hr_throw( E_INVALIDARG, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"Invalid property description argument passed to CMEvaluateCAOBefore", "", 0 , "");
		if (pObject == NULL || *pObject == NULL)
			hr_throw( E_INVALIDARG, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"Invalid object argument passed to CMEvaluateCAOBefore", "", 0 , "");

		CComPtr<IDispatch> oOccurrenceDisp;
		hr = (*pPropertyDescriptions)->get_CAO( &oOccurrenceDisp );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"(*pPropertyDescriptions)->get_CAO failed", "", 0 , "");
		CComPtr<IJSmartOccurrence> oSmartOcc;
		hr = oOccurrenceDisp->QueryInterface( IID_IJSmartOccurrence, (void**) &oSmartOcc);
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oOccurrenceDisp->QueryInterface( IID_IJSmartOccurrenc failed", "", 0 , "");
		CComPtr<IJDAttributes> oAttrs; 
		hr = oOccurrenceDisp->QueryInterface( IID_IJDAttributes, (void**) &oAttrs);
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oOccurrenceDisp->QueryInterface( IID_IJDAttributes failed", "", 0 , "");
		CComPtr<IJDAttributesCol> oAttrCol;
		hr = oAttrs->get_CollectionOfAttributes(CComVariant(INTERFACE_Sphere), &oAttrCol );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"(*oAttrs->get_CollectionOfAttributes failed", "", 0 , "");
		hr = S_OK;
		if ( oAttrCol != NULL )
		{
			CComPtr<IJDAttribute> oAttr;
			hr = oAttrCol->get_Item(CComVariant(PROPERTY_SphereDiameter), &oAttr );
			hr = S_OK;
			if ( oAttr != NULL )
			{
				CComVariant varDiameterValue(0.2);
				hr = oAttr->put_Value( varDiameterValue );
				hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oAttr->put_Value( &varDiameterValue ) failed", "", 0 , "");
			}
			oAttr.Release();

			hr = oAttrCol->get_Item(CComVariant(PROPERTY_SphereOriginX), &oAttr );
			hr = S_OK;
			if ( oAttr != NULL )
			{
				CComVariant varOriginXValue(0.0);
				hr = oAttr->put_Value( varOriginXValue );
				hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oAttr->put_Value( &varDiameterValue ) failed", "", 0 , "");
			}
			oAttr.Release();

			hr = oAttrCol->get_Item(CComVariant(PROPERTY_SphereOriginY), &oAttr );
			hr = S_OK;
			if ( oAttr != NULL )
			{
				CComVariant varOriginYValue(0.0);
				hr = oAttr->put_Value( varOriginYValue );
				hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oAttr->put_Value( &varDiameterValue ) failed", "", 0 , "");
			}
			oAttr.Release();

			hr = oAttrCol->get_Item(CComVariant(PROPERTY_SphereOriginZ), &oAttr );
			hr = S_OK;
			if ( oAttr != NULL )
			{
				CComVariant varOriginZValue(-1.0);
				hr = oAttr->put_Value( varOriginZValue );
				hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oAttr->put_Value( &varDiameterValue ) failed", "", 0 , "");
			}
			oAttr.Release();
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

STDMETHODIMP CCAEvaluateAfterSymbolDef::CMEvaluateCAOAfter( IJDPropertyDescription** pPropertyDescriptions, IDispatch** pObject)
{
	const CComBSTR INTERFACE_Sphere(L"IJUATestDotNetSphere");
	const CComBSTR  PROPERTY_SphereDiameter(L"Diameter");
	const CComBSTR  PROPERTY_SphereOriginX(L"OriginX");
	const CComBSTR  PROPERTY_SphereOriginY(L"OriginY");

	HRESULT hr = S_OK;

	try
	{
		if (pPropertyDescriptions == NULL || *pPropertyDescriptions == NULL)
			hr_throw( E_INVALIDARG, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"Invalid property description argument passed to CMEvaluateCAOAfter", "", 0 , "");
		if (pObject == NULL || *pObject == NULL)
			hr_throw( E_INVALIDARG, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"Invalid object argument passed to CMEvaluateCAOAfter", "", 0 , "");

		CComPtr<IDispatch> oOccurrenceDisp;
		hr = (*pPropertyDescriptions)->get_CAO( &oOccurrenceDisp );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"(*pPropertyDescriptions)->get_CAO failed", "", 0 , "");
		CComPtr<IJSmartOccurrence> oSmartOcc;
		hr = oOccurrenceDisp->QueryInterface( IID_IJSmartOccurrence, (void**) &oSmartOcc);
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oOccurrenceDisp->QueryInterface( IID_IJSmartOccurrenc failed", "", 0 , "");
		CComPtr<IJDAttributes> oAttrs; 
		hr = oOccurrenceDisp->QueryInterface( IID_IJDAttributes, (void**) &oAttrs);
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oOccurrenceDisp->QueryInterface( IID_IJDAttributes failed", "", 0 , "");
		CComPtr<IJDAttributesCol> oAttrCol;
		hr = oAttrs->get_CollectionOfAttributes(CComVariant(INTERFACE_Sphere), &oAttrCol );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"(*oAttrs->get_CollectionOfAttributes failed", "", 0 , "");
		hr = S_OK;
		if ( oAttrCol != NULL )
		{
			CComPtr<IJDAttribute> oAttr;
			hr = oAttrCol->get_Item(CComVariant(PROPERTY_SphereDiameter), &oAttr );
			hr = S_OK;
			if ( oAttr != NULL )
			{
				CComVariant varDiameterValue(1.12);
				hr = oAttr->put_Value( varDiameterValue );
				hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oAttr->put_Value( &varDiameterValue ) failed", "", 0 , "");
			}
			oAttr.Release();

			hr = oAttrCol->get_Item(CComVariant(PROPERTY_SphereOriginX), &oAttr );
			hr = S_OK;
			if ( oAttr != NULL )
			{
				CComVariant varOriginXValue(1.0);
				hr = oAttr->put_Value( varOriginXValue );
				hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oAttr->put_Value( &varDiameterValue ) failed", "", 0 , "");
			}
			oAttr.Release();

			hr = oAttrCol->get_Item(CComVariant(PROPERTY_SphereOriginY), &oAttr );
			hr = S_OK;
			if ( oAttr != NULL )
			{
				CComVariant varOriginYValue(2.0);
				hr = oAttr->put_Value( varOriginYValue );
				hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolDef, __FILE__, L"oAttr->put_Value( &varDiameterValue ) failed", "", 0 , "");
			}
			oAttr.Release();
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

STDMETHODIMP CCAEvaluateAfterSymbolDef::CMConstructSphere( IJDMemberDescription* pMemberDescription, IUnknown* pResourceManager, IDispatch** pObj)
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

STDMETHODIMP CCAEvaluateAfterSymbolDef::CMSetInputSphere( IJDMemberDescription** pMemberDesc)
{
	return S_OK;
}

STDMETHODIMP CCAEvaluateAfterSymbolDef::CMFinalConstructSphere( IJDMemberDescription** pMemberDesc)
{
	return S_OK;
}

STDMETHODIMP CCAEvaluateAfterSymbolDef::CMReleaseSphere( IJDMemberDescription** pMemberDesc)
{
	return S_OK;
}

STDMETHODIMP CCAEvaluateAfterSymbolDef::CMEvaluateSphereProperties( IJDPropertyDescription** pPropertyDescriptions, IDispatch** pObject)
{
	return S_OK;
}

STDMETHODIMP CCAEvaluateAfterSymbolDef::CMEvaluateSphereGeometry( IJDPropertyDescription** pPropertyDescriptions, IDispatch** pObject)
{
	return S_OK;
}

double CCAEvaluateAfterSymbolDef::GetCAOAttributeDouble(IJDMemberDescription* pMemberDescription, BSTR InterfaceName, BSTR AttributeName)
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



