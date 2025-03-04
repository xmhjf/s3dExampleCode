// CAEvaluateAfterSymbolSym.cpp : Implementation of CCAEvaluateAfterSymbolSym

#include "stdafx.h"
#include "CAEvaluateAfterSymbolSym.h"
#include "Geom3d.h"

// CCAEvaluateAfterSymbolSym
STDMETHODIMP CCAEvaluateAfterSymbolSym::InstanciateDefinition(BSTR CodeBase, VARIANT definitionParameters, LPDISPATCH pResourceMgr, LPDISPATCH * ppSymbolDefDisp)
{
	const CComBSTR cProgId(L"ATPSO.CAEvaluateAfterSymbolSym");
	HRESULT hr = S_OK;
	
	try
	{
		if (ppSymbolDefDisp == NULL)
			hr_throw( E_INVALIDARG, IID_IJDUserSymbolServices, __FILE__, L"Invalid symbol definition argument passed to InstanciateDefinition", "", 0 , "");
		*ppSymbolDefDisp = NULL;

		// This method is in charge of the creation of the symbol definition object
		CComPtr<IJDSymbolEntitiesFactory>	pFact;

		hr = pFact.CoCreateInstance( CLSID_DSymbolEntitiesFactory, NULL, CLSCTX_ALL );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pFact.CoCreateInstance( IID_CAFactory failed", "", 0 , "");
	    
		CComPtr<IDispatch>	 pDefinitionDisp;
		hr = pFact->CreateEntity(definition, pResourceMgr, &pDefinitionDisp);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pFact->CreateCAD failed", "", 0 , "");
		CComPtr<IJDSymbolDefinition>	pDefinition;
		hr = pDefinitionDisp->QueryInterface(IID_IJDSymbolDefinition, (void**)&pDefinition);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pDefinitionDisp->QueryInterface(IID_IJSymbolDefinition failed", "", 0 , "");
	    
		hr = pDefinition->put_Name( cProgId );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pDefinition->put_Name failed", "", 0 , "");

		// Set definition progId and codebase
		hr = pDefinition->put_ProgID( cProgId );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pDefinition->put_ProgID failed", "", 0 , "");

		hr = pDefinition->put_CodeBase( CodeBase );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pDefinition->put_CodeBase failed", "", 0 , "");
	    
		// Initialize the definition
		hr = InitializeDefinition( pDefinition );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"InitializeSymbolDefinition failed", "", 0 , "");

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

STDMETHODIMP CCAEvaluateAfterSymbolSym::InitializeSymbolDefinition(IJDSymbolDefinition * * ppSymbolDefDisp)
{
	HRESULT hr = InitializeDefinition( *ppSymbolDefDisp );

	return hr;
}
	
HRESULT CCAEvaluateAfterSymbolSym::InitializeDefinition(IJDSymbolDefinition* pSymbolDefDisp)
{
	const CComBSTR CLSID_AssemblyMembers1Relationship = "{45E4020F-F8D8-47A1-9B00-C9570C1E0B17}";
	const CComBSTR cProgId(L"ATPSO.CAEvaluateAfterSymbolSym");

	const CComBSTR  INTERFACE_IJDAttributes = "{B25FD387-CFEB-11D1-850B-080036DE8E03}";
	const CComBSTR  INTERFACE_IJGeometry = "{96eb9676-6530-11d1-977f-080036754203}";
	const CComBSTR  INTERFACE_Sphere = "IJUATestDotNetSphere";
	const CComBSTR  INTERFACE_ModifyOutput = "IJUAModifyOutput";
	const CComBSTR  INTERFACE_IJSmartOccurrence = "IJSmartOccurrence";

	const CComBSTR  PROPERTY_SphereDiameter = "Diameter";
	const CComBSTR  PROPERTY_SphereOriginX = "OriginX";
	const CComBSTR  PROPERTY_SphereOriginY = "OriginY";
	const CComBSTR  PROPERTY_SphereOriginZ = "OriginZ";
	const long PARAMETERINDEX_Diameter = 2;
	const long PARAMETERINDEX_OriginX = 3;
	const long PARAMETERINDEX_OriginY = 4;
	const long PARAMETERINDEX_OriginZ = 5;
	const long PARAMETERINDEX_DeleteOutput = 6;
	const REPID	REPID_simplePhysical = 1;
	const REPID	REPID_DetailPhysical = 2;

	HRESULT hr = S_OK;
	try
	{
		if (pSymbolDefDisp == NULL)
			hr_throw( E_INVALIDARG, IID_IJDUserSymbolServices, __FILE__, L"Invalid symbol definition argument passed to InitializeDefinition", "", 0 , "");
    
		// Define the inputs -
		CComPtr<IJDInputs> oInputs;
		hr = pSymbolDefDisp->get_IJDInputs( &oInputs );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->get_IJDInputs failed", "", 0 , "");

		hr = oInputs->RemoveAllInput();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInputs->RemoveAllInput failed", "", 0 , "");

		CComPtr<IJDRepresentations> oIJDRepresentations;
		hr = pSymbolDefDisp->get_IJDRepresentations( &oIJDRepresentations );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->get_IJDRepresentations failed", "", 0 , "");
		hr = oIJDRepresentations->RemoveAllRepresentation();

		CComPtr<IJDRepresentationEvaluations> oIJDRepresentationsEvaluations;
		hr = pSymbolDefDisp->get_IJDRepresentationEvaluations( &oIJDRepresentationsEvaluations );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->get_IJDRepresentationEvaluations failed", "", 0 , "");
		hr = oIJDRepresentationsEvaluations->RemoveAllRepresentationEvaluations();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oIJDRepresentationsEvaluations->RemoveAllRepresentationEvaluations failed", "", 0 , "");
		
		// Set the inputs on the definition
		CComPtr<IJDUserMethods> iUM;
		hr = pSymbolDefDisp->QueryInterface( IID_IJDUserMethods, (void**)&iUM);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->QueryInterface( IID_IJDUserMethods failed", "", 0 , "");

		CComPtr<IJDLibraryDescription> libDesc;
		hr = libDesc.CoCreateInstance( CLSID_DLibraryDescription, NULL, CLSCTX_INPROC_SERVER );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"libDesc.CoCreateInstance( CLSID_DLibraryDescription failed", "", 0 , "");
    
		hr = libDesc->put_Name(CComBSTR(L"mySelfAsLib"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"libDesc->put_Name failed", "", 0 , "");
		hr = libDesc->put_Type(imsLIBRARY_IS_ACTIVEX);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"libDesc->put_Type failed", "", 0 , "");
		hr = libDesc->put_Properties(imsLIBRARY_AUTO_EXTRACT_METHOD_COOKIES);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"libDesc->put_Properties failed", "", 0 , "");
		hr = libDesc->put_Source(cProgId);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"libDesc->put_Source failed", "", 0 , "");
		hr = iUM->SetLibrary( libDesc, CComVariant() );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"iUM->SetLibrary failed", "", 0 , "");

		//Get the lib cookie
		long libCookie = -1;
		hr = libDesc->get_Cookie( &libCookie );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"libDesc->get_Cookie failed", "", 0 , "");
    
		// set to variable number of inputs for suppored equipments and supporting plane
		IMSDescriptionProperties inputsProp;
		hr = oInputs->get_Property( &inputsProp );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInputs->get_Property failed", "", 0 , "");
		inputsProp =  (IMSDescriptionProperties) (inputsProp | igCOLLECTION_VARIABLE);
		hr = oInputs->put_Property( inputsProp );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInputs->set_Property failed", "", 0 , "");
    
        CComPtr<IJDParameterContent> PC;
		hr = PC.CoCreateInstance(CLSID_DParameterContent, NULL, CLSCTX_INPROC_SERVER);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC.CoCreateInstance(CLSID_DParameterContent failed", "", 0 , "");
		hr = PC->put_Type(igValue);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L" PC->put_Type failed", "", 0 , "");

		CComPtr<IJDInput> oInput;
		hr = oInput.CoCreateInstance(CLSID_DInput, NULL, CLSCTX_INPROC_SERVER);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput.CoCreateInstance(CLSID_DInput failed", "", 0 , "");
    
		// SupportingObject input
		hr = oInput->put_Name(CComBSTR(L"SupportingObject"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Name failed", "", 0 , "");
		hr = oInput->put_Description(CComBSTR(L"Supporting Object"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Description failed", "", 0 , "");
		hr = oInputs->SetInput( oInput, CComVariant(1));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInputs->SetInput failed", "", 0 , "");
		hr = oInput->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->Reset() failed", "", 0 , "");
    
		// Diameter input
		hr = oInput->put_Name(CComBSTR(L"Diameter"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Name failed", "", 0 , "");
		hr = oInput->put_Description(CComBSTR(L"Sphere Diameter"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Description failed", "", 0 , "");
		hr = oInput->put_Properties( igINPUT_IS_A_PARAMETER );
		hr = PC->put_UomValue( 0.55 );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC->put_UomValue failed", "", 0 , "");
		hr = oInput->put_DefaultParameterValue( PC );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_DefaultParameterValue failed", "", 0 , "");
		hr = oInputs->SetInput( oInput, CComVariant(PARAMETERINDEX_Diameter));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInputs->SetInput failed", "", 0 , "");
		hr = oInput->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->Reset failed", "", 0 , "");
		hr = PC->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC->Reset failed", "", 0 , "");
    
		// OriginX input
		hr = oInput->put_Name(CComBSTR(L"OriginX"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Name failed", "", 0 , "");
		hr = oInput->put_Description(CComBSTR(L"Sphere X origin"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Description failed", "", 0 , "");
		hr = oInput->put_Properties( igINPUT_IS_A_PARAMETER );
		hr = PC->put_UomValue( 0.5 );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC->put_UomValue failed", "", 0 , "");
		hr = oInput->put_DefaultParameterValue( PC );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_DefaultParameterValue failed", "", 0 , "");
		hr = oInputs->SetInput( oInput, CComVariant(PARAMETERINDEX_OriginX));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInputs->SetInput failed", "", 0 , "");
		hr = oInput->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->Reset failed", "", 0 , "");
		hr = PC->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC->Reset failed", "", 0 , "");
    
		// OriginY input
		hr = oInput->put_Name(CComBSTR(L"OriginY"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Name failed", "", 0 , "");
		hr = oInput->put_Description(CComBSTR(L"Sphere Y origin"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Description failed", "", 0 , "");
		hr = oInput->put_Properties( igINPUT_IS_A_PARAMETER );
		hr = PC->put_UomValue( 0.5 );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC->put_UomValue failed", "", 0 , "");
		hr = oInput->put_DefaultParameterValue( PC );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_DefaultParameterValue failed", "", 0 , "");
		hr = oInputs->SetInput( oInput, CComVariant(PARAMETERINDEX_OriginY));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInputs->SetInput failed", "", 0 , "");
		hr = oInput->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->Reset failed", "", 0 , "");
		hr = PC->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC->Reset failed", "", 0 , "");
    
		// OriginZ input
		hr = oInput->put_Name(CComBSTR(L"OriginZ"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Name failed", "", 0 , "");
		hr = oInput->put_Description(CComBSTR(L"Sphere Z origin"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Description failed", "", 0 , "");
		hr = oInput->put_Properties( igINPUT_IS_A_PARAMETER );
		hr = PC->put_UomValue( 0.5 );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC->put_UomValue failed", "", 0 , "");
		hr = oInput->put_DefaultParameterValue( PC );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_DefaultParameterValue failed", "", 0 , "");
		hr = oInputs->SetInput( oInput, CComVariant(PARAMETERINDEX_OriginZ));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInputs->SetInput failed", "", 0 , "");
		hr = oInput->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->Reset failed", "", 0 , "");
		hr = PC->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC->Reset failed", "", 0 , "");
    
		// DeleteOutput input
		hr = oInput->put_Name(CComBSTR(L"DeleteOutput"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Name failed", "", 0 , "");
		hr = oInput->put_Description(CComBSTR(L"Delete Output?"));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_Description failed", "", 0 , "");
		hr = oInput->put_Properties( igINPUT_IS_A_PARAMETER );
		hr = PC->put_UomValue( 0.5 );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC->put_UomValue failed", "", 0 , "");
		hr = oInput->put_DefaultParameterValue( PC );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->put_DefaultParameterValue failed", "", 0 , "");
		hr = oInputs->SetInput( oInput, CComVariant(PARAMETERINDEX_DeleteOutput));
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInputs->SetInput failed", "", 0 , "");
		hr = oInput->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"oInput->Reset failed", "", 0 , "");
		hr = PC->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"PC->Reset failed", "", 0 , "");
    
		oInput.Release();
		oInputs.Release();
    
		// Representations
        CComPtr<IJDRepresentations> pIReps;
		hr = pSymbolDefDisp->QueryInterface( IID_IJDRepresentations, (void**)&pIReps);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->QueryInterface( IID_IJDRepresentations failed", "", 0 , "");
		CComPtr<IJDRepresentation> pIRep;
		hr = pIRep.CoCreateInstance(CLSID_DRepresentation, NULL, CLSCTX_INPROC_SERVER);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pIRep.CoCreateInstance(CLSID_DRepresentation failed", "", 0 , "");

		// Physical aspect
		hr = pIRep->put_Name( CComBSTR(L"Physical") );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pIRep->put_Name failed", "", 0 , "");
		hr = pIRep->put_Description( CComBSTR(L"Physical representation") );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pIRep->put_Description failed", "", 0 , "");
		hr = pIRep->put_RepresentationId( REPID_simplePhysical );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pIRep->put_RepresentationId failed", "", 0 , "");
		long methodCookie = -1;
		hr = iUM->GetMethodCookie(CComBSTR(L"Physical"), CComVariant(libCookie), &methodCookie );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"iUM->GetMethodCookie failed", "", 0 , "");
		CComPtr<IJDRepresentationStdCustomMethod> pCustomMethod;
		hr = pIRep->get_IJDRepresentationStdCustomMethod( &pCustomMethod );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pIRep->get_IJDRepresentationStdCustomMethod failed", "", 0 , "");
		hr = pCustomMethod->SetCMEvaluate( libCookie, methodCookie );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pCustomMethod->SetCMEvaluate failed", "", 0 , "");
    
		// Define outputs
		CComPtr<IJDOutputs> pOutputs;
 		hr = pIRep->QueryInterface( IID_IJDOutputs, (void**)&pOutputs);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pIRep->QueryInterface( IID_IJDOutputs failed", "", 0 , "");
		hr = pOutputs->put_Property( igCOLLECTION_VARIABLE ); // declare that the number of outputs is variable
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pOutputs->put_Property( igCOLLECTION_VARIABLE ) failed", "", 0 , "");
		CComPtr<IJDOutput> output;
		hr = output.CoCreateInstance(CLSID_DOutput, NULL, CLSCTX_INPROC_SERVER);
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"output.CoCreateInstance(CLSID_DOutput failed", "", 0 , "");
    
		// SymbolSphere output
		hr = output->put_Name( CComBSTR(L"SymbolSphere") );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L" output->put_Name failed", "", 0 , "");
		hr = output->put_Description( CComBSTR(L"Sphere from symbol") );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L" output->put_Description failed", "", 0 , "");
		hr = pOutputs->SetOutput( output );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pOutputs->SetOutput failed", "", 0 , "");
		hr = output->Reset();
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"output->Reset failed", "", 0 , "");
    
		hr = pIReps->SetRepresentation( pIRep ); // Add representation to definition
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pIReps->SetRepresentation failed", "", 0 , "");
     
		// Set definition cache properties
		hr = pSymbolDefDisp->put_CacheOption( igSYMBOL_CACHE_OPTION_NOT_SHARED );
		hr_onFail_throw( hr, IID_IJDUserSymbolServices, __FILE__, L"pSymbolDefDisp->put_CacheOption( igSYMBOL_CACHE_OPTION_NOT_SHARED ) failed", "", 0 , "");
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
	
STDMETHODIMP CCAEvaluateAfterSymbolSym::InvokeRepresentation(LPDISPATCH pSymbolOccurrence, BSTR pRepName, LPDISPATCH pOutputColl, SAFEARRAY * * arrayOfInputs)
{
	return S_OK;
}
	
STDMETHODIMP CCAEvaluateAfterSymbolSym::EditOccurence(LPDISPATCH * pSymbolOccurrence, LPDISPATCH pTransactionMgr, VARIANT_BOOL * pbHasOwnEditForm)
{
	return S_OK;
}
	
STDMETHODIMP CCAEvaluateAfterSymbolSym::GetDefinitionName(VARIANT definitionParameters, BSTR * pDefName)
{
	const CComBSTR cName(L"ATPSO.CAEvaluateAfterSymbolSym");

	*pDefName = cName.Copy();
	return S_OK;
}

HRESULT CCAEvaluateAfterSymbolSym::GetSymbolVersion( BSTR *pstrVersion )
{
    if (  !pstrVersion )              
        return E_INVALIDARG;

	*pstrVersion = SysAllocString(L"1.0.0.0");
	if ( ! *pstrVersion )
		return E_OUTOFMEMORY;

	return S_OK;
}

// CCAEvaluateAfterSymbolSym
STDMETHODIMP CCAEvaluateAfterSymbolSym::Physical(IJDRepresentationStdCustomMethod** pIRepSCM)
{
	const long PARAMETERINDEX_Diameter = 2;
	const long PARAMETERINDEX_OriginX = 3;
	const long PARAMETERINDEX_OriginY = 4;
	const long PARAMETERINDEX_OriginZ = 5;
	const long PARAMETERINDEX_DeleteOutput = 6;

	HRESULT hr = S_OK;

	try
	{
		if ( pIRepSCM == NULL )
			hr_throw( E_INVALIDARG, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"Invalid representation argument passed to Physical method.", "", 0 , "");

		CComPtr<IJDRepresentationDuringGame> pRepDG;
		hr = (*pIRepSCM)->QueryInterface( IID_IJDRepresentationDuringGame, (void**)&pRepDG );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pIRepSCM->QueryInterface( IID_IJDRepresentationDuringGame failed", "", 0 , "");

		CComPtr<IJDOutputCollection> pOC;
		hr = pRepDG->get_OutputCollection( &pOC );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pRepDG->get_OutputCollection failed", "", 0 , "");

		// remove all outputs
		CComPtr<IJDSymbolDefinition> oDefinition;
		hr = pOC->get_Definition( &oDefinition );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pOC->get_Definition failed", "", 0 , "");
		CComPtr<IJDRepresentations> oRepresentations;
		hr = oDefinition->get_IJDRepresentations( &oRepresentations );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pOC->get_Definition failed", "", 0 , "");
		CComPtr<IJDRepresentation> oRep;
		hr = oRepresentations->GetRepresentationByName( CComBSTR(L"Physical"), &oRep );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oRepresentations->GetRepresentationByName( CComBSTR(Physical) failed", "", 0 , "");
		CComPtr<IJDOutputs> oOutputs;
		hr = oRep->QueryInterface( IID_IJDOutputs, (void**)&oOutputs );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oRep->QueryInterface( IID_IJDOutputs failed", "", 0 , "");
		hr = oOutputs->RemoveAllOutput();
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oOutputs->RemoveAllOutput failed", "", 0 , "");

		CComPtr<IJDInputs> pInputs;
		hr = oDefinition->get_IJDInputs( &pInputs );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oOutputs->RemoveAllOutput failed", "", 0 , "");

        // assign to meaningful variables from the input array
		double lDiameter = 0.0;
		hr = GetInputValue( pInputs, PARAMETERINDEX_Diameter, &lDiameter );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"GetInputValue( pInputs, PARAMETERINDEX_Diameter failed", "", 0 , "");

		double lOriginX = 0.0;
		hr = GetInputValue( pInputs, PARAMETERINDEX_OriginX, &lOriginX );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"GetInputValue( pInputs, PARAMETERINDEX_OriginX failed", "", 0 , "");

		double lOriginY = 0.0;
		hr = GetInputValue( pInputs, PARAMETERINDEX_OriginY, &lOriginY );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"GetInputValue( pInputs, PARAMETERINDEX_OriginY failed", "", 0 , "");

		double lOriginZ = 0.0;
		hr = GetInputValue( pInputs, PARAMETERINDEX_OriginZ, &lOriginZ );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"GetInputValue( pInputs, PARAMETERINDEX_OriginZ failed", "", 0 , "");

		CComPtr<IJGeometryFactory> oGeomFactory;
		hr = oGeomFactory.CoCreateInstance(CLSID_GeometryFactory, NULL, CLSCTX_INPROC_SERVER);
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oGeomFactory.CoCreateInstance(CLSID_GeometryFactory failed", "", 0 , "");
		CComPtr<ISpheres3d>	pSpheres;
		hr = oGeomFactory->get_Spheres3d( &pSpheres );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oGeomFactory->get_Spheres3d failed", "", 0 , "");
		CComPtr<IUnknown>	pResourceManager;
		hr = pOC->get_ResourceManager( &pResourceManager );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pOC->get_ResourceManager failed", "", 0 , "");
		CComPtr<IJSphere> pSphere;
		hr = pSpheres->CreateByCenterRadius(pResourceManager, lOriginX, lOriginY, lOriginZ, lDiameter / 2.0, VARIANT_TRUE, &pSphere);
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pSpheres->CreateByCenterRadius failed", "", 0 , "");

		CComPtr<IJDOutput>	oOutput;
		hr = InitNewOutput(pOC, CComBSTR(L"SymbolSphere"), &oOutput);
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"InitNewOutput failed", "", 0 , "");
		CComBSTR	pAutoName;
		hr = pOC->AddOutput( CComBSTR(L"SymbolSphere"), pSphere, &pAutoName );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pOC->AddOutput failed", "", 0 , "");
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

HRESULT CCAEvaluateAfterSymbolSym::GetInputValue( IJDInputs *pInputs, long inputIndex, double *value )
{
	HRESULT hr = S_OK;

	try
	{
		if ( pInputs == NULL )
			hr_throw( E_INVALIDARG, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"Invalid pInputs argument passed to GetInputValue.", "", 0 , "");
		if ( value == NULL )
			hr_throw( E_INVALIDARG, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"Invalid value argument passed to GetInputValue.", "", 0 , "");

		CComPtr<IJDInput>	pInput;
		hr = pInputs->GetInputByIndex(inputIndex, &pInput);
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pInputs->GetInputByIndex failed", "", 0 , "");
		CComPtr<IJDInputDuringGame>	pInputDuringGame;
		hr = pInput->get_IJDInputDuringGame( &pInputDuringGame );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pInput->get_IJDInputDuringGame failed", "", 0 , "");
		CComPtr<IDispatch>	pResult;
		hr = pInputDuringGame->get_Result( &pResult );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pInput->get_IJDInputDuringGame failed", "", 0 , "");
		CComPtr<IJDParameterContent>	pc;
		hr = pResult->QueryInterface( IID_IJDParameterContent, (void**) &pc );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pResult->QueryInterface( IID_IJDParameterContent failed", "", 0 , "");
		hr = pc->get_UomValue( value );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pc->get_UomValue failed", "", 0 , "");
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


HRESULT CCAEvaluateAfterSymbolSym::InitNewOutput(IJDOutputCollection *pOC, BSTR outputName, IJDOutput **pOutput)
{
	HRESULT hr = S_OK;

	try
	{
		if ( pOC == NULL )
			hr_throw( E_INVALIDARG, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"Invalid pOC argument passed to InitNewOutput", "", 0 , "");
		if ( pOutput == NULL )
			hr_throw( E_INVALIDARG, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"Invalid pOutput argument passed to InitNewOutput", "", 0 , "");

		CComPtr<IJDRepresentation> oRep;
		CComPtr<IJDSymbolDefinition> oDefinition;
		hr = pOC->get_Definition( &oDefinition );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"pOC->get_Definition failed", "", 0 , "");
		CComPtr<IJDRepresentations> oRepresentations;
		hr = oDefinition->get_IJDRepresentations( &oRepresentations );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oDefinition->get_IJDRepresentations failed", "", 0 , "");
		hr = oRepresentations->GetRepresentationByName(CComBSTR(L"Physical"), &oRep);
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oRepresentations->GetRepresentationByName failed", "", 0 , "");
		CComPtr<IJDOutputs> oOutputs;
		hr = oRep->get_IJDOutputs( &oOutputs );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oRep->get_IJDOutputs failed", "", 0 , "");
		
		CComPtr<IJDOutput>	oOutput;
		hr = oOutput.CoCreateInstance(CLSID_DOutput, NULL, CLSCTX_INPROC_SERVER);
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oOutput.CoCreateInstance(CLSID_DOutput failed", "", 0 , "");
		hr = oOutput->put_Name( outputName );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oOutput->put_Name failed", "", 0 , "");
		hr = oOutput->put_Description( outputName );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oOutput->put_Description failed", "", 0 , "");
		hr = oOutputs->SetOutput( oOutput );
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oOutputs->SetOutput failed", "", 0 , "");
		hr = oOutput->Reset();
		hr_onFail_throw( hr, IID_ICAEvaluateAfterSymbolSym, __FILE__, L"oOutput.Reset failed", "", 0 , "");
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


