// COM2ndSelectionRuleSel.cpp : Implementation of CCOM2ndSelectionRuleSel

#include "stdafx.h"
#include "COM2ndSelectionRuleSel.h"

// CCOM2ndSelectionRuleSel
const CComBSTR CCOM2ndSelectionRuleSel::cSelectorProgid(L"ATPSO.COM2ndSelectionRuleSel");
const CComBSTR CCOM2ndSelectionRuleSel::cSelectorName(L"ATPSO.COM2ndSelectionRuleSel");

STDMETHODIMP CCOM2ndSelectionRuleSel::InstanciateDefinition(BSTR CodeBase, VARIANT definitionParameters, LPDISPATCH pResourceMgr, LPDISPATCH * ppSymbolDefDisp)
{
	HRESULT hr = S_OK;
	hr = InstanciateSelector(cSelectorProgid, CodeBase, cSelectorName, pResourceMgr, ppSymbolDefDisp);
	return hr;
}
	
STDMETHODIMP CCOM2ndSelectionRuleSel::InitializeSymbolDefinition(IJDSymbolDefinition * * ppSymbolDefDisp)
{
	HRESULT hr = S_OK;
	_ASSERT(ppSymbolDefDisp);

	hr = InitAbstractSelector(*ppSymbolDefDisp);

	// Add the questions
	if ( SUCCEEDED(hr) )
	{
		hr = AddQuestion(*ppSymbolDefDisp, 1, CComBSTR(L"QuestionQ"), CComBSTR(L"Q"), CComBSTR(L""));
	}

	return hr;
}
	
STDMETHODIMP CCOM2ndSelectionRuleSel::InvokeRepresentation(LPDISPATCH pSymbolOccurrence, BSTR pRepName, LPDISPATCH pOutputColl, SAFEARRAY * * arrayOfInputs)
{
	return S_OK;
}
	
STDMETHODIMP CCOM2ndSelectionRuleSel::EditOccurence(LPDISPATCH * pSymbolOccurrence, LPDISPATCH pTransactionMgr, VARIANT_BOOL * pbHasOwnEditForm)
{
	return S_OK;
}
	
STDMETHODIMP CCOM2ndSelectionRuleSel::GetDefinitionName(VARIANT definitionParameters, BSTR * pDefName)
{
	 *pDefName = cSelectorName.Copy();
	 return S_OK;
}

HRESULT CCOM2ndSelectionRuleSel::InstanciateSelector(BSTR Selector_Progid, BSTR Selector_CodeBase, BSTR Selector_Name, LPDISPATCH pResourceMgr, LPDISPATCH *ppInstanciatedDefinition)
{
	HRESULT hr = S_OK;

	 try 
	 {	
		CComPtr<IDispatch> pSelectorDisp;
		CComQIPtr<IJDSymbolDefinition,&IID_IJDSymbolDefinition> pSelector;
		CComPtr<IJDSymbolEntitiesFactory> pSymbolFactory;
	
		hr = pSymbolFactory.CoCreateInstance(CLSID_DSymbolEntitiesFactory, NULL, CLSCTX_ALL);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"CoCreateInstance(CLSID_DSymbolEntitiesFactory", "", 0 , "");
	
		hr = pSymbolFactory->CreateEntity(definition, pResourceMgr,&pSelectorDisp);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSymbolFactory->CreateEntity", "", 0 , "");

		pSelector = pSelectorDisp;
		
		// Set definition progId and codebase
		hr = pSelector->put_ProgID(Selector_Progid);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSelector->put_ProgID", "", 0 , "");
		
		hr = pSelector->put_CodeBase(Selector_CodeBase);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSelector->put_CodeBase", "", 0 , "");

		hr = pSelector->put_Name(Selector_Name);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSelector->put_Name", "", 0 , "");

		// The symbol definition data are persisted in order to store the Answer interface IID.
		hr = pSelector->put_MetaDataOption(igSYMBOL_DYNAMIC_METADATA);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSelector->put_MetaDataOption", "", 0 , "");

		hr = pSelector->put_SupportOnlyOption(igSYMBOL_NOT_SUPPORT_ONLY);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSelector->put_SupportOnlyOption", "", 0 , "");

		hr = pSelector->Update();
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSelector->Update", "", 0 , "");

		// return the selector
		*ppInstanciatedDefinition = pSelectorDisp.Detach();
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


HRESULT CCOM2ndSelectionRuleSel::InitAbstractSelector( IJDSymbolDefinition *pSelector )
{
	HRESULT hr = S_OK;

	 try 
	 {	
		const CComBSTR SELECTOR_REPRESENTATION_NAME(L"Results");
		const CComBSTR SELECTOR_DESCRIPTION(L"Returns the list of available Smart Objects");
		long mthCookie = 0;

		CComPtr<IJDInputs> pIJDInputs;
		CComPtr<IJDRepresentations> pIJDRepresentations;
		CComPtr<IJDUserMethods> pIJDUserMethods;

		if(!pSelector)
			hr_throw( E_INVALIDARG, IID_NULL, __FILE__, L"pSelector is NULL", "", 0 , "");

		// Clean previous data
		hr = pSelector->get_IJDInputs(&pIJDInputs);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSelector->get_IJDInputs", "", 0 , "");
		
		hr = pIJDInputs->RemoveAllInput();
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSelector->RemoveAllInput", "", 0 , "");

		hr = pSelector->get_IJDRepresentations(&pIJDRepresentations);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSelector->get_IJDRepresentations", "", 0 , "");

		hr = pIJDRepresentations->RemoveAllRepresentation();
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pIJDRepresentations->RemoveAllRepresentation", "", 0 , "");


		// ---------------------------------------------------
		// -         "Results" representation                -
		// 
		// It returns the list of available Smart Objects
		// regarding the input arguments.
		// 
		// Note : The first output is the system default
		// ---------------------------------------------------
		CComPtr<IJDRepresentation> pIJDRepresentation;

		hr = pIJDRepresentation.CoCreateInstance(CLSID_DRepresentation, NULL, CLSCTX_ALL);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"CoCreateInstance CLSID_DRepresentation", "", 0 , "");

		hr = pIJDRepresentation->put_Name(SELECTOR_REPRESENTATION_NAME);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pIJDRepresentation->put_Name", "", 0 , "");

		hr = pIJDRepresentation->put_Description(SELECTOR_DESCRIPTION);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pIJDRepresentation->put_Description", "", 0 , "");

		hr = pSelector->get_IJDUserMethods(&pIJDUserMethods);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pSelector->get_IJDUserMethods", "", 0 , "");

		hr = pIJDUserMethods->GetMethodCookie(CComBSTR(L"CMSelector"),CComVariant(imsCOOKIE_ID_USS_LIB),&mthCookie);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pIJDUserMethods->GetMethodCookie", "", 0 , "");
	
		CComPtr<IJDRepresentationStdCustomMethod> pIJDRepresentationStdCustomMethod;
		hr = pIJDRepresentation->get_IJDRepresentationStdCustomMethod(&pIJDRepresentationStdCustomMethod);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pIJDRepresentation->get_IJDRepresentationStdCustomMethod", "", 0 , "");

		hr = pIJDRepresentationStdCustomMethod->SetCMEvaluate(imsCOOKIE_ID_USS_LIB,mthCookie);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pIJDRepresentationStdCustomMethod->SetCMEvaluate", "", 0 , "");

		CComPtr<IJDOutputs> pIJDOutputs;
		hr = pIJDRepresentation->get_IJDOutputs(&pIJDOutputs);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pIJDRepresentation->get_IJDOutputs", "", 0 , "");

		IMSDescriptionProperties pDescriptionProperties;
		hr = pIJDOutputs->get_Property(&pDescriptionProperties);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pIJDOutputs->get_Property", "", 0 , "");

		pDescriptionProperties =  static_cast<IMSDescriptionProperties>(pDescriptionProperties | igCOLLECTION_VARIABLE);
		hr = pIJDOutputs->put_Property(pDescriptionProperties);
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pIJDOutputs->put_Property", "", 0 , "");

		hr = pIJDRepresentations->Add(pIJDRepresentation,CComVariant());
		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pIJDRepresentations->Add", "", 0 , "");

		pIJDRepresentation->Reset();
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

STDMETHODIMP CCOM2ndSelectionRuleSel::CMSelector(IJDRepresentation * pRepresentation)
{
	const CComBSTR SELECTOR_OUTPUT_NAME(L"Choice_");

	HRESULT hr = S_OK;

	try
	{
		if (pRepresentation == NULL)
			hr_throw( E_INVALIDARG, IID_ICOMSelectionRuleSel, __FILE__, L"Invalid pRepresentation argument passed to CMSelector", "", 0 , "");

		CComPtr<IJDRepresentationDuringGame>	oRepresentationDuringGame;
		hr = pRepresentation->get_IJDRepresentationDuringGame( &oRepresentationDuringGame );
		hr_onFail_throw( hr, IID_ICOMSelectionRuleSel, __FILE__, L"pRepresentation->get_IJDRepresentationDuringGame failed", "", 0 , "");
		CComPtr<IJDOutputCollection>	oOutputColl;
		hr = oRepresentationDuringGame->get_OutputCollection( &oOutputColl );
		hr_onFail_throw( hr, IID_ICOMSelectionRuleSel, __FILE__, L"oRepresentationDuringGame->get_OutputCollection failed", "", 0 , "");
		CComBSTR sAutoName;
		
		CComPtr<IJDParameterContent>	 pChoice;
		hr = pChoice.CoCreateInstance( CLSID_DParameterContent, NULL, CLSCTX_INPROC_SERVER );
		hr_onFail_throw( hr, IID_ICOMSelectionRuleSel, __FILE__, L"pChoice.CoCreateInstance( CLSID_DParameterContent failed", "", 0 , "");

		hr = pChoice->put_Type( igString );
		hr_onFail_throw( hr, IID_ICOMSelectionRuleSel, __FILE__, L"pChoice->put_Type failed", "", 0 , "");
		hr = pChoice->put_String( CComBSTR(L"TestCOMSelectionRules") );
		hr_onFail_throw( hr, IID_ICOMSelectionRuleSel, __FILE__, L"pChoice->put_String( CComBSTR(TestCOMSelectionRules failed", "", 0 , "");

		hr = oOutputColl->AddOutput( SELECTOR_OUTPUT_NAME, pChoice, &sAutoName );
		hr_onFail_throw( hr, IID_ICOMSelectionRuleSel, __FILE__, L"oOutputColl->AddOutput failed", "", 0 , "");
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
HRESULT CCOM2ndSelectionRuleSel::GetSymbolVersion( BSTR *pstrVersion )
{
    if (  !pstrVersion )              
        return E_INVALIDARG;

	*pstrVersion = SysAllocString(L"1.0.0.0");
	if ( ! *pstrVersion )
		return E_OUTOFMEMORY;

	return S_OK;
}

HRESULT CCOM2ndSelectionRuleSel::AddQuestion(IJDSymbolDefinition *pDefinition, long index, BSTR question, BSTR defaultValue, BSTR CodelistTableName)
{
	HRESULT hr = S_OK;

	try
	{
		CComPtr<IJDInput>  oInput;
		hr = oInput.CoCreateInstance( CLSID_DInput, NULL, CLSCTX_INPROC_SERVER );
  		hr_onFail_throw( hr, IID_NULL, __FILE__, L"oInput.CoCreateInstance( CLSID_DInput failed", "", 0 , "");

		hr = oInput->put_Name( question );
  		hr_onFail_throw( hr, IID_NULL, __FILE__, L"oInput->put_Name failed", "", 0 , "");
		hr = oInput->put_Properties( igINPUT_IS_A_PARAMETER );
  		hr_onFail_throw( hr, IID_NULL, __FILE__, L"oInput->put_Properties failed", "", 0 , "");
		CComPtr<IJDParameterContent> oPC; 
		hr = oPC.CoCreateInstance( CLSID_DParameterContent, NULL, CLSCTX_INPROC_SERVER );
  		hr_onFail_throw( hr, IID_NULL, __FILE__, L"oPC.CoCreateInstance( CLSID_DParameterContent failed", "", 0 , "");
		// type is string in this Test symbol
		hr = oPC->put_Type( igString );
  		hr_onFail_throw( hr, IID_NULL, __FILE__, L"oPC->put_Type failed", "", 0 , "");
		hr = oPC->put_String( defaultValue );
  		hr_onFail_throw( hr, IID_NULL, __FILE__, L"oPC->put_String failed", "", 0 , "");
		hr = oInput->put_DefaultParameterValue( oPC );
  		hr_onFail_throw( hr, IID_NULL, __FILE__, L"oInput->put_DefaultParameterValue failed", "", 0 , "");
		// Handle code list if any
		CComBSTR sDescription(L"<ThisIsAQuestion>");
		if ( CodelistTableName && wcslen(CodelistTableName) > 0 )
		{
			sDescription += L"<JCodeListTable><Name>";
			sDescription += CodelistTableName;
			sDescription += L"</Name></JCodeListTable>";
		}  
		sDescription += L"</ThisIsAQuestion>";
		hr = oInput->put_Description( sDescription );
  		hr_onFail_throw( hr, IID_NULL, __FILE__, L"oInput->put_Description failed", "", 0 , "");
		CComPtr<IJDInputs>  oInputs;
		hr = pDefinition->get_IJDInputs( &oInputs );
  		hr_onFail_throw( hr, IID_NULL, __FILE__, L"pDefinition->get_IJDInputs failed", "", 0 , "");
		hr = oInputs->Add( oInput, CComVariant(index) );
  		hr_onFail_throw( hr, IID_NULL, __FILE__, L"oInputs->Add failed", "", 0 , "");
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
