// COM2ndSelectionRuleSel.h : Declaration of the CCOM2ndSelectionRuleSel

#pragma once
#include "resource.h"       // main symbols
#include "SmartOccurrence.h"
#include "SblEntities.h"

#include "ATPSO_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CCOM2ndSelectionRuleSel

class ATL_NO_VTABLE CCOM2ndSelectionRuleSel :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCOM2ndSelectionRuleSel, &CLSID_COM2ndSelectionRuleSel>,
	public IDispatchImpl<ICOM2ndSelectionRuleSel, &IID_ICOM2ndSelectionRuleSel, &LIBID_ATPSO,/*wMajor =*/ 0xFFFF, /*wMinor =*/ 0xFFFF>,
	public IDispatchImpl<IJDUserSymbolServices, &IID_IJDUserSymbolServices, &LIBID_GSCADSmartOccurrence, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispatchImpl<IJSymbolVersion,&IID_IJSymbolVersion, &LIBID_GSCADSmartOccurrence>
{
public:
	CCOM2ndSelectionRuleSel()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_COM2NDSELECTIONRULESEL)


BEGIN_COM_MAP(CCOM2ndSelectionRuleSel)
	COM_INTERFACE_ENTRY2(IDispatch, ICOM2ndSelectionRuleSel)
	COM_INTERFACE_ENTRY(ICOM2ndSelectionRuleSel)
	COM_INTERFACE_ENTRY(IJDUserSymbolServices)
	COM_INTERFACE_ENTRY(IJSymbolVersion)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:
	// IJDUserSymbolServices Methods
	STDMETHOD(InstanciateDefinition)(BSTR CodeBase, VARIANT definitionParameters, LPDISPATCH pResourceMgr, LPDISPATCH * ppSymbolDefDisp);
	STDMETHOD(InitializeSymbolDefinition)(IJDSymbolDefinition * * ppSymbolDefDisp);
	STDMETHOD(InvokeRepresentation)(LPDISPATCH pSymbolOccurrence, BSTR pRepName, LPDISPATCH pOutputColl, SAFEARRAY * * arrayOfInputs);;
	STDMETHOD(EditOccurence)(LPDISPATCH * pSymbolOccurrence, LPDISPATCH pTransactionMgr, VARIANT_BOOL * pbHasOwnEditForm);
	STDMETHOD(GetDefinitionName)(VARIANT definitionParameters, BSTR * pDefName);

	// IJSymbolVersion
	STDMETHOD(GetSymbolVersion)(BSTR *pstrVersion );

private:
	HRESULT InstanciateSelector(BSTR Selector_Progid, BSTR Selector_CodeBase, BSTR Selector_Name, LPDISPATCH pResourceMgr, LPDISPATCH *ppInstanciatedDefinition);
	HRESULT InitAbstractSelector( IJDSymbolDefinition *pSelector );
	HRESULT AddQuestion(IJDSymbolDefinition *pDefinition, long index, BSTR question, BSTR defaultValue, BSTR CodelistTableName);
	
	static const CComBSTR cSelectorName;

public:
	static const CComBSTR cSelectorProgid; //Used by smart occurrence

public:
	//Note: methods has dispatchId 0x60030005 to be compatible with old VB code.
	STDMETHOD(CMSelector)(IJDRepresentation * pRepresentation);
};

OBJECT_ENTRY_AUTO(__uuidof(COM2ndSelectionRuleSel), CCOM2ndSelectionRuleSel)
