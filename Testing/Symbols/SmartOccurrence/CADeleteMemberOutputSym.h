// CADeleteMemberOutputSym.h : Declaration of the CCADeleteMemberOutputSym

#pragma once
#include "resource.h"       // main symbols
#include "SmartOccurrence.h"
#include "SblEntities.h"

#include "ATPSO_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CCADeleteMemberOutputSym

class ATL_NO_VTABLE CCADeleteMemberOutputSym :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCADeleteMemberOutputSym, &CLSID_CADeleteMemberOutputSym>,
	public IDispatchImpl<ICADeleteMemberOutputSym, &IID_ICADeleteMemberOutputSym, &LIBID_ATPSO,/*wMajor =*/ 0xFFFF, /*wMinor =*/ 0xFFFF>,
	public IDispatchImpl<IJDUserSymbolServices, &IID_IJDUserSymbolServices, &LIBID_GSCADSmartOccurrence, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispatchImpl<IJSymbolVersion,&IID_IJSymbolVersion, &LIBID_GSCADSmartOccurrence>
{
public:
	CCADeleteMemberOutputSym()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CADELETEMEMBEROUTPUTSYM)


BEGIN_COM_MAP(CCADeleteMemberOutputSym)
	COM_INTERFACE_ENTRY2(IDispatch,ICADeleteMemberOutputSym)
	COM_INTERFACE_ENTRY(ICADeleteMemberOutputSym)
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


public:
	// ICADeleteMemberOutputSym
    STDMETHOD(Physical)(/*[in, out]*/ IJDRepresentationStdCustomMethod** pIRepSCM);

private:
	HRESULT InitializeDefinition(IJDSymbolDefinition* pSymbolDefDisp);
	HRESULT GetInputValue( IJDInputs *pInputs, long inputIndex, double *value );
	HRESULT InitNewOutput(IJDOutputCollection *pOC, BSTR outputName, IJDOutput** pOutput);
};

OBJECT_ENTRY_AUTO(__uuidof(CADeleteMemberOutputSym), CCADeleteMemberOutputSym)
