// CAEvaluateAfterSymbolDef.h : Declaration of the CCAEvaluateAfterSymbolDef

#pragma once
#include "resource.h"       // main symbols
#include "SmartOccurrence.h"
#include "SblEntities.h"

#include "ATPSO_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CCAEvaluateAfterSymbolDef

class ATL_NO_VTABLE CCAEvaluateAfterSymbolDef :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCAEvaluateAfterSymbolDef, &CLSID_CAEvaluateAfterSymbolDef>,
	public IDispatchImpl<ICAEvaluateAfterSymbolDef, &IID_ICAEvaluateAfterSymbolDef, &LIBID_ATPSO,/*wMajor =*/ 0xFFFF, /*wMinor =*/ 0xFFFF>,
	public IDispatchImpl<IJDUserSymbolServices, &IID_IJDUserSymbolServices, &LIBID_GSCADSmartOccurrence, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispatchImpl<IJSymbolVersion,&IID_IJSymbolVersion, &LIBID_GSCADSmartOccurrence>
{
public:
	CCAEvaluateAfterSymbolDef()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CAEVALUATEAFTERSYMBOLDEF)


BEGIN_COM_MAP(CCAEvaluateAfterSymbolDef)
	COM_INTERFACE_ENTRY2(IDispatch, ICAEvaluateAfterSymbolDef)
	COM_INTERFACE_ENTRY(ICAEvaluateAfterSymbolDef)
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
	// ICAEvaluateAfterSymbolDef
	STDMETHOD(CMFinalConstructAsm)(/*[in, out]*/ IJDAggregatorDescription** pAggregatorDescription);
    STDMETHOD(CMConstructAsm)(/*[in, out]*/ IJDAggregatorDescription** pAggregatorDescription);
    STDMETHOD(CMEvaluateCAOBefore)(/*[in, out]*/ IJDPropertyDescription** pPropertyDescriptions, 
                                   /*[in, out]*/ IDispatch** pObject);
    STDMETHOD(CMEvaluateCAOAfter)(/*[in, out]*/ IJDPropertyDescription** pPropertyDescriptions, 
                                  /*[in, out]*/ IDispatch** pObject);
    STDMETHOD(CMConstructSphere)(/*[in]*/ IJDMemberDescription* pMemberDescription, 
                                 /*[in]*/ IUnknown* pResourceManager, 
                                 /*[in, out]*/ IDispatch** pObj);
    STDMETHOD(CMSetInputSphere)(/*[in, out]*/ IJDMemberDescription** pMemberDesc);
    STDMETHOD(CMFinalConstructSphere)(/*[in, out]*/ IJDMemberDescription** pMemberDesc);
    STDMETHOD(CMReleaseSphere)(/*[in, out]*/ IJDMemberDescription** pMemberDesc);
    STDMETHOD(CMEvaluateSphereProperties)(/*[in, out]*/ IJDPropertyDescription** pPropertyDescriptions, 
                                          /*[in, out]*/ IDispatch** pObject);
    STDMETHOD(CMEvaluateSphereGeometry)(/*[in, out]*/ IJDPropertyDescription** pPropertyDescriptions, 
                                        /*[in, out]*/ IDispatch** pObject);

private:
	HRESULT InitializeCustomAssemblyDefinition(IJDSymbolDefinition *pSymbolDefDisp);
	double GetCAOAttributeDouble(IJDMemberDescription* pMemberDescription, BSTR InterfaceName, BSTR AttributeName);
};

OBJECT_ENTRY_AUTO(__uuidof(CAEvaluateAfterSymbolDef), CCAEvaluateAfterSymbolDef)
