// CADeleteMemberOutputDef.h : Declaration of the CCADeleteMemberOutputDef

#pragma once
#include "resource.h"       // main symbols
#include "SmartOccurrence.h"
#include "SblEntities.h"

#include "ATPSO_i.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CCADeleteMemberOutputDef

class ATL_NO_VTABLE CCADeleteMemberOutputDef :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCADeleteMemberOutputDef, &CLSID_CADeleteMemberOutputDef>,
	public IDispatchImpl<ICADeleteMemberOutputDef, &IID_ICADeleteMemberOutputDef, &LIBID_ATPSO,/*wMajor =*/ 0xFFFF, /*wMinor =*/ 0xFFFF>,
	public IDispatchImpl<IJDUserSymbolServices, &IID_IJDUserSymbolServices, &LIBID_GSCADSmartOccurrence, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispatchImpl<IJSymbolVersion,&IID_IJSymbolVersion, &LIBID_GSCADSmartOccurrence>
{
public:
	CCADeleteMemberOutputDef()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CADELETEMEMBEROUTPUTDEF)


BEGIN_COM_MAP(CCADeleteMemberOutputDef)
	COM_INTERFACE_ENTRY2(IDispatch, ICADeleteMemberOutputDef)
	COM_INTERFACE_ENTRY(ICADeleteMemberOutputDef)
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
	STDMETHOD(CMFinalConstructAsm)(/*[in, out]*/ IJDAggregatorDescription** pAggregatorDescription);
	STDMETHOD(CMConstructAsm)(/*[in, out]*/ IJDAggregatorDescription** pAggregatorDescription);
	STDMETHOD(CMEvaluateCAO)(/*[in, out]*/ IJDPropertyDescription** pPropertyDescriptions, 
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
	HRESULT InitializeCustomAssemblyDefinition(IJDSymbolDefinition* pSymbolDefDisp);
	double GetCAOAttributeDouble(IJDMemberDescription* pMemberDescription, BSTR InterfaceName, BSTR AttributeName);

};

OBJECT_ENTRY_AUTO(__uuidof(CADeleteMemberOutputDef), CCADeleteMemberOutputDef)
