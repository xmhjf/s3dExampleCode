

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Thu Oct 13 16:55:15 2016
 */
/* Compiler settings for ATPSO.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __ATPSO_i_h__
#define __ATPSO_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ICADeleteMemberOutputDef_FWD_DEFINED__
#define __ICADeleteMemberOutputDef_FWD_DEFINED__
typedef interface ICADeleteMemberOutputDef ICADeleteMemberOutputDef;

#endif 	/* __ICADeleteMemberOutputDef_FWD_DEFINED__ */


#ifndef __ICADeleteMemberOutputSym_FWD_DEFINED__
#define __ICADeleteMemberOutputSym_FWD_DEFINED__
typedef interface ICADeleteMemberOutputSym ICADeleteMemberOutputSym;

#endif 	/* __ICADeleteMemberOutputSym_FWD_DEFINED__ */


#ifndef __ICAEvaluateAfterSymbolSym_FWD_DEFINED__
#define __ICAEvaluateAfterSymbolSym_FWD_DEFINED__
typedef interface ICAEvaluateAfterSymbolSym ICAEvaluateAfterSymbolSym;

#endif 	/* __ICAEvaluateAfterSymbolSym_FWD_DEFINED__ */


#ifndef __ICAEvaluateAfterSymbolDef_FWD_DEFINED__
#define __ICAEvaluateAfterSymbolDef_FWD_DEFINED__
typedef interface ICAEvaluateAfterSymbolDef ICAEvaluateAfterSymbolDef;

#endif 	/* __ICAEvaluateAfterSymbolDef_FWD_DEFINED__ */


#ifndef __ICOMSelectionRuleSel_FWD_DEFINED__
#define __ICOMSelectionRuleSel_FWD_DEFINED__
typedef interface ICOMSelectionRuleSel ICOMSelectionRuleSel;

#endif 	/* __ICOMSelectionRuleSel_FWD_DEFINED__ */


#ifndef __ICOMSelectionRuleRootSel_FWD_DEFINED__
#define __ICOMSelectionRuleRootSel_FWD_DEFINED__
typedef interface ICOMSelectionRuleRootSel ICOMSelectionRuleRootSel;

#endif 	/* __ICOMSelectionRuleRootSel_FWD_DEFINED__ */


#ifndef __ICOM2ndSelectionRuleSel_FWD_DEFINED__
#define __ICOM2ndSelectionRuleSel_FWD_DEFINED__
typedef interface ICOM2ndSelectionRuleSel ICOM2ndSelectionRuleSel;

#endif 	/* __ICOM2ndSelectionRuleSel_FWD_DEFINED__ */


#ifndef __ITestSelRuleWithQs_FWD_DEFINED__
#define __ITestSelRuleWithQs_FWD_DEFINED__
typedef interface ITestSelRuleWithQs ITestSelRuleWithQs;

#endif 	/* __ITestSelRuleWithQs_FWD_DEFINED__ */


#ifndef __ICOMTestParameterRule_FWD_DEFINED__
#define __ICOMTestParameterRule_FWD_DEFINED__
typedef interface ICOMTestParameterRule ICOMTestParameterRule;

#endif 	/* __ICOMTestParameterRule_FWD_DEFINED__ */


#ifndef __ICAForParameterRuleDef_FWD_DEFINED__
#define __ICAForParameterRuleDef_FWD_DEFINED__
typedef interface ICAForParameterRuleDef ICAForParameterRuleDef;

#endif 	/* __ICAForParameterRuleDef_FWD_DEFINED__ */


#ifndef __ICAForParameterRuleSym_FWD_DEFINED__
#define __ICAForParameterRuleSym_FWD_DEFINED__
typedef interface ICAForParameterRuleSym ICAForParameterRuleSym;

#endif 	/* __ICAForParameterRuleSym_FWD_DEFINED__ */


#ifndef __ISONoGraphics_FWD_DEFINED__
#define __ISONoGraphics_FWD_DEFINED__
typedef interface ISONoGraphics ISONoGraphics;

#endif 	/* __ISONoGraphics_FWD_DEFINED__ */


#ifndef __ISO1GraphicOutput_FWD_DEFINED__
#define __ISO1GraphicOutput_FWD_DEFINED__
typedef interface ISO1GraphicOutput ISO1GraphicOutput;

#endif 	/* __ISO1GraphicOutput_FWD_DEFINED__ */


#ifndef __ICANestedOutputParentSym_FWD_DEFINED__
#define __ICANestedOutputParentSym_FWD_DEFINED__
typedef interface ICANestedOutputParentSym ICANestedOutputParentSym;

#endif 	/* __ICANestedOutputParentSym_FWD_DEFINED__ */


#ifndef __ICANestedOutputParentDef_FWD_DEFINED__
#define __ICANestedOutputParentDef_FWD_DEFINED__
typedef interface ICANestedOutputParentDef ICANestedOutputParentDef;

#endif 	/* __ICANestedOutputParentDef_FWD_DEFINED__ */


#ifndef __ICABasicDef_FWD_DEFINED__
#define __ICABasicDef_FWD_DEFINED__
typedef interface ICABasicDef ICABasicDef;

#endif 	/* __ICABasicDef_FWD_DEFINED__ */


#ifndef __ICABasicSym_FWD_DEFINED__
#define __ICABasicSym_FWD_DEFINED__
typedef interface ICABasicSym ICABasicSym;

#endif 	/* __ICABasicSym_FWD_DEFINED__ */


#ifndef __CADeleteMemberOutputDef_FWD_DEFINED__
#define __CADeleteMemberOutputDef_FWD_DEFINED__

#ifdef __cplusplus
typedef class CADeleteMemberOutputDef CADeleteMemberOutputDef;
#else
typedef struct CADeleteMemberOutputDef CADeleteMemberOutputDef;
#endif /* __cplusplus */

#endif 	/* __CADeleteMemberOutputDef_FWD_DEFINED__ */


#ifndef __CADeleteMemberOutputSym_FWD_DEFINED__
#define __CADeleteMemberOutputSym_FWD_DEFINED__

#ifdef __cplusplus
typedef class CADeleteMemberOutputSym CADeleteMemberOutputSym;
#else
typedef struct CADeleteMemberOutputSym CADeleteMemberOutputSym;
#endif /* __cplusplus */

#endif 	/* __CADeleteMemberOutputSym_FWD_DEFINED__ */


#ifndef __CAEvaluateAfterSymbolSym_FWD_DEFINED__
#define __CAEvaluateAfterSymbolSym_FWD_DEFINED__

#ifdef __cplusplus
typedef class CAEvaluateAfterSymbolSym CAEvaluateAfterSymbolSym;
#else
typedef struct CAEvaluateAfterSymbolSym CAEvaluateAfterSymbolSym;
#endif /* __cplusplus */

#endif 	/* __CAEvaluateAfterSymbolSym_FWD_DEFINED__ */


#ifndef __CAEvaluateAfterSymbolDef_FWD_DEFINED__
#define __CAEvaluateAfterSymbolDef_FWD_DEFINED__

#ifdef __cplusplus
typedef class CAEvaluateAfterSymbolDef CAEvaluateAfterSymbolDef;
#else
typedef struct CAEvaluateAfterSymbolDef CAEvaluateAfterSymbolDef;
#endif /* __cplusplus */

#endif 	/* __CAEvaluateAfterSymbolDef_FWD_DEFINED__ */


#ifndef __COMSelectionRuleSel_FWD_DEFINED__
#define __COMSelectionRuleSel_FWD_DEFINED__

#ifdef __cplusplus
typedef class COMSelectionRuleSel COMSelectionRuleSel;
#else
typedef struct COMSelectionRuleSel COMSelectionRuleSel;
#endif /* __cplusplus */

#endif 	/* __COMSelectionRuleSel_FWD_DEFINED__ */


#ifndef __COMSelectionRuleRootSel_FWD_DEFINED__
#define __COMSelectionRuleRootSel_FWD_DEFINED__

#ifdef __cplusplus
typedef class COMSelectionRuleRootSel COMSelectionRuleRootSel;
#else
typedef struct COMSelectionRuleRootSel COMSelectionRuleRootSel;
#endif /* __cplusplus */

#endif 	/* __COMSelectionRuleRootSel_FWD_DEFINED__ */


#ifndef __COM2ndSelectionRuleSel_FWD_DEFINED__
#define __COM2ndSelectionRuleSel_FWD_DEFINED__

#ifdef __cplusplus
typedef class COM2ndSelectionRuleSel COM2ndSelectionRuleSel;
#else
typedef struct COM2ndSelectionRuleSel COM2ndSelectionRuleSel;
#endif /* __cplusplus */

#endif 	/* __COM2ndSelectionRuleSel_FWD_DEFINED__ */


#ifndef __TestSelRuleWithQs_FWD_DEFINED__
#define __TestSelRuleWithQs_FWD_DEFINED__

#ifdef __cplusplus
typedef class TestSelRuleWithQs TestSelRuleWithQs;
#else
typedef struct TestSelRuleWithQs TestSelRuleWithQs;
#endif /* __cplusplus */

#endif 	/* __TestSelRuleWithQs_FWD_DEFINED__ */


#ifndef __COMTestParameterRule_FWD_DEFINED__
#define __COMTestParameterRule_FWD_DEFINED__

#ifdef __cplusplus
typedef class COMTestParameterRule COMTestParameterRule;
#else
typedef struct COMTestParameterRule COMTestParameterRule;
#endif /* __cplusplus */

#endif 	/* __COMTestParameterRule_FWD_DEFINED__ */


#ifndef __CAForParameterRuleDef_FWD_DEFINED__
#define __CAForParameterRuleDef_FWD_DEFINED__

#ifdef __cplusplus
typedef class CAForParameterRuleDef CAForParameterRuleDef;
#else
typedef struct CAForParameterRuleDef CAForParameterRuleDef;
#endif /* __cplusplus */

#endif 	/* __CAForParameterRuleDef_FWD_DEFINED__ */


#ifndef __CAForParameterRuleSym_FWD_DEFINED__
#define __CAForParameterRuleSym_FWD_DEFINED__

#ifdef __cplusplus
typedef class CAForParameterRuleSym CAForParameterRuleSym;
#else
typedef struct CAForParameterRuleSym CAForParameterRuleSym;
#endif /* __cplusplus */

#endif 	/* __CAForParameterRuleSym_FWD_DEFINED__ */


#ifndef __SONoGraphics_FWD_DEFINED__
#define __SONoGraphics_FWD_DEFINED__

#ifdef __cplusplus
typedef class SONoGraphics SONoGraphics;
#else
typedef struct SONoGraphics SONoGraphics;
#endif /* __cplusplus */

#endif 	/* __SONoGraphics_FWD_DEFINED__ */


#ifndef __SO1GraphicOutput_FWD_DEFINED__
#define __SO1GraphicOutput_FWD_DEFINED__

#ifdef __cplusplus
typedef class SO1GraphicOutput SO1GraphicOutput;
#else
typedef struct SO1GraphicOutput SO1GraphicOutput;
#endif /* __cplusplus */

#endif 	/* __SO1GraphicOutput_FWD_DEFINED__ */


#ifndef __CANestedOutputParentDef_FWD_DEFINED__
#define __CANestedOutputParentDef_FWD_DEFINED__

#ifdef __cplusplus
typedef class CANestedOutputParentDef CANestedOutputParentDef;
#else
typedef struct CANestedOutputParentDef CANestedOutputParentDef;
#endif /* __cplusplus */

#endif 	/* __CANestedOutputParentDef_FWD_DEFINED__ */


#ifndef __CANestedOutputParentSym_FWD_DEFINED__
#define __CANestedOutputParentSym_FWD_DEFINED__

#ifdef __cplusplus
typedef class CANestedOutputParentSym CANestedOutputParentSym;
#else
typedef struct CANestedOutputParentSym CANestedOutputParentSym;
#endif /* __cplusplus */

#endif 	/* __CANestedOutputParentSym_FWD_DEFINED__ */


#ifndef __CABasicDef_FWD_DEFINED__
#define __CABasicDef_FWD_DEFINED__

#ifdef __cplusplus
typedef class CABasicDef CABasicDef;
#else
typedef struct CABasicDef CABasicDef;
#endif /* __cplusplus */

#endif 	/* __CABasicDef_FWD_DEFINED__ */


#ifndef __CABasicSym_FWD_DEFINED__
#define __CABasicSym_FWD_DEFINED__

#ifdef __cplusplus
typedef class CABasicSym CABasicSym;
#else
typedef struct CABasicSym CABasicSym;
#endif /* __cplusplus */

#endif 	/* __CABasicSym_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "CustomAssembly.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __ICADeleteMemberOutputDef_INTERFACE_DEFINED__
#define __ICADeleteMemberOutputDef_INTERFACE_DEFINED__

/* interface ICADeleteMemberOutputDef */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICADeleteMemberOutputDef;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D6E7D86E-1140-4A65-AD1D-9C545161B2D0")
    ICADeleteMemberOutputDef : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMFinalConstructAsm( 
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMConstructAsm( 
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateCAO( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMConstructSphere( 
            /* [in] */ IJDMemberDescription *pMemberDescription,
            /* [in] */ IUnknown *pResourceManager,
            /* [out][in] */ IDispatch **pObj) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMSetInputSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMFinalConstructSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMReleaseSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateSphereProperties( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateSphereGeometry( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICADeleteMemberOutputDefVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICADeleteMemberOutputDef * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICADeleteMemberOutputDef * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICADeleteMemberOutputDef * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICADeleteMemberOutputDef * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICADeleteMemberOutputDef * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICADeleteMemberOutputDef * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICADeleteMemberOutputDef * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMFinalConstructAsm )( 
            ICADeleteMemberOutputDef * This,
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMConstructAsm )( 
            ICADeleteMemberOutputDef * This,
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateCAO )( 
            ICADeleteMemberOutputDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMConstructSphere )( 
            ICADeleteMemberOutputDef * This,
            /* [in] */ IJDMemberDescription *pMemberDescription,
            /* [in] */ IUnknown *pResourceManager,
            /* [out][in] */ IDispatch **pObj);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMSetInputSphere )( 
            ICADeleteMemberOutputDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMFinalConstructSphere )( 
            ICADeleteMemberOutputDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMReleaseSphere )( 
            ICADeleteMemberOutputDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateSphereProperties )( 
            ICADeleteMemberOutputDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateSphereGeometry )( 
            ICADeleteMemberOutputDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        END_INTERFACE
    } ICADeleteMemberOutputDefVtbl;

    interface ICADeleteMemberOutputDef
    {
        CONST_VTBL struct ICADeleteMemberOutputDefVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICADeleteMemberOutputDef_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICADeleteMemberOutputDef_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICADeleteMemberOutputDef_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICADeleteMemberOutputDef_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICADeleteMemberOutputDef_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICADeleteMemberOutputDef_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICADeleteMemberOutputDef_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICADeleteMemberOutputDef_CMFinalConstructAsm(This,pAggregatorDescription)	\
    ( (This)->lpVtbl -> CMFinalConstructAsm(This,pAggregatorDescription) ) 

#define ICADeleteMemberOutputDef_CMConstructAsm(This,pAggregatorDescription)	\
    ( (This)->lpVtbl -> CMConstructAsm(This,pAggregatorDescription) ) 

#define ICADeleteMemberOutputDef_CMEvaluateCAO(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateCAO(This,pPropertyDescriptions,pObject) ) 

#define ICADeleteMemberOutputDef_CMConstructSphere(This,pMemberDescription,pResourceManager,pObj)	\
    ( (This)->lpVtbl -> CMConstructSphere(This,pMemberDescription,pResourceManager,pObj) ) 

#define ICADeleteMemberOutputDef_CMSetInputSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMSetInputSphere(This,pMemberDesc) ) 

#define ICADeleteMemberOutputDef_CMFinalConstructSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMFinalConstructSphere(This,pMemberDesc) ) 

#define ICADeleteMemberOutputDef_CMReleaseSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMReleaseSphere(This,pMemberDesc) ) 

#define ICADeleteMemberOutputDef_CMEvaluateSphereProperties(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateSphereProperties(This,pPropertyDescriptions,pObject) ) 

#define ICADeleteMemberOutputDef_CMEvaluateSphereGeometry(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateSphereGeometry(This,pPropertyDescriptions,pObject) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICADeleteMemberOutputDef_INTERFACE_DEFINED__ */


#ifndef __ICADeleteMemberOutputSym_INTERFACE_DEFINED__
#define __ICADeleteMemberOutputSym_INTERFACE_DEFINED__

/* interface ICADeleteMemberOutputSym */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICADeleteMemberOutputSym;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("18A2C589-D381-46DC-A7C9-9A7DF6A5AA63")
    ICADeleteMemberOutputSym : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Physical( 
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICADeleteMemberOutputSymVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICADeleteMemberOutputSym * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICADeleteMemberOutputSym * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICADeleteMemberOutputSym * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICADeleteMemberOutputSym * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICADeleteMemberOutputSym * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICADeleteMemberOutputSym * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICADeleteMemberOutputSym * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Physical )( 
            ICADeleteMemberOutputSym * This,
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM);
        
        END_INTERFACE
    } ICADeleteMemberOutputSymVtbl;

    interface ICADeleteMemberOutputSym
    {
        CONST_VTBL struct ICADeleteMemberOutputSymVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICADeleteMemberOutputSym_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICADeleteMemberOutputSym_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICADeleteMemberOutputSym_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICADeleteMemberOutputSym_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICADeleteMemberOutputSym_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICADeleteMemberOutputSym_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICADeleteMemberOutputSym_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICADeleteMemberOutputSym_Physical(This,pIRepSCM)	\
    ( (This)->lpVtbl -> Physical(This,pIRepSCM) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICADeleteMemberOutputSym_INTERFACE_DEFINED__ */


#ifndef __ICAEvaluateAfterSymbolSym_INTERFACE_DEFINED__
#define __ICAEvaluateAfterSymbolSym_INTERFACE_DEFINED__

/* interface ICAEvaluateAfterSymbolSym */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICAEvaluateAfterSymbolSym;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("6AE5114E-7A57-4885-BB71-99EB4AABA5EC")
    ICAEvaluateAfterSymbolSym : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Physical( 
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICAEvaluateAfterSymbolSymVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICAEvaluateAfterSymbolSym * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICAEvaluateAfterSymbolSym * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICAEvaluateAfterSymbolSym * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICAEvaluateAfterSymbolSym * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICAEvaluateAfterSymbolSym * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICAEvaluateAfterSymbolSym * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICAEvaluateAfterSymbolSym * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Physical )( 
            ICAEvaluateAfterSymbolSym * This,
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM);
        
        END_INTERFACE
    } ICAEvaluateAfterSymbolSymVtbl;

    interface ICAEvaluateAfterSymbolSym
    {
        CONST_VTBL struct ICAEvaluateAfterSymbolSymVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICAEvaluateAfterSymbolSym_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICAEvaluateAfterSymbolSym_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICAEvaluateAfterSymbolSym_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICAEvaluateAfterSymbolSym_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICAEvaluateAfterSymbolSym_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICAEvaluateAfterSymbolSym_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICAEvaluateAfterSymbolSym_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICAEvaluateAfterSymbolSym_Physical(This,pIRepSCM)	\
    ( (This)->lpVtbl -> Physical(This,pIRepSCM) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICAEvaluateAfterSymbolSym_INTERFACE_DEFINED__ */


#ifndef __ICAEvaluateAfterSymbolDef_INTERFACE_DEFINED__
#define __ICAEvaluateAfterSymbolDef_INTERFACE_DEFINED__

/* interface ICAEvaluateAfterSymbolDef */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICAEvaluateAfterSymbolDef;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("0A87A52A-41A6-4A6B-B331-DD964C2ACF60")
    ICAEvaluateAfterSymbolDef : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMFinalConstructAsm( 
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMConstructAsm( 
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateCAOBefore( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateCAOAfter( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMConstructSphere( 
            /* [in] */ IJDMemberDescription *pMemberDescription,
            /* [in] */ IUnknown *pResourceManager,
            /* [out][in] */ IDispatch **pObj) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMSetInputSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMFinalConstructSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMReleaseSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateSphereProperties( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateSphereGeometry( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICAEvaluateAfterSymbolDefVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICAEvaluateAfterSymbolDef * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICAEvaluateAfterSymbolDef * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMFinalConstructAsm )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMConstructAsm )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateCAOBefore )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateCAOAfter )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMConstructSphere )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [in] */ IJDMemberDescription *pMemberDescription,
            /* [in] */ IUnknown *pResourceManager,
            /* [out][in] */ IDispatch **pObj);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMSetInputSphere )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMFinalConstructSphere )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMReleaseSphere )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateSphereProperties )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateSphereGeometry )( 
            ICAEvaluateAfterSymbolDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        END_INTERFACE
    } ICAEvaluateAfterSymbolDefVtbl;

    interface ICAEvaluateAfterSymbolDef
    {
        CONST_VTBL struct ICAEvaluateAfterSymbolDefVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICAEvaluateAfterSymbolDef_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICAEvaluateAfterSymbolDef_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICAEvaluateAfterSymbolDef_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICAEvaluateAfterSymbolDef_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICAEvaluateAfterSymbolDef_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICAEvaluateAfterSymbolDef_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICAEvaluateAfterSymbolDef_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICAEvaluateAfterSymbolDef_CMFinalConstructAsm(This,pAggregatorDescription)	\
    ( (This)->lpVtbl -> CMFinalConstructAsm(This,pAggregatorDescription) ) 

#define ICAEvaluateAfterSymbolDef_CMConstructAsm(This,pAggregatorDescription)	\
    ( (This)->lpVtbl -> CMConstructAsm(This,pAggregatorDescription) ) 

#define ICAEvaluateAfterSymbolDef_CMEvaluateCAOBefore(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateCAOBefore(This,pPropertyDescriptions,pObject) ) 

#define ICAEvaluateAfterSymbolDef_CMEvaluateCAOAfter(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateCAOAfter(This,pPropertyDescriptions,pObject) ) 

#define ICAEvaluateAfterSymbolDef_CMConstructSphere(This,pMemberDescription,pResourceManager,pObj)	\
    ( (This)->lpVtbl -> CMConstructSphere(This,pMemberDescription,pResourceManager,pObj) ) 

#define ICAEvaluateAfterSymbolDef_CMSetInputSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMSetInputSphere(This,pMemberDesc) ) 

#define ICAEvaluateAfterSymbolDef_CMFinalConstructSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMFinalConstructSphere(This,pMemberDesc) ) 

#define ICAEvaluateAfterSymbolDef_CMReleaseSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMReleaseSphere(This,pMemberDesc) ) 

#define ICAEvaluateAfterSymbolDef_CMEvaluateSphereProperties(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateSphereProperties(This,pPropertyDescriptions,pObject) ) 

#define ICAEvaluateAfterSymbolDef_CMEvaluateSphereGeometry(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateSphereGeometry(This,pPropertyDescriptions,pObject) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICAEvaluateAfterSymbolDef_INTERFACE_DEFINED__ */


#ifndef __ICOMSelectionRuleSel_INTERFACE_DEFINED__
#define __ICOMSelectionRuleSel_INTERFACE_DEFINED__

/* interface ICOMSelectionRuleSel */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICOMSelectionRuleSel;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("1ECEAC32-4154-4C1C-8535-55F5847FBD2A")
    ICOMSelectionRuleSel : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMSelector( 
            /* [in] */ IJDRepresentation *pRepresentation) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICOMSelectionRuleSelVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICOMSelectionRuleSel * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICOMSelectionRuleSel * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICOMSelectionRuleSel * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICOMSelectionRuleSel * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICOMSelectionRuleSel * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICOMSelectionRuleSel * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICOMSelectionRuleSel * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMSelector )( 
            ICOMSelectionRuleSel * This,
            /* [in] */ IJDRepresentation *pRepresentation);
        
        END_INTERFACE
    } ICOMSelectionRuleSelVtbl;

    interface ICOMSelectionRuleSel
    {
        CONST_VTBL struct ICOMSelectionRuleSelVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICOMSelectionRuleSel_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICOMSelectionRuleSel_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICOMSelectionRuleSel_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICOMSelectionRuleSel_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICOMSelectionRuleSel_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICOMSelectionRuleSel_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICOMSelectionRuleSel_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICOMSelectionRuleSel_CMSelector(This,pRepresentation)	\
    ( (This)->lpVtbl -> CMSelector(This,pRepresentation) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICOMSelectionRuleSel_INTERFACE_DEFINED__ */


#ifndef __ICOMSelectionRuleRootSel_INTERFACE_DEFINED__
#define __ICOMSelectionRuleRootSel_INTERFACE_DEFINED__

/* interface ICOMSelectionRuleRootSel */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICOMSelectionRuleRootSel;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("EF9D05C0-893D-4084-8F6F-0B1CBFBF750F")
    ICOMSelectionRuleRootSel : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMSelector( 
            /* [in] */ IJDRepresentation *pRepresentation) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICOMSelectionRuleRootSelVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICOMSelectionRuleRootSel * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICOMSelectionRuleRootSel * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICOMSelectionRuleRootSel * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICOMSelectionRuleRootSel * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICOMSelectionRuleRootSel * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICOMSelectionRuleRootSel * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICOMSelectionRuleRootSel * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMSelector )( 
            ICOMSelectionRuleRootSel * This,
            /* [in] */ IJDRepresentation *pRepresentation);
        
        END_INTERFACE
    } ICOMSelectionRuleRootSelVtbl;

    interface ICOMSelectionRuleRootSel
    {
        CONST_VTBL struct ICOMSelectionRuleRootSelVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICOMSelectionRuleRootSel_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICOMSelectionRuleRootSel_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICOMSelectionRuleRootSel_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICOMSelectionRuleRootSel_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICOMSelectionRuleRootSel_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICOMSelectionRuleRootSel_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICOMSelectionRuleRootSel_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICOMSelectionRuleRootSel_CMSelector(This,pRepresentation)	\
    ( (This)->lpVtbl -> CMSelector(This,pRepresentation) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICOMSelectionRuleRootSel_INTERFACE_DEFINED__ */


#ifndef __ICOM2ndSelectionRuleSel_INTERFACE_DEFINED__
#define __ICOM2ndSelectionRuleSel_INTERFACE_DEFINED__

/* interface ICOM2ndSelectionRuleSel */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICOM2ndSelectionRuleSel;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("270F5D5A-6F54-497D-8DA1-212B909292A5")
    ICOM2ndSelectionRuleSel : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMSelector( 
            /* [in] */ IJDRepresentation *pRepresentation) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICOM2ndSelectionRuleSelVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICOM2ndSelectionRuleSel * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICOM2ndSelectionRuleSel * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICOM2ndSelectionRuleSel * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICOM2ndSelectionRuleSel * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICOM2ndSelectionRuleSel * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICOM2ndSelectionRuleSel * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICOM2ndSelectionRuleSel * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMSelector )( 
            ICOM2ndSelectionRuleSel * This,
            /* [in] */ IJDRepresentation *pRepresentation);
        
        END_INTERFACE
    } ICOM2ndSelectionRuleSelVtbl;

    interface ICOM2ndSelectionRuleSel
    {
        CONST_VTBL struct ICOM2ndSelectionRuleSelVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICOM2ndSelectionRuleSel_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICOM2ndSelectionRuleSel_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICOM2ndSelectionRuleSel_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICOM2ndSelectionRuleSel_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICOM2ndSelectionRuleSel_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICOM2ndSelectionRuleSel_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICOM2ndSelectionRuleSel_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICOM2ndSelectionRuleSel_CMSelector(This,pRepresentation)	\
    ( (This)->lpVtbl -> CMSelector(This,pRepresentation) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICOM2ndSelectionRuleSel_INTERFACE_DEFINED__ */


#ifndef __ITestSelRuleWithQs_INTERFACE_DEFINED__
#define __ITestSelRuleWithQs_INTERFACE_DEFINED__

/* interface ITestSelRuleWithQs */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ITestSelRuleWithQs;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("750BBA63-407F-42BF-AA8E-66763B6DB3BA")
    ITestSelRuleWithQs : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMSelector( 
            /* [in] */ IJDRepresentation *pRepresentation) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Question1MethodCM( 
            /* [in] */ IJDInputStdCustomMethod *pInput,
            /* [out][in] */ IDispatch **ppArgument) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Question2MethodCM( 
            /* [in] */ IJDInputStdCustomMethod *pInput,
            /* [out][in] */ IDispatch **ppArgument) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ITestSelRuleWithQsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITestSelRuleWithQs * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITestSelRuleWithQs * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITestSelRuleWithQs * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITestSelRuleWithQs * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITestSelRuleWithQs * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITestSelRuleWithQs * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITestSelRuleWithQs * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMSelector )( 
            ITestSelRuleWithQs * This,
            /* [in] */ IJDRepresentation *pRepresentation);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Question1MethodCM )( 
            ITestSelRuleWithQs * This,
            /* [in] */ IJDInputStdCustomMethod *pInput,
            /* [out][in] */ IDispatch **ppArgument);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Question2MethodCM )( 
            ITestSelRuleWithQs * This,
            /* [in] */ IJDInputStdCustomMethod *pInput,
            /* [out][in] */ IDispatch **ppArgument);
        
        END_INTERFACE
    } ITestSelRuleWithQsVtbl;

    interface ITestSelRuleWithQs
    {
        CONST_VTBL struct ITestSelRuleWithQsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITestSelRuleWithQs_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITestSelRuleWithQs_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITestSelRuleWithQs_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITestSelRuleWithQs_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITestSelRuleWithQs_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITestSelRuleWithQs_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITestSelRuleWithQs_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITestSelRuleWithQs_CMSelector(This,pRepresentation)	\
    ( (This)->lpVtbl -> CMSelector(This,pRepresentation) ) 

#define ITestSelRuleWithQs_Question1MethodCM(This,pInput,ppArgument)	\
    ( (This)->lpVtbl -> Question1MethodCM(This,pInput,ppArgument) ) 

#define ITestSelRuleWithQs_Question2MethodCM(This,pInput,ppArgument)	\
    ( (This)->lpVtbl -> Question2MethodCM(This,pInput,ppArgument) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITestSelRuleWithQs_INTERFACE_DEFINED__ */


#ifndef __ICOMTestParameterRule_INTERFACE_DEFINED__
#define __ICOMTestParameterRule_INTERFACE_DEFINED__

/* interface ICOMTestParameterRule */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICOMTestParameterRule;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("8CC1C108-0BEB-4CEB-9208-7FBD660DF09F")
    ICOMTestParameterRule : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMParameterRule( 
            /* [in] */ IJDRepresentation *pRepresentation) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICOMTestParameterRuleVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICOMTestParameterRule * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICOMTestParameterRule * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICOMTestParameterRule * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICOMTestParameterRule * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICOMTestParameterRule * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICOMTestParameterRule * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICOMTestParameterRule * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMParameterRule )( 
            ICOMTestParameterRule * This,
            /* [in] */ IJDRepresentation *pRepresentation);
        
        END_INTERFACE
    } ICOMTestParameterRuleVtbl;

    interface ICOMTestParameterRule
    {
        CONST_VTBL struct ICOMTestParameterRuleVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICOMTestParameterRule_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICOMTestParameterRule_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICOMTestParameterRule_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICOMTestParameterRule_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICOMTestParameterRule_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICOMTestParameterRule_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICOMTestParameterRule_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICOMTestParameterRule_CMParameterRule(This,pRepresentation)	\
    ( (This)->lpVtbl -> CMParameterRule(This,pRepresentation) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICOMTestParameterRule_INTERFACE_DEFINED__ */


#ifndef __ICAForParameterRuleDef_INTERFACE_DEFINED__
#define __ICAForParameterRuleDef_INTERFACE_DEFINED__

/* interface ICAForParameterRuleDef */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICAForParameterRuleDef;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("1C20CEB2-B64E-4946-A93D-6D8E4B1F2116")
    ICAForParameterRuleDef : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMFinalConstructAsm( 
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMConstructAsm( 
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateCAO( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMConstructSphere( 
            /* [in] */ IJDMemberDescription *pMemberDescription,
            /* [in] */ IUnknown *pResourceManager,
            /* [out][in] */ IDispatch **pObj) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMSetInputSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMFinalConstructSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMReleaseSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateSphereProperties( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateSphereGeometry( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICAForParameterRuleDefVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICAForParameterRuleDef * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICAForParameterRuleDef * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICAForParameterRuleDef * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICAForParameterRuleDef * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICAForParameterRuleDef * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICAForParameterRuleDef * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICAForParameterRuleDef * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMFinalConstructAsm )( 
            ICAForParameterRuleDef * This,
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMConstructAsm )( 
            ICAForParameterRuleDef * This,
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateCAO )( 
            ICAForParameterRuleDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMConstructSphere )( 
            ICAForParameterRuleDef * This,
            /* [in] */ IJDMemberDescription *pMemberDescription,
            /* [in] */ IUnknown *pResourceManager,
            /* [out][in] */ IDispatch **pObj);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMSetInputSphere )( 
            ICAForParameterRuleDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMFinalConstructSphere )( 
            ICAForParameterRuleDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMReleaseSphere )( 
            ICAForParameterRuleDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateSphereProperties )( 
            ICAForParameterRuleDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateSphereGeometry )( 
            ICAForParameterRuleDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        END_INTERFACE
    } ICAForParameterRuleDefVtbl;

    interface ICAForParameterRuleDef
    {
        CONST_VTBL struct ICAForParameterRuleDefVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICAForParameterRuleDef_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICAForParameterRuleDef_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICAForParameterRuleDef_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICAForParameterRuleDef_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICAForParameterRuleDef_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICAForParameterRuleDef_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICAForParameterRuleDef_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICAForParameterRuleDef_CMFinalConstructAsm(This,pAggregatorDescription)	\
    ( (This)->lpVtbl -> CMFinalConstructAsm(This,pAggregatorDescription) ) 

#define ICAForParameterRuleDef_CMConstructAsm(This,pAggregatorDescription)	\
    ( (This)->lpVtbl -> CMConstructAsm(This,pAggregatorDescription) ) 

#define ICAForParameterRuleDef_CMEvaluateCAO(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateCAO(This,pPropertyDescriptions,pObject) ) 

#define ICAForParameterRuleDef_CMConstructSphere(This,pMemberDescription,pResourceManager,pObj)	\
    ( (This)->lpVtbl -> CMConstructSphere(This,pMemberDescription,pResourceManager,pObj) ) 

#define ICAForParameterRuleDef_CMSetInputSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMSetInputSphere(This,pMemberDesc) ) 

#define ICAForParameterRuleDef_CMFinalConstructSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMFinalConstructSphere(This,pMemberDesc) ) 

#define ICAForParameterRuleDef_CMReleaseSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMReleaseSphere(This,pMemberDesc) ) 

#define ICAForParameterRuleDef_CMEvaluateSphereProperties(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateSphereProperties(This,pPropertyDescriptions,pObject) ) 

#define ICAForParameterRuleDef_CMEvaluateSphereGeometry(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateSphereGeometry(This,pPropertyDescriptions,pObject) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICAForParameterRuleDef_INTERFACE_DEFINED__ */


#ifndef __ICAForParameterRuleSym_INTERFACE_DEFINED__
#define __ICAForParameterRuleSym_INTERFACE_DEFINED__

/* interface ICAForParameterRuleSym */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICAForParameterRuleSym;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("D835FF24-0B30-4A9D-AA86-E5D867ADDD58")
    ICAForParameterRuleSym : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Physical( 
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICAForParameterRuleSymVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICAForParameterRuleSym * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICAForParameterRuleSym * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICAForParameterRuleSym * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICAForParameterRuleSym * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICAForParameterRuleSym * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICAForParameterRuleSym * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICAForParameterRuleSym * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Physical )( 
            ICAForParameterRuleSym * This,
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM);
        
        END_INTERFACE
    } ICAForParameterRuleSymVtbl;

    interface ICAForParameterRuleSym
    {
        CONST_VTBL struct ICAForParameterRuleSymVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICAForParameterRuleSym_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICAForParameterRuleSym_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICAForParameterRuleSym_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICAForParameterRuleSym_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICAForParameterRuleSym_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICAForParameterRuleSym_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICAForParameterRuleSym_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICAForParameterRuleSym_Physical(This,pIRepSCM)	\
    ( (This)->lpVtbl -> Physical(This,pIRepSCM) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICAForParameterRuleSym_INTERFACE_DEFINED__ */


#ifndef __ISONoGraphics_INTERFACE_DEFINED__
#define __ISONoGraphics_INTERFACE_DEFINED__

/* interface ISONoGraphics */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ISONoGraphics;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("273050D7-BEA4-4DFB-9B0C-411EF56CD9AC")
    ISONoGraphics : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Physical( 
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ISONoGraphicsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ISONoGraphics * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ISONoGraphics * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ISONoGraphics * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ISONoGraphics * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ISONoGraphics * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ISONoGraphics * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ISONoGraphics * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Physical )( 
            ISONoGraphics * This,
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM);
        
        END_INTERFACE
    } ISONoGraphicsVtbl;

    interface ISONoGraphics
    {
        CONST_VTBL struct ISONoGraphicsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISONoGraphics_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ISONoGraphics_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ISONoGraphics_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ISONoGraphics_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ISONoGraphics_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ISONoGraphics_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ISONoGraphics_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ISONoGraphics_Physical(This,pIRepSCM)	\
    ( (This)->lpVtbl -> Physical(This,pIRepSCM) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ISONoGraphics_INTERFACE_DEFINED__ */


#ifndef __ISO1GraphicOutput_INTERFACE_DEFINED__
#define __ISO1GraphicOutput_INTERFACE_DEFINED__

/* interface ISO1GraphicOutput */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ISO1GraphicOutput;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("16F840AA-9135-42B1-A4E1-3832AEF010E1")
    ISO1GraphicOutput : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Physical( 
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ISO1GraphicOutputVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ISO1GraphicOutput * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ISO1GraphicOutput * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ISO1GraphicOutput * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ISO1GraphicOutput * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ISO1GraphicOutput * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ISO1GraphicOutput * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ISO1GraphicOutput * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Physical )( 
            ISO1GraphicOutput * This,
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM);
        
        END_INTERFACE
    } ISO1GraphicOutputVtbl;

    interface ISO1GraphicOutput
    {
        CONST_VTBL struct ISO1GraphicOutputVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ISO1GraphicOutput_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ISO1GraphicOutput_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ISO1GraphicOutput_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ISO1GraphicOutput_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ISO1GraphicOutput_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ISO1GraphicOutput_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ISO1GraphicOutput_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ISO1GraphicOutput_Physical(This,pIRepSCM)	\
    ( (This)->lpVtbl -> Physical(This,pIRepSCM) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ISO1GraphicOutput_INTERFACE_DEFINED__ */


#ifndef __ICANestedOutputParentSym_INTERFACE_DEFINED__
#define __ICANestedOutputParentSym_INTERFACE_DEFINED__

/* interface ICANestedOutputParentSym */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICANestedOutputParentSym;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("266C38F3-421F-4491-9923-75CA40458D4F")
    ICANestedOutputParentSym : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Physical( 
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICANestedOutputParentSymVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICANestedOutputParentSym * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICANestedOutputParentSym * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICANestedOutputParentSym * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICANestedOutputParentSym * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICANestedOutputParentSym * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICANestedOutputParentSym * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICANestedOutputParentSym * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Physical )( 
            ICANestedOutputParentSym * This,
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM);
        
        END_INTERFACE
    } ICANestedOutputParentSymVtbl;

    interface ICANestedOutputParentSym
    {
        CONST_VTBL struct ICANestedOutputParentSymVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICANestedOutputParentSym_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICANestedOutputParentSym_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICANestedOutputParentSym_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICANestedOutputParentSym_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICANestedOutputParentSym_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICANestedOutputParentSym_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICANestedOutputParentSym_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICANestedOutputParentSym_Physical(This,pIRepSCM)	\
    ( (This)->lpVtbl -> Physical(This,pIRepSCM) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICANestedOutputParentSym_INTERFACE_DEFINED__ */


#ifndef __ICANestedOutputParentDef_INTERFACE_DEFINED__
#define __ICANestedOutputParentDef_INTERFACE_DEFINED__

/* interface ICANestedOutputParentDef */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICANestedOutputParentDef;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("B00BD839-FAF5-466A-8CF9-827138DAA42F")
    ICANestedOutputParentDef : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMFinalConstructAsm( 
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMConstructAsm( 
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateCAO( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMConstructNestedOutput( 
            /* [in] */ IJDMemberDescription *pMemberDescription,
            /* [in] */ IUnknown *pResourceManager,
            /* [out][in] */ IDispatch **pObj) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMSetInputNestedOutput( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMFinalConstructNestedOutput( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMReleaseNestedOutput( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateNestedOutputProperties( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateNestedOutputGeometry( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICANestedOutputParentDefVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICANestedOutputParentDef * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICANestedOutputParentDef * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICANestedOutputParentDef * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICANestedOutputParentDef * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICANestedOutputParentDef * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICANestedOutputParentDef * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICANestedOutputParentDef * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMFinalConstructAsm )( 
            ICANestedOutputParentDef * This,
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMConstructAsm )( 
            ICANestedOutputParentDef * This,
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateCAO )( 
            ICANestedOutputParentDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMConstructNestedOutput )( 
            ICANestedOutputParentDef * This,
            /* [in] */ IJDMemberDescription *pMemberDescription,
            /* [in] */ IUnknown *pResourceManager,
            /* [out][in] */ IDispatch **pObj);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMSetInputNestedOutput )( 
            ICANestedOutputParentDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMFinalConstructNestedOutput )( 
            ICANestedOutputParentDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMReleaseNestedOutput )( 
            ICANestedOutputParentDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateNestedOutputProperties )( 
            ICANestedOutputParentDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateNestedOutputGeometry )( 
            ICANestedOutputParentDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        END_INTERFACE
    } ICANestedOutputParentDefVtbl;

    interface ICANestedOutputParentDef
    {
        CONST_VTBL struct ICANestedOutputParentDefVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICANestedOutputParentDef_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICANestedOutputParentDef_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICANestedOutputParentDef_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICANestedOutputParentDef_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICANestedOutputParentDef_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICANestedOutputParentDef_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICANestedOutputParentDef_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICANestedOutputParentDef_CMFinalConstructAsm(This,pAggregatorDescription)	\
    ( (This)->lpVtbl -> CMFinalConstructAsm(This,pAggregatorDescription) ) 

#define ICANestedOutputParentDef_CMConstructAsm(This,pAggregatorDescription)	\
    ( (This)->lpVtbl -> CMConstructAsm(This,pAggregatorDescription) ) 

#define ICANestedOutputParentDef_CMEvaluateCAO(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateCAO(This,pPropertyDescriptions,pObject) ) 

#define ICANestedOutputParentDef_CMConstructNestedOutput(This,pMemberDescription,pResourceManager,pObj)	\
    ( (This)->lpVtbl -> CMConstructNestedOutput(This,pMemberDescription,pResourceManager,pObj) ) 

#define ICANestedOutputParentDef_CMSetInputNestedOutput(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMSetInputNestedOutput(This,pMemberDesc) ) 

#define ICANestedOutputParentDef_CMFinalConstructNestedOutput(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMFinalConstructNestedOutput(This,pMemberDesc) ) 

#define ICANestedOutputParentDef_CMReleaseNestedOutput(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMReleaseNestedOutput(This,pMemberDesc) ) 

#define ICANestedOutputParentDef_CMEvaluateNestedOutputProperties(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateNestedOutputProperties(This,pPropertyDescriptions,pObject) ) 

#define ICANestedOutputParentDef_CMEvaluateNestedOutputGeometry(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateNestedOutputGeometry(This,pPropertyDescriptions,pObject) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICANestedOutputParentDef_INTERFACE_DEFINED__ */


#ifndef __ICABasicDef_INTERFACE_DEFINED__
#define __ICABasicDef_INTERFACE_DEFINED__

/* interface ICABasicDef */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICABasicDef;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("F0AFC39A-F4C9-408D-8FF5-7E2C922CCB0E")
    ICABasicDef : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMFinalConstructAsm( 
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMConstructAsm( 
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateCAO( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMConstructSphere( 
            /* [in] */ IJDMemberDescription *pMemberDescription,
            /* [in] */ IUnknown *pResourceManager,
            /* [out][in] */ IDispatch **pObj) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMSetInputSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMFinalConstructSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMReleaseSphere( 
            /* [out][in] */ IJDMemberDescription **pMemberDesc) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateSphereProperties( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE CMEvaluateSphereGeometry( 
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICABasicDefVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICABasicDef * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICABasicDef * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICABasicDef * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICABasicDef * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICABasicDef * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICABasicDef * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICABasicDef * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMFinalConstructAsm )( 
            ICABasicDef * This,
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMConstructAsm )( 
            ICABasicDef * This,
            /* [out][in] */ IJDAggregatorDescription **pAggregatorDescription);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateCAO )( 
            ICABasicDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMConstructSphere )( 
            ICABasicDef * This,
            /* [in] */ IJDMemberDescription *pMemberDescription,
            /* [in] */ IUnknown *pResourceManager,
            /* [out][in] */ IDispatch **pObj);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMSetInputSphere )( 
            ICABasicDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMFinalConstructSphere )( 
            ICABasicDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMReleaseSphere )( 
            ICABasicDef * This,
            /* [out][in] */ IJDMemberDescription **pMemberDesc);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateSphereProperties )( 
            ICABasicDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *CMEvaluateSphereGeometry )( 
            ICABasicDef * This,
            /* [out][in] */ IJDPropertyDescription **pPropertyDescriptions,
            /* [out][in] */ IDispatch **pObject);
        
        END_INTERFACE
    } ICABasicDefVtbl;

    interface ICABasicDef
    {
        CONST_VTBL struct ICABasicDefVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICABasicDef_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICABasicDef_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICABasicDef_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICABasicDef_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICABasicDef_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICABasicDef_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICABasicDef_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICABasicDef_CMFinalConstructAsm(This,pAggregatorDescription)	\
    ( (This)->lpVtbl -> CMFinalConstructAsm(This,pAggregatorDescription) ) 

#define ICABasicDef_CMConstructAsm(This,pAggregatorDescription)	\
    ( (This)->lpVtbl -> CMConstructAsm(This,pAggregatorDescription) ) 

#define ICABasicDef_CMEvaluateCAO(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateCAO(This,pPropertyDescriptions,pObject) ) 

#define ICABasicDef_CMConstructSphere(This,pMemberDescription,pResourceManager,pObj)	\
    ( (This)->lpVtbl -> CMConstructSphere(This,pMemberDescription,pResourceManager,pObj) ) 

#define ICABasicDef_CMSetInputSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMSetInputSphere(This,pMemberDesc) ) 

#define ICABasicDef_CMFinalConstructSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMFinalConstructSphere(This,pMemberDesc) ) 

#define ICABasicDef_CMReleaseSphere(This,pMemberDesc)	\
    ( (This)->lpVtbl -> CMReleaseSphere(This,pMemberDesc) ) 

#define ICABasicDef_CMEvaluateSphereProperties(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateSphereProperties(This,pPropertyDescriptions,pObject) ) 

#define ICABasicDef_CMEvaluateSphereGeometry(This,pPropertyDescriptions,pObject)	\
    ( (This)->lpVtbl -> CMEvaluateSphereGeometry(This,pPropertyDescriptions,pObject) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICABasicDef_INTERFACE_DEFINED__ */


#ifndef __ICABasicSym_INTERFACE_DEFINED__
#define __ICABasicSym_INTERFACE_DEFINED__

/* interface ICABasicSym */
/* [unique][helpstring][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICABasicSym;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("9944D0A9-D898-41D1-BD81-A0A26E725FDD")
    ICABasicSym : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Physical( 
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct ICABasicSymVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICABasicSym * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICABasicSym * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICABasicSym * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICABasicSym * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICABasicSym * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICABasicSym * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICABasicSym * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Physical )( 
            ICABasicSym * This,
            /* [out][in] */ IJDRepresentationStdCustomMethod **pIRepSCM);
        
        END_INTERFACE
    } ICABasicSymVtbl;

    interface ICABasicSym
    {
        CONST_VTBL struct ICABasicSymVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICABasicSym_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICABasicSym_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICABasicSym_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICABasicSym_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICABasicSym_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICABasicSym_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICABasicSym_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ICABasicSym_Physical(This,pIRepSCM)	\
    ( (This)->lpVtbl -> Physical(This,pIRepSCM) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICABasicSym_INTERFACE_DEFINED__ */



#ifndef __ATPSO_LIBRARY_DEFINED__
#define __ATPSO_LIBRARY_DEFINED__

/* library ATPSO */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_ATPSO;

EXTERN_C const CLSID CLSID_CADeleteMemberOutputDef;

#ifdef __cplusplus

class DECLSPEC_UUID("0F66A04C-4E3C-40F0-B40E-E9E1D51A1F95")
CADeleteMemberOutputDef;
#endif

EXTERN_C const CLSID CLSID_CADeleteMemberOutputSym;

#ifdef __cplusplus

class DECLSPEC_UUID("1C8970B0-92D3-4AAA-BF5A-72EFE4DA8CB3")
CADeleteMemberOutputSym;
#endif

EXTERN_C const CLSID CLSID_CAEvaluateAfterSymbolSym;

#ifdef __cplusplus

class DECLSPEC_UUID("9FD62BD4-B0FF-47D9-9203-46A09799824F")
CAEvaluateAfterSymbolSym;
#endif

EXTERN_C const CLSID CLSID_CAEvaluateAfterSymbolDef;

#ifdef __cplusplus

class DECLSPEC_UUID("7C04C448-F2A6-4E39-880C-26E8EE6ADD25")
CAEvaluateAfterSymbolDef;
#endif

EXTERN_C const CLSID CLSID_COMSelectionRuleSel;

#ifdef __cplusplus

class DECLSPEC_UUID("AD0F81C3-941F-4A4A-B020-B24AF3A0252F")
COMSelectionRuleSel;
#endif

EXTERN_C const CLSID CLSID_COMSelectionRuleRootSel;

#ifdef __cplusplus

class DECLSPEC_UUID("64D91A79-5CB7-441F-933D-9A0C2D564D09")
COMSelectionRuleRootSel;
#endif

EXTERN_C const CLSID CLSID_COM2ndSelectionRuleSel;

#ifdef __cplusplus

class DECLSPEC_UUID("71664078-A2CF-471B-9EA8-0C314CE2617D")
COM2ndSelectionRuleSel;
#endif

EXTERN_C const CLSID CLSID_TestSelRuleWithQs;

#ifdef __cplusplus

class DECLSPEC_UUID("6233ABA4-5635-4F58-B51D-5E4DF34F92B1")
TestSelRuleWithQs;
#endif

EXTERN_C const CLSID CLSID_COMTestParameterRule;

#ifdef __cplusplus

class DECLSPEC_UUID("47E60D53-F154-481D-85F4-9EEB88E2C162")
COMTestParameterRule;
#endif

EXTERN_C const CLSID CLSID_CAForParameterRuleDef;

#ifdef __cplusplus

class DECLSPEC_UUID("8072837C-2979-4124-BC12-88AB3892EF7D")
CAForParameterRuleDef;
#endif

EXTERN_C const CLSID CLSID_CAForParameterRuleSym;

#ifdef __cplusplus

class DECLSPEC_UUID("E75458B5-167F-46FE-BCF1-7227F8B848E3")
CAForParameterRuleSym;
#endif

EXTERN_C const CLSID CLSID_SONoGraphics;

#ifdef __cplusplus

class DECLSPEC_UUID("61661D07-CCD3-4787-9946-4A0971EBCCA3")
SONoGraphics;
#endif

EXTERN_C const CLSID CLSID_SO1GraphicOutput;

#ifdef __cplusplus

class DECLSPEC_UUID("11EBB6C3-11B4-4C26-AAC0-D6307F0A1539")
SO1GraphicOutput;
#endif

EXTERN_C const CLSID CLSID_CANestedOutputParentDef;

#ifdef __cplusplus

class DECLSPEC_UUID("96B0D8F3-BF28-4993-8504-808F1B2D1FA2")
CANestedOutputParentDef;
#endif

EXTERN_C const CLSID CLSID_CANestedOutputParentSym;

#ifdef __cplusplus

class DECLSPEC_UUID("61F36747-A466-42E6-BAFD-ED7C5FF77E0A")
CANestedOutputParentSym;
#endif

EXTERN_C const CLSID CLSID_CABasicDef;

#ifdef __cplusplus

class DECLSPEC_UUID("E81D4319-D466-41ED-AD82-72BDA654745C")
CABasicDef;
#endif

EXTERN_C const CLSID CLSID_CABasicSym;

#ifdef __cplusplus

class DECLSPEC_UUID("68432C7A-5E6C-4C3B-9BF8-7008CCF85635")
CABasicSym;
#endif
#endif /* __ATPSO_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


