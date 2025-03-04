// dllmain.h : Declaration of module class.

class CATPSOModule : public CAtlDllModuleT< CATPSOModule >
{
public :
	DECLARE_LIBID(LIBID_ATPSO)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_ATPSO, "{25D78ECB-22F6-44E6-BD37-5AA6818067DE}")
};

extern class CATPSOModule _AtlModule;
