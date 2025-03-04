using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class OglaendAssembliesLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, OglaendAssembliesResourceIDs.DEFAULT_RESOURCE, OglaendAssembliesResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
