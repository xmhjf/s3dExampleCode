using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class PipeHgrAssemblies2Localizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, PipeHgrAssemblies2ResourceIDs.DEFAULT_RESOURCE, PipeHgrAssemblies2ResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
