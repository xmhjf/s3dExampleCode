using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class OglaendPartLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, OglaendPartSymbolResourceIDs.DEFAULT_RESOURCE, OglaendPartSymbolResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
