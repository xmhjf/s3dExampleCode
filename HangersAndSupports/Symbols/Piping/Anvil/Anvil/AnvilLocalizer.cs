using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class AnvilLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, AnvilSymbolResourceIDs.DEFAULT_RESOURCE, AnvilSymbolResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
