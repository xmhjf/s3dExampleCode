using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class BracketSymbolsLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, BracketSymbolsResourceIDs.DEFAULT_RESOURCE, BracketSymbolsResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
