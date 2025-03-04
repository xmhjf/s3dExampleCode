using Ingr.SP3D.Common.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    class GeneralProfilesLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, GeneralProfilesSymbolsResourceIDs.DEFAULT_RESOURCE, GeneralProfilesSymbolsResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
