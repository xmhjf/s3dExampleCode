using Ingr.SP3D.Common.Middle.Services;
namespace Ingr.SP3D.Content.Support.Symbols
{
    public class PipeHgrAssemblySymbolsLocalizer
    {
        public static string GetString(int iID, string defMsgStr)
        {
            return CmnLocalizer.GetString(iID, defMsgStr, PipeHgrAssemblySymbolsResourceIDs.DEFAULT_RESOURCE, PipeHgrAssemblySymbolsResourceIDs.DEFAULT_ASSEMBLY);
        }
    }
}
