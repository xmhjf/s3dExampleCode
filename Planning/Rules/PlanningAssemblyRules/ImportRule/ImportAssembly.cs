using Ingr.SP3D.Common.Middle.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ingr.SP3D.Content.Planning
{
    public class ImportAssembly : PlnAssemblyImportRuleBase
    {
        /// <summary>
        /// Imports the assembly data
        /// </summary>
        /// <param name="importData"></param>
        /// <returns>void</returns>
        public override void Evaluate(ImportData importData)
        {
            #region Input Error Handling

            if (importData == null)
            {
                throw new ArgumentNullException("importData Information is null");
            }
            #endregion //Input Error Handling

            try
            {
                ImportAssemblyData(importData);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }
    }
}
