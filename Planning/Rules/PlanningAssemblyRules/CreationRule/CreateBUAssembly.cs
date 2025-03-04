using Ingr.SP3D.Common.Middle.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ingr.SP3D.Content.Planning
{
    public class CreateBUAssembly : PlnAssemblyCreateRuleBase
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="creationData"></param>
        public override void Evaluate(CreationData creationData)
        {
            #region Input Error Handling

            if (creationData == null)
            {
                throw new ArgumentNullException("creationData Information is null");
            }

            #endregion //Input Error Handling

            try
            {
                CreateBUMemberAssemblies(creationData.Blocks, creationData.AssemblyCreateOption);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }
    }
}
