//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Default class 
//                
//                 
//                    
//
//      History:
//      Sept 25th, 2014   Created by VIBGYOR
//
//-----------------------------------------------------------------------------------


using System;
using System.Collections.Generic;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using System.Runtime.InteropServices;
using Ingr.SP3D.Planning.Middle;
//using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Planning
{

    /// <summary>
    /// Create Assembly rule class Implements Planning Assembly Create Rule Base Class
    /// </summary>
    public class CreateAssembly : PlnAssemblyCreateRuleBase
    {
        #region Methods

        /// <summary>
        /// Creates assemblies automatically based on input data that contains block collection
        /// </summary>
        /// <param name="creationData"></param>
        /// <returns>void<ComplexString3d></returns>
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
                CreateAssemblies(creationData);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        # endregion
    }

   
}
