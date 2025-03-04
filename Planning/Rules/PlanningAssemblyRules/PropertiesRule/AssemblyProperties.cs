//----------------------------------------------------------------------------------
//      Copyright (C) 2016 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   AssemblyProperties rule class   
//
//      History:
//      17 Jan 2017     VIBGYOR         Created 
//
//-----------------------------------------------------------------------------------

using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Planning.Middle;

namespace Ingr.SP3D.Content.Planning
{
    /// <summary>
    /// AssemblyProperties rule class
    /// </summary>
    public class AssemblyProperties : PlanningAssemblyPropertiesRuleBase
    {
        /// <summary>
        /// Evaluate the property values for the given property based on the rule xml.
        /// </summary>
        /// <param name="componentBase"></param>
        /// <param name="propName"></param>
        /// <param name="defaultValue"></param>
        /// <param name="allowedValues"></param>
        public override void EvaluateProperties(AssemblyBase componentBase, string propertyName, out object defaultValue, out ReadOnlyCollection<object> allowedValues)
        {
            #region Input Validation

            if (componentBase == null)
            {
                throw new CmnArgumentNullException("Components");
            }
            if (string.IsNullOrEmpty(propertyName))
            {
                throw new CmnArgumentNullException("propertyName");
            }
            #endregion Input Validation

            Evaluate(componentBase, propertyName, out defaultValue, out allowedValues);
        }
    }
}
