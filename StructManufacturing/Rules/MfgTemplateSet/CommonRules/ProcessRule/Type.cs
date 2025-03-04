//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Type rule returns the type and interval of the TemplateSet.    
//
//      Author:  
//
//      History:
//      May 28th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using TemplateProcessInfo = Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// TemplateSet Type rule.
    /// </summary>
    public class Type : UserDefinedValuesRule
    {    
        /// <summary>
        /// Gets the TemplateSet type and interval.
        /// </summary>
        /// <param name="processInfo">The process info.</param>
        public override void Evaluate(ProcessInformation processInfo)
        {
            try
            {
                if (processInfo == null)
                    throw new ArgumentNullException("Input ProcessInfo is empty");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                TemplateSet mfgTemplateSet = null;
                BusinessObject parentPart = null;

                PlatePartBase platePart = null;
                ProfilePart profilePart = null;

                if (processInfo.ManufacturingPart != null)
                {
                    mfgTemplateSet = (TemplateSet)processInfo.ManufacturingPart;
                }
                if (processInfo.ManufacturingParent != null)
                {
                    parentPart = (BusinessObject)processInfo.ManufacturingParent;

                    if (mfgTemplateSet != null)
                    {

                        if (mfgTemplateSet.Type == TemplateSetType.Plate)
                        {
                            platePart = (PlatePartBase)parentPart;
                        }
                        else if (mfgTemplateSet.Type == TemplateSetType.ProfileFace)
                        {
                            profilePart = (ProfilePart)parentPart;
                        }
                    }
                }

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                Dictionary<string, object> results = new Dictionary<string, object>();

                if(mfgTemplateSet != null)
                {
                    if (mfgTemplateSet.Type == TemplateSetType.Plate || mfgTemplateSet.Type == TemplateSetType.ProfileFace)
                    {
                        // ***Note: Uncomment the below lines of code if the values defined in XML need to be overwritten.

                        //int templateType = 5140;
                        //double interval = 0.2;

                        //results.Add("TemplateType", templateType);
                        //results.Add("Interval", interval);
                    }
                }

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                processInfo.SetAttribute((int)TemplateProcessInfo.ProcessValues.Type, results);

                #endregion Set Outputs

            }        
            catch (Exception e)
            {
                LogForToDoList(e, 3004, "Call to TemplateSet Type Rule failed with the error" + e.Message);
            }
        }
    }
}
