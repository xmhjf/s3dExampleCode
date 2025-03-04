//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   MaximumHeight rule returns the maximum height of the TemplateSet.    
//
//      Author:  
//
//      History:
//      May 28th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using System.Linq;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using TemplateProcessInfo = Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// MaximumHeight Rule.
    /// </summary>
    public class MaximumHeight : UserDefinedValuesRule
    {
        /// <summary>
        /// Gets the TemplateSet maximum height.
        /// </summary>
        /// <param name="processInfo">The process info.</param>
        public override void Evaluate(ProcessInformation processInfo)
        {
            try
            {
                if (processInfo == null)
                    throw new ArgumentNullException("Input ProcessInfo is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                TemplateSet mfgTemplateSet = null;
                BusinessObject parentPart = null;

                PlatePartBase platePart = null;
                ProfilePart profilePart = null;
                MemberPart memberPart = null;

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
                        else if (mfgTemplateSet.Type == TemplateSetType.ProfileFace || mfgTemplateSet.Type == TemplateSetType.ProfileEdge)
                        {
                            profilePart = (ProfilePart)parentPart;
                        }
                        else if (mfgTemplateSet.Type == TemplateSetType.Tube)
                        {
                            memberPart = (MemberPart)parentPart;
                        }
                    }
                }

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                // Set the Maximum Template Height depending on the TemplateSet Type.
                // Currently only returns a fixed value but this could be
                // in the future dependent on the workcenter assignment

                double maximumHeight = 0.0;

                if (mfgTemplateSet != null)
                {
                    switch (mfgTemplateSet.Type)
                    {
                        case TemplateSetType.Plate:
                            maximumHeight = 1.0;
                            break;
                        case TemplateSetType.ProfileFace:
                            maximumHeight = 0.6;
                            break;
                        case TemplateSetType.ProfileEdge:
                            maximumHeight = 1.0;
                            break;
                        case TemplateSetType.Tube:
                            maximumHeight = 1.0;
                            break;
                    }
                }

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                processInfo.SetAttribute((int)TemplateProcessInfo.ProcessValues.MaximumHeight, maximumHeight);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3007, "Call to TemplateSet Maximum Height Rule failed with the error" + e.Message);

            }
        }
    }
}
