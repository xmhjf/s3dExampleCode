//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   UserDefinedValues rule overrides user defined value for built-in attribute.  
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
    /// UserDefinedValues Rule
    /// </summary>
    public class UserDefinedValues : UserDefinedValuesRule
    {
        /// <summary>
        /// Override the UserDefined Values.
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

                #endregion

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                int unfoldAlgorithm = 0, applyMarginInProcess = 0,userDefinedNamingRule = 0;

                if (mfgTemplateSet != null)
                {
                    if (mfgTemplateSet.Type == TemplateSetType.Plate || mfgTemplateSet.Type == TemplateSetType.ProfileFace)
                    {
                        //' Applying margin to an edges based on the template type.
                        //' applyMarginInProcess = 0 - Don't apply margin to Template
                        //' applyMarginInProcess = 1 - Do Apply Margin to Template 

                        string templateType = Convert.ToString(processInfo.GetAttribute((int)TemplateProcessInfo.ProcessValues.Type, "TemplateType"));

                        if (templateType == "BOX" ||
                            templateType =="USERDEFINEDBOX" ||
                            templateType == "USERDEFINEDBOXWITHEDGES")
                        {
                            applyMarginInProcess = 0;
                        }
                        else
                        {
                            applyMarginInProcess = 1;
                        }

                        if (templateType == "USERDEFINED" ||
                            templateType == "USERDEFINEDBOX" ||
                            templateType == "USERDEFINEDBOXWITHEDGES")
                        {
                            userDefinedNamingRule = 1;
                        }
                    }
                    else if (mfgTemplateSet.Type == TemplateSetType.ProfileEdge)
                    {
                        unfoldAlgorithm = 1;
                    }
                    else //if (mfgTemplateSet.Type == TemplateSetType.Tube)
                    {
                        unfoldAlgorithm = 2;
                    }
                }

                #endregion

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                if (mfgTemplateSet != null)
                {
                    if (mfgTemplateSet.Type == TemplateSetType.Plate || mfgTemplateSet.Type == TemplateSetType.ProfileFace)
                    {
                        processInfo.SetAttribute((int)TemplateProcessInfo.UserDefinedValues.ApplyMarginInProcess, applyMarginInProcess);
                    }
                    else if (mfgTemplateSet.Type == TemplateSetType.ProfileEdge || mfgTemplateSet.Type == TemplateSetType.Tube)
                    {
                        processInfo.SetAttribute((int)TemplateProcessInfo.UserDefinedValues.UnfoldAlgorithm, unfoldAlgorithm);
                    }

                    processInfo.SetAttribute((int)TemplateProcessInfo.UserDefinedValues.UserDefinedNamingRule, userDefinedNamingRule);
                }

                #endregion
            }
            catch (Exception e)
            {               
                LogForToDoList(e, 3025, "Call to TemplateSet UserDefined Rule failed with the error" + e.Message);
            }

        }
    }
}
