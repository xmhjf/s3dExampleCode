// Copyright (C) 2013 Intergraph Corporation.  All rights reserved.
// 
// Class NameMapRule
//Delivered out of the box Property Comparision Rule ignores any Name Property mismatches in Compare with DesignBasis 
//for Correlated Pipelines and PipeRuns. However, the Comparision behavior can be overridden using this rule.
//This rule compares the Name property for S3D Pipeline System with Correlated Design Basis PIDPipeline and 
//S3D PipeRun with Correlated Design Basis PIDPipingConnector. Name property mismatches between S3D and DesignBasis object being compared would show that 
//property as different (in red color) in Compare With Design Basis. 
// Callback ProgID for Pipeline and PipeRun mappings in SP3DPunblishMap.xml file should be updated with this ProgID for this to take effect.
 
// Author:  Sai
//
// Date created:  Feb 16, 2013
// 
// History:
// 
// Hailong        Feb 16, 2013 
//                CR-CP-222051  Pipe run name during correlation is not showing any mismatch and is not updated  
//
//

using System;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Route.Middle;
using Ingr.SP3D.Systems.Middle;
using Ingr.SP3D.Common.Middle.Services;

namespace PropertyMapRule
{
    public class NameMapRule : DesignBasisProperyMapBase
    {
        private enum ErrorLevel
        {
            ERRORLEVEL_ErrInformation = 1,
            ERRORLEVEL_ErrWarning = 2,
            ERRORLEVEL_ErrCritical = 3,
        }

        // returning false would consider the S3D and DesignBasis property value as different and show it as red in Compare with DesignBasis
        public override bool CompareProperty(BusinessObject s3dBO, DesignBasis designbasisBO, DesignBasis designbasisRelatedBO, string s3dInterfaceName, string s3dPropertyName, string designbasisInterfaceName, string designbasisPropertyName, ref object s3dPropVal, ref object designbasisPropVal)
        {
            bool ignoreOne2ManyNameDiff = false;
            //set ignoreOne2ManyNameDiff to True if Name differences should be ignored when more than one S3D object (Pipeline or Piperun)
            //is correlated to a single design basis object (PIDPipeline or PIDPipingConnector)
            bool comparisionResult = false;
            try
            {
                if (ignoreOne2ManyNameDiff == true)
                {
                    RelationCollection correlatedObjectCollection = designbasisBO.GetRelationship("SP3DCorrelationToDesignBasis", "CorrelatedObject");
                    if (correlatedObjectCollection.TargetObjects.Count > 1)
                    {
                        comparisionResult = true;
                    }
                }
                else
                {
                    String s3dName = s3dPropVal as string;
                    String designbasisName = designbasisPropVal as string;
                    if (String.Compare(s3dName, designbasisName) == 0)
                    {
                        comparisionResult = true;
                    }
                    else
                    {
                        comparisionResult = false;
                    };
                }
            }
            catch (Exception exp)
            {
                LogErrorMessage(ErrorLevel.ERRORLEVEL_ErrInformation, this.ToString() + ".CompareProperty", exp);

            }
            return comparisionResult;

        }

        // returned value is what gets displayed for DesgignBasis object in the Compare with DesignBasis
        public override object ConvertProperty(DesignBasis designbasisBO, DesignBasis designbasisRelatedBO, string s3dInterfaceName, string s3dPropertyName, string designbasisInterfaceName, string designbasisPropertyName)
        {
            PropertyValueString designbasisPropertyValueString = null;

            try
            {
                if (designbasisRelatedBO != null)
                {
                    designbasisPropertyValueString = designbasisRelatedBO.GetPropertyValue(designbasisInterfaceName, designbasisPropertyName) as PropertyValueString;
                }
                else if (designbasisBO != null)
                {
                    designbasisPropertyValueString = designbasisBO.GetPropertyValue(designbasisInterfaceName, designbasisPropertyName) as PropertyValueString;
                }
            }
            catch (Exception exp)
            {
                LogErrorMessage(ErrorLevel.ERRORLEVEL_ErrInformation, this.ToString() + ".ConvertProperty", exp);

            }
            
            if(designbasisPropertyValueString != null)
                return designbasisPropertyValueString.PropValue;
            else
            {
                return null;
            }
        }

        // this method would update the S3D object name from the DesignBasis object name 
        public override bool SetProperty(BusinessObject s3dBO, BusinessObject s3dRelatedBO, DesignBasis designbasisBO, DesignBasis designbasisRelatedBO, string s3dInterfaceName, string s3dPropertyName, string designbasisInterfaceName, string designbasisPropertyName)
        {
            bool updatePropertyValue;
            try
            {
                updatePropertyValue = false;
                BusinessObject nameditemBO = null;
                NameRuleHelper s3dNameRuleHelper = null;
                NamedItemHelper s3dNamedItemHelper;

                if (designbasisBO == null)
                    throw new ArgumentNullException("design basis object is null.");

                PropertyValue propertyValue = designbasisBO.GetPropertyValue(designbasisInterfaceName, designbasisPropertyName);
                PropertyValueString oPropString = propertyValue as PropertyValueString;
                    
                // set property only when name is set on the designbasis object
                if (oPropString != null)
                {
                    if (oPropString.ToString().Length != 0)
                    {
                        if (s3dRelatedBO is Pipeline)
                        {
                            nameditemBO = s3dRelatedBO;
                        }
                        else if (s3dBO is Pipeline)
                        {
                            nameditemBO = s3dBO;
                        }
                        else if (s3dBO is PipeRun)
                        {
                            nameditemBO = s3dBO;
                        }
                    }
                }
                
                if (nameditemBO != null)
                {
                    s3dNameRuleHelper = new NameRuleHelper(nameditemBO);
                    s3dNameRuleHelper.ActiveNameRule = null;
                    s3dNamedItemHelper = new NamedItemHelper(nameditemBO);
                    s3dNamedItemHelper.SetUserDefinedName(oPropString.PropValue);
                    updatePropertyValue = true;
                }
            }
            catch (Exception exp)
            {
                updatePropertyValue = false;
                LogErrorMessage(ErrorLevel.ERRORLEVEL_ErrCritical, this.ToString() + ".SetProperty", exp);
            }

            return updatePropertyValue;
        }

        // Log error info to log file
        private void LogErrorMessage(ErrorLevel errLevel, string sMethod, Exception exp)
        {
            // todo add error logging.
            LogError logError = MiddleServiceProvider.ErrorLogger;
            string componentName = "NameMapRule : ";
            logError.Log(componentName + sMethod + "" + exp.Message + "", (int)errLevel);
            return;
        }
    }
}