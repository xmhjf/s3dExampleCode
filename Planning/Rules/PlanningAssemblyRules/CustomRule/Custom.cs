//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Custom class 
//                
//
//      History:
//      April 24th, 2015   Created by VIBGYOR
//
//-----------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using Ingr.SP3D.Common.Middle.Services;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace Ingr.SP3D.Content.Planning
{
    /// <summary>
    /// Custom rule class Implements PlnAssemblyCustomRuleBase Class
    /// </summary>
    public class CustomRule : PlnAssemblyCustomRuleBase
    {
        # region Global Variables

        S3D_COMPONENT rulesXML;
        private bool xmlParsed = false;

        # endregion

        #region Methods

        /// <summary>
        /// Get argument value from argument category and name in XML
        /// </summary>
        /// <param name="argsName"></param>
        /// <param name="argName"></param>
        private string GetArgumentValue(string actionName, string argsName, string argName)
        {
            IEnumerable<S3D_ARGS[]> argsColl = from actions in rulesXML.S3D_ACTION
                                               where actions.NAME == actionName
                                               select actions.S3D_PROPERTY.S3D_VALUE.S3D_ARGS;

            foreach (S3D_ARGS[] args in argsColl)
            {
                IEnumerable<S3D_ARG[]> argColl = from customArgs in args where customArgs.NAME == argsName select customArgs.S3D_ARG;                
                foreach (S3D_ARG[] arg in argColl)
                {
                    IEnumerable<string> argValues = from customArg in arg where customArg.NAME == argName select customArg.VALUE;
                    foreach (string value in argValues)
                    {
                        return value;
                    }
                }
                break;
            }
            return string.Empty;
        }

        /// <summary>
        /// Deserialize XML for computing the attributes
        /// </summary>
        /// <param name="xmlPath"></param>
        private void DeserializeXML(string xmlPath)
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Parse;
            settings.ValidationType = ValidationType.Schema;
            settings.IgnoreComments = true;
            settings.CloseInput = true;
            XmlReader xmlreader = XmlReader.Create(xmlPath, settings);
            XmlSerializer myserializer = new XmlSerializer(typeof(S3D_COMPONENT));
            rulesXML = (S3D_COMPONENT)myserializer.Deserialize(xmlreader);
            xmlParsed = true;
        }

        /// <summary>
        /// Compute the assembly custom data
        /// </summary>
        /// <param name="customData"></param>
        /// <returns>void</returns>
        public override void Evaluate(CustomData customData)
        {
            #region Input Error Handling

            if (customData == null)
            {
                throw new ArgumentNullException("customData Information is null");
            }

            #endregion //Input Error Handling

            try
            {
                if (xmlParsed == false)
                {
                    string xmlPath = string.Empty;
                    string sharedContentPath = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog.SymbolShare;

                    if (!string.IsNullOrEmpty(base.Argument) && (Directory.Exists(sharedContentPath + "\\" + base.Argument) == true) && (File.Exists(sharedContentPath + "\\" + base.Argument) == true))
                     {
                         xmlPath = sharedContentPath + "\\" + base.Argument;
                     }
                     else
                     {
                         xmlPath = sharedContentPath + "\\" + "Planning\\Assembly\\XML\\PlanningAssemblyRules.xml";
                     }

                    DeserializeXML(xmlPath);
                 }

                string argumetName = string.Empty;
                if (customData.CustomProcess == AssemblyCustomProcess.UpdateAssembly)
                {
                    argumetName = "UpdateAssembly";
                }
                else if(customData.CustomProcess == AssemblyCustomProcess.UpdatePlanningJoints)
                {
                    argumetName = "UpdatePlanningJoints";
                }

                bool customFlag = Convert.ToBoolean(GetArgumentValue("AssemblyCustomRules", "General", argumetName));
                customData.CustomFlag = customFlag;
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        # endregion
    }
}
