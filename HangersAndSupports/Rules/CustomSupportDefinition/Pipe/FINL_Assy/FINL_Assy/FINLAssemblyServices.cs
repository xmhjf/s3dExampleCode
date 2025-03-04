//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   FINLAssemblyServices.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.FINLAssemblyServices
//   Author       :  BS
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     BS     CR224491,224492,224485- Convert FINL_Assy to C# .Net 
//   11-Dec-2014     PVK    TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using System.Reflection;
using Ingr.SP3D.Common.Exceptions;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    public  static class FINLAssemblyServices
    {
        /// <summary>
        /// Gets the sub assembly joints.
        /// </summary>
        /// <param name="subAssembly">The sub assembly.</param>
        /// <param name="supCompColl">The support component collection.</param>
        /// <returns></returns>
        public static ReadOnlyCollection<object> GetSubAssemblyJoints(CustomSupportDefinition supAssembly, CustomSupportDefinition subAssembly, Collection<SupportComponent> supCompColl)
        {

            try
            {
                subAssembly.SupportHelper.SupportComponentDictionary = supAssembly.SupportHelper.SupportComponentDictionary;
                subAssembly.ConfigureSupport(supCompColl);
                return subAssembly.JointHelper.Joints();
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get GetSubAssemblyJoints of FINLAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// Create the assembly and returns the support
        /// </summary>
        /// <param name="sAssemblyName">sAssemblyName</param>
        /// <param name="sClassName">sClassName</param>
        public static CustomSupportDefinition GetAssembly(string progID, Ingr.SP3D.Support.Middle.Support support)
        {
            string sAssemblyName = String.Empty, sClassName = String.Empty;
            if (string.IsNullOrEmpty(progID))
            {
                throw new ArgumentNullException("progID");
            }

            ////the .NET progID not in the correct format
            int posOfDelimiter = progID.IndexOf(",");
            if (!(posOfDelimiter > 0))
            {
                throw new IncorrectProgIDFormatException(CmnLocalizer.GetString(FINL_AssyResourceIDs.CmnIncorrectProgIDFormat,
                    "Given ProgId is not in the correct format. .NET class ProgID should be in the following format: AssemblyName,Namespace.ClassName"));
            }

            ////get the active catalog to get the SymbolShare 
            Catalog catalog = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog;
            string symbolShare = catalog.SymbolShare;

            ////Full path of SystemSymbolConfig.xml
            string systemSymbolConfigFullPath = symbolShare + "\\Xml\\SystemSymbolConfig.xml";

            ////if SystemSymbolConfig.xml exist in symbol share then load the file into the XML document in memory
            string dllFullPath = string.Empty;
            if (File.Exists(systemSymbolConfigFullPath))
            {
                dllFullPath = GetDllFullPathFromXml(progID, systemSymbolConfigFullPath);

                ////may be it will present in CustomSymbolConfig.xml 
                if (string.IsNullOrEmpty(dllFullPath))
                {
                    ////Full path of CustomSymbolConfig.xml
                    string customSymbolConfigFullPath = symbolShare + "\\Xml\\CustomSymbolConfig.xml";

                    ////if CustomSymbolConfig.xml exist in symbol share then load the file into the XML document in memory
                    if (File.Exists(customSymbolConfigFullPath))
                    {
                        dllFullPath = GetDllFullPathFromXml(progID, customSymbolConfigFullPath);

                        ////we don't got the right one so it is error
                        if (string.IsNullOrEmpty(dllFullPath))
                        {
                            throw new ProgIDNotFoundInSymbolMapException(CmnLocalizer.GetString(FINL_AssyResourceIDs.CmnProgIDNotFoundInSymbolMap,
                               "Given ProgID not found in the System or Custom Symbol Configuration XML file in symbol share. Please make sure the assembly for the rule or symbol exists in symbol share and run the UpdateSymbolConfiguratoin command from Project Management."));
                        }
                    }
                    else
                    {
                        throw new FileNotFoundException(CmnLocalizer.GetString(FINL_AssyResourceIDs.CmnCustomSymbolConfigFileNotFound,
                            "CustomSymbolConfig.xml file is not found at symbol share."));
                    }
                }
            }
            else
            {
                throw new FileNotFoundException(CmnLocalizer.GetString(FINL_AssyResourceIDs.CmnSystemSymbolConfigFileNotFound,
                    "SystemSymbolConfig.xml file is not found at symbol share."));
            }

            ////Get the assembly full path
            string assemblyFullPath = dllFullPath.Replace("%OLE_SERVER%", symbolShare);

            ////Need to make sure the assembly is in symbol share or not
            if (!File.Exists(assemblyFullPath))
            {
                throw new AssemblyNotFoundInSymbolShareException(CmnLocalizer.GetString(FINL_AssyResourceIDs.CmnAssemblyNotFoundInSymbolShare,
                    "An assembly mapping to the given ProgID was found in the symbol configuration file, but the assembly cannot be found in symbol share. This can happen if the assembly was deleted later after updating symbol configuration or if the mapping file was copied from another symbol share."));
            }

            ////Get the class name including namespace 
            sClassName = progID.Substring(posOfDelimiter + 1);
            CustomSupportDefinition customSupportDefinition = null;
            bool isAssembly = false;
            try
            {
                // Load the assembly using LoadFrom SampleAssembly = Assembly.LoadFrom("c:\\Sample.Assembly.dll");
                Assembly assembly;
                try
                {
                    // @-quoting a string enables that, the escape sequences are not processed, which makes it easy to write, for example, a fully qualified file name:
                    // @"c:\Docs\Source\a.txt" rather than "c:\\Docs\\Source\\a.txt"
                    assembly = Assembly.LoadFrom(assemblyFullPath);
                    isAssembly = true;
                }
                catch (Exception oException)
                {
                    isAssembly = false;
                    // "Failed to load NameRule assembly : $1."
                    String[] paramsArray = new string[1];
                    paramsArray[0] = sAssemblyName;
                    String sExceptionMessage = CmnLocalizer.GetStringWithParameters(CmnResourceIDs.CmnLoadNameRuleAssemblyException, paramsArray, "SP3DSOMResources", "FINL_Assy");
                    throw new CmnLoadNameRuleAssemblyException(assemblyFullPath, oException);
                }

                // Get type of the class to load from assembly
                if (isAssembly)
                {
                    Type typeToLoad = null;
                    try
                    {
                        typeToLoad = assembly.GetType(sClassName);
                    }
                    catch (Exception oException)
                    {
                        // "Failed to get NameRule class [$1] type from assembly [$2]."
                        String[] paramsArray = new string[2];
                        paramsArray[0] = sClassName;
                        paramsArray[1] = sAssemblyName;
                        String sExceptionMessage = CmnLocalizer.GetStringWithParameters(CmnResourceIDs.CmnGetNameRuleTypeException, paramsArray, "SP3DSOMResources", "FINL_Assy");
                        throw new CmnGetNameRuleTypeException(sExceptionMessage, oException);
                    }
                    if (typeToLoad == null)
                    {
                        // "Failed to get NameRule class [$1] type from assembly [$2]."
                        String[] paramsArray = new string[2];
                        paramsArray[0] = sClassName;
                        paramsArray[1] = sAssemblyName;
                        String sExceptionMessage = CmnLocalizer.GetStringWithParameters(CmnResourceIDs.CmnGetNameRuleTypeException, paramsArray, "SP3DSOMResources", "FINL_Assy");
                        throw new CmnGetNameRuleTypeException(sExceptionMessage);
                    }

                    // Create an instance of sClassName
                    object classFromAssembly = null;
                    try
                    {
                        classFromAssembly = Activator.CreateInstance(typeToLoad);
                    }
                    catch (Exception oException)
                    {
                        // "Failed to create NameRule instance. Class:$1, Assembly:$2."
                        String[] paramsArray = new string[2];
                        paramsArray[0] = sClassName;
                        paramsArray[1] = sAssemblyName;
                        String sExceptionMessage = CmnLocalizer.GetStringWithParameters(CmnResourceIDs.CmnNameRuleBaseSubClassedTypeException, paramsArray, "SP3DSOMResources", "FINL_Assy");
                        throw new CmnNameRuleBaseSubClassedTypeException(sExceptionMessage, oException);
                    }


                    customSupportDefinition = null;

                    // Check to see if NameRule is a new implementation ... derives from NameRuleBase class
                    customSupportDefinition = classFromAssembly as CustomSupportDefinition;
                    if (customSupportDefinition == null)
                    {
                        // "Class [$1] from assembly [$2] is not a namerule."
                        String[] paramsArray = new string[2];
                        paramsArray[0] = sClassName;
                        paramsArray[1] = sAssemblyName;
                        String sExceptionMessage = CmnLocalizer.GetStringWithParameters(CmnResourceIDs.CmnNotNameRuleException, paramsArray, "SP3DSOMResources", "FINL_Assy");
                        throw new CmnNotNameRuleException(sExceptionMessage);
                    }
                }
                //Initialize the properties on CSD
                customSupportDefinition.SupportHelper = new SupportHelper(support);
                customSupportDefinition.JointHelper = new JointHelper(support);
                customSupportDefinition.SupportingHelper = new SupportingHelper(support);
                customSupportDefinition.SupportedHelper = new SupportedHelper(support);
                customSupportDefinition.RefPortHelper = new RefPortHelper(support);
                customSupportDefinition.GenericHelper = new GenericHelper(support);
                customSupportDefinition.BoundingBoxHelper = new BoundingBoxHelper(support);

                return customSupportDefinition;

            }
            catch (Exception e)
            {
                // Report and raise the exception.
                CmnException e1 = new CmnException("Error in GetAssembly of FINLAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }

        }
        /// <summary>
        /// Gets the DLL full path from XML.
        /// </summary>
        /// <param name="progID">The prog ID.</param>
        /// <param name="xmlFileFullPath">The XML file full path.</param>
        /// <returns></returns>
        private static string GetDllFullPathFromXml(string progID, string xmlFileFullPath)
        {
            if (string.IsNullOrEmpty(progID))
            {
                throw new ArgumentNullException("progID");
            }
            if (string.IsNullOrEmpty(xmlFileFullPath))
            {
                throw new ArgumentNullException("xmlFileFullPath");
            }

            string dllFullPath = string.Empty;
            string progidsNode = "progids", progidNode = "progid", progidNameAttribute = "name", progidDllAttribute = "dll";

            ////creating new XElement in memory
            XElement xElement = XElement.Load(xmlFileFullPath);

            ////get all the progids nodes
            IEnumerable<XElement> progIDsNodes = xElement.Elements(progidsNode);

            ////get the required node 
            IEnumerable<XElement> requiredNode = from c in progIDsNodes.Elements(progidNode) where (string)c.Attribute(progidNameAttribute) == progID select c;
            if (requiredNode != null)
            {
                foreach (XElement requiredXElement in requiredNode)
                {
                    ////get the dll full path from xml
                    dllFullPath = (string)requiredXElement.Attribute(progidDllAttribute);
                }
            }

            return dllFullPath;
        }
        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name .</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <returns> double value</returns>
        /// <example>
        /// <code>
        ///     PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
        ///     pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
        ///     double dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862","IJUAFINL_C", "C","IJUAFINL_PipeND_mm", "PipeND",pipeDiameter)   
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, double referenceValue)
        {
            IEnumerable<BusinessObject> finlCmpParts = null;
            try
            {
                double distance = 0;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass finlCmpPartClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (finlCmpPartClass.PartClassType.Equals("HgrServiceClass"))
                    finlCmpParts = finlCmpPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    finlCmpParts = finlCmpPartClass.Parts;

                finlCmpParts = finlCmpParts.Where(part => Ingr.SP3D.Content.Support.Symbols.HgrCompareDoubleService.cmpdbl((double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue, referenceValue) == true);
                if (finlCmpParts.Count() > 0)
                    distance = ((double)((PropertyValueDouble)finlCmpParts.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);

                return distance;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get DataByCondition of FINLAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (finlCmpParts is IDisposable)
                {
                    ((IDisposable)finlCmpParts).Dispose(); // This line will be executed
                }
            }
        }
        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name .</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <returns> double value</returns>
        /// <example>
        /// <code>
        ///     PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
        ///     pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
        ///     double dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862","IJUAFINL_C", "C","IJUAFINL_PipeND_mm", "PipeND",pipeDiameter)   
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, int referenceValue)
        {
            IEnumerable<BusinessObject> finlCmpParts = null;
            try
            {
                double distance = 0;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass finlCmpPartClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);
                if (finlCmpPartClass.PartClassType.Equals("HgrServiceClass"))
                    finlCmpParts = finlCmpPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    finlCmpParts = finlCmpPartClass.Parts;

                finlCmpParts = finlCmpParts.Where(part => (int)((PropertyValueInt)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue == referenceValue);
                if (finlCmpParts.Count() > 0)
                    distance = ((double)((PropertyValueDouble)finlCmpParts.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);

                return distance;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get DataByConditionof FINLAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (finlCmpParts is IDisposable)
                {
                    ((IDisposable)finlCmpParts).Dispose(); // This line will be executed
                }
            }
        }
        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>
        ///     PropertyValueCodelist loadClassCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass");
        ///     loadClass = loadClassCodeList.PropertyInfo.DisplayName;
        ///     double dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862","IJUAFINL_C", "C","IJUAFINL_LoadClass", "LoadClass",loadClass)   
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, string referenceValue)
        {
            IEnumerable<BusinessObject> finlCmpParts = null;
            try
            {
                double distance;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass finlCmpPartClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (finlCmpPartClass.PartClassType.Equals("HgrServiceClass"))
                    finlCmpParts = finlCmpPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    finlCmpParts = finlCmpPartClass.Parts;

                finlCmpParts = finlCmpParts.Where(part => (string)((PropertyValueString)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue == referenceValue);
                if (finlCmpParts.Count() > 0)
                    distance = ((double)((PropertyValueDouble)finlCmpParts.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                else
                    distance = 0;
                return distance;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get DataByConditionof FINLAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (finlCmpParts is IDisposable)
                {
                    ((IDisposable)finlCmpParts).Dispose(); // This line will be executed
                }
            }
        }
        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <param name="comparisionavalue">The comparision value</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>
        ///     PropertyValueCodelist loadClassCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass");
        ///     loadClass = loadClassCodeList.PropertyInfo.DisplayName;
        ///     double dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862","IJUAFINL_C", "C","IJUAFINL_LoadClass", "LoadClass",loadClass,0.001)   
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, double minimumReferencevalue, double maximumReferencevalue)
        {
            IEnumerable<BusinessObject> finlCmpParts = null;
            try
            {
                double distance = 0;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass finlCmpPartClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (finlCmpPartClass.PartClassType.Equals("HgrServiceClass"))
                    finlCmpParts = finlCmpPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    finlCmpParts = finlCmpPartClass.Parts;

                finlCmpParts = finlCmpParts.Where(part => (double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue > minimumReferencevalue && (double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue < maximumReferencevalue);
                if (finlCmpParts.Count() > 0)
                    distance = ((double)((PropertyValueDouble)finlCmpParts.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                return distance;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get DataByConditionof FINLAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (finlCmpParts is IDisposable)
                {
                    ((IDisposable)finlCmpParts).Dispose(); // This line will be executed
                }
            }
        }
        /// <summary>
        /// When overridden in a derived class, sets the property value for a specified object
        /// </summary>
        /// <param name="csd">The object whose property value will be set.</param>
        /// <param name="property">The property name to be set.</param>
        /// <param name="value">The new property value.</param>
        public static void SetValueOnPropertyType(Object csd, string property, Object value)
        {
            Type type = csd.GetType();
            PropertyInfo propInfo = type.GetProperty(property);
            propInfo.SetValue(csd, value, null);
        }

        /// <summary>
        /// Get the dimensions of the cross section.
        /// </summary>
        /// <param name="SectionStandard">The SectionStandard.</param>
        /// <param name="sectionType">Section Type of the cross section.</param>
        /// <param name="SectionName">Name of the section.</param>
        /// <param name="width">width of the cross section.</param>
        /// <param name="flangeThickness">flangeThickness of the cross section.</param>
        /// <param name="webThickness">webThickness of the cross section.</param>
        /// <param name="depth">depth of the cross section.</param>        
        /// <example>
        /// <code>
        ///  double widthL, thicknessL, depthL, l webThicknessL;
        ///  FINLAssemblyServices.GetCrossSectionDimenstions("Euro", "L", sSectionSize, out widthL, out thicknessL, out webThicknessL, out depthL);   
        /// </code>
        /// </example>      
        public static void GetCrossSectionDimensions(string SectionStandard, string sectionType, string SectionName, out double width, out double flangeThickness, out double webThickness, out double depth)
        {
            CatalogStructHelper catalogHelper = new CatalogStructHelper();
            CrossSection crosssection = catalogHelper.GetCrossSection(SectionStandard, sectionType, SectionName);

            //Get the C Section Data
            width = crosssection.Width;
            try
            {
                flangeThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
            }
            catch
            {
                flangeThickness = crosssection.Width;
            }
            try
            {
                webThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tw")).PropValue;
            }
            catch
            {
                webThickness = flangeThickness;
            }
            depth = crosssection.Depth;
        }
    }
}
