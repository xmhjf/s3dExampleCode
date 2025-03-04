//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   FrameAssemblyServices.cs
//  
//   Author       :  Rajeswari
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who                change description
//   -----------     ---                ------------------
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET  
//   17/12/2015     Ramya     TR 284319	Multiple Record exception dumps are created on copy pasting supports
//   06/06/2016     Vinay     TR-CP-296065	Fix new coverity issues found in H&S Content 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    public static class OglaendAssemblyServices
    {
        /// <summary>
        /// Used in the AddPart function to get the Rule Type
        /// </summary>
        public enum RuleType { HgrPartSelectionRule = 1, HgrSupportRule = 2, None = 0 };
        /// <summary>
        /// Valid Orientation Angles for Steel Parts
        /// Values match hsSteelOrientationAngle CodeList from HS_System_Codelist.xls
        /// </summary>
        public enum SteelOrientationAngle { SteelOrientationAngle_0 = 0, SteelOrientationAngle_90 = 90, SteelOrientationAngle_180 = 180, SteelOrientationAngle_270 = 270 };

        /// <summary>
        /// Used in the AddPart function to get the Rule Type
        /// </summary>
        public enum FromCatalogOrOccurrence { FromCatalog = 1, FromOccurrence=2 };

        /// <summary>
        /// Data Type to hold configuration data for steel parts.
        /// </summary>
        public struct HSSteelConfig
        {
            public SteelOrientationAngle Orient;
            public int CardinalPoint;
            public double OffsetX;
            public double OffsetY;
            //public double dOffsetZ  (Removed this because there is no z Offset on Begin/End Cap ports)
        }
        /// <summary>
        /// Data type to hold steel cross section attributes
        /// </summary>
        public struct HSSteelMember
        {
            // Names and Descriptions
            public string sectionStandard;
            public string sectionType;
            public string sectionName;
            public string sectionDescription;
            public string partNumber;
            // Dimensions and Data
            public double unitWeight;
            public double depth;
            public double width;
            public double webThickness;
            public double webDepth;
            public double flangeThickness;
            public double flangeWidth;
            public double centroidX;
            public double centroidY;
            // Back to Back Data
            public int B2B_Config;
            public double B2B_Spacing;
            public double B2B_SingleFlangeWidth;
            // HSS
            public double HSS_NominalWallThickness;
            public double HSS_DesignWallThickness;
            // HSSR
            public double HSSR_RatioWidthperThickness;
            public double HSSR_RatioHeightperThickness;
            // HSSC
            public double HSSC_OuterDiameter;
            public double HSSC_RatioDepthPerThickness;
            // Flanged Bolt Gage
            public double FB_FlangeGage;
            public double FB_WebGage;
            // Angle Bolt Gage
            public double AB_LongSideGage;
            public double AB_LongSideGage1;
            public double AB_LongSideGage2;
            public double AB_ShortSideGage;
            public double AB_ShortSideGage1;
            public double AB_ShortSideGage2;
        }
        /// <summary>
        /// Data Type For Weld Parts
        /// </summary>
        public struct WeldData
        {
            public string partKey;
            public string partNumber;
            public string partRule;
            public string connection;
            public int location;
            public double offsetXValue;
            public double offsetYValue;
            public double offsetZValue;
        }
        // Different BBX Orientations for Frame Specific Bounding Boxes
        public enum FrameBBXOrientation { FrameBBXOrientation_Direct = 1, FrameBBXOrientation_Orthogonal = 2, FrameBBXOrientation_Tangent = 3 };
        /// <summary>
        /// Used for getting SteelConnectionAngle
        /// </summary>
        public enum SteelConnectionAngle { SteelConnectionAngle_90 = 90, SteelConnectionAngle_270 = 270 };

        /// <summary>
        /// Used for getting SteelConnection
        /// </summary>
        public enum SteelConnection { SteelConnection_Butted = 0, SteelConnection_Lapped = 1, SteelConnection_Nested = 2, SteelConnection_Coped = 3, SteelConnection_Fitted = 4, SteelConnection_Mitered = 5, SteelConnection_UserDefined = 999 };

        /// <summary>
        /// Used for getting SteelJointType
        /// </summary>
        public enum SteelJointType { SteelJoint_PRISMATIC = 4, SteelJoint_RIGID = 5 };
        /// <summary>
        /// Function to add a part to the collection of part occurences for the assembly.  This function will add the part and return the corresponding index for that part in the collection.
        /// The index can then be used later to access the part from the part occurence collection.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="partKey">Name of the PartKey</param>
        /// <param name="part">String containing a Part-Number or Part-Class</param>
        ///  <param name="rule">Case 1 - A Hanger Rule that returns a Part-Class or Part-Number Case 2 - A Part Selection Rule</param>
        ///  <param name="parts">colllection to hold PartInfo</param>
        ///  <param name="objectForRule">The object to be used by the Hanger Rule Specified in 'rule' </param>
        ///  <param name="ruletype">Rule type enum default value is "none"</param>
        /// <returns>Index of the part in the Part Occurence Collection.Returns null if no part is added</returns>
        /// NOTE: If a Hanger Rule is specified in the sRule parameter, then the sPart parameter MUST be left blank.
        /// <code>
        /// AddPart(this, partKey, partNo or partClassName, rule, parts)
        /// </code>
        public static Boolean AddPart(CustomSupportDefinition customSupportDefinition, string partKey, string part, string rule, Collection<PartInfo> parts, BusinessObject objectForRule = null, RuleType ruletype = RuleType.HgrPartSelectionRule)
        {
            try
            {
                bool isPartAdded = false; ;
                string partClassValue = string.Empty;
                Collection<object> ruleResults1 = new Collection<object>();
                if (ruletype == RuleType.HgrSupportRule)
                {
                    customSupportDefinition.GenericHelper.GetDataByRule(rule, objectForRule, out ruleResults1);
                    if (ruleResults1 != null)
                    {
                        if (ruleResults1.Count > 0 && ruleResults1[0] != (object)"" && ruleResults1[0] != null)
                        {
                            parts.Add(new PartInfo(partKey, ruleResults1[0].ToString()));
                            isPartAdded = true;
                        }
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(part))
                    {
                        if (rule == "")
                        {
                            GetPartClassValue(part, ref partClassValue);
                            parts.Add(new PartInfo(partKey, part, partClassValue));
                            isPartAdded = true;
                        }
                        else
                        {
                            parts.Add(new PartInfo(partKey, part, rule));
                            isPartAdded = true;
                        }
                    }
                    else if (!string.IsNullOrEmpty(rule))
                    {
                        // Just a Rule is specifed
                        customSupportDefinition.GenericHelper.GetDataByRule(rule, objectForRule, out ruleResults1);
                        if (ruleResults1 != null)
                        {
                            if (ruleResults1.Count > 0 && ruleResults1[0] != (object)"" && ruleResults1[0] != null)
                            {
                                parts.Add(new PartInfo(partKey, ruleResults1[0].ToString()));
                                isPartAdded = true;
                            }
                            else
                            {
                                isPartAdded = false;
                            }
                        }
                    }
                    else
                    {
                        isPartAdded = false;
                    }
                }
                return isPartAdded;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in AddPart." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// Function to add an implied part to the collection of part occurences for the assembly
        /// The index can then be used later to access the part from the part occurence collection.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="partKey">Name of the PartKey</param>
        /// <param name="part">String containing a Part-Number or Part-Class</param>
        ///  <param name="rule">Case 1 - A Hanger Rule that returns a Part-Class or Part-Number Case 2 - A Part Selection Rule</param>
        ///  <param name="impliedparts"> colllection to hold implied PartInfo</param>
        ///  <param name="objectForRule">The object to be used by the Hanger Rule Specified in 'rule' </param>
        ///  <param name="quantity">Number of parts to add</param>
        /// NOTE: If a Hanger Rule is specified in the sRule parameter, then the sPart parameter MUST be left blank.
        /// <code>
        /// AddImpliedPart(this,partKey[], partNo or partClassName, rule, impliedParts)
        /// </code>
        public static Boolean AddImpliedPart(CustomSupportDefinition customSupportDefinition, string part, string rule, Collection<PartInfo> impliedparts, BusinessObject objectForRule = null, int quantity=1)
        {
            try
            {
                bool isImpliedPartAdded = false; ;
                string partClassValue = string.Empty, partRule = string.Empty;
                Collection<object> ruleResults1 = new Collection<object>();
                string[] partKey = new string[quantity];

                if (!string.IsNullOrEmpty(part))
                {
                    if (rule == "")
                    {
                        GetPartClassValue(part, ref partClassValue);
                        partRule = partClassValue;
                        isImpliedPartAdded = true;
                    }
                    else
                    {
                        partRule = rule;
                        isImpliedPartAdded = true;
                    }
                }
                else if (!string.IsNullOrEmpty(rule))
                {
                    // Just a Rule is specifed
                    customSupportDefinition.GenericHelper.GetDataByRule(rule, objectForRule, out ruleResults1);
                    if (ruleResults1 != null)
                    {
                        if (ruleResults1.Count > 0 && ruleResults1[0] != (object)"" && ruleResults1[0] != null)
                        {
                            partRule = ruleResults1[0].ToString();
                            isImpliedPartAdded = true;
                        }
                    }
                }
                else
                {
                    isImpliedPartAdded = false;
                }

                if (isImpliedPartAdded == true)
                {
                    for (int index = 0; index < quantity; index++)
                    {
                        partKey[index] = "IMPLIEDPART_" + index;
                        impliedparts.Add(new PartInfo(partKey[index], part, partRule));
                    }
                }
                return isImpliedPartAdded;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in AddImpliedPart." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method will Check either its Part or PartClass and return PSL(PartSelectionRule) if it is PartClass or return empty string for Part.
        /// </summary>
        /// <param name="partOrPartClassName">Name of the PartClass</param>
        /// <param name="partSelectionRule">Return the PartSelectionRule</param>
        /// <returns></returns>
        /// <code>
        /// GetPartClassValue(partClassName, ref partClassValue)
        /// </code>
        public static void GetPartClassValue(string partOrPartClassName, ref string partSelectionRule)
        {
            try
            {
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                BusinessObject partclass = catalogBaseHelper.GetPartClass(partOrPartClassName);
                if (partclass is PartClass)
                {
                    partSelectionRule = partclass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                }
                else
                {
                    partSelectionRule = partOrPartClassName;
                }
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetPartClassValue."  + "ERROR:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// Using the supplied part Key, looks up all the relevent data for the steel cross section.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="partKey">The part Key of a steel member, either a Hanger Beam or a Smart Steel part.</param>
        ///<code>
        /// Returns
        ///The cross section information in the form of a hsSteelMember user defined type.
        ///GetSectionDataFromPartIndex(this,partKey)
        ///</code>
        public static HSSteelMember GetSectionDataFromPartIndex(CustomSupportDefinition customSupportDefinition, string partKey)
        {
            try
            {
                HSSteelMember resultMember = new HSSteelMember();
                Dictionary<string, SupportComponent> componentDictionary = customSupportDefinition.SupportHelper.SupportComponentDictionary;
                IPart part = (IPart)componentDictionary[partKey].GetRelationship("madeFrom", "part").TargetObjects[0];

                if (partKey == string.Empty)
                    resultMember.partNumber = "";

                BusinessObject tSectionPart = componentDictionary[partKey].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)tSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                if (part != null)
                {
                    resultMember.partNumber = part.PartNumber;
                    resultMember.sectionDescription = part.PartDescription;
                }
                if (crossSection != null)
                {
                    resultMember.sectionName = crossSection.Name;
                    resultMember.sectionType = crossSection.CrossSectionClass.Name;
                    GetDoublePropertyValue(crossSection, "IStructCrossSectionDimensions", "Depth", ref resultMember.depth);
                    GetDoublePropertyValue(crossSection, "IStructCrossSectionDimensions", "Width", ref resultMember.width);
                    GetDoublePropertyValue(crossSection, "IStructFlangedSectionDimensions", "tf", ref resultMember.flangeThickness);
                    GetDoublePropertyValue(crossSection, "IStructFlangedSectionDimensions", "tw", ref resultMember.webThickness);
                    GetDoublePropertyValue(crossSection, "IStructCrossSectionUnitWeight", "UnitWeight", ref resultMember.unitWeight);
                    GetDoublePropertyValue(crossSection, "IStructFlangedSectionDimensions", "d", ref resultMember.webDepth);
                    GetDoublePropertyValue(crossSection, "IStructFlangedSectionDimensions", "bf", ref resultMember.flangeWidth);
                    GetDoublePropertyValue(crossSection, "IStructCrossSectionDesignProperties", "CentroidX", ref resultMember.centroidX);
                    GetDoublePropertyValue(crossSection, "IStructCrossSectionDesignProperties", "CentroidY", ref resultMember.centroidY);
                    GetIntPropertyValue(crossSection, "IJUA2L", "bbConfiguration", ref resultMember.B2B_Config);
                    GetDoublePropertyValue(crossSection, "IJUA2L", "bb", ref resultMember.B2B_Spacing);
                    GetDoublePropertyValue(crossSection, "IJUA2L", "b", ref resultMember.B2B_SingleFlangeWidth);
                    GetDoublePropertyValue(crossSection, "IJUAHSS", "tnom", ref resultMember.HSS_NominalWallThickness);
                    GetDoublePropertyValue(crossSection, "IJUAHSS", "tdes", ref resultMember.HSS_DesignWallThickness);
                    GetDoublePropertyValue(crossSection, "IJUAHSSR", "b_t", ref resultMember.HSSR_RatioWidthperThickness);
                    GetDoublePropertyValue(crossSection, "IJUAHSSR", "h_t", ref resultMember.HSSR_RatioHeightperThickness);
                    GetDoublePropertyValue(crossSection, "IJUAHSSC", "OD", ref resultMember.HSSC_OuterDiameter);
                    GetDoublePropertyValue(crossSection, "IJUAHSSC", "D_t", ref resultMember.HSSC_RatioDepthPerThickness);
                    GetDoublePropertyValue(crossSection, "IStructFlangedBoltGage", "gf", ref resultMember.FB_FlangeGage);
                    GetDoublePropertyValue(crossSection, "IStructFlangedBoltGage", "gw", ref resultMember.FB_WebGage);
                    GetDoublePropertyValue(crossSection, "IStructAngleBoltGage", "lsg", ref resultMember.AB_LongSideGage);
                    GetDoublePropertyValue(crossSection, "IStructAngleBoltGage", "lsg1", ref resultMember.AB_LongSideGage1);
                    GetDoublePropertyValue(crossSection, "IStructAngleBoltGage", "lsg2", ref resultMember.AB_LongSideGage2);
                    GetDoublePropertyValue(crossSection, "IStructAngleBoltGage", "ssg", ref resultMember.AB_ShortSideGage);
                    GetDoublePropertyValue(crossSection, "IStructAngleBoltGage", "ssg1", ref resultMember.AB_ShortSideGage1);
                    GetDoublePropertyValue(crossSection, "IStructAngleBoltGage", "ssg2", ref resultMember.AB_ShortSideGage2);
                }
                if (HgrCompareDoubleService.cmpdbl(resultMember.webThickness, 0) == true)
                    resultMember.webThickness = resultMember.flangeThickness;
                return resultMember;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetSectionDataFromPartIndex." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }


        /// <summary>
        /// Function to add all the welds supplied in the catalog for the given Support PartNumber.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="customSupportDefinition"></param>
        /// <param name="parts">The partinfo collection</param>
        /// <param name="interfaceName">The interface on which the weld information is stored on</param>
        /// <param name="supportPartNumber">Optional. The Support PartNumber to look up the welds for. If it is not specified, the current support part number will be used</param>
        ///<code>
        /// Returns
        ///A collection of hsWeldData Types, containing all the information for each individual weld, including the index that the partocc can be accesed from the part occurence collection.
        ///AddWeldsFromCatalog(this,parts,interfacename)
        ///</code>
        public static Collection<WeldData> AddWeldsFromCatalog(CustomSupportDefinition customSupportDefinition, Collection<PartInfo> parts, string interfaceName, string supportPartNumber = "")
        {
            try
            {
                string offsetZRule = string.Empty;
                WeldData weld = new WeldData();
                Collection<WeldData> weldCollection = new Collection<WeldData>();

                if (supportPartNumber == "")
                    supportPartNumber = customSupportDefinition.SupportHelper.Support.SupportDefinition.PartNumber;

                BusinessObject partClass = customSupportDefinition.SupportHelper.Support.SupportDefinition.PartClass;

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass S3DIframeWelds = (PartClass)catalogBaseHelper.GetPartClass(partClass + "_Welds");
                IEnumerable<BusinessObject> weldParts = S3DIframeWelds.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                weldParts = weldParts.Where(part => (string)((PropertyValueString)part.GetPropertyValue(interfaceName, "SupportPartNumber")).PropValue == supportPartNumber);
                int i = 1;
                foreach (BusinessObject part in weldParts)
                {
                    if (part.SupportsInterface("IJUAhsIFrameWelds"))
                    {
                        weld.partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsIFrameWelds", "WeldPartNumber")).PropValue;
                        weld.partRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsIFrameWelds", "WeldPartRule")).PropValue;
                        weld.connection = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsIFrameWelds", "Connection")).PropValue;
                        weld.location = (int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsIFrameWelds", "Location")).PropValue;

                        GetDoublePropertyValue(part, "IJUAhsIFrameWelds", "OffsetXValue", ref weld.offsetXValue);
                        string offsetXRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsIFrameWelds", "OffsetXRule")).PropValue;
                        if (!string.IsNullOrEmpty(offsetXRule))
                            customSupportDefinition.GenericHelper.GetDataByRule(offsetXRule, null, out weld.offsetXValue);

                        GetDoublePropertyValue(part, "IJUAhsIFrameWelds", "OffsetYValue", ref weld.offsetYValue);
                        string offsetYRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsIFrameWelds", "OffsetYRule")).PropValue;
                        if (!string.IsNullOrEmpty(offsetYRule))
                            customSupportDefinition.GenericHelper.GetDataByRule(offsetYRule, null, out weld.offsetYValue);

                        GetDoublePropertyValue(part, "IJUAhsIFrameWelds", "OffsetZValue", ref weld.offsetZValue);
                        offsetZRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsIFrameWelds", "OffsetZRule")).PropValue;
                    }
                    else if (part.SupportsInterface("IJUAhsLFrameWelds"))
                    {
                        weld.partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsLFrameWelds", "WeldPartNumber")).PropValue;
                        weld.partRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsLFrameWelds", "WeldPartRule")).PropValue;
                        weld.connection = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsLFrameWelds", "Connection")).PropValue;
                        weld.location = (int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsLFrameWelds", "Location")).PropValue;

                        GetDoublePropertyValue(part, "IJUAhsLFrameWelds", "OffsetXValue", ref weld.offsetXValue);
                        string offsetXRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsLFrameWelds", "OffsetXRule")).PropValue;
                        if (!string.IsNullOrEmpty(offsetXRule))
                            customSupportDefinition.GenericHelper.GetDataByRule(offsetXRule, null, out weld.offsetXValue);

                        GetDoublePropertyValue(part, "IJUAhsLFrameWelds", "OffsetYValue", ref weld.offsetYValue);
                        string offsetYRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsLFrameWelds", "OffsetYRule")).PropValue;
                        if (!string.IsNullOrEmpty(offsetYRule))
                            customSupportDefinition.GenericHelper.GetDataByRule(offsetYRule, null, out weld.offsetYValue);

                        GetDoublePropertyValue(part, "IJUAhsLFrameWelds", "OffsetZValue", ref weld.offsetZValue);
                        offsetZRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsLFrameWelds", "OffsetZRule")).PropValue;
                    }
                    else if (part.SupportsInterface("IJUAhsTFrameWelds"))
                    {
                        weld.partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsTFrameWelds", "WeldPartNumber")).PropValue;
                        weld.partRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsTFrameWelds", "WeldPartRule")).PropValue;
                        weld.connection = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsTFrameWelds", "Connection")).PropValue;
                        weld.location = (int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsTFrameWelds", "Location")).PropValue;

                        GetDoublePropertyValue(part, "IJUAhsTFrameWelds", "OffsetXValue", ref weld.offsetXValue);
                        string offsetXRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsTFrameWelds", "OffsetXRule")).PropValue;
                        if (!string.IsNullOrEmpty(offsetXRule))
                            customSupportDefinition.GenericHelper.GetDataByRule(offsetXRule, null, out weld.offsetXValue);

                        GetDoublePropertyValue(part, "IJUAhsTFrameWelds", "OffsetYValue", ref weld.offsetYValue);
                        string offsetYRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsTFrameWelds", "OffsetYRule")).PropValue;
                        if (!string.IsNullOrEmpty(offsetYRule))
                            customSupportDefinition.GenericHelper.GetDataByRule(offsetYRule, null, out weld.offsetYValue);

                        GetDoublePropertyValue(part, "IJUAhsTFrameWelds", "OffsetZValue", ref weld.offsetZValue);
                        offsetZRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsTFrameWelds", "OffsetZRule")).PropValue;
                    }
                    else if (part.SupportsInterface("IJUAhsUFrameWelds"))
                    {
                        weld.partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsUFrameWelds", "WeldPartNumber")).PropValue;
                        weld.partRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsUFrameWelds", "WeldPartRule")).PropValue;
                        weld.connection = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsUFrameWelds", "Connection")).PropValue;
                        weld.location = (int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsUFrameWelds", "Location")).PropValue;

                        GetDoublePropertyValue(part, "IJUAhsUFrameWelds", "OffsetXValue", ref weld.offsetXValue);
                        string offsetXRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsUFrameWelds", "OffsetXRule")).PropValue;
                        if (!string.IsNullOrEmpty(offsetXRule))
                            customSupportDefinition.GenericHelper.GetDataByRule(offsetXRule, null, out weld.offsetXValue);

                        GetDoublePropertyValue(part, "IJUAhsUFrameWelds", "OffsetYValue", ref weld.offsetYValue);
                        string offsetYRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsUFrameWelds", "OffsetYRule")).PropValue;
                        if (!string.IsNullOrEmpty(offsetYRule))
                            customSupportDefinition.GenericHelper.GetDataByRule(offsetYRule, null, out weld.offsetYValue);

                        GetDoublePropertyValue(part, "IJUAhsUFrameWelds", "OffsetZValue", ref weld.offsetZValue);
                        offsetZRule = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsUFrameWelds", "OffsetZRule")).PropValue;
                    }
                    if (!string.IsNullOrEmpty(offsetZRule))
                        customSupportDefinition.GenericHelper.GetDataByRule(offsetZRule, null, out weld.offsetZValue);

                    // Add the part to the Part Collection
                    weld.partKey = "weld" + i;
                    i++;
                    parts.Add(new PartInfo(weld.partKey, weld.partNumber, weld.partRule));

                    // Add the Weld Object to the Weld Collection
                    weldCollection.Add(weld);
                }
                return weldCollection;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in AddWeldsFromCatalog." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method returns the direct angle between two vectors in Radians.
        /// </summary>
        /// <param name="vector1">The first Vector</param>
        /// <param name="vector2">The Second Vector</param>
        /// <returns>Angle in Radians As Double</returns>
        ///<code>
        ///AngleBetweenVectors(vector1,vector2)
        ///</code>
        public static double AngleBetweenVectors(Vector vector1, Vector vector2)
        {
            return Math.Acos(vector1.Dot(vector2) / ((vector1.Length * vector2.Length)));
        }
        /// <summary>
        /// This method returns the resultant vector projection of one vector onto a plane defined by the normal.
        /// </summary>
        /// <param name="vector1">The Vector</param>
        /// <param name="vector2">The planeNormal</param>
        /// <returns>scalar vector</returns>
        ///<code>
        ///ProjectVectorIntoPlane(vector,normal)
        ///</code>
        public static Vector ProjectVectorIntoPlane(Vector vector, Vector planeNormal)
        {
            Vector resultVector = new Vector(0, 0, 0);

            // Make sure normVector is a Unit Vector
            Vector normal = new Vector(planeNormal.X, planeNormal.Y, planeNormal.Z);
            normal.Length = vector.Dot(normal);
            resultVector = vector.Subtract(normal);

            return resultVector;
        }
        /// <summary>
        /// This method is used to Create the BoundingBox according to the frameDimensions.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="boundingBoxName">name of the BoundingBox</param>
        /// <param name="bbxOrientation">direction of the BoundingBox</param>
        /// <param name="includeInsulation"> insulation option</param>
        /// <param name="mirrorFlag">mirror configuration</param>
        /// <param name="isPipeVertical">boolean ispipevertical or not</param>
        ///<code>
        ///CreateFrameBoundingBox(this, boundingBoxName,(FrameAssemblyServices.FrameBBXOrientation)frameOrientation, includeInsulation, mirrorFrame, SupportedHelper.IsSupportedObjectVertical(1, 45))
        ///</code>
        public static void CreateFrameBoundingBox(CustomSupportDefinition customSupportDefinition, string boundingBoxName, FrameBBXOrientation bbxOrientation, Boolean includeInsulation, Boolean mirrorFlag, Boolean isPipeVertical)
        {
            try
            {
                customSupportDefinition.BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                // Create Vectors to define the plane of the BBX
                Vector GlobalZ = new Vector(0, 0, 0), GlobalX = new Vector(0, 0, 0), GlobalY = new Vector(0, 0, 0), StructZ = new Vector(0, 0, 0), RouteZ = new Vector(0, 0, 0), RouteToStuct = new Vector(0, 0, 0), StructToStruct = new Vector(0, 0, 0), BB_X = new Vector(0, 0, 0), BB_Z = new Vector(0, 0, 0), BB_Y = new Vector(0, 0, 0);
                int supportingCount = 0;
                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    supportingCount = 1;
                else
                    supportingCount = customSupportDefinition.SupportHelper.Support.SupportingObjects.Count;

                string supportingType = String.Empty;

                if ((customSupportDefinition.SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                

                Matrix4X4 hangerPort = new Matrix4X4();
                Position pointObject = new Position(0, 0, 0);
                if (customSupportDefinition.SupportHelper.PlacementType != PlacementType.PlaceByReference && (supportingType == "Steel") && supportingCount == 2 && customSupportDefinition.SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                {
                    double origin_X1 = 0, origin_Y1 = 0, origin_Z1 = 0, origin_X2 = 0, origin_Y2 = 0, origin_Z2 = 0, origin_X = 0, origin_Y = 0, origin_Z = 0;


                    int numberOfRoutes = customSupportDefinition.SupportHelper.SupportedObjects.Count;
                    for (int i = 0; i < numberOfRoutes; i++)
                    {
                        hangerPort = customSupportDefinition.RefPortHelper.PortLCS("Structure");
                        origin_X1 = hangerPort.Origin.X;
                        origin_Y1 = hangerPort.Origin.Y;
                        origin_Z1 = hangerPort.Origin.Z;
                        hangerPort = new Matrix4X4();
                        hangerPort = customSupportDefinition.RefPortHelper.PortLCS("Struct_2");
                        origin_X2 = hangerPort.Origin.X;
                        origin_Y2 = hangerPort.Origin.Y;
                        origin_Z2 = hangerPort.Origin.Z;
                        hangerPort = new Matrix4X4();
                        hangerPort = customSupportDefinition.RefPortHelper.PortLCS("Route");
                        origin_X = hangerPort.Origin.X;
                        origin_Y = hangerPort.Origin.Y;
                        origin_Z = hangerPort.Origin.Z;
                    }

                    StructToStruct = new Vector(origin_X2 - origin_X1, origin_Y2 - origin_Y1, origin_Z2 - origin_Z1);
                    MemberSystem memberSystem1;
                    MemberPart member1;
                    double struct1StartX = 0, struct1StartY = 0, struct1StartZ = 0, struct1EndX = 0, struct1EndY = 0, struct1EndZ = 0;
                    double struct2StartX = 0, struct2StartY = 0, struct2StartZ = 0, struct2EndX = 0, struct2EndY = 0, struct2EndZ = 0;

                    object structObject;
                    structObject = customSupportDefinition.SupportHelper.SupportingObjects[0];
                    member1 = (MemberPart)structObject;
                    memberSystem1 = member1.MemberSystem;
                    Position position1Start = new Position(struct1StartX, struct1StartY, struct1StartZ);
                    Position position1End = new Position(struct1EndX, struct1EndY, struct1EndZ);
                    member1.Axis.EndPoints(out position1Start, out position1End);

                    structObject = new object();
                    structObject = customSupportDefinition.SupportHelper.SupportingObjects[1];
                    member1 = (MemberPart)structObject;
                    memberSystem1 = member1.MemberSystem;
                    Position position2Start = new Position(struct2StartX, struct2StartY, struct2StartZ);
                    Position position2End = new Position(struct2EndX, struct2EndY, struct2EndZ);
                    member1.Axis.EndPoints(out position2Start, out position2End);

                    double structDirectionAngle = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, "Struct_2", PortAxisType.X, OrientationAlong.Direct);
                    double tempStartX = 0, tempStartY = 0, tempStartZ = 0, tempEndX = 0, tempEndY = 0, tempEndZ = 0;
                    if ((structDirectionAngle * 180 / Math.PI) > 180 - 0.0001 && (structDirectionAngle * 180 / Math.PI) < 180 + 0.0001)
                    {
                        tempEndX = position1End.X;
                        tempEndY = position2End.Y;
                        tempEndZ = position2End.Z;
                        tempStartX = position2Start.X;
                        tempStartY = position2Start.Y;
                        tempStartZ = position2Start.Z;
                    }
                    else
                    {
                        tempStartX = position2End.X;
                        tempStartY = position2End.Y;
                        tempStartZ = position2End.Z;
                        tempEndX = position2Start.X;
                        tempEndY = position2Start.Y;
                        tempEndZ = position2Start.Z;
                    }
                    Collection<Position> points = new Collection<Position>();
                    points.Add(new Position(position1Start.X, position1Start.Y, position1Start.Z));
                    points.Add(new Position(position1End.X, position1End.Y, position1End.Z));
                    points.Add(new Position(tempStartX, tempStartY, tempStartZ));
                    points.Add(new Position(tempEndX, tempEndY, tempEndZ));

                    Plane3d plane = new Plane3d(points);
                    Vector pointVector = new Vector(0, 0, 0);
                    ISurface surface = plane;
                    RouteZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteZ);
                    try
                    {
                        customSupportDefinition.SupportingHelper.GetProjectedPointOnSurface(new Position(origin_X, origin_Y, origin_Z), new Vector(RouteZ.X, RouteZ.Y, RouteZ.Z), (BusinessObject)surface, out pointObject, out pointVector);
                    }
                    catch
                    {
                        pointObject = null;
                    }

                }

                // A couple special vectors for the tangent orientation
                // Tangent Enum Doesn't work well for different sized pipes, so for tangent
                // orientation I will grab the Y-Vector from the Tangent_BBR
                Vector tangent_X = new Vector(0, 0, 0); Vector tangent_Y = new Vector(0, 0, 0); Vector tangent_Z = new Vector(0, 0, 0);

                Boolean insulationOption;
                if (includeInsulation == true)
                    insulationOption = true;
                else
                    insulationOption = false;
                customSupportDefinition.BoundingBoxHelper.CreateStandardBoundingBoxes(insulationOption);

                if (isPipeVertical)
                {
                    switch (bbxOrientation)
                    {
                        case FrameBBXOrientation.FrameBBXOrientation_Direct:
                            {
                                // If Equipment need to do something different
                                BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ); // Use GlobalZ as BBX X-Axis
                                BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project vector from Route to Struct into plane of BBX

                                if ((supportingType == "Steel") && supportingCount == 2 && customSupportDefinition.SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                                {
                                    if (pointObject != null)
                                        BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                    else
                                    {
                                        BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteY, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                        if (Math.Round((customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y) * 180 / Math.PI)) > 90)
                                            BB_Z.Set(-BB_Z.X, -BB_Z.Y, -BB_Z.Z);
                                    }
                                }

                                customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                                break;
                            }
                        case FrameBBXOrientation.FrameBBXOrientation_Orthogonal:
                            {
                                BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ); // Use GlobalZ as BBX X-Axis
                                BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project vector from Route to Struct into plane of BBX
                                GlobalX = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalX, BB_X); // Project GlobalX into plane of BBX
                                GlobalY = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalY, BB_X); // Project GlobalY into plane of BBX

                                if ((supportingType == "Steel") && supportingCount == 2 && customSupportDefinition.SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                                {
                                    if (pointObject != null)
                                        BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                    else
                                    {
                                        BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteY, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                        if (Math.Round((customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y) * 180 / Math.PI)) > 90)
                                            BB_Z.Set(-BB_Z.X, -BB_Z.Y, -BB_Z.Z);
                                    }
                                }

                                // Get Orthogonal Vector in the plane of the BBX (Gx, -Gx, Gy or -Gy) depending on angle
                                if (AngleBetweenVectors(GlobalX, BB_Z) < (Math.Atan(1) * 4.0) / 4)
                                    BB_Z = GlobalX;
                                else if (AngleBetweenVectors(GlobalX, BB_Z) > 3 * (Math.Atan(1) * 4.0) / 4)
                                    BB_Z.Set(-GlobalX.X, -GlobalX.Y, -GlobalX.Z);
                                else if (AngleBetweenVectors(GlobalY, BB_Z) < (Math.Atan(1) * 4.0) / 4)
                                    BB_Z = GlobalY;
                                else
                                    BB_Z.Set(-GlobalY.X, -GlobalY.Y, -GlobalY.Z);

                                customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                                break;
                            }
                        case FrameBBXOrientation.FrameBBXOrientation_Tangent:
                            {
                                int frameConfiguration;
                                try
                                {
                                    frameConfiguration = (int)((PropertyValueCodelist)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                                }
                                catch
                                {
                                    frameConfiguration = 0;
                                }
                                if ((frameConfiguration == 1 && customSupportDefinition.Configuration == 1) || (frameConfiguration == 2 && customSupportDefinition.Configuration == 2))
                                {
                                    hangerPort = new Matrix4X4();
                                    hangerPort = customSupportDefinition.RefPortHelper.PortLCS("TANGENT_BBR_Low");
                                }
                                else
                                {
                                    hangerPort = new Matrix4X4();
                                    hangerPort = customSupportDefinition.RefPortHelper.PortLCS("TANGENT_ALT_BBR_Low");
                                }
                                tangent_X.Set(hangerPort.XAxis.X, hangerPort.XAxis.Y, hangerPort.XAxis.Z);
                                tangent_Z.Set(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);
                                tangent_Y = tangent_Z.Cross(tangent_X);
                                GlobalZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);

                                BB_Y = ProjectVectorIntoPlane(tangent_Y, GlobalZ);
                                BB_Z = ProjectVectorIntoPlane(tangent_Z, GlobalZ);
                                BB_X = BB_Y.Cross(BB_Z);

                                customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                                break;
                            }
                    }
                }
                else
                {
                    if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        switch (bbxOrientation)
                        {
                            case FrameBBXOrientation.FrameBBXOrientation_Direct:
                                {
                                    // Vertical Plane Normal Along Route - Z Axis Towards Structure
                                    GlobalZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ); // Get Global Z
                                    BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.StructureY, GlobalZ); // Project Route X-Axis into Horizontal Plane
                                    BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.StructureZ, BB_X);
                                    //BB_Z = new Vector(-BB_Z.X, -BB_Z.Y, -BB_Z.Z); 
                                    BB_Z.Set(-BB_Z.X, -BB_Z.Y, -BB_Z.Z);
                                    customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                                    break;
                                }
                            case FrameBBXOrientation.FrameBBXOrientation_Orthogonal:
                                {
                                    // Vertical Plane Normal Along Route - Z Axis Orthogonal to Global CS
                                    GlobalZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ); // Project GlobalX into plane of BBX
                                    BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.StructureY, GlobalZ); // Project Route X-Axis into Horizontal Plane
                                    BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project vector from Route to Struct into plane of BBX
                                    BB_Y = GlobalZ.Cross(BB_X); // Get Horizontal Vector in the BBX Plane (For Orthogonal direction)

                                    // Get Orthogonal Vector in the plane of the BBX (Gz, -Gz, By, -By) depending on Angle

                                    if (AngleBetweenVectors(GlobalZ, BB_Z) < (Math.Atan(1) * 4.0) / 4)
                                        BB_Z = GlobalZ;
                                    else if (AngleBetweenVectors(GlobalZ, BB_Z) > 3 * (Math.Atan(1) * 4.0) / 4)
                                        BB_Z.Set(-GlobalZ.X, -GlobalZ.Y, -GlobalZ.Z);
                                    else if (AngleBetweenVectors(BB_Y, BB_Z) < (Math.Atan(1) * 4.0) / 4)
                                        BB_Z = BB_Y;
                                    else
                                        BB_Z.Set(-BB_Y.X, -BB_Y.Y, -BB_Y.Z);

                                    customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                                    break;
                                }
                            case FrameBBXOrientation.FrameBBXOrientation_Tangent:
                                {
                                    int frameConfiguration;
                                    try
                                    {
                                        frameConfiguration = (int)((PropertyValueCodelist)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                                    }
                                    catch
                                    {
                                        frameConfiguration = 0;
                                    }
                                    if ((frameConfiguration == 1 && customSupportDefinition.Configuration == 1) || (frameConfiguration == 2 && customSupportDefinition.Configuration == 2))
                                    {
                                        hangerPort = new Matrix4X4();
                                        hangerPort = customSupportDefinition.RefPortHelper.PortLCS("TANGENT_BBR_Low");
                                    }
                                    else
                                    {
                                        hangerPort = new Matrix4X4();
                                        hangerPort = customSupportDefinition.RefPortHelper.PortLCS("TANGENT_ALT_BBR_Low");
                                    }
                                    tangent_X.Set(hangerPort.XAxis.X, hangerPort.XAxis.Y, hangerPort.XAxis.Z);
                                    tangent_Z.Set(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);
                                    tangent_Y = tangent_Z.Cross(tangent_X);
                                    GlobalZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);

                                    BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.StructureY, GlobalZ);
                                    BB_Z = ProjectVectorIntoPlane(tangent_Z, BB_X);

                                    customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                                    break;
                                }
                        }
                    }

                    else
                    {
                        switch (bbxOrientation)
                        {
                            case FrameBBXOrientation.FrameBBXOrientation_Direct:
                                {
                                    RouteToStuct = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteToStructure);

                                    BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, RouteToStuct); // Project Route X-Axis into Horizontal Plane
                                    BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X);

                                    if (customSupportDefinition.SupportHelper.PlacementType != PlacementType.PlaceByReference && (supportingType == "Steel") && supportingCount == 2 && customSupportDefinition.SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                                    {
                                        if (pointObject != null)
                                            BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                        else
                                        {
                                            BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteY, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                            if (Math.Round((customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y) * 180 / Math.PI)) > 90)
                                                BB_Z.Set(-BB_Z.X, -BB_Z.Y, -BB_Z.Z);
                                        }
                                    }

                                    customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                                    break;
                                }
                            case FrameBBXOrientation.FrameBBXOrientation_Orthogonal:
                                {
                                    StructZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.StructureZ);
                                    GlobalZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ); // Get Global Z
                                    BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, GlobalZ); // Use GlobalZ as BBX X-Axis
                                    BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project Route into Horizontal Plane

                                    BB_Y = GlobalZ.Cross(BB_X); // Get Horizontal Vector in the BBX Plane (For Orthogonal direction)
                                    if ((supportingType == "Steel") && supportingCount == 2 && customSupportDefinition.SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                                    {
                                        if (pointObject != null)
                                            BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                        else
                                        {
                                            BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteY, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                            if (Math.Round((customSupportDefinition.RefPortHelper.PortConfigurationAngle ("Route", "Structure", PortAxisType.Y) * 180 / Math.PI)) > 90)
                                                BB_Z.Set(-BB_Z.X, -BB_Z.Y, -BB_Z.Z);
                                        }
                                    }
                                   
                                    // Get Orthogonal Vector in the plane of the BBX (Gx, -Gx, Gy or -Gy) depending on angle
                                    if (AngleBetweenVectors(GlobalZ, BB_Z) < (Math.Atan(1) * 4.0) / 4)
                                        BB_Z = GlobalZ;
                                    else if (AngleBetweenVectors(GlobalZ, BB_Z) > 3 * (Math.Atan(1) * 4.0) / 4)
                                        BB_Z.Set(-GlobalZ.X, -GlobalZ.Y, -GlobalZ.Z);
                                    else if (AngleBetweenVectors(BB_Y, BB_Z) < (Math.Atan(1) * 4.0) / 4)
                                        BB_Z = BB_Y;
                                    else
                                        BB_Z.Set(-BB_Y.X, -BB_Y.Y, -BB_Y.Z);

                                    customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                                    break;
                                }
                            case FrameBBXOrientation.FrameBBXOrientation_Tangent:
                                {
                                    int frameConfiguration;
                                    try
                                    {
                                        frameConfiguration = (int)((PropertyValueCodelist)customSupportDefinition.SupportHelper.Support.GetPropertyValue("IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                                    }
                                    catch
                                    {
                                        frameConfiguration = 0;
                                    }
                                    if ((frameConfiguration == 1 && customSupportDefinition.Configuration == 1) || (frameConfiguration == 2 && customSupportDefinition.Configuration == 2))
                                    {
                                        hangerPort = new Matrix4X4();
                                        hangerPort = customSupportDefinition.RefPortHelper.PortLCS("TANGENT_BBR_Low");
                                    }
                                    else
                                    {
                                        hangerPort = new Matrix4X4();
                                        hangerPort = customSupportDefinition.RefPortHelper.PortLCS("TANGENT_ALT_BBR_Low");
                                    }
                                    tangent_X.Set(hangerPort.XAxis.X, hangerPort.XAxis.Y, hangerPort.XAxis.Z);
                                    tangent_Z.Set(hangerPort.ZAxis.X, hangerPort.ZAxis.Y, hangerPort.ZAxis.Z);
                                    tangent_Y = tangent_Z.Cross(tangent_X);
                                    GlobalZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);

                                    BB_X = ProjectVectorIntoPlane(tangent_X, GlobalZ);
                                    BB_Z = tangent_Z;

                                    customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                                    break;
                                }
                        }
                    }
                }
            }
            catch (Exception e) 
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in CreateFrameBoundingBox." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method returns the scalar projection of one vector onto another
        /// </summary>
        /// <param name="vector1">The Vector1</param>
        /// <param name="vector2">The Vector2</param>
        /// <returns>scalar</returns>
        ///<code>
        ///GetVectorProjection(vector1,vector2)
        ///</code>
        public static Double GetVectorProjection(Vector vector1, Vector vector2)
        {
            return (vector1.Dot(vector2)) / vector2.Length;
        }
        /// <summary>
        /// Using the supplied part number, looks up all the relevent data for the steel cross section.
        /// </summary>
        /// <param name="partNumber">The part number of a steel member, either a Hanger Beam or a Smart Steel part</param>
        /// <param name="vector1">The Vector1</param>
        /// <param name="vector2">The Vector2</param>
        /// <returns>The cross section information in the form of a hsSteelMember user defined type.</returns>
        ///<code>
        ///GetSectionDataFromPart(partNumber)
        ///</code>
        public static HSSteelMember GetSectionDataFromPart(string partNumber)
        {
            try
            {
                HSSteelMember resultMember = new HSSteelMember();
                if (partNumber == "")
                    resultMember.partNumber = "";

                Catalog m_oPlantCatalog = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantCatalog;
                BusinessObject sectionData = m_oPlantCatalog.GetNamedObject(partNumber);
                Part part = sectionData as Part;
                if (part!=null)
                    resultMember.sectionDescription = part.PartDescription;

                CrossSection crossSection = (CrossSection)sectionData.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                if (crossSection != null)
                {
                    resultMember.sectionName = crossSection.Name;
                    resultMember.sectionType = crossSection.CrossSectionClass.Name;
                    resultMember.width = crossSection.Width;
                    resultMember.depth = crossSection.Depth;
                    GetDoublePropertyValue(crossSection, "IStructFlangedSectionDimensions", "tw", ref resultMember.webThickness);
                }

                return resultMember;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in GetSectionDataFromPart." + "ERROR:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// Formats a value into a specific unit with specified precision.
        /// </summary>
        /// <param name="businessObject">The object(Part) on which the attribute is present, the attributeshould be an occurence attribute specified in the catalog. If Nothing then the support object(Assembly) is used.</param>
        /// <param name="interfaceName">Inteface that the attribute is defined on.</param>
        /// <param name="attribute">Name of the attribute</param>
        /// <param name="value">The value of the attribute</param>
        /// <param name="primaryUnits"> Optional The units to convert the value into. If not specified then the default catalog units are used.</param>
        /// <param name="secondaryUnits">Optional The secondary units.  Can be used to combine different unit types.IE: PrimaryUnits = DISTANCE_FEET , SecondaryUnits = DISTANCE_INCE results in a formated value combining feet and inches - (4ft 3in)</param>
        /// <param name="precisionType"> Optional Enum specifying the required Precision Type</param>
        /// <param name="precision"> Optional Integer value specifying the required precision. For Decimal this is the number of digits after the decimal. For Fractional this is the precision of the denomonator. IE: 16 = 1/16th</param>
        /// /// <code>
        /// FormatValueWithUnits(SupportOrComponent, "IJUAhsFrameSpan", "SpanValue", spanValue, primaryUnits, secondaryUnits, precisionType, precision);
        /// </code>
        public static string FormatValueWithUnits(BusinessObject businessObject, string interfaceName, string attribute, double value, UnitName primaryUnits = UnitName.UNIT_NOT_SET, UnitName secondaryUnits = UnitName.UNIT_NOT_SET, PrecisionType precisionType = PrecisionType.PRECISIONTYPE_DECIMAL, int precision = 16)
        {
            try
            {
                string formatedString = string.Empty;

                PropertyValue attributeInfo = businessObject.GetPropertyValue(interfaceName, attribute);
                if (primaryUnits == UnitName.UNIT_NOT_SET)
                {
                    UOMManager uomManager = MiddleServiceProvider.UOMMgr;
                    primaryUnits = uomManager.GetDefaultPrimaryUnit(attributeInfo.PropertyInfo.UOMType);
                }

                UOMFormat uomFormat = MiddleServiceProvider.UOMMgr.GetDefaultUnitFormat(attributeInfo.PropertyInfo.UOMType);
                uomFormat.PrecisionType = precisionType;
                if (precisionType == PrecisionType.PRECISIONTYPE_FRACTIONAL)
                    uomFormat.FractionalPrecision = (short)precision;
                else
                    uomFormat.DecimalPrecision = (short)precision;

                uomFormat.LeadingZero = false;
                uomFormat.ReduceFraction = true;
                uomFormat.TrailingZeros = false;
                uomFormat.UnitsDisplayed = true;

                if (secondaryUnits == UnitName.UNIT_NOT_SET)
                    formatedString = MiddleServiceProvider.UOMMgr.FormatUnit(attributeInfo.PropertyInfo.UOMType, value, uomFormat, primaryUnits);
                else
                    formatedString = MiddleServiceProvider.UOMMgr.FormatUnit(attributeInfo.PropertyInfo.UOMType, value, uomFormat, primaryUnits, secondaryUnits, UnitName.UNIT_NOT_SET);

                return formatedString;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in FormatValueWithUnits." + "ERROR:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method will be called to check property is available on the supportcomponent and support return its value.
        /// </summary>
        /// <param name="part">supportcomponent or support is of type bussiness object</param>
        /// <param name="interfaceName">Interface to which the property belongs.</param>
        /// <param name="propertyName">Name of property.</param>
        /// <param name="propertyValue">Returns Property Value</param>
        /// <returns></returns>
        /// <code>
        /// GetIntPropertyValue(part, "IJUAhsHeight1","Height1",ref propertyValue)
        /// </code>
        public static void GetIntPropertyValue(BusinessObject part, string interfaceName, string propertyName, ref int propertyValue)
        {
            if (part.SupportsInterface(interfaceName))
            {
                try
                {
                    propertyValue = (int)((PropertyValueInt)part.GetPropertyValue(interfaceName, propertyName)).PropValue;
                }
                catch
                {
                    propertyValue = 0;
                }
            }
        }
        /// <summary>
        /// This method will be called to check property is available on the supportcomponent and support return its value.
        /// </summary>
        /// <param name="part">supportcomponent or support is of type bussiness object</param>
        /// <param name="interfaceName">Interface to which the property belongs.</param>
        /// <param name="propertyName">Name of property.</param>
        /// <param name="propertyValue">Returns Property Value</param>
        /// <returns></returns>
        /// <code>
        /// GetDoublePropertyValue(part, "IJUAhsHeight1","Height1",ref propertyValue)
        /// </code>
        public static void GetDoublePropertyValue(BusinessObject part, string interfaceName, string propertyName, ref double propertyValue)
        {
            if (part.SupportsInterface(interfaceName))
            {
                try
                {
                    propertyValue = (double)((PropertyValueDouble)part.GetPropertyValue(interfaceName, propertyName)).PropValue;
                }
                catch
                {
                    propertyValue = 0;
                }
            }
        }
        /// <summary>
        /// Looks up all the relevent data for the steel cross section of the supporting object
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="index"> The index of the supporting object you want to look up the cross sectional data of</param>
        /// <returns>The cross section information in the form of a hsSteelMember user defined type.</returns>
        ///  /// <code>
        /// GetSupportingSectionData(this,1)
        /// </code>
        public static HSSteelMember GetSupportingSectionData(CustomSupportDefinition customSupportDefinition, int index)
        {
            try
            {
                HSSteelMember resultMember = new HSSteelMember();
                string supportingType = String.Empty;

                if ((customSupportDefinition.SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member))
                        supportingType = "Steel";    //Steel                      
                    else if(customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                        supportingType = "HangerBeam";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                
                if (supportingType == "Steel")
                {
                    resultMember.sectionName = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).SectionName;
                    resultMember.sectionType = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).SectionType;
                    resultMember.webThickness = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).WebThickness;
                    resultMember.width = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).Width;
                    resultMember.depth = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).Depth;
                    resultMember.flangeThickness = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).FlangeThickness;
                    resultMember.flangeWidth = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).FlangeWidth;
                    resultMember.webDepth = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).WebDepth;
                }
                else if (supportingType == "HangerBeam")
                    resultMember = GetSectionDataFromPart(customSupportDefinition.SupportHelper.Support.SupportDefinition.PartNumber);

                return resultMember;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetSupportingSectionData." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method is used to Create the BoundingBox according to the frameDimensions.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="boundingBoxName">name of the BoundingBox</param>
        /// <param name="bbxOrientation">direction of the BoundingBox</param>
        /// <param name="includeInsulation"> insulation option</param>
        /// <param name="mirrorFlag">mirror configuration</param>
        /// <param name="isPipeVertical">boolean ispipevertical or not</param>
        ///<code>
        ///CreateVerticalFrameBoundingBox(this, boundingBoxName,(FrameAssemblyServices.FrameBBXOrientation)frameOrientation, includeInsulation, mirrorFlag, SupportedHelper.IsSupportedObjectVertical(1, 45))
        ///</code>
        public static void CreateVerticalFrameBoundingBox(CustomSupportDefinition customSupportDefinition, string boundingBoxName, FrameBBXOrientation frameBBXOrientation, Boolean includeInsulation, Boolean mirrorFlag, Boolean rotate90Deg)
        {
            try
            {
                // Create Vectors to define the plane of the BBX
                Vector GlobalX = new Vector(0, 0, 0), GlobalY = new Vector(0, 0, 0), BB_X = new Vector(0, 0, 0), BB_Z = new Vector(0, 0, 0);

                // A couple special vectors for the tangent orientation
                // Tangent Enum Doesn't work well for different sized pipes, so for tangent
                // orientation I will grab the Y-Vector from the Tangent_BBR

                Matrix4X4 tangentBBPort = new Matrix4X4();
                Vector tangentX = new Vector(0, 0, 0), tangentY = new Vector(0, 0, 0), tangentZ = new Vector(0, 0, 0);

                Boolean insulationOption;
                if (includeInsulation == true)
                    insulationOption = true;
                else
                    insulationOption = false;
                customSupportDefinition.BoundingBoxHelper.CreateStandardBoundingBoxes(insulationOption);

                switch (frameBBXOrientation)
                {
                    case FrameBBXOrientation.FrameBBXOrientation_Direct:
                        {
                            // If Equipment need to do something different
                            BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ); // Use GlobalZ as BBX X-Axis
                            BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteToStructure);

                            if (HgrCompareDoubleService.cmpdbl(Math.Round((AngleBetweenVectors(BB_X, BB_Z) * 180 / Math.PI), 0) ,0) == true || HgrCompareDoubleService.cmpdbl(Math.Round((AngleBetweenVectors(BB_X, BB_Z) * 180 / Math.PI),0),180)==true)
                                // Vector from Route To Struct is Parallel the BB_X
                                // Use Vector Perpendicular to Route
                                BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X);
                            else
                                // Use Vector from Route to Struct
                                BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteToStructure, BB_X);

                            if (rotate90Deg == true)
                                BB_Z = BB_Z.Cross(BB_X);

                            customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                        }
                        break;
                    case FrameBBXOrientation.FrameBBXOrientation_Orthogonal:
                        {
                            BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ); // Use GlobalZ as BBX X-Axis
                            BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteToStructure); // Project vector from Route to Struct into plane of BBX

                            if (HgrCompareDoubleService.cmpdbl(Math.Round((AngleBetweenVectors(BB_X, BB_Z) * 180 / Math.PI), 0),0)==true || HgrCompareDoubleService.cmpdbl(Math.Round((AngleBetweenVectors(BB_X, BB_Z) * 180 / Math.PI), 0),180)==true)
                                // Vector from Route To Struct is Parallel the BB_X
                                // Use Vector Perpendicular to Route
                                BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X);
                            else
                                // Use Vector from Route to Struct
                                BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteToStructure, BB_X);

                            GlobalX = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalX, BB_X); // Project GlobalX into plane of BBX
                            GlobalY = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalY, BB_X); // Project GlobalY into plane of BBX
                            // Get Orthogonal Vector in the plane of the BBX (Gx, -Gx, Gy or -Gy) depending on angle
                            if (AngleBetweenVectors(GlobalX, BB_Z) < (Math.Atan(1) * 4.0) / 4)
                                BB_Z = GlobalX;
                            else if (AngleBetweenVectors(GlobalX, BB_Z) > 3 * (Math.Atan(1) * 4.0) / 4)
                                BB_Z.Set(-GlobalX.X, -GlobalX.Y, -GlobalX.Z);
                            else if (AngleBetweenVectors(GlobalY, BB_Z) < (Math.Atan(1) * 4.0) / 4)
                                BB_Z = GlobalY;
                            else
                                BB_Z.Set(-GlobalY.X, -GlobalY.Y, -GlobalY.Z);

                            if (rotate90Deg == true)
                                BB_Z = BB_Z.Cross(BB_X);

                            customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                        }
                        break;
                    case FrameBBXOrientation.FrameBBXOrientation_Tangent:
                        {
                            if (customSupportDefinition.Configuration == 1 || customSupportDefinition.Configuration == 3)
                                tangentBBPort = customSupportDefinition.RefPortHelper.PortLCS("TANGENT_BBR_Low");
                            else
                                tangentBBPort = customSupportDefinition.RefPortHelper.PortLCS("TANGENT_ALT_BBR_Low");

                            tangentX.Set(tangentBBPort.XAxis.X, tangentBBPort.XAxis.Y, tangentBBPort.XAxis.Z);
                            tangentZ.Set(tangentBBPort.ZAxis.X, tangentBBPort.ZAxis.Y, tangentBBPort.ZAxis.Z);
                            tangentY = tangentZ.Cross(tangentX);

                            BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);
                            BB_Z = ProjectVectorIntoPlane(tangentY, BB_X); // Use Vector Parallel to Pipes as BB_Z

                            if (rotate90Deg == true)
                                BB_Z = BB_Z.Cross(BB_X);

                            customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                        }
                        break;
                }
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in CreateVerticalFrameBoundingBox.." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// Used to determine the values that should be used to set the Cardinal Point, OverLength,
        /// Rotation and Reflection of smart steel sections where they are connected to other smart
        /// steel section.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="steel1PartKey">Stores the partkey of first steel part- String</param>
        /// <param name="steel1Port">Stores the port of first steel part-String</param>
        /// <param name="steel1Data">storing the orientation data of the first steel part -HSSteelConfig</param>
        /// <param name="overLengthSteel1">An overlength adjustment to be added to the OverLength of the first steel part-Double</param>
        /// <param name="steel2PartKey">Stores the partkey of second steel part- String</param>
        /// <param name="steel2Port">Stores the port of second steel part-String</param>
        /// <param name="steel2Data">storing the orientation data of the second steel part -HSSteelConfig</param>
        /// <param name="overLengthSteel2">An overlength adjustment to be added to the OverLength of the second steel part-Double</param>
        /// <param name="angle">gets angle-SteelConnectionAngle</param>
        /// <param name="axialOffset">sets the axials offset value-double</param>
        /// <param name="connectionType">gets the type of connection-SteelConnection</param>
        /// <param name="mirrorConnection">gets the mirror connection-boolean</param>
        /// <param name="steelJoint">gets the steel joint type-SteelJointType</param>
        /// <param name="weldDataCollection">Returns the collection of welds- Collection<WeldData></param>
        /// <code>
        /// FrameAssemblyServices.SetSteelConnection(LEG, "BeginCap",ref legData, leg1BeginOverHang, SECTION, "BeginCap", sectionData, 0, (FrameAssemblyServices.SteelConnectionAngle)90, -axialOffset, (FrameAssemblyServices.SteelConnection)connection1Type, connection1Mirror, FrameAssemblyServices.SteelJointType.SteelJoint_RIGID,frameConnection1Welds);
        /// </code>
        public static void SetSteelConnection(CustomSupportDefinition customSupportDefinition,string steel1PartKey, string steel1Port, ref HSSteelConfig steel1Data, double overLengthSteel1, string steel2PartKey, string steel2Port, ref HSSteelConfig steel2Data, double overLengthSteel2, SteelConnectionAngle angle, double axialOffset, SteelConnection connectionType, Boolean mirrorConnection = false, SteelJointType steelJoint = SteelJointType.SteelJoint_RIGID, Collection<WeldData> weldDataCollection = null)
        {
            try
            {
            int steel1CP = 5, steel2CP = 5;
            double steel1OffsetX = 0, steel1OffsetY = 0, steel1OverLength = 0, steel2OffsetX = 0, steel2OffsetY = 0, steel2OverLength = 0;
            HSSteelMember steel1Member, steel2Member;
            steel1Member = GetSectionDataFromPartIndex(customSupportDefinition, steel1PartKey);
            steel2Member = GetSectionDataFromPartIndex(customSupportDefinition, steel2PartKey);
            string steel1Type = string.Empty, steel2Type = string.Empty;
            Dictionary<string, SupportComponent> componentDictionary = customSupportDefinition.SupportHelper.SupportComponentDictionary;
            PropertyValueCodelist endCutbackAnchorPoint ;
            PropertyValueCodelist beginCutbackAnchorPoint;
            string strIJOAhsCutback = "";

            if (componentDictionary[steel1PartKey].SupportsInterface("IJOAhsCutback"))
                strIJOAhsCutback = "IJOAhsCutback";
            else
                strIJOAhsCutback = "IJOAhsCutbak";

            endCutbackAnchorPoint = (PropertyValueCodelist)componentDictionary[steel1PartKey].GetPropertyValue(strIJOAhsCutback, "EndCutbackAnchorPoint");
            beginCutbackAnchorPoint = (PropertyValueCodelist)componentDictionary[steel1PartKey].GetPropertyValue(strIJOAhsCutback, "BeginCutbackAnchorPoint");
            CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
            PartClass hangerSteelType = (PartClass)cataloghelper.GetPartClass("HgrSteelTypes");
            ReadOnlyCollection<BusinessObject> partsHangerSteelType = hangerSteelType.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            foreach (BusinessObject part in partsHangerSteelType)
            {
                if ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsSteelTypes", "SteelType")).PropValue == steel1Member.sectionType)
                    steel1Type = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsSteelTypes", "StdSteelType")).PropValue;
                if ((string)((PropertyValueString)part.GetPropertyValue("IJUAhsSteelTypes", "SteelType")).PropValue == steel2Member.sectionType)
                {
                    steel2Type = (string)((PropertyValueString)part.GetPropertyValue("IJUAhsSteelTypes", "StdSteelType")).PropValue;
                    break;
                }
            }
            //NOTE: Reflect functionality on Rich Steel will not be implemented for v2011. Therefore, all lines
            //that deal with the reflected cross sections have been commented out. If in the future a reflect
            //toggle gets added to Rich steel, then we can uncomment these lines and extend the SetSteelConnection
            //function to work with reflected cross sections.
            if (HgrCompareDoubleService.cmpdbl(Math.Round(axialOffset, 5),0)==true && steelJoint == SteelJointType.SteelJoint_RIGID)
            {
                switch (steel1Data.Orient)
                {
                    case 0:
                        steel2OverLength = steel1Member.depth / 2;
                        break;
                    case (SteelOrientationAngle)90:
                        steel2OverLength = steel1Member.width / 2;
                        break;
                    case (SteelOrientationAngle)180:
                        steel2OverLength = steel1Member.depth / 2;
                        break;
                    case (SteelOrientationAngle)270:
                        steel2OverLength = steel1Member.width / 2;
                        break;
                }
            }
            else
                steel2OverLength = 0;
            //Steel Section 1
            switch (connectionType)
            {
                case SteelConnection.SteelConnection_Butted:
                    switch (steel2Type)
                    {
                        case "W":
                            if (steel2Data.Orient == (SteelOrientationAngle)90 || steel2Data.Orient == (SteelOrientationAngle)270)
                                steel1OverLength = -steel2Member.webThickness / 2;
                            else
                                steel1OverLength = -steel2Member.depth / 2;
                            break;
                        case "L":
                            if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                            {
                                switch (steel2Data.Orient)
                                {
                                    case 0:
                                        steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                        break;
                                    case (SteelOrientationAngle)90:
                                        steel1OverLength = -steel2Member.width / 2;
                                        break;
                                    case (SteelOrientationAngle)180:
                                        steel1OverLength = -steel2Member.depth / 2;
                                        break;
                                    case (SteelOrientationAngle)270:
                                        steel1OverLength = steel2Member.width / 2 - steel2Member.flangeThickness;
                                        break;
                                }
                            }
                            else
                            {
                                switch (steel2Data.Orient)
                                {
                                    case 0:
                                        steel1OverLength = -steel2Member.depth / 2;
                                        break;
                                    case (SteelOrientationAngle)90:
                                        steel1OverLength = -steel2Member.width / 2 - steel2Member.flangeThickness;
                                        break;
                                    case (SteelOrientationAngle)180:
                                        steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                        break;
                                    case (SteelOrientationAngle)270:
                                        steel1OverLength = -steel2Member.width / 2;
                                        break;
                                }
                            }
                            break;
                        case "C":
                            if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                            {
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                    steel1OverLength = -steel2Member.depth / 2;
                                else if (steel2Data.Orient == (SteelOrientationAngle)90)
                                    steel1OverLength = -steel2Member.width / 2;
                                else
                                    steel1OverLength = steel2Member.width / 2 - steel2Member.webThickness;
                            }
                            else
                            {
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                    steel1OverLength = -steel2Member.depth / 2;
                                else if (steel2Data.Orient == (SteelOrientationAngle)90)
                                    steel1OverLength = steel2Member.width / 2 - steel2Member.webThickness;
                                else
                                    steel1OverLength = -steel2Member.width / 2;
                            }
                            break;
                        case "T":
                            if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                steel1OverLength = -steel2Member.depth / 2;
                            else if (steel2Data.Orient == (SteelOrientationAngle)90 || steel2Data.Orient == (SteelOrientationAngle)270)
                                steel1OverLength = -steel2Member.webThickness / 2;
                            break;
                        case "HSSR":
                        case "2C":
                        case "2L":
                        case "PIPE":
                            if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                steel1OverLength = -steel2Member.depth / 2;
                            else
                                steel1OverLength = -steel2Member.width / 2;
                            break;
                        default:
                            if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                steel1OverLength = -steel2Member.depth / 2;
                            else
                                steel1OverLength = -steel2Member.width / 2;
                            break;
                    }
                    break;

                case SteelConnection.SteelConnection_Lapped:
                    if (mirrorConnection)
                    {
                        switch (steel1Data.Orient)
                        {
                            case 0:
                                steel1CP = 4;
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                {
                                    steel1OverLength = steel2Member.depth / 2;
                                    steel1OffsetX = -steel2Member.width / 2;
                                }
                                else
                                {
                                    steel1OverLength = steel2Member.width / 2;
                                    steel1OffsetX = -steel2Member.depth / 2;
                                }
                                break;
                            case (SteelOrientationAngle)90:
                                steel1CP = 2;
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                {
                                    steel1OverLength = steel2Member.depth / 2;
                                    steel1OffsetY = -steel2Member.width / 2;
                                }
                                else
                                {
                                    steel1OverLength = steel2Member.width / 2;
                                    steel1OffsetY = -steel2Member.depth / 2;
                                }
                                break;
                            case (SteelOrientationAngle)180:
                                steel1CP = 6;
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                {
                                    steel1OverLength = steel2Member.depth / 2;
                                    steel1OffsetX = steel2Member.width / 2;
                                }
                                else
                                {
                                    steel1OverLength = steel2Member.width / 2;
                                    steel1OffsetX = steel2Member.depth / 2;
                                }
                                break;
                            case (SteelOrientationAngle)270:
                                steel1CP = 8;
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                {
                                    steel1OverLength = steel2Member.depth / 2;
                                    steel1OffsetY = steel2Member.width / 2;
                                }
                                else
                                {
                                    steel1OverLength = steel2Member.width / 2;
                                    steel1OffsetY = steel2Member.depth / 2;
                                }
                                break;
                        }
                    }
                    else
                    {
                        switch (steel1Data.Orient)
                        {
                            case 0:
                                steel1CP = 6;
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                {
                                    steel1OverLength = steel2Member.depth / 2;
                                    steel1OffsetX = steel2Member.width / 2;
                                }
                                else
                                {
                                    steel1OverLength = steel2Member.width / 2;
                                    steel1OffsetX = steel2Member.depth / 2;
                                }
                                break;
                            case (SteelOrientationAngle)90:
                                steel1CP = 8;
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                {
                                    steel1OverLength = steel2Member.depth / 2;
                                    steel1OffsetY = steel2Member.width / 2;
                                }
                                else
                                {
                                    steel1OverLength = steel2Member.width / 2;
                                    steel1OffsetY = steel2Member.depth / 2;
                                }
                                break;
                            case (SteelOrientationAngle)180:
                                steel1CP = 4;
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                {
                                    steel1OverLength = steel2Member.depth / 2;
                                    steel1OverLength = steel2Member.depth / 2;
                                    steel1OffsetX = -steel2Member.width / 2;
                                }
                                else
                                {
                                    steel1OverLength = steel2Member.width / 2;
                                    steel1OffsetX = -steel2Member.depth / 2;
                                }
                                break;
                            case (SteelOrientationAngle)270:
                                steel1CP = 2;
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                {
                                    steel1OverLength = steel2Member.depth / 2;
                                    steel1OffsetY = -steel2Member.width / 2;
                                }
                                else
                                {
                                    steel1OverLength = steel2Member.width / 2;
                                    steel1OffsetY = -steel2Member.depth / 2;
                                }
                                break;
                        }
                    }
                    break;
                case SteelConnection.SteelConnection_Nested:
                    {
                        switch (steel2Type)
                        {
                            case "W":
                                {
                                    switch (steel2Data.Orient)
                                    {
                                        case 0:
                                            steel1OverLength = -steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)90:
                                            steel1OverLength = -steel2Member.webThickness / 2;
                                            if (mirrorConnection)
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 4;
                                                        steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 2;
                                                        steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 6;
                                                        steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 8;
                                                        steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                }
                                            }
                                            else
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 6;
                                                        steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 8;
                                                        steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 4;
                                                        steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 2;
                                                        steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                }
                                            }
                                            break;
                                        case (SteelOrientationAngle)180:
                                            steel1OverLength = -steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)270:
                                            steel1OverLength = -steel2Member.webThickness / 2;
                                            if (mirrorConnection)
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 4;
                                                        steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 2;
                                                        steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 6;
                                                        steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 8;
                                                        steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                }
                                            }
                                            else
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 6;
                                                        steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 8;
                                                        steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 4;
                                                        steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 2;
                                                        steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                }
                                            }
                                            break;
                                    }
                                }
                                break;
                            case "L":
                                switch (steel2Data.Orient)
                                {
                                    case 0:
                                        if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                        {
                                            steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                            switch (steel1Data.Orient)
                                            {
                                                case 0:
                                                    steel1CP = 4;
                                                    steel1OffsetX = steel2Member.width / 2 - steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)90:
                                                    steel1CP = 2;
                                                    steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)180:
                                                    steel1CP = 6;
                                                    steel1OffsetX = -steel2Member.width / 2 + steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)270:
                                                    steel1CP = 8;
                                                    steel1OffsetY = -steel2Member.width / 2 + steel2Member.flangeThickness;
                                                    break;
                                            }

                                        }
                                        else
                                            steel1OverLength = -steel2Member.depth / 2;
                                        break;
                                    case (SteelOrientationAngle)90:
                                        if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                            steel1OverLength = -steel2Member.width / 2;
                                        else
                                        {
                                            steel1OverLength = steel2Member.width / 2 - steel2Member.flangeThickness;
                                            switch (steel1Data.Orient)
                                            {
                                                case 0:
                                                    steel1CP = 4;
                                                    steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)90:
                                                    steel1CP = 2;
                                                    steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)180:
                                                    steel1CP = 6;
                                                    steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)270:
                                                    steel1CP = 8;
                                                    steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                    break;
                                            }
                                        }
                                        break;
                                    case (SteelOrientationAngle)180:
                                        if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                            steel1OverLength = -steel2Member.depth / 2;
                                        else
                                        {
                                            steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                            switch (steel1Data.Orient)
                                            {
                                                case 0:
                                                    steel1CP = 6;
                                                    steel1OffsetX = -steel2Member.width / 2 + steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)90:
                                                    steel1CP = 8;
                                                    steel1OffsetY = -steel2Member.width / 2 + steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)180:
                                                    steel1CP = 4;
                                                    steel1OffsetX = steel2Member.width / 2 - steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)270:
                                                    steel1CP = 2;
                                                    steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                    break;
                                            }
                                        }
                                        break;
                                    case (SteelOrientationAngle)270:
                                        if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                        {
                                            steel1OverLength = steel2Member.width / 2 - steel2Member.flangeThickness;
                                            switch (steel1Data.Orient)
                                            {
                                                case 0:
                                                    steel1CP = 6;
                                                    steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)90:
                                                    steel1CP = 8;
                                                    steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)180:
                                                    steel1CP = 4;
                                                    steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                    break;
                                                case (SteelOrientationAngle)270:
                                                    steel1CP = 2;
                                                    steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                    break;
                                            }
                                        }
                                        else
                                            steel1OverLength = -steel2Member.width / 2;
                                        break;
                                }
                                break;
                            case "C":
                                switch (steel2Data.Orient)
                                {
                                    case 0:
                                        steel1OverLength = -steel2Member.depth / 2;
                                        break;
                                    case (SteelOrientationAngle)90:
                                        if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                            steel1OverLength = -steel2Member.width / 2;
                                        else
                                        {
                                            steel1OverLength = steel2Member.width / 2 - steel2Member.webThickness;
                                            if (mirrorConnection)
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 4;
                                                        steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 2;
                                                        steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 6;
                                                        steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 8;
                                                        steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                }
                                            }
                                            else
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 6;
                                                        steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 8;
                                                        steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 4;
                                                        steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 2;
                                                        steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                }
                                            }
                                        }
                                        break;
                                    case (SteelOrientationAngle)180:
                                        steel1OverLength = -steel2Member.depth / 2;
                                        break;
                                    case (SteelOrientationAngle)270:
                                        if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                        {
                                            steel1OverLength = steel2Member.width / 2 - steel2Member.webThickness;
                                            if (mirrorConnection)
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 4;
                                                        steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 2;
                                                        steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 6;
                                                        steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 8;
                                                        steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                }
                                            }
                                            else
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 6;
                                                        steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 8;
                                                        steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 4;
                                                        steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 2;
                                                        steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                        break;
                                                }
                                            }
                                        }
                                        else
                                            steel1OverLength = -steel2Member.width / 2;
                                        break;
                                }
                                break;
                            case "T":
                                switch (steel2Data.Orient)
                                {
                                    case 0:
                                        if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                            steel1OverLength = -steel2Member.depth / 2;
                                        else
                                        {
                                            steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                            if (mirrorConnection)
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 4;
                                                        steel1OffsetX = -steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 2;
                                                        steel1OffsetY = -steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 6;
                                                        steel1OffsetX = steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 8;
                                                        steel1OffsetY = steel2Member.webThickness / 2;
                                                        break;
                                                }
                                            }
                                            else
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 6;
                                                        steel1OffsetX = steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 8;
                                                        steel1OffsetY = steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 4;
                                                        steel1OffsetX = -steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 2;
                                                        steel1OffsetY = -steel2Member.webThickness / 2;
                                                        break;
                                                }
                                            }
                                        }
                                        break;
                                    case (SteelOrientationAngle)90:
                                        steel1OverLength = -steel2Member.webThickness / 2;
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                break;
                                        }
                                        break;
                                    case (SteelOrientationAngle)180:
                                        if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                        {
                                            steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                            if (mirrorConnection)
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 4;
                                                        steel1OffsetX = -steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 2;
                                                        steel1OffsetY = -steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 6;
                                                        steel1OffsetX = steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 8;
                                                        steel1OffsetY = steel2Member.webThickness / 2;
                                                        break;
                                                }
                                            }
                                            else
                                            {
                                                switch (steel1Data.Orient)
                                                {
                                                    case 0:
                                                        steel1CP = 6;
                                                        steel1OffsetX = steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)90:
                                                        steel1CP = 8;
                                                        steel1OffsetY = steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)180:
                                                        steel1CP = 4;
                                                        steel1OffsetX = -steel2Member.webThickness / 2;
                                                        break;
                                                    case (SteelOrientationAngle)270:
                                                        steel1CP = 2;
                                                        steel1OffsetY = -steel2Member.webThickness / 2;
                                                        break;
                                                }
                                            }
                                            break;
                                        }
                                        else
                                            steel1OverLength = -steel2Member.depth / 2;
                                        break;
                                    case (SteelOrientationAngle)270:
                                        steel1OverLength = -steel2Member.webThickness / 2;
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.depth / 2 + steel2Member.flangeThickness;
                                                break;
                                        }
                                        break;
                                }
                                break;
                            case "HSSR":
                            case "2C":
                            case "2L":
                            case "PIPE":
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                    steel1OverLength = -steel2Member.depth / 2;
                                else
                                    steel1OverLength = -steel2Member.width / 2;
                                break;
                            default:
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                    steel1OverLength = -steel2Member.depth / 2;
                                else
                                    steel1OverLength = -steel2Member.width / 2;
                                break;
                        }
                    }
                    break;
                case SteelConnection.SteelConnection_Coped:
                    switch (steel2Type)
                    {
                        case "W":
                            switch (steel2Data.Orient)
                            {
                                case 0:
                                    steel1OverLength = -steel2Member.depth / 2;
                                    break;
                                case (SteelOrientationAngle)90:
                                    steel1OverLength = -steel2Member.webThickness / 2;
                                    if (mirrorConnection)
                                    {
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.depth / 2;
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.depth / 2;
                                                break;
                                        }
                                    }
                                    break;
                                case (SteelOrientationAngle)180:
                                    steel1OverLength = -steel2Member.depth / 2;
                                    break;
                                case (SteelOrientationAngle)270:
                                    steel1OverLength = -steel2Member.webThickness / 2;
                                    if (mirrorConnection)
                                    {
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.depth / 2;
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.depth / 2;
                                                break;
                                        }
                                    }
                                    break;
                            }
                            break;
                        case "L":
                            switch (steel2Data.Orient)
                            {
                                case 0:
                                    if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                    {
                                        steel1OverLength = steel2Member.depth / 2;
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.width / 2;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.width / 2;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.width / 2;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.width / 2;
                                                break;
                                        }
                                    }
                                    else
                                        steel1OverLength = -steel2Member.depth / 2;
                                    break;
                                case (SteelOrientationAngle)90:
                                    if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                        steel1OverLength = -steel2Member.width / 2;
                                    else
                                        steel1OverLength = steel2Member.width / 2;
                                    switch (steel1Data.Orient)
                                    {
                                        case 0:
                                            steel1CP = 4;
                                            steel1OffsetX = steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)90:
                                            steel1CP = 2;
                                            steel1OffsetY = steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)180:
                                            steel1CP = 6;
                                            steel1OffsetX = -steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)270:
                                            steel1CP = 8;
                                            steel1OffsetY = -steel2Member.depth / 2;
                                            break;
                                    }
                                    break;
                                case (SteelOrientationAngle)180:
                                    if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                        steel1OverLength = -steel2Member.depth / 2;
                                    else
                                    {
                                        steel1OverLength = steel2Member.depth / 2;
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.width / 2;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.width / 2;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.width / 2;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.width / 2;
                                                break;
                                        }
                                    }
                                    break;
                                case (SteelOrientationAngle)270:
                                    if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                    {
                                        steel1OverLength = steel2Member.width / 2;
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.depth / 2;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.depth / 2;
                                                break;
                                        }
                                    }
                                    else
                                        steel1OverLength = -steel2Member.width / 2;
                                    break;
                            }
                            break;
                        case "C":
                            switch (steel2Data.Orient)
                            {
                                case 0:
                                    steel1OverLength = -steel2Member.depth / 2;
                                    break;
                                case (SteelOrientationAngle)90:
                                    if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                        steel1OverLength = -steel2Member.width / 2;
                                    else
                                    {
                                        steel1OverLength = steel2Member.width / 2;
                                        if (mirrorConnection)
                                        {
                                            switch (steel1Data.Orient)
                                            {
                                                case 0:
                                                    steel1CP = 4;
                                                    steel1OffsetX = steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)90:
                                                    steel1CP = 2;
                                                    steel1OffsetY = steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)180:
                                                    steel1CP = 6;
                                                    steel1OffsetX = -steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)270:
                                                    steel1CP = 8;
                                                    steel1OffsetY = -steel2Member.depth / 2;
                                                    break;
                                            }
                                        }
                                        else
                                        {
                                            switch (steel1Data.Orient)
                                            {
                                                case 0:
                                                    steel1CP = 6;
                                                    steel1OffsetX = -steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)90:
                                                    steel1CP = 8;
                                                    steel1OffsetY = -steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)180:
                                                    steel1CP = 4;
                                                    steel1OffsetX = steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)270:
                                                    steel1CP = 2;
                                                    steel1OffsetY = steel2Member.depth / 2;
                                                    break;
                                            }
                                        }
                                    }
                                    break;
                                case (SteelOrientationAngle)180:
                                    steel1OverLength = -steel2Member.depth / 2;
                                    break;
                                case (SteelOrientationAngle)270:
                                    if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                    {
                                        steel1OverLength = steel2Member.width / 2;
                                        if (mirrorConnection)
                                        {
                                            switch (steel1Data.Orient)
                                            {
                                                case 0:
                                                    steel1CP = 4;
                                                    steel1OffsetX = steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)90:
                                                    steel1CP = 2;
                                                    steel1OffsetY = steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)180:
                                                    steel1CP = 6;
                                                    steel1OffsetX = -steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)270:
                                                    steel1CP = 8;
                                                    steel1OffsetY = -steel2Member.depth / 2;
                                                    break;
                                            }
                                        }
                                        else
                                        {
                                            switch (steel1Data.Orient)
                                            {
                                                case 0:
                                                    steel1CP = 6;
                                                    steel1OffsetX = -steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)90:
                                                    steel1CP = 8;
                                                    steel1OffsetY = -steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)180:
                                                    steel1CP = 4;
                                                    steel1OffsetX = steel2Member.depth / 2;
                                                    break;
                                                case (SteelOrientationAngle)270:
                                                    steel1CP = 2;
                                                    steel1OffsetY = steel2Member.depth / 2;
                                                    break;
                                            }
                                        }
                                    }
                                    else
                                        steel1OverLength = -steel2Member.width / 2;
                                    break;
                            }
                            break;
                        case "T":
                            switch (steel2Data.Orient)
                            {
                                case 0:

                                    if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                        steel1OverLength = -steel2Member.depth / 2;
                                    else
                                        steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                    if (mirrorConnection)
                                    {
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.webThickness / 2;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.webThickness / 2;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.webThickness / 2;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.webThickness / 2;
                                                break;
                                        }
                                    }
                                    else
                                    {
                                        switch (steel1Data.Orient)
                                        {
                                            case 0:
                                                steel1CP = 6;
                                                steel1OffsetX = -steel2Member.webThickness / 2;
                                                break;
                                            case (SteelOrientationAngle)90:
                                                steel1CP = 8;
                                                steel1OffsetY = -steel2Member.webThickness / 2;
                                                break;
                                            case (SteelOrientationAngle)180:
                                                steel1CP = 4;
                                                steel1OffsetX = steel2Member.webThickness / 2;
                                                break;
                                            case (SteelOrientationAngle)270:
                                                steel1CP = 2;
                                                steel1OffsetY = steel2Member.webThickness / 2;
                                                break;
                                        }
                                    }
                                    break;
                                case (SteelOrientationAngle)90:
                                    steel1OverLength = -steel2Member.webThickness / 2;
                                    switch (steel1Data.Orient)
                                    {
                                        case 0:
                                            steel1CP = 6;
                                            steel1OffsetX = -steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)90:
                                            steel1CP = 8;
                                            steel1OffsetY = -steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)180:
                                            steel1CP = 4;
                                            steel1OffsetX = steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)270:
                                            steel1CP = 2;
                                            steel1OffsetY = steel2Member.depth / 2;
                                            break;
                                    }
                                    break;
                                case (SteelOrientationAngle)180:
                                    if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                    {
                                        steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                        if (mirrorConnection)
                                        {
                                            switch (steel1Data.Orient)
                                            {
                                                case 0:
                                                    steel1CP = 4;
                                                    steel1OffsetX = steel2Member.webThickness / 2;
                                                    break;
                                                case (SteelOrientationAngle)90:
                                                    steel1CP = 2;
                                                    steel1OffsetY = steel2Member.webThickness / 2;
                                                    break;
                                                case (SteelOrientationAngle)180:
                                                    steel1CP = 6;
                                                    steel1OffsetX = -steel2Member.webThickness / 2;
                                                    break;
                                                case (SteelOrientationAngle)270:
                                                    steel1CP = 8;
                                                    steel1OffsetY = -steel2Member.webThickness / 2;
                                                    break;
                                            }
                                        }
                                        else
                                        {
                                            switch (steel1Data.Orient)
                                            {
                                                case 0:
                                                    steel1CP = 6;
                                                    steel1OffsetX = -steel2Member.webThickness / 2;
                                                    break;
                                                case (SteelOrientationAngle)90:
                                                    steel1CP = 8;
                                                    steel1OffsetY = -steel2Member.webThickness / 2;
                                                    break;
                                                case (SteelOrientationAngle)180:
                                                    steel1CP = 4;
                                                    steel1OffsetX = steel2Member.webThickness / 2;
                                                    break;
                                                case (SteelOrientationAngle)270:
                                                    steel1CP = 2;
                                                    steel1OffsetY = steel2Member.webThickness / 2;
                                                    break;
                                            }
                                        }
                                    }
                                    else
                                        steel1OverLength = -steel2Member.depth / 2;
                                    break;
                                case (SteelOrientationAngle)270:
                                    steel1OverLength = -steel2Member.webThickness / 2;
                                    switch (steel1Data.Orient)
                                    {
                                        case 0:
                                            steel1CP = 4;
                                            steel1OffsetX = steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)90:
                                            steel1CP = 2;
                                            steel1OffsetY = steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)180:
                                            steel1CP = 6;
                                            steel1OffsetX = -steel2Member.depth / 2;
                                            break;
                                        case (SteelOrientationAngle)270:
                                            steel1CP = 8;
                                            steel1OffsetY = -steel2Member.depth / 2;
                                            break;
                                    }
                                    break;
                            }
                            break;
                        case "HSSR":
                        case "2C":
                        case "2L":
                        case "PIPE":
                        default:
                            if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                steel1OverLength = -steel2Member.depth / 2;
                            else
                                steel1OverLength = -steel2Member.width / 2;
                            break;
                    }
                    break;
                case SteelConnection.SteelConnection_Fitted:
                    switch (steel2Type)
                    {
                        case "W":
                            if (steel2Data.Orient == (SteelOrientationAngle)90 || steel2Data.Orient == (SteelOrientationAngle)270)
                                steel1OverLength = -steel2Member.webThickness / 2;
                            else
                                steel1OverLength = -steel2Member.depth / 2;
                            break;
                        case "L":
                            if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                            {
                                switch (steel2Data.Orient)
                                {
                                    case 0:
                                        steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                        break;
                                    case (SteelOrientationAngle)90:
                                        steel1OverLength = -steel2Member.width / 2;
                                        break;
                                    case (SteelOrientationAngle)180:
                                        steel1OverLength = -steel2Member.depth / 2;
                                        break;
                                    case (SteelOrientationAngle)270:
                                        steel1OverLength = steel2Member.width / 2 - steel2Member.flangeThickness;
                                        break;
                                }
                            }
                            else
                            {
                                switch (steel2Data.Orient)
                                {
                                    case 0:
                                        steel1OverLength = -steel2Member.depth / 2;
                                        break;
                                    case (SteelOrientationAngle)90:
                                        steel1OverLength = steel2Member.width / 2 - steel2Member.flangeThickness;
                                        break;
                                    case (SteelOrientationAngle)180:
                                        steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                        break;
                                    case (SteelOrientationAngle)270:
                                        steel1OverLength = -steel2Member.width / 2;
                                        break;
                                }
                            }
                            break;
                        case "C":
                            if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                            {
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                    steel1OverLength = -steel2Member.depth / 2;
                                else if (steel2Data.Orient == (SteelOrientationAngle)90)
                                    steel1OverLength = -steel2Member.width / 2;
                                else
                                    steel1OverLength = steel2Member.width / 2 - steel2Member.webThickness;
                            }
                            else
                            {
                                if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                    steel1OverLength = -steel2Member.depth / 2;
                                else if (steel2Data.Orient == (SteelOrientationAngle)90)
                                    steel1OverLength = steel2Member.width / 2 - steel2Member.webThickness;
                                else
                                    steel1OverLength = -steel2Member.width / 2;
                            }
                            break;
                        case "T":
                            if (steel2Data.Orient == 0)
                            {
                                if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                    steel1OverLength = -steel2Member.depth / 2;
                                else
                                    steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                            }
                            else if (steel2Data.Orient == (SteelOrientationAngle)180)
                            {
                                if ((angle == (SteelConnectionAngle)90 && steel1Port.Substring(0, 3).Equals("End")) || (angle == (SteelConnectionAngle)270 && steel1Port.Substring(0, 5).Equals("Begin")))
                                    steel1OverLength = steel2Member.depth / 2 - steel2Member.flangeThickness;
                                else
                                    steel1OverLength = -steel2Member.depth / 2;
                            }
                            else if (steel2Data.Orient == (SteelOrientationAngle)270)
                                steel1OverLength = -steel2Member.webThickness / 2;
                            break;
                        case "HSSR":
                        case "2C":
                        case "2L":
                        case "PIPE":
                        default:
                            if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                steel1OverLength = -steel2Member.depth / 2;
                            else
                                steel1OverLength = -steel2Member.width / 2;
                            break;
                    }
                    break;
                case SteelConnection.SteelConnection_Mitered:
                    PropertyValueCodelist steel1IndexHangerType = (PropertyValueCodelist)componentDictionary[steel1PartKey].GetPropertyValue("IJOAHsHgrBeamType", "HgrBeamType");
                    PropertyValueCodelist steel2IndexHangerType = (PropertyValueCodelist)componentDictionary[steel2PartKey].GetPropertyValue("IJOAHsHgrBeamType", "HgrBeamType");
                    if (steel1Type == "L")
                        componentDictionary[steel1PartKey].SetPropertyValue(steel1IndexHangerType.PropValue = 1, "IJOAHsHgrBeamType", "HgrBeamType");
                    if (steel2Type == "L")
                        componentDictionary[steel1PartKey].SetPropertyValue(steel2IndexHangerType.PropValue = 1, "IJOAHsHgrBeamType", "HgrBeamType");
                    //  break;

                    switch (steel1Data.Orient)
                    {
                        case 0:
                            steel2OverLength = steel1Member.depth / 2;
                            if (steel1Port.Substring(0, 3).Equals("End"))
                            {
                                if (steel2Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                            }
                            else
                            {
                                if (steel2Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                            }
                            break;
                        case (SteelOrientationAngle)90:
                            steel2OverLength = steel1Member.width / 2;
                            if (steel1Port.Substring(0, 3).Equals("End"))
                            {
                                if (steel2Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                            }
                            else
                            {
                                if (steel2Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                            }
                            break;
                        case (SteelOrientationAngle)180:
                            steel2OverLength = steel1Member.depth / 2;
                            if (steel1Port.Substring(0, 3).Equals("End"))
                            {
                                if (steel2Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                            }
                            else
                            {
                                if (steel2Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                            }
                            break;
                        case (SteelOrientationAngle)270:
                            steel2OverLength = steel1Member.width / 2;
                            if (steel1Port.Substring(0, 3).Equals("End"))
                            {
                                if (steel2Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                            }
                            else
                            {
                                if (steel2Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                            }
                            break;
                    }
                    switch (steel2Data.Orient)
                    {
                        case 0:
                            steel1OverLength = steel2Member.depth / 2;
                            if (steel2Port.Substring(0, 3).Equals("End"))
                            {
                                if (steel1Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                            }
                            else
                            {
                                if (steel1Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                            }
                            break;
                        case (SteelOrientationAngle)90:
                            steel1OverLength = steel2Member.width / 2;
                            if (steel2Port.Substring(0, 3).Equals("End"))
                            {
                                if (steel1Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 *(Math.Atan(1) * 4.0)/ 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                            }
                            else
                            {
                                if (steel1Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                            }
                            break;
                        case (SteelOrientationAngle)180:
                            steel1OverLength = steel2Member.depth / 2;
                            if (steel2Port.Substring(0, 3).Equals("End"))
                            {
                                if (steel1Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                            }
                            else
                            {
                                if (steel1Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                            }
                            break;
                        case (SteelOrientationAngle)270:
                            steel1OverLength = steel2Member.width / 2;
                            if (steel2Port.Substring(0, 3).Equals("End"))
                            {
                                if (steel1Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "EndCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackEndAngle");
                                    }
                                }
                            }
                            else
                            {
                                if (steel1Port.Substring(0, 3).Equals("End"))
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                                else
                                {
                                    if (angle == (SteelConnectionAngle)90)
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                    else
                                    {
                                        componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, strIJOAhsCutback, "BeginCutbackAnchorPoint");
                                        componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, strIJOAhsCutback, "CutbackBeginAngle");
                                    }
                                }
                            }
                            break;
                    }
                    break;
                default:
                    //Either no Connection is specified or it is UserDefined, user can specify connection directly
                    //by setting the CP, orientation and offset values on the hsSteelConfig member.

                    //Steel1CPReflect = Steel1Data.iCardinalPoint
                    steel1CP = steel1Data.CardinalPoint;
                    steel1OffsetX = steel1Data.OffsetX;
                    steel1OffsetY = steel1Data.OffsetY;
                    steel1OverLength = 0;

                    //Steel2CPReflect = Steel2Data.iCardinalPoint
                    steel2CP = steel2Data.CardinalPoint;
                    steel2OffsetX = steel2Data.OffsetX;
                    steel2OffsetY = steel2Data.OffsetY;
                    steel2OverLength = 0;
                    break;
            }
            if (connectionType != SteelConnection.SteelConnection_Mitered)
            {
                //Need to set the Cutback angles to zero, so they won't keep their value if the
                //connection type is changed from Mitered to another type
                if (steel1Port.Substring(0, 3).Equals("End"))
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(0), strIJOAhsCutback, "CutbackEndAngle");
                else
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(0), strIJOAhsCutback, "CutbackBeginAngle");

                if (steel2Port.Substring(0, 3).Equals("End"))
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(0), strIJOAhsCutback, "CutbackEndAngle");
                else
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(0), strIJOAhsCutback, "CutbackBeginAngle");
            }
            steel1Data.CardinalPoint = steel1CP;
            steel2Data.CardinalPoint = steel2CP;

            steel1Data.OffsetX = steel1OffsetX;
            steel1Data.OffsetY = steel1OffsetY;

            steel2Data.OffsetX = steel2OffsetX;
            steel2Data.OffsetY = steel2OffsetY;

            //Set the Attributes on Steel 1.
            PropertyValueCodelist CP1 = (PropertyValueCodelist)componentDictionary[steel1PartKey].GetPropertyValue("IJOAhsSteelCP", "CP1");
            PropertyValueCodelist CP2 = (PropertyValueCodelist)componentDictionary[steel1PartKey].GetPropertyValue("IJOAhsSteelCP", "CP2");
            PropertyValueCodelist CP6 = (PropertyValueCodelist)componentDictionary[steel1PartKey].GetPropertyValue("IJOAhsSteelCP", "CP6");
            PropertyValueCodelist CP7 = (PropertyValueCodelist)componentDictionary[steel1PartKey].GetPropertyValue("IJOAhsSteelCPFlexPort", "CP7");
            switch (steel1Port)
            {
                case "BeginCap":
                    componentDictionary[steel1PartKey].SetPropertyValue(CP1.PropValue = steel1Data.CardinalPoint, "IJOAhsSteelCP", "CP1");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(steel1Data.Orient) * Math.Atan(1) * 4.0 / 180), "IJOAhsBeginCap", "BeginCapRotZ");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1Data.OffsetX), "IJOAhsBeginCap", "BeginCapXOffset");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1Data.OffsetY), "IJOAhsBeginCap", "BeginCapYOffset");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1OverLength + overLengthSteel1), "IJUAHgrOccOverLength", "BeginOverLength");
                    break;
                case "EndCap":
                    componentDictionary[steel1PartKey].SetPropertyValue(CP2.PropValue = steel1Data.CardinalPoint, "IJOAhsSteelCP", "CP2");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(steel1Data.Orient) * Math.Atan(1) * 4.0 / 180), "IJOAhsEndCap", "EndCapRotZ");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1Data.OffsetX), "IJOAhsEndCap", "EndCapXOffset");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1Data.OffsetY), "IJOAhsEndCap", "EndCapYOffset");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1OverLength + overLengthSteel1), "IJUAHgrOccOverLength", "EndOverLength");
                    break;
                case "BeginFlex":
                    componentDictionary[steel1PartKey].SetPropertyValue(CP6.PropValue = steel1Data.CardinalPoint, "IJOAhsSteelCP", "CP6");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(steel1Data.Orient) * Math.Atan(1) * 4.0 / 180), "IJOAhsFlexPort", "FlexPortRotZ");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1Data.OffsetX), "IJOAhsFlexPort", "FlexPortXOffset");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1Data.OffsetY), "IJOAhsFlexPort", "FlexPortYOffset");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1OverLength + overLengthSteel1), "IJUAHgrOccOverLength", "BeginOverLength");
                    break;
                case "EndFlex":
                    componentDictionary[steel1PartKey].SetPropertyValue(CP7.PropValue = steel1Data.CardinalPoint , "IJOAhsSteelCP", "CP7");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(steel1Data.Orient) * Math.Atan(1) * 4.0 / 180), "IJOAhsEndFlexPort", "EndFlexPortRotZ");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1Data.OffsetX), "IJOAhsEndFlexPort", "EndFlexPortXOffset");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1Data.OffsetY), "IJOAhsEndFlexPort", "EndFlexPortYOffset");
                    componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1OverLength + overLengthSteel1), "IJUAHgrOccOverLength", "EndOverLength");
                    break;
            }
            //Set the Attributes on Steel 2
            CP1 = (PropertyValueCodelist)componentDictionary[steel2PartKey].GetPropertyValue("IJOAhsSteelCP", "CP1");
            CP2 = (PropertyValueCodelist)componentDictionary[steel2PartKey].GetPropertyValue("IJOAhsSteelCP", "CP2");
            CP6 = (PropertyValueCodelist)componentDictionary[steel2PartKey].GetPropertyValue("IJOAhsSteelCP", "CP6");
            CP7 = (PropertyValueCodelist)componentDictionary[steel2PartKey].GetPropertyValue("IJOAhsSteelCPFlexPort", "CP7");
            switch (steel2Port)
            {

                case "BeginCap":
                    componentDictionary[steel2PartKey].SetPropertyValue(CP1.PropValue = steel2Data.CardinalPoint, "IJOAhsSteelCP", "CP1");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(steel2Data.Orient) * Math.Atan(1) * 4.0 / 180), "IJOAhsBeginCap", "BeginCapRotZ");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2Data.OffsetX), "IJOAhsBeginCap", "BeginCapXOffset");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2Data.OffsetY), "IJOAhsBeginCap", "BeginCapYOffset");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2OverLength + overLengthSteel2), "IJUAHgrOccOverLength", "BeginOverLength");
                    break;
                case "EndCap":
                    componentDictionary[steel2PartKey].SetPropertyValue(CP2.PropValue = steel2Data.CardinalPoint, "IJOAhsSteelCP", "CP2");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(steel2Data.Orient) * Math.Atan(1) * 4.0 / 180), "IJOAhsEndCap", "EndCapRotZ");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2Data.OffsetX), "IJOAhsEndCap", "EndCapXOffset");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2Data.OffsetY), "IJOAhsEndCap", "EndCapYOffset");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2OverLength + overLengthSteel2), "IJUAHgrOccOverLength", "EndOverLength");
                    break;
                case "BeginFlex":
                    componentDictionary[steel2PartKey].SetPropertyValue(CP6.PropValue = steel2Data.CardinalPoint , "IJOAhsSteelCP", "CP6");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(steel2Data.Orient) * Math.Atan(1) * 4.0 / 180), "IJOAhsFlexPort", "FlexPortRotZ");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2Data.OffsetX), "IJOAhsFlexPort", "FlexPortXOffset");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2Data.OffsetY), "IJOAhsFlexPort", "FlexPortYOffset");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2OverLength + overLengthSteel2), "IJUAHgrOccOverLength", "BeginOverLength");
                    break;
                case "EndFlex":
                    componentDictionary[steel2PartKey].SetPropertyValue(CP7.PropValue = steel2Data.CardinalPoint, "IJOAhsSteelCP", "CP7");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(steel2Data.Orient) * Math.Atan(1) * 4.0 / 180), "IJOAhsEndFlexPort", "EndFlexPortRotZ");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2Data.OffsetX), "IJOAhsEndFlexPort", "EndFlexPortXOffset");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2Data.OffsetY), "IJOAhsEndFlexPort", "EndFlexPortYOffset");
                    componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(steel2OverLength + overLengthSteel2), "IJUAHgrOccOverLength", "EndOverLength");
                    break;
            }
            //WHENEVER THE BELOW CALLS ARE REQUIRED UNCOMMENT THEM.didn't find where these calls are required.
            switch (angle)
            {
                case (SteelConnectionAngle)90:
                    if (steelJoint == SteelJointType.SteelJoint_RIGID)
                        customSupportDefinition.JointHelper.CreateRigidJoint(steel2PartKey, steel2Port, steel1PartKey, steel1Port, Plane.XY, Plane.ZX, Axis.X, Axis.X, axialOffset, 0, 0);
                    else
                        customSupportDefinition.JointHelper.CreatePrismaticJoint(steel1PartKey, steel1Port, steel2PartKey, steel2Port, Plane.YZ, Plane.YZ, Axis.Y, Axis.Z, 0, 0);
                    break;
                case (SteelConnectionAngle)270:
                    if (steelJoint == SteelJointType.SteelJoint_RIGID)
                        customSupportDefinition.JointHelper.CreateRigidJoint(steel2PartKey, steel2Port, steel1PartKey, steel1Port, Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, axialOffset, 0, 0);
                    else
                        customSupportDefinition.JointHelper.CreatePrismaticJoint(steel1PartKey, steel1Port, steel2PartKey, steel2Port, Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeZ, 0, 0);
                    break;
            }
            //Add the Weld Objects
            if (weldDataCollection != null)
            {
                int count;
                WeldData weld;
                string steelPort;
                for (count = 0; count < weldDataCollection.Count; count++)
                {
                    weld = weldDataCollection[count];
                    if (steel1Port.Substring(0, 3).Equals("End"))
                        steelPort = "EndFace";
                    else
                        steelPort = "BeginFace";
                    if (connectionType == SteelConnection.SteelConnection_Lapped)
                    {
                        int sign;
                        steelPort = steel1Port;
                        if (mirrorConnection)
                            sign = 1;
                        else
                            sign = -1;

                        switch (weld.location)
                        {
                            case 2:
                                switch (steel1Data.Orient)
                                {
                                    case 0:
                                    case (SteelOrientationAngle)180:
                                        customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, -steel1Member.depth / 2 + weld.offsetYValue, sign * Math.Abs(steel1OffsetX) + weld.offsetXValue);
                                        break;
                                    case (SteelOrientationAngle)90:
                                    case (SteelOrientationAngle)270:
                                        customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, -steel1Member.width / 2 + weld.offsetYValue, sign * Math.Abs(steel1OffsetY) + weld.offsetXValue);
                                        break;
                                }
                                break;
                            case 4:
                                switch (steel1Data.Orient)
                                {
                                    case 0:
                                    case (SteelOrientationAngle)180:
                                        if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                            customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, -steel2Member.depth / 2 + weld.offsetZValue, weld.offsetYValue, sign * Math.Abs(steel1OffsetX) + weld.offsetXValue);
                                        else
                                            customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, -steel2Member.width / 2 + weld.offsetZValue, weld.offsetYValue, sign * Math.Abs(steel1OffsetX) + weld.offsetXValue);
                                        break;
                                    case (SteelOrientationAngle)90:
                                    case (SteelOrientationAngle)270:
                                        if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                            customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, -steel2Member.depth / 2 + weld.offsetZValue, weld.offsetYValue, sign * Math.Abs(steel1OffsetY) + weld.offsetXValue);
                                        else
                                            customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, -steel2Member.width / 2 + weld.offsetZValue, weld.offsetYValue, sign * Math.Abs(steel1OffsetY) + weld.offsetXValue);
                                        break;
                                }
                                break;
                            case 6:
                                switch (steel1Data.Orient)
                                {
                                    case 0:
                                    case (SteelOrientationAngle)180:
                                        if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                            customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, steel2Member.depth / 2 + weld.offsetZValue, weld.offsetYValue, sign * Math.Abs(steel1OffsetX) + weld.offsetXValue);
                                        else
                                            customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, steel2Member.width / 2 + weld.offsetZValue, weld.offsetYValue, sign * Math.Abs(steel1OffsetX) + weld.offsetXValue);
                                        break;
                                    case (SteelOrientationAngle)90:
                                    case (SteelOrientationAngle)270:
                                        if (steel2Data.Orient == 0 || steel2Data.Orient == (SteelOrientationAngle)180)
                                            customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, steel2Member.depth / 2 + weld.offsetZValue, weld.offsetYValue, sign * Math.Abs(steel1OffsetY) + weld.offsetXValue);
                                        else
                                            customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, steel2Member.width / 2 + weld.offsetZValue, weld.offsetYValue, sign * Math.Abs(steel1OffsetY) + weld.offsetXValue);
                                        break;
                                }
                                break;
                            case 8:
                                switch (steel1Data.Orient)
                                {
                                    case 0:
                                    case (SteelOrientationAngle)180:
                                        customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, steel1Member.depth / 2, sign * Math.Abs(steel1OffsetX));
                                        break;
                                    case (SteelOrientationAngle)90:
                                    case (SteelOrientationAngle)270:
                                        customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, steel1Member.width / 2, sign * Math.Abs(steel1OffsetY));
                                        break;
                                }
                                break;
                        }
                    }
                    else if (connectionType == SteelConnection.SteelConnection_Mitered)
                    {
                        switch (weld.location)
                        {
                            case 2:
                                if (steel1Data.Orient == 0 || steel1Data.Orient == (SteelOrientationAngle)180)
                                    customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, -(steel1Member.depth / 2) / Math.Cos(45 * Math.PI / 180) + weld.offsetYValue, weld.offsetXValue);
                                else
                                    customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, -steel1Member.depth / 2 + weld.offsetYValue, weld.offsetXValue);
                                break;
                            case 4:
                                if (steel1Data.Orient == 0 || steel1Data.Orient == (SteelOrientationAngle)180)
                                    customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, -steel1Member.width / 2 + weld.offsetXValue);
                                else
                                    customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, -(steel1Member.width / 2) / Math.Cos(45 * Math.PI / 180) + weld.offsetXValue);
                                break;
                            case 6:
                                if (steel1Data.Orient == 0 || steel1Data.Orient == (SteelOrientationAngle)180)
                                    customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, steel1Member.width / 2);
                                else
                                    customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, (steel1Member.width / 2) / Math.Cos(45 * Math.PI / 180) + weld.offsetXValue);
                                break;
                            case 8:
                                if (steel1Data.Orient == 0 || steel1Data.Orient == (SteelOrientationAngle)180)
                                    customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, (steel1Member.depth / 2) / Math.Cos(45 * Math.PI / 180) + weld.offsetYValue, weld.offsetXValue);
                                else
                                    customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, steel1Member.depth / 2 + weld.offsetYValue, weld.offsetXValue);
                                break;
                        }

                    }
                    else
                    {
                        switch (weld.location)
                        {
                            case 2:
                                customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, -steel1Member.depth / 2 + weld.offsetYValue, weld.offsetXValue);
                                break;
                            case 4:
                                customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, -steel1Member.width / 2 + weld.offsetXValue);
                                break;
                            case 6:
                                customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, weld.offsetYValue, steel1Member.width / 2 + weld.offsetXValue);
                                break;
                            case 8:
                                customSupportDefinition.JointHelper.CreateRigidJoint(steel1PartKey, steelPort, weld.partKey, "Other", Plane.XY, Plane.XY, Axis.X, Axis.X, weld.offsetZValue, steel1Member.depth / 2 + weld.offsetYValue, weld.offsetXValue);
                                break;
                        }
                    }
                }
            }
        }

        catch (Exception e)
        {
            Type myType = customSupportDefinition.GetType();
            CmnException e1 = new CmnException("Error in SetSteelConnection.." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
            throw e1;
        }
        }
        

        /// <summary>
        /// This method will be called to check property is available on the Support(Part) or SupportOccurrence(PartOccurrence) and return its value.
        /// </summary>
        /// <param name="part">supportcomponent or support is of type bussiness object</param>
        /// <param name="interfaceName">Interface to which the property belongs.</param>
        /// <param name="propertyName">Name of property.</param>
        /// <param name="bFromCatalogOrOccurrence">Specifies the property from Catalog or occurrence.</param>
        /// <param name="bGetAutomatic">If automatic it checks occurrence and then catalog</param>
        /// <param name="propertyValue">Returns Property Value</param>
        /// <returns></returns>
        /// <code>
        /// GetIntPropertyValue(part, "IJUAhsHeight1","Height1",ref propertyValue)
        /// </code>
        public static PropertyValue GetAttributeValue(BusinessObject part, string interfaceName, string propertyName, FromCatalogOrOccurrence bFromCatalogOrOccurrence, bool bGetAutomatic)
        {
            PropertyValue returnPropertyValue = null;
            if (bGetAutomatic == false)
            {
                if (bFromCatalogOrOccurrence == FromCatalogOrOccurrence.FromCatalog)
                {
                    if (part.SupportsInterface(interfaceName))
                        returnPropertyValue = part.GetPropertyValue(interfaceName, propertyName);
                    else
                    {
                        Part oPart = (Part)part.GetRelationship("madeFrom", "part").TargetObjects[0];
                        returnPropertyValue = oPart.GetPropertyValue(interfaceName, propertyName);
                    }
                }
                else
                {
                    if (part.SupportsInterface(interfaceName))
                        returnPropertyValue = part.GetPropertyValue(interfaceName, propertyName);
                    else
                    {
                        BusinessObject oPart = (BusinessObject)part.GetRelationship("madeFrom", "partOcc").TargetObjects[0];
                        returnPropertyValue = oPart.GetPropertyValue(interfaceName, propertyName);
                    }
                }
            }
            else
            {
                try
                {
                    returnPropertyValue = part.GetPropertyValue(interfaceName, propertyName);
                }
                catch
                {
                    Part oPart = (Part)part.GetRelationship("madeFrom", "part").TargetObjects[0];
                    returnPropertyValue = oPart.GetPropertyValue(interfaceName, propertyName);
                }
            }
            return returnPropertyValue;
        }

        public static double GetVectorProjectioBetweenPorts(CustomSupportDefinition customSupportDefinition, String Port1, String Port2, PortAxisType Port1Axis)
        {
            Matrix4X4 PortName1 = customSupportDefinition.RefPortHelper.PortLCS(Port1);

            Matrix4X4 PortName2 = customSupportDefinition.RefPortHelper.PortLCS(Port2);

            Vector PortToPortVector;

            Vector RefPort1Axis = null;

            Position PortLoc1;

            PortLoc1 = PortName1.Origin;

            Position PortLoc2;

            PortLoc2 = PortName2.Origin;

            PortToPortVector = PortLoc2.Subtract(PortLoc1);

            if (Port1Axis == PortAxisType.X)
                RefPort1Axis = PortName1.XAxis;

            if (Port1Axis == PortAxisType.Y)
                RefPort1Axis = PortName1.YAxis;

            if (Port1Axis == PortAxisType.Z)
                RefPort1Axis = PortName1.ZAxis;

            double VectorResultant = 0;
            if (RefPort1Axis != null)
                VectorResultant = GetVectorProjection(PortToPortVector, RefPort1Axis);

            return VectorResultant;

        }
        public static Boolean AddImpliedPartbyInterface(CustomSupportDefinition customSupportDefinition, Collection<PartInfo> impliedparts, string Interface, String SupportNumber = "", int Catalog = 5)
        {
            try
            {

                bool isPartAdded = false; ;
                string partClassValue = string.Empty;
                Collection<object> ImpParts = new Collection<object>();
                string[] partKey = new string[100];
                if (SupportNumber == "")
                    SupportNumber = customSupportDefinition.SupportHelper.Support.SupportDefinition.PartNumber;
                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass auxTable = (PartClass)cataloghelper.GetPartClass("HS_OGAssy_FrImpParts");
                ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in classItems)
                {

                    bool isEqual = String.Equals(SupportNumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsFrImpPart", "SupportPartNumber")).PropValue), StringComparison.Ordinal);
                    if (isEqual == true)
                    {
                        //MessageBox.Show("1");
                        string partnumber = ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsFrImpPart", "ImpPartNumber")).PropValue);
                        int quantity = ((int)((PropertyValueInt)classItem.GetPropertyValue("IJUAhsFrImpPart", "ImpPartQuantity")).PropValue);

                        for (int index = 0; index < quantity; index++)
                        {
                            partKey[index] = "IMPLIEDPART_" + index;
                            impliedparts.Add(new PartInfo(partKey[index], partnumber, null));
                        }
                    }

                }
                return isPartAdded;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in Get Implied Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method will be called to check property is available on the SupportOccurrence(PartOccurrence) and Support(Part) and return its value.
        /// </summary>
        /// <param name="part">support or supportcomponent is of type bussiness object</param>
        /// <param name="interfaceName">Interface to which the property belongs.</param>
        /// <param name="propertyName">Name of property.</param>
        /// <param name="fromCatalog">Specifies whether the property is from occurence or catalog, by default its from occurence.</param>
        /// <param name="enforceFailure">Specifies whether to terminate or continue executing if a particular attribute is not present</param>
        /// <returns>Returns Property Value</returns>
        /// <code>
        /// GetIntPropertyValue(part, "IJUAhsHeight1","Height1",ref propertyValue)
        /// </code>
        public static PropertyValue GetAttributeValue(BusinessObject part, string interfaceName, string propertyName, bool fromCatalog = false, bool enforceFailure = false)
        {
            PropertyValue returnPropertyValue = null;

            if (fromCatalog == true)
            {
                try
                {
                    if (part is Part)
                    {
                        returnPropertyValue = part.GetPropertyValue(interfaceName, propertyName);
                    }
                }
                catch (Exception e)
                {
                    Type myType = typeof(OglaendAssemblyServices);
                    CmnException e1 = new CmnException("Unable to get Attribute value." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    if (enforceFailure)
                        throw e1;
                }

                try
                {
                    if (part is SP3D.Support.Middle.Support)
                    {
                        BusinessObject supportPart = part.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                        returnPropertyValue = supportPart.GetPropertyValue(interfaceName, propertyName);
                    }
                    else
                    {
                        BusinessObject supportComponentPart = part.GetRelationship("madeFrom", "part").TargetObjects[0];
                        returnPropertyValue = supportComponentPart.GetPropertyValue(interfaceName, propertyName);
                    }
                }
                catch (Exception e)
                {
                    Type myType = typeof(OglaendAssemblyServices);
                    CmnException e1 = new CmnException("Unable to get Attribute value." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    if (enforceFailure)
                        throw e1;
                }
            }
            else
            {
                try
                {
                    if (part is Part)
                    {
                        returnPropertyValue = part.GetPropertyValue(interfaceName, propertyName);
                    }
                }
                catch (Exception e)
                {
                    Type myType = typeof(OglaendAssemblyServices);
                    CmnException e1 = new CmnException("Unable to get Attribute value." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    if (enforceFailure)
                        throw e1;
                }

                if (part is SP3D.Support.Middle.Support)
                {
                    try
                    {
                        if (part.SupportsInterface(interfaceName))
                        {
                            returnPropertyValue = part.GetPropertyValue(interfaceName, propertyName);
                        }
                        else
                        {
                            BusinessObject supportPart = part.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                            returnPropertyValue = supportPart.GetPropertyValue(interfaceName, propertyName);
                        }
                    }
                    catch (Exception e)
                    {
                        Type myType = typeof(OglaendAssemblyServices);
                        CmnException e1 = new CmnException("Unable to get Attribute value." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                        if (enforceFailure)
                            throw e1;
                    }
                }
                else
                {
                    try
                    {
                        if (part.SupportsInterface(interfaceName))
                        {
                            returnPropertyValue = part.GetPropertyValue(interfaceName, propertyName);
                        }
                        else
                        {
                            BusinessObject supportComponentPart = part.GetRelationship("madeFrom", "part").TargetObjects[0];
                            returnPropertyValue = supportComponentPart.GetPropertyValue(interfaceName, propertyName);
                        }
                    }
                    catch (Exception e)
                    {
                        Type myType = typeof(OglaendAssemblyServices);
                        CmnException e1 = new CmnException("Unable to get Attribute value." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                        if (enforceFailure)
                            throw e1;
                    }
                }
            }
            return returnPropertyValue;
        }

        /// <summary>
        /// Function to add a part to the collection of part occurences for the assembly.  This function will add the part.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="partKey">Name of the PartKey</param>
        /// <param name="partNumber">String containing a PartNumber</param>
        /// <param name="parts">collection to hold PartInfo</param>
        /// <code>
        /// AddPart(this, partKey, partNumber,parts)
        /// </code>
        public static void AddPart(CustomSupportDefinition customSupportDefinition, string partKey, string partNumber, Collection<PartInfo> parts)
        {
            try
            {
                if (!string.IsNullOrEmpty(partNumber))
                {
                    parts.Add(new PartInfo(partKey, partNumber));
                }
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in AddPart." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}
