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
//   20-Jun-2013     Rajeswari,Hema     CR-CP-224491- Convert HS_FINL_Assy to C# .Net
//   31-Oct-2014     PVK                TR-CP-260301  Resolve coverity issues found in August 2014 report 
//   22-Jan-2015     PVK                TR-CP-264951  Resolve coverity issues found in November 2014 report
//   30-Mar-2015     PVK                CR-CP-245789  Modify the exsisting .Net Frame_Assy to behave like URS Frame supports
//   06-May-2015     PVK                CR-CP-270669	Modify the exsisting .Net Frame_Assy to honor attributes which are not in OA
//   16-Jul-2015     PVK       		    Resolve coverity issues found in July 2015 report
//   26-Oct-2015     PVK       		    Resolve coverity issues found in Octpber 2015 report
//   17-Dec-2015     Ramya              TR 284319	Multiple Record exception dumps are created on copy pasting supports
//   10-05-2016      PVK                TR-CP-291676	Bug in AddImpliedPart function when using Hanger Rule
//   07-Jun-2016     PVK                TR-CP-293408	Delivered HS_S3DAssy L Frame and Corner Frames Fail with Cap Plate 3/Added General error message 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Support.Symbols;
using Ingr.SP3D.Route.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    public static class FrameAssemblyServices
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
        /// Connection Types for steel parts.
        /// Values must match hsSteelTeeConnection and hsSteelCornerConnection Codelists from HS_System_Codelist.xls
        /// </summary>
        public enum SteelConnection { SteelConnection_Butted = 0, SteelConnection_Lapped = 1, SteelConnection_Nested = 2, SteelConnection_Coped = 3, SteelConnection_Fitted = 4, SteelConnection_Mitered = 5, SteelConnection_UserDefined = 999 };
        /// <summary>
        /// Used in GetPropertyValue method while retrieving the Property Value
        /// </summary>
        public enum QueryPropertyOn { CatalogOnly = 1, OccurrenceOnly = 2, Both = 3 };
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
        public static Boolean AddImpliedPart(CustomSupportDefinition customSupportDefinition, string part, string rule, Collection<PartInfo> impliedparts, BusinessObject objectForRule = null, int quantity = 1)
        {
            try
            {
                bool isImpliedPartAdded = false; ;
                string partClassValue = string.Empty, partRule = string.Empty;
                Collection<object> ruleResults1 = new Collection<object>();
                string[] partKey = new string[quantity];
                string partNumber = null;

                if (!string.IsNullOrEmpty(part))
                {
                    if (rule == "")
                    {
                        GetPartClassValue(part, ref partClassValue);
                        partRule = partClassValue;
                    }
                    else
                    {
                        partRule = rule;
                    }

                    for (int index = 0; index < quantity; index++)
                    {
                        partKey[index] = "IMPLIEDPART_" + index;
                        impliedparts.Add(new PartInfo(partKey[index], part, partRule));
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
                            partKey = new string[ruleResults1.Count];
                            for (int index = 0; index < ruleResults1.Count; index++)
                            {
                                partNumber = ruleResults1[index].ToString();
                                partKey[index] = "IMPLIEDPART_" + index;
                                impliedparts.Add(new PartInfo(partKey[index], partNumber, null));
                            }
                            isImpliedPartAdded = true;
                        }
                    }
                }
                else
                {
                    isImpliedPartAdded = false;
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
                CmnException e1 = new CmnException("Error in GetPartClassValue." + "ERROR:" + e.Message, e);
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

                resultMember.partNumber = part.PartNumber;
                resultMember.sectionDescription = part.PartDescription;
                resultMember.sectionName = crossSection.Name;
                resultMember.sectionType = crossSection.CrossSectionClass.Name;

                GetDoublePropertyValue(crossSection, "IStructCrossSectionDimensions", "Depth", ref resultMember.depth);
                GetDoublePropertyValue(crossSection, "IStructCrossSectionDimensions", "Width", ref resultMember.width);
                GetDoublePropertyValue(crossSection, "IStructFlangedSectionDimensions", "tw", ref resultMember.webThickness);
                GetDoublePropertyValue(crossSection, "IStructFlangedSectionDimensions", "tf", ref resultMember.flangeThickness);
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
        public static Collection<WeldData> AddWeldsFromCatalog(CustomSupportDefinition customSupportDefinition, Collection<PartInfo> parts, string interfaceName, string supportPartNumber = "", string weldServiceClassName ="")
        {
            try
            {
                string offsetZRule = string.Empty;
                WeldData weld = new WeldData();
                Collection<WeldData> weldCollection = new Collection<WeldData>();

                if (supportPartNumber == "")
                    supportPartNumber = customSupportDefinition.SupportHelper.Support.SupportDefinition.PartNumber;

                BusinessObject partClass = customSupportDefinition.SupportHelper.Support.SupportDefinition.PartClass;
                PartClass S3DframeWelds = null;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                if (weldServiceClassName != string.Empty)
                {
                    try
                    {
                        S3DframeWelds = (PartClass)catalogBaseHelper.GetPartClass(weldServiceClassName);
                    }
                    catch
                    {
                        S3DframeWelds = null;
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "AddWeldsFromCatalog", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "The Given Weld parts Service Class does not exist in Database.", "", "FrameAssemblyServices.cs", 417);
                    }
                    
                }
                else
                {
                    try
                    {
                        S3DframeWelds = (PartClass)catalogBaseHelper.GetPartClass(partClass + "_Welds");
                    }
                    catch
                    {
                        S3DframeWelds = null;
                    }
                }                
                
                if (S3DframeWelds != null)
                {
                    IEnumerable<BusinessObject> weldParts = S3DframeWelds.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    weldParts = weldParts.Where(part => (string)((PropertyValueString)GetPropertyValue(part,interfaceName, "SupportPartNumber")).PropValue == supportPartNumber);
                    int i = 1;
                    foreach (BusinessObject part in weldParts)
                    {
                        if (part.SupportsInterface("IJUAhsIFrameWelds"))
                        {
                            weld.partNumber = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsIFrameWelds", "WeldPartNumber")).PropValue;
                            weld.partRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsIFrameWelds", "WeldPartRule")).PropValue;
                            weld.connection = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsIFrameWelds", "Connection")).PropValue;
                            weld.location = (int)((PropertyValueCodelist)GetPropertyValue(part,"IJUAhsIFrameWelds", "Location")).PropValue;

                            weld.offsetXValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsIFrameWelds", "OffsetXValue")).PropValue;
                            string offsetXRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsIFrameWelds", "OffsetXRule")).PropValue;
                            if (!string.IsNullOrEmpty(offsetXRule))
                                customSupportDefinition.GenericHelper.GetDataByRule(offsetXRule, null, out weld.offsetXValue);

                            weld.offsetYValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsIFrameWelds", "OffsetYValue")).PropValue;
                            string offsetYRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsIFrameWelds", "OffsetYRule")).PropValue;
                            if (!string.IsNullOrEmpty(offsetYRule))
                                customSupportDefinition.GenericHelper.GetDataByRule(offsetYRule, null, out weld.offsetYValue);

                            weld.offsetZValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsIFrameWelds", "OffsetZValue")).PropValue;
                            offsetZRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsIFrameWelds", "OffsetZRule")).PropValue;
                        }
                        else if (part.SupportsInterface("IJUAhsLFrameWelds"))
                        {
                            weld.partNumber = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsLFrameWelds", "WeldPartNumber")).PropValue;
                            weld.partRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsLFrameWelds", "WeldPartRule")).PropValue;
                            weld.connection = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsLFrameWelds", "Connection")).PropValue;
                            weld.location = (int)((PropertyValueCodelist)GetPropertyValue(part,"IJUAhsLFrameWelds", "Location")).PropValue;

                            weld.offsetXValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsLFrameWelds", "OffsetXValue")).PropValue;
                            string offsetXRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsLFrameWelds", "OffsetXRule")).PropValue;
                            if (!string.IsNullOrEmpty(offsetXRule))
                                customSupportDefinition.GenericHelper.GetDataByRule(offsetXRule, null, out weld.offsetXValue);

                            weld.offsetYValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsLFrameWelds", "OffsetYValue")).PropValue;
                            string offsetYRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsLFrameWelds", "OffsetYRule")).PropValue;
                            if (!string.IsNullOrEmpty(offsetYRule))
                                customSupportDefinition.GenericHelper.GetDataByRule(offsetYRule, null, out weld.offsetYValue);

                            weld.offsetZValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsLFrameWelds", "OffsetZValue")).PropValue;
                            offsetZRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsLFrameWelds", "OffsetZRule")).PropValue;
                        }
                        else if (part.SupportsInterface("IJUAhsTFrameWelds"))
                        {
                            weld.partNumber = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsTFrameWelds", "WeldPartNumber")).PropValue;
                            weld.partRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsTFrameWelds", "WeldPartRule")).PropValue;
                            weld.connection = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsTFrameWelds", "Connection")).PropValue;
                            weld.location = (int)((PropertyValueCodelist)GetPropertyValue(part,"IJUAhsTFrameWelds", "Location")).PropValue;

                            weld.offsetXValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsTFrameWelds", "OffsetXValue")).PropValue;
                            string offsetXRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsTFrameWelds", "OffsetXRule")).PropValue;
                            if (!string.IsNullOrEmpty(offsetXRule))
                                customSupportDefinition.GenericHelper.GetDataByRule(offsetXRule, null, out weld.offsetXValue);

                            weld.offsetYValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsTFrameWelds", "OffsetYValue")).PropValue;
                            string offsetYRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsTFrameWelds", "OffsetYRule")).PropValue;
                            if (!string.IsNullOrEmpty(offsetYRule))
                                customSupportDefinition.GenericHelper.GetDataByRule(offsetYRule, null, out weld.offsetYValue);

                            weld.offsetZValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsTFrameWelds", "OffsetZValue")).PropValue;;
                            offsetZRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsTFrameWelds", "OffsetZRule")).PropValue;
                        }
                        else if (part.SupportsInterface("IJUAhsUFrameWelds"))
                        {
                            weld.partNumber = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsUFrameWelds", "WeldPartNumber")).PropValue;
                            weld.partRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsUFrameWelds", "WeldPartRule")).PropValue;
                            weld.connection = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsUFrameWelds", "Connection")).PropValue;
                            weld.location = (int)((PropertyValueCodelist)GetPropertyValue(part,"IJUAhsUFrameWelds", "Location")).PropValue;

                            weld.offsetXValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsUFrameWelds", "OffsetXValue")).PropValue;
                            string offsetXRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsUFrameWelds", "OffsetXRule")).PropValue;
                            if (!string.IsNullOrEmpty(offsetXRule))
                                customSupportDefinition.GenericHelper.GetDataByRule(offsetXRule, null, out weld.offsetXValue);

                            weld.offsetYValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsUFrameWelds", "OffsetYValue")).PropValue;
                            string offsetYRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsUFrameWelds", "OffsetYRule")).PropValue;
                            if (!string.IsNullOrEmpty(offsetYRule))
                                customSupportDefinition.GenericHelper.GetDataByRule(offsetYRule, null, out weld.offsetYValue);

                            weld.offsetZValue = (double)((PropertyValueDouble)GetPropertyValue(part, "IJUAhsUFrameWelds", "OffsetZValue")).PropValue;
                            offsetZRule = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsUFrameWelds", "OffsetZRule")).PropValue;
                        }
                        else if (part.SupportsInterface(interfaceName))
                        {
                            weld.partNumber = (string)((PropertyValueString)GetPropertyValue(part,interfaceName, "WeldPartNumber")).PropValue;
                            weld.partRule = (string)((PropertyValueString)GetPropertyValue(part,interfaceName, "WeldPartRule")).PropValue;
                            weld.connection = (string)((PropertyValueString)GetPropertyValue(part,interfaceName, "Connection")).PropValue;
                            weld.location = (int)((PropertyValueCodelist)GetPropertyValue(part,interfaceName, "Location")).PropValue;

                            weld.offsetXValue = (double)((PropertyValueDouble)GetPropertyValue(part, interfaceName, "OffsetXValue")).PropValue;
                            string offsetXRule = (string)((PropertyValueString)GetPropertyValue(part,interfaceName, "OffsetXRule")).PropValue;
                            if (!string.IsNullOrEmpty(offsetXRule))
                                customSupportDefinition.GenericHelper.GetDataByRule(offsetXRule, null, out weld.offsetXValue);

                            weld.offsetYValue = (double)((PropertyValueDouble)GetPropertyValue(part, interfaceName, "OffsetYValue")).PropValue;
                            string offsetYRule = (string)((PropertyValueString)GetPropertyValue(part,interfaceName, "OffsetYRule")).PropValue;
                            if (!string.IsNullOrEmpty(offsetYRule))
                                customSupportDefinition.GenericHelper.GetDataByRule(offsetYRule, null, out weld.offsetYValue);

                            weld.offsetZValue = (double)((PropertyValueDouble)GetPropertyValue(part, interfaceName, "OffsetZValue")).PropValue;
                            offsetZRule = (string)((PropertyValueString)GetPropertyValue(part,interfaceName, "OffsetZRule")).PropValue;
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

                Matrix4X4 hangerPort = new Matrix4X4();
                Position pointObject = new Position(0, 0, 0);
                double structDirectionAngle = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, "Struct_2", PortAxisType.X, OrientationAlong.Direct); 
                int frameType = ((PropertyValueCodelist)customSupportDefinition.SupportHelper.Support.SupportDefinition.GetPropertyValue("IJUAhsFrameType", "FrameType")).PropValue;

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

                if (supportingCount == 2 && customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByPoint && frameType == 1 && (supportingType == "Steel"))
                {
                    if (HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(structDirectionAngle)) , 180)==true || HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(structDirectionAngle)) , 0)==true)
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

                        structDirectionAngle = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Structure", PortAxisType.X, "Struct_2", PortAxisType.X, OrientationAlong.Direct);
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

                                if (supportingCount == 2 && customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByPoint && frameType == 1 && (supportingType == "Steel"))
                                {
                                    if (pointObject != null)
                                        BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                    else
                                    {
                                        if (HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(structDirectionAngle)), 180) == true || HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(structDirectionAngle)), 0) == true)
                                        {
                                            BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteY, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                            if (Math.Round((customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y) * 180 / Math.PI)) > 90)
                                                BB_Z.Set(-BB_Z.X, -BB_Z.Y, -BB_Z.Z);
                                        }
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

                                if (supportingCount == 2 && customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByPoint && frameType == 1 && (supportingType == "Steel"))
                                {
                                    if (pointObject != null)
                                        BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                    else
                                    {
                                        if (HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(structDirectionAngle)), 180) == true || HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(structDirectionAngle)), 0) == true)
                                        {
                                            BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteY, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                            if (Math.Round((customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y) * 180 / Math.PI)) > 90)
                                                BB_Z.Set(-BB_Z.X, -BB_Z.Y, -BB_Z.Z);
                                        }
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
                                    frameConfiguration = (int)((PropertyValueCodelist)GetPropertyValue(customSupportDefinition.SupportHelper.Support, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                                }
                                catch
                                {
                                    frameConfiguration = 1;
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
                                        frameConfiguration = (int)((PropertyValueCodelist)GetPropertyValue(customSupportDefinition.SupportHelper.Support, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                                    }
                                    catch
                                    {
                                        frameConfiguration = 1;
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
                                    GlobalZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ);
                                    BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, GlobalZ); // Project Route X-Axis into Horizontal Plane
                                    BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X);

                                    if (supportingCount == 2 && customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByPoint && frameType == 1 && (supportingType == "Steel"))
                                    {
                                        if (pointObject != null)
                                            BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                        else
                                        {
                                            if (HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(structDirectionAngle)), 180) == true || HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(structDirectionAngle)), 0) == true)
                                            {
                                                BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteY, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                                if (Math.Round((customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y) * 180 / Math.PI)) > 90)
                                                    BB_Z.Set(-BB_Z.X, -BB_Z.Y, -BB_Z.Z);
                                            }
                                        }
                                    }

                                    customSupportDefinition.BoundingBoxHelper.CreateBoundingBox(BB_Z, BB_X, boundingBoxName, insulationOption, mirrorFlag, true, false);
                                    break;
                                }
                            case FrameBBXOrientation.FrameBBXOrientation_Orthogonal:
                                {
                                    GlobalZ = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.GlobalZ); // Get Global Z
                                    BB_X = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.AlongRoute, GlobalZ); // Use GlobalZ as BBX X-Axis
                                    BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project Route into Horizontal Plane

                                    BB_Y = GlobalZ.Cross(BB_X); // Get Horizontal Vector in the BBX Plane (For Orthogonal direction)
                                    if (supportingCount == 2 && customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByPoint && frameType == 1 && (supportingType == "Steel"))
                                    {
                                        if (pointObject != null)
                                            BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.OrthogonalToRoute, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                        else
                                        {
                                            if (HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(structDirectionAngle)), 180) == true || HgrCompareDoubleService.cmpdbl(Math.Abs(Math.Round(structDirectionAngle)), 0) == true)
                                            {
                                                BB_Z = customSupportDefinition.BoundingBoxHelper.GetVectorForBBXDirection(BoundingBoxDirection.RouteY, BB_X); // Project Vector From Route to Structure into the BBX Plane
                                                if (Math.Round((customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", "Structure", PortAxisType.Y) * 180 / Math.PI)) > 90)
                                                    BB_Z.Set(-BB_Z.X, -BB_Z.Y, -BB_Z.Z);
                                            }
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
                                        frameConfiguration = (int)((PropertyValueCodelist)FrameAssemblyServices.GetPropertyValue(customSupportDefinition.SupportHelper.Support, "IJUAhsFrameConfiguration", "FrameConfiguration")).PropValue;
                                    }
                                    catch
                                    {
                                        frameConfiguration = 1;
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
                if (part != null)
                    resultMember.sectionDescription = part.PartDescription;

                CrossSection crossSection = (CrossSection)sectionData.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                if (crossSection != null)
                {
                    resultMember.sectionName = crossSection.Name;
                    resultMember.sectionType = crossSection.CrossSectionClass.Name;
                    resultMember.width = crossSection.Width;
                    resultMember.depth = crossSection.Depth;
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
        private static void GetIntPropertyValue(BusinessObject part, string interfaceName, string propertyName, ref int propertyValue)
        {
            if (part.SupportsInterface(interfaceName))
            {
                try
                {
                    propertyValue = (int)((PropertyValueInt)part.GetPropertyValue(interfaceName, propertyName)).PropValue;
                }
                catch (System.InvalidCastException ex)
                {
                    CmnException e1 = new CmnException("Invalid Typecast while calling GetIntPropertyValue on property: " + propertyValue + " present on Interface: " + interfaceName, ex);
                    throw e1;
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
        private static void GetDoublePropertyValue(BusinessObject part, string interfaceName, string propertyName, ref double propertyValue)
        {
            if (part.SupportsInterface(interfaceName))
            {
                try
                {
                    propertyValue = (double)((PropertyValueDouble)part.GetPropertyValue(interfaceName, propertyName)).PropValue;
                }
                catch (System.InvalidCastException ex)
                {
                    CmnException e1 = new CmnException("Invalid Typecast while calling GetDoublePropertyValue on property: " + propertyValue + " present on Interface: " + interfaceName, ex);
                    throw e1;
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
                    else if( (customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "HangerBeam";    //HgrBeam                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                if (supportingType == "Steel")
                {
                    try
                    {
                        resultMember.sectionName = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).SectionName;
                    }
                    catch { resultMember.sectionName = string.Empty;  }
                    try
                    {
                        resultMember.sectionType = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).SectionType;
                    }
                    catch { resultMember.sectionType = string.Empty; }
                    try
                    {
                        resultMember.webThickness = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).WebThickness;
                    }
                    catch { resultMember.webThickness = 0; }
                    try
                    {
                        resultMember.width = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).Width;
                    }
                    catch { resultMember.width = 0; }
                    try
                    {
                        resultMember.depth = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).Depth;
                    }
                    catch { resultMember.depth = 0; }
                    try
                    {
                        resultMember.flangeThickness = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).FlangeThickness;
                    }
                    catch { resultMember.flangeThickness = 0; }
                    try
                    {
                        resultMember.flangeWidth = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).FlangeWidth;
                    }
                    catch { resultMember.flangeWidth = 0;  }
                    try
                    {
                        resultMember.webDepth = customSupportDefinition.SupportingHelper.SupportingObjectInfo(index).WebDepth;
                    }
                    catch { resultMember.webDepth = 0; }
                    resultMember.HSS_NominalWallThickness = GetDoubleValueFromSupportingObject("IJUAHSS", "tnom", customSupportDefinition, 0);
                    resultMember.HSS_DesignWallThickness = GetDoubleValueFromSupportingObject("IJUAHSS", "tdes", customSupportDefinition, 0);

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

        public static double GetDoubleValueFromSupportingObject(string interfaceName, string propertyName, CustomSupportDefinition customSupportDefinition, int structureIndex)
        {
            Ingr.SP3D.Support.Middle.Support support = customSupportDefinition.SupportHelper.Support;
            double value = 0;
            if (customSupportDefinition.SupportHelper.PlacementType != PlacementType.PlaceByReference &&
                structureIndex < support.SupportingObjects.Count &&
                support.SupportingObjects[structureIndex] != null)
            {
                
                ConnectionComponent oconnComp = support.SupportingObjects[structureIndex] as ConnectionComponent;
                if (oconnComp != null)
                {
                    if (oconnComp.CrossSection.SupportsInterface(interfaceName))
                    {
                        PropertyValueDouble oPVD = (PropertyValueDouble)(oconnComp.CrossSection.GetPropertyValue(interfaceName, propertyName));
                        value = (double)oPVD.PropValue;
                    }
                }
                MemberPart oMemberart = support.SupportingObjects[structureIndex] as MemberPart;
                if (oMemberart != null)
                {
                    if (oMemberart.CrossSection.SupportsInterface(interfaceName))
                    {
                        PropertyValueDouble oPVD = (PropertyValueDouble)(oMemberart.CrossSection.GetPropertyValue(interfaceName, propertyName));
                        value = (double)oPVD.PropValue;
                    }
                }
            }

            return value;
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

                            if (HgrCompareDoubleService.cmpdbl(Math.Round((AngleBetweenVectors(BB_X, BB_Z) * 180 / Math.PI), 0), 0) == true || HgrCompareDoubleService.cmpdbl(Math.Round((AngleBetweenVectors(BB_X, BB_Z) * 180 / Math.PI), 0), 180) == true)
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

                            if (HgrCompareDoubleService.cmpdbl(Math.Round((AngleBetweenVectors(BB_X, BB_Z) * 180 / Math.PI), 0), 0) == true || HgrCompareDoubleService.cmpdbl(Math.Round((AngleBetweenVectors(BB_X, BB_Z) * 180 / Math.PI), 0), 180) == true)
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
        public static void SetSteelConnection(CustomSupportDefinition customSupportDefinition, string steel1PartKey, string steel1Port, ref HSSteelConfig steel1Data, double overLengthSteel1, string steel2PartKey, string steel2Port, ref HSSteelConfig steel2Data, double overLengthSteel2, SteelConnectionAngle angle, double axialOffset, SteelConnection connectionType, Boolean mirrorConnection = false, SteelJointType steelJoint = SteelJointType.SteelJoint_RIGID, Collection<WeldData> weldDataCollection = null)
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
                PropertyValueCodelist endCutbackAnchorPoint = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel1PartKey],"IJOAhsCutback", "EndCutbackAnchorPoint");
                PropertyValueCodelist beginCutbackAnchorPoint = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel1PartKey],"IJOAhsCutback", "BeginCutbackAnchorPoint");

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass hangerSteelType = (PartClass)cataloghelper.GetPartClass("HgrSteelTypes");
                ReadOnlyCollection<BusinessObject> partsHangerSteelType = hangerSteelType.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject part in partsHangerSteelType)
                {
                    if ((string)((PropertyValueString)GetPropertyValue(part,"IJUAhsSteelTypes", "SteelType")).PropValue == steel1Member.sectionType)
                        steel1Type = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsSteelTypes", "StdSteelType")).PropValue;
                    if ((string)((PropertyValueString)GetPropertyValue(part,"IJUAhsSteelTypes", "SteelType")).PropValue == steel2Member.sectionType)
                    {
                        steel2Type = (string)((PropertyValueString)GetPropertyValue(part,"IJUAhsSteelTypes", "StdSteelType")).PropValue;
                        break;
                    }
                }
                //NOTE: Reflect functionality on Rich Steel will not be implemented for v2011. Therefore, all lines
                //that deal with the reflected cross sections have been commented out. If in the future a reflect
                //toggle gets added to Rich steel, then we can uncomment these lines and extend the SetSteelConnection
                //function to work with reflected cross sections.
                if (HgrCompareDoubleService.cmpdbl(Math.Round(axialOffset, 5) , 0) == true && steelJoint == SteelJointType.SteelJoint_RIGID)
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
                        PropertyValueCodelist steel1IndexHangerType = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel1PartKey],"IJOAHsHgrBeamType", "HgrBeamType");
                        PropertyValueCodelist steel2IndexHangerType = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel2PartKey],"IJOAHsHgrBeamType", "HgrBeamType");
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
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                }
                                else
                                {
                                    if (steel2Port.Substring(0, 3).Equals("End"))
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
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
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                }
                                else
                                {
                                    if (steel2Port.Substring(0, 3).Equals("End"))
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
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
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                }
                                else
                                {
                                    if (steel2Port.Substring(0, 3).Equals("End"))
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
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
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                }
                                else
                                {
                                    if (steel2Port.Substring(0, 3).Equals("End"))
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel1PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel1PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
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
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                }
                                else
                                {
                                    if (steel1Port.Substring(0, 3).Equals("End"))
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
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
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                }
                                else
                                {
                                    if (steel1Port.Substring(0, 3).Equals("End"))
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
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
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                }
                                else
                                {
                                    if (steel1Port.Substring(0, 3).Equals("End"))
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 2, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 8, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
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
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(endCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "EndCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackEndAngle");
                                        }
                                    }
                                }
                                else
                                {
                                    if (steel1Port.Substring(0, 3).Equals("End"))
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                    }
                                    else
                                    {
                                        if (angle == (SteelConnectionAngle)90)
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 6, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(-45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
                                        }
                                        else
                                        {
                                            componentDictionary[steel2PartKey].SetPropertyValue(beginCutbackAnchorPoint.PropValue = 4, "IJOAhsCutback", "BeginCutbackAnchorPoint");
                                            componentDictionary[steel2PartKey].SetPropertyValue(45 * (Math.Atan(1) * 4.0) / 180, "IJOAhsCutback", "CutbackBeginAngle");
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
                        componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(0), "IJOAhsCutback", "CutbackEndAngle");
                    else
                        componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(0), "IJOAhsCutback", "CutbackBeginAngle");

                    if (steel2Port.Substring(0, 3).Equals("End"))
                        componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(0), "IJOAhsCutback", "CutbackEndAngle");
                    else
                        componentDictionary[steel2PartKey].SetPropertyValue(Convert.ToDouble(0), "IJOAhsCutback", "CutbackBeginAngle");
                }
                steel1Data.CardinalPoint = steel1CP;
                steel2Data.CardinalPoint = steel2CP;

                steel1Data.OffsetX = steel1OffsetX;
                steel1Data.OffsetY = steel1OffsetY;

                steel2Data.OffsetX = steel2OffsetX;
                steel2Data.OffsetY = steel2OffsetY;

                //Set the Attributes on Steel 1.
                PropertyValueCodelist CP1 = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel1PartKey],"IJOAhsSteelCP", "CP1");
                PropertyValueCodelist CP2 = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel1PartKey],"IJOAhsSteelCP", "CP2");
                PropertyValueCodelist CP6 = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel1PartKey],"IJOAhsSteelCP", "CP6");
                PropertyValueCodelist CP7 = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel1PartKey],"IJOAhsSteelCPFlexPort", "CP7");
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
                        componentDictionary[steel1PartKey].SetPropertyValue(CP7.PropValue = steel1Data.CardinalPoint, "IJOAhsSteelCP", "CP7");
                        componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(Convert.ToDouble(steel1Data.Orient) * Math.Atan(1) * 4.0 / 180), "IJOAhsEndFlexPort", "EndFlexPortRotZ");
                        componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1Data.OffsetX), "IJOAhsEndFlexPort", "EndFlexPortXOffset");
                        componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1Data.OffsetY), "IJOAhsEndFlexPort", "EndFlexPortYOffset");
                        componentDictionary[steel1PartKey].SetPropertyValue(Convert.ToDouble(steel1OverLength + overLengthSteel1), "IJUAHgrOccOverLength", "EndOverLength");
                        break;
                }
                //Set the Attributes on Steel 2
                CP1 = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel2PartKey],"IJOAhsSteelCP", "CP1");
                CP2 = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel2PartKey],"IJOAhsSteelCP", "CP2");
                CP6 = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel2PartKey],"IJOAhsSteelCP", "CP6");
                CP7 = (PropertyValueCodelist)GetPropertyValue(componentDictionary[steel2PartKey],"IJOAhsSteelCPFlexPort", "CP7");
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
                        componentDictionary[steel2PartKey].SetPropertyValue(CP6.PropValue = steel2Data.CardinalPoint, "IJOAhsSteelCP", "CP6");
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
        /// This method will be called to check the support with the given family and type attributes
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="family"> the value of family attribute from support</param>
        /// <param name="type">the value of type attribute from support</param>
        /// <returns></returns>
        /// <code>
        /// CheckSupportWithFamilyAndType(this, family, type)
        /// </code>
        public static void CheckSupportWithFamilyAndType(CustomSupportDefinition customSupportDefinition, string family, string type)
        {
            try
            {
                RouteFeature routeFeature = customSupportDefinition.SupportHelper.SupportedObjects[0] as RouteFeature;
                ReadOnlyCollection<RoutePart> collection = null;
                if (routeFeature != null)
                    collection = routeFeature.Parts;
                RoutePart genPart = null;
                if (collection != null)
                    genPart = collection[0] as RoutePart;
                if (family == "" || type == "")
                    return;

                double familyNDPFrom = 0, familyNDPTo = 0;
                string geStdWeight = string.Empty, lStdWeight = string.Empty, carbonSteel = string.Empty, alloySteel = string.Empty, stainlessSteel = string.Empty, LE3 = string.Empty, G3 = string.Empty, PP = string.Empty, FP = string.Empty, HC = string.Empty, CC = string.Empty, freezeProt = string.Empty, procHeat = string.Empty, controTrace = string.Empty, t300 = string.Empty, t650 = string.Empty, gt650 = string.Empty, t750 = string.Empty, gt750 = string.Empty;
                // Get configuration variable from excel sheet.
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass ursWGFamilyAux = (PartClass)catalogBaseHelper.GetPartClass("hsURS_WG_Family");
                ReadOnlyCollection<BusinessObject> ursWGFamilyItems = ursWGFamilyAux.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                Boolean isWgClass = false;
                if (ursWGFamilyItems != null)
                {
                    foreach (BusinessObject part1 in ursWGFamilyItems)
                    {
                        if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "Family")).PropValue == family) && ((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "Type")).PropValue == type))
                        {
                            familyNDPFrom = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrWGFamily", "PipeDiaFrom")).PropValue;
                            familyNDPTo = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrWGFamily", "PipeDiaTo")).PropValue;
                            geStdWeight = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GeStdWeight")).PropValue;
                            lStdWeight = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "LStdWeight")).PropValue;
                            carbonSteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "CS")).PropValue;
                            alloySteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "AS")).PropValue;
                            stainlessSteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "SS")).PropValue;
                            LE3 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "LE3")).PropValue;
                            G3 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "G3")).PropValue;
                            PP = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "PP")).PropValue;
                            FP = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "FP")).PropValue;
                            HC = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "HC")).PropValue;
                            CC = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "CC")).PropValue;
                            freezeProt = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "FreezeProt")).PropValue;
                            procHeat = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "ProcHeat")).PropValue;
                            controTrace = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "ControTrace")).PropValue;
                            t300 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T300")).PropValue;
                            t650 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T650")).PropValue;
                            gt650 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GT650")).PropValue;
                            t750 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T750")).PropValue;
                            gt750 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GT750")).PropValue;
                            isWgClass = true;
                            break;
                        }
                    }
                }
                if (isWgClass == false)
                {
                    ursWGFamilyAux = (PartClass)catalogBaseHelper.GetPartClass("hsURSAssy_WG_NFamily");
                    ursWGFamilyItems = ursWGFamilyAux.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    if (ursWGFamilyItems != null)
                    {
                        foreach (BusinessObject part1 in ursWGFamilyItems)
                        {
                            if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "Family")).PropValue == family) && ((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "Type")).PropValue == type))
                            {
                                familyNDPFrom = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrWGFamily", "PipeDiaFrom")).PropValue;
                                familyNDPTo = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrWGFamily", "PipeDiaTo")).PropValue;
                                geStdWeight = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GeStdWeight")).PropValue;
                                lStdWeight = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "LStdWeight")).PropValue;
                                carbonSteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "CS")).PropValue;
                                alloySteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "AS")).PropValue;
                                stainlessSteel = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "SS")).PropValue;
                                LE3 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "LE3")).PropValue;
                                G3 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "G3")).PropValue;
                                PP = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "PP")).PropValue;
                                FP = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "FP")).PropValue;
                                HC = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "HC")).PropValue;
                                CC = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "CC")).PropValue;
                                freezeProt = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "FreezeProt")).PropValue;
                                procHeat = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "ProcHeat")).PropValue;
                                controTrace = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "ControTrace")).PropValue;
                                t300 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T300")).PropValue;
                                t650 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T650")).PropValue;
                                gt650 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GT650")).PropValue;
                                t750 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "T750")).PropValue;
                                gt750 = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrWGFamily", "GT750")).PropValue;
                                break;
                            }
                        }
                    }
                }
                // Get all pipes on the support.
                // *********************************************************************************************
                // Filter based on Pipe Size, command type, discipline type
                // *********************************************************************************************
                PipeObjectInfo pipeInfo = (PipeObjectInfo)customSupportDefinition.SupportedHelper.SupportedObjectInfo(1);
                double primaryPipeSize = pipeInfo.NominalDiameter.Size;
                string unit = pipeInfo.NominalDiameter.Units;

                if (unit == "mm")
                    primaryPipeSize = Math.Round(MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, primaryPipeSize, UnitName.NPD_MILLIMETER, UnitName.NPD_INCH), 4);
                // check valid pipe size
                if (HgrCompareDoubleService.cmpdbl(familyNDPFrom, 0) == false && HgrCompareDoubleService.cmpdbl(familyNDPTo, 0) == false)
                {
                    if (familyNDPFrom > primaryPipeSize || familyNDPTo < primaryPipeSize)
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Pipe size not valid.", "", "FrameAssemblyServices.cs", 557);
                }

                // Setting up to get pipe information that we will use to determine what supports can be used.
                // Get the insulation thickness of the pipe that we are placing on.
                double inslatThickness = pipeInfo.InsulationThickness;
                double compareThickness = Math.Round(MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, inslatThickness, UnitName.DISTANCE_METER, UnitName.DISTANCE_INCH), 0);
                string insulCompare = string.Empty;
                if (compareThickness <= 3)
                    insulCompare = "LE3";
                else
                    insulCompare = "G3";

                int inslatCodeTemp = 0, inslatGroupCode = 0;
                string inslatCode = string.Empty;
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (HgrCompareDoubleService.cmpdbl(inslatThickness, 0) == false)
                {
                    // Get the index number for the inslation purpose
                    inslatCodeTemp = pipeInfo.InsulationPurpose;
                    // Convert the index number into the string representing the purpose
                    inslatGroupCode = metadataManager.GetCodelistInfo("InsulationPurpose", "REFDAT").GetCodelistItem(inslatCodeTemp).ParentValue;
                    inslatCode = metadataManager.GetCodelistInfo("InsulationType", "REFDAT").GetCodelistItem(inslatCodeTemp).ShortDisplayName.Trim();

                    // Format the insulation code to match what we have in the excel sheet.
                    if (inslatCode == "Heat conservation")
                        inslatCode = "HC";
                    else if (inslatCode == "Cold conservation")
                        inslatCode = "CC";
                    else if (inslatCode == "Fire proofing")
                        inslatCode = "FP";
                    else if (inslatCode == "Safety")
                        inslatCode = "PP";
                }

                string lineMaterial = customSupportDefinition.SupportedHelper.SupportedObjectInfo(1).MaterialCategory;

                if (lineMaterial == "Carbon Steels")
                    lineMaterial = "CarbonSteel";
                else
                {
                    if (lineMaterial == "Stainless Steels")
                        lineMaterial = "StainlessSteel";
                    else
                        lineMaterial = "AlloySteel";
                }

                // Get the pipe temperature
                double pipeTemp = pipeInfo.MaxDesignTemperature;
                ReadOnlyCollection<IPort> portcollection;
                IPipePort pipeport;
                int scheduleValue = 0;
                string schedule = string.Empty;
                double pipeWallThickness = 0;
                if (genPart.GetType().FullName.Equals("Ingr.SP3D.Route.Middle.PipeComponent"))
                {
                    Ingr.SP3D.Route.Middle.PipeComponent pipecomponent = (Ingr.SP3D.Route.Middle.PipeComponent)genPart;
                    portcollection = pipecomponent.GetConnectedPorts(PortType.Piping);
                    pipeport = (IPipePort)portcollection[1];
                    scheduleValue = pipeport.ScheduleThickness;
                    schedule = metadataManager.GetCodelistInfo("ScheduleThickness", "REFDAT").GetCodelistItem(scheduleValue).ShortDisplayName.Trim();
                    pipeWallThickness = pipeport.WallThicknessOrGrooveSetback;
                }
                else
                {
                    Ingr.SP3D.Route.Middle.PipeStockPart pipeStockPart = (Ingr.SP3D.Route.Middle.PipeStockPart)genPart;
                    portcollection = pipeStockPart.GetConnectablePorts(PortType.Piping);
                    pipeport = (IPipePort)portcollection[1];
                    scheduleValue = pipeport.ScheduleThickness;
                    schedule = metadataManager.GetCodelistInfo("ScheduleThickness", "REFDAT").GetCodelistItem(scheduleValue).ShortDisplayName.Trim();
                    pipeWallThickness = pipeport.WallThicknessOrGrooveSetback;
                }


                string schedualCompare = string.Empty;
                double wallThickness = 0;
                if (schedule.ToUpper() == "UNDEFINED")
                    schedualCompare = "Unknown";
                else if (schedule.ToUpper() != "S-STD")
                {
                    catalogBaseHelper = new CatalogBaseHelper();
                    ursWGFamilyAux = (PartClass)catalogBaseHelper.GetPartClass("REFDATPlainPipeEndData");
                    ursWGFamilyItems = ursWGFamilyAux.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    if (ursWGFamilyItems != null)
                    {
                        foreach (BusinessObject part1 in ursWGFamilyItems)
                        {
                            if ((HgrCompareDoubleService.cmpdbl((double)((PropertyValueDouble)part1.GetPropertyValue("IJPlainPipeEndData", "NominalPipingDiameter")).PropValue, primaryPipeSize) == true) && ((string)((PropertyValueString)part1.GetPropertyValue("IJPlainPipeEndData", "Schedule")).PropValue == "100"))
                            {
                                wallThickness = (double)((PropertyValueDouble)part1.GetPropertyValue("IJPlainPipeEndData", "WallThickness")).PropValue;
                                if (pipeWallThickness < wallThickness)
                                    schedualCompare = "LStdWeight";
                                else
                                    schedualCompare = "GeStdWeight";
                            }
                        }
                    }
                }
                else
                    schedualCompare = "GeStdWeight";

                string tempCode = string.Empty;
                // these will compare the tempature of the pipe with the allowed tempature.  We will be using the Kelvin units of mesure to compare with.
                if ((string.IsNullOrEmpty(t650) || t650.ToUpper() != "YES") && (string.IsNullOrEmpty(gt650) || gt650.ToUpper() != "YES"))
                {
                    if (pipeTemp <= 422.0389)
                        tempCode = "T300";
                    else if (pipeTemp > 422.0389 && pipeTemp <= 672.0389)
                        tempCode = "T750";
                    else
                        tempCode = "GT750";
                }
                else
                {
                    if (pipeTemp <= 422.0389)
                        tempCode = "T300";
                    else if (pipeTemp > 422.0389 && pipeTemp <= 616.4833)
                        tempCode = "T650";
                    else
                        tempCode = "GT650";
                }

                // Added a check for Unknown because the schedule is 'Undefined' for Tee's - RCM
                if (schedualCompare == "Unknown")
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Wall Thickness: Unable to determine Wall Thickness.", "", "FrameAssemblyServices.cs", 659);
                else if (schedualCompare == "LStdWeight")
                {
                    if (string.IsNullOrEmpty(lStdWeight) || lStdWeight.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(lStdWeight) && lStdWeight.Trim() != "")
                        {
                            if (lStdWeight.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Wall Thickness.", "", "FrameAssemblyServices.cs", 667);
                            else if (lStdWeight.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Wall Thickness.", "", "FrameAssemblyServices.cs", 669);
                        }
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(geStdWeight) || geStdWeight.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(geStdWeight) && geStdWeight.Trim() != "")
                        {
                            if (geStdWeight.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Wall Thickness.", "", "FrameAssemblyServices.cs", 680);
                            else if (geStdWeight.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Wall Thickness.", "", "FrameAssemblyServices.cs", 682);
                        }
                    }
                }

                if (lineMaterial == "CarbonSteel")
                {
                    if (string.IsNullOrEmpty(carbonSteel) || carbonSteel.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(carbonSteel) && carbonSteel.Trim() != "")
                        {
                            if (carbonSteel.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Line Material.", "", "FrameAssemblyServices.cs", 694);
                            else if (carbonSteel.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Line Material.", "", "FrameAssemblyServices.cs", 696);
                        }
                    }
                }
                else if (lineMaterial == "StainlessSteel")
                {
                    if (string.IsNullOrEmpty(stainlessSteel) || stainlessSteel.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(stainlessSteel) && stainlessSteel.Trim() != "")
                        {
                            if (stainlessSteel.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Line Material.", "", "FrameAssemblyServices.cs", 707);
                            else if (stainlessSteel.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Line Material.", "", "FrameAssemblyServices.cs", 709);
                        }
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(alloySteel) || alloySteel.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(alloySteel) && alloySteel.Trim() != "")
                        {
                            if (alloySteel.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Line Material.", "", "FrameAssemblyServices.cs", 720);
                            else if (alloySteel.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Line Material.", "", "FrameAssemblyServices.cs", 722);
                        }
                    }
                }

                if (!((!string.IsNullOrEmpty(LE3) && LE3.ToUpper() == "NO") && (!string.IsNullOrEmpty(G3) && G3.ToUpper() == "NO") && HgrCompareDoubleService.cmpdbl(inslatThickness, 0) == true))
                //{
                //}
                //else
                {
                    if (insulCompare == "LE3")
                    {
                        if (string.IsNullOrEmpty(LE3) || LE3.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(LE3) && LE3.Trim() != "")
                            {
                                if (LE3.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Thickness.", "", "FrameAssemblyServices.cs", 739);
                                else if (LE3.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Thickness.", "", "FrameAssemblyServices.cs", 741);
                            }
                        }
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(G3) || G3.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(G3) && G3.Trim() != "")
                            {
                                if (G3.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Thickness.", "", "FrameAssemblyServices.cs", 752);
                                else if (G3.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Thickness.", "", "FrameAssemblyServices.cs", 754);
                            }
                        }
                    }

                    if (inslatCode == "HC")
                    {
                        if (string.IsNullOrEmpty(HC) || HC.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(HC) && HC.Trim() != "")
                            {
                                if (HC.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Code.", "", "FrameAssemblyServices.cs", 766);
                                else if (HC.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Code.", "", "FrameAssemblyServices.cs", 768);

                            }
                        }
                    }
                    else if (inslatCode == "CC")
                    {
                        if (string.IsNullOrEmpty(CC) || CC.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(CC) && CC.Trim() != "")
                            {
                                if (CC.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Code.", "", "FrameAssemblyServices.cs", 780);
                                else if (CC.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Code.", "", "FrameAssemblyServices.cs", 782);

                            }
                        }
                    }
                    else if (inslatCode == "FP")
                    {
                        if (string.IsNullOrEmpty(FP) || FP.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(FP) && FP.Trim() != "")
                            {
                                if (FP.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Code.", "", "FrameAssemblyServices.cs", 794);
                                else if (FP.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Code.", "", "FrameAssemblyServices.cs", 796);

                            }
                        }
                    }
                    else if (inslatCode == "PP")
                    {
                        if (string.IsNullOrEmpty(PP) || PP.ToUpper() != "YES")
                        {
                            if (!string.IsNullOrEmpty(PP) && PP.Trim() != "")
                            {
                                if (PP.ToUpper() == "NO")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Insulation Code.", "", "FrameAssemblyServices.cs", 808);
                                else if (PP.Trim() == "*")
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Insulation Code.", "", "FrameAssemblyServices.cs", 810);

                            }
                        }
                    }
                }

                if (tempCode == "T300")
                {
                    if (string.IsNullOrEmpty(t300) || t300.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(t300) && t300.Trim() != "")
                        {
                            if (t300.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Temperature.", "", "FrameAssemblyServices.cs", 824);
                            else if (t300.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Temperature.", "", "FrameAssemblyServices.cs", 826);
                        }
                    }
                }
                else if (tempCode == "T650")
                {
                    if (string.IsNullOrEmpty(t650) || t650.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(t650) && t650.Trim() != "")
                        {
                            if (t650.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Temperature.", "", "FrameAssemblyServices.cs", 837);
                            else if (t650.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Temperature.", "", "FrameAssemblyServices.cs", 839);
                        }
                    }
                }
                else if (tempCode == "GT650")
                {
                    if (string.IsNullOrEmpty(gt650) || gt650.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(gt650) && gt650.Trim() != "")
                        {
                            if (gt650.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Temperature.", "", "FrameAssemblyServices.cs", 850);
                            else if (gt650.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Temperature.", "", "FrameAssemblyServices.cs", 852);
                        }
                    }
                }
                else if (tempCode == "T750")
                {
                    if (string.IsNullOrEmpty(t750) || t750.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(t750) && t750.Trim() != "")
                        {
                            if (t750.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Temperature.", "", "FrameAssemblyServices.cs", 863);
                            else if (t750.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Temperature.", "", "FrameAssemblyServices.cs", 865);
                        }
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(gt750) || gt750.ToUpper() != "YES")
                    {
                        if (!string.IsNullOrEmpty(gt750) && gt750.Trim() != "")
                        {
                            if (gt750.ToUpper() == "NO")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Selected Support is not valid for Given Temperature.", "", "FrameAssemblyServices.cs", 876);
                            else if (gt750.Trim() == "*")
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "CheckSupportWithFamilyAndType", "USERWARNINGMESSAGE", " " + ": " + "WARNING: " + "Selected Support may not be valid for Given Temperature.", "", "FrameAssemblyServices.cs", 878);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in CheckSupportWithFamilyAndType." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// Function to add all the Implied Parts supplied in the catalog for the given Support PartNumber.
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="customSupportDefinition"></param>
        /// <param name="impliedparts">The partinfo collection</param>
        /// <param name="impliedPartsServiceClassName">The Serviceclass on which the ImpliedParts information is stored on</param>
        ///<code>
        /// Returns
        ///A collection of hsWeldData Types, containing all the information for each individual weld, including the index that the partocc can be accesed from the part occurence collection.
        ///AddImpliedPartFromCatalog(this,parts,impliedPartsServiceClassName)
        ///</code>
        public static void AddImpliedPartFromCatalog(CustomSupportDefinition customSupportDefinition, Collection<PartInfo> impliedparts, string impliedPartsServiceClassName)
        {
            try
            {

                if (impliedPartsServiceClassName != string.Empty)
                {
                    string partClassValue = string.Empty;
                    Collection<object> ImpParts = new Collection<object>();
                    string[] partKey = new string[100];
                    Collection<object> ruleResults1 = new Collection<object>();
                    string SupportNumber = customSupportDefinition.SupportHelper.Support.SupportDefinition.PartNumber;
                    CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                    PartClass auxTable = null;
                    string partnumber = "";

                    if (impliedPartsServiceClassName != null)
                    {
                        try
                        {
                            auxTable = (PartClass)cataloghelper.GetPartClass(impliedPartsServiceClassName);
                        }
                        catch
                        {
                            auxTable = null;
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "AddImpliedPartFromCatalog", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "The given Implied parts Service Class does not exist in Database.", "", "FrameAssemblyServices.cs", 4010);
                        }
                    }
                    if (auxTable != null)
                    {
                        ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        foreach (BusinessObject classItem in classItems)
                        {
                            bool isEqual = String.Equals(SupportNumber, ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsFrImpPart", "SupportPartNumber")).PropValue), StringComparison.Ordinal);
                            if (isEqual == true)
                            {
                                string rule = ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsFrImpPart", "ImpPartRule")).PropValue);
                                if (rule == "" || rule == string.Empty)
                                    rule = null;
                                if (rule != null)
                                {
                                    customSupportDefinition.GenericHelper.GetDataByRule(rule, null, out ruleResults1);
                                    if (ruleResults1 != null)
                                    {
                                        partKey = new string[ruleResults1.Count];
                                        for (int index = 0; index < ruleResults1.Count; index++)
                                        {
                                            partnumber = ruleResults1[index].ToString();
                                            partKey[index] = "IMPLIEDPART_" + index;
                                            impliedparts.Add(new PartInfo(partKey[index], partnumber, null));
                                        }
                                    }
                                }
                                else
                                {
                                    partnumber = ((string)((PropertyValueString)classItem.GetPropertyValue("IJUAhsFrImpPart", "ImpPartNumber")).PropValue);
                                    int quantity = ((int)((PropertyValueInt)classItem.GetPropertyValue("IJUAhsFrImpPart", "ImpPartQuantity")).PropValue);
                                    partKey = new string[quantity];
                                    for (int index = 0; index < quantity; index++)
                                    {
                                        partKey[index] = "IMPLIEDPART_" + index;
                                        impliedparts.Add(new PartInfo(partKey[index], partnumber, null));
                                    }
                                }
                            }

                        }
                    }
                }
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in Get Implied Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method will be called to check if property is available on the BusinessObject passed and return its PropertValue.
        /// </summary>
        /// <param name="businessObject">Support or SupportOccurence or SupportComponent or SupportComponentOccurence of type bussiness object</param>
        /// <param name="interfaceName">Interface to which the property belongs.</param>
        /// <param name="propertyName">Name of property.</param>
        /// <param name="propertyQueryType">Specifies whether the property is to be queried on Occurence or Catalog or Both, the default value passed is Both.</param>
        /// <code>
        /// GetPropertyValue(businessObject, "IJUAhsHeight1","Height1")
        /// </code>
        public static PropertyValue GetPropertyValue(BusinessObject businessObject, string interfaceName, string propertyName, QueryPropertyOn propertyQueryType = QueryPropertyOn.Both)
        {
            PropertyValue returnPropertyValue = null;
            try
            {
                returnPropertyValue = GetPropertyValueInternal(businessObject, interfaceName, propertyName, propertyQueryType);
                if (returnPropertyValue == null)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "GetPropertyValue", "USERERRORMESSAGE", " " + ": " + "ERROR: " + "Unable to get the value of " + propertyName + " property on " + interfaceName + " interface", "", "FrameAssemblyServices.cs", 4189);
                }
                else if (businessObject != null && returnPropertyValue != null)
                {
                    PropertyEmptyCheck(returnPropertyValue);
                }
            }
            catch (Exception e)
            {
                Type myType = typeof(FrameAssemblyServices);
                CmnException e1 = new CmnException("Unable to get the value of " + propertyName + " property on" + interfaceName + " interface" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);

                throw e1;
            }
            return (returnPropertyValue);
        }
        /// <summary>
        /// This method will be called to check property is available on the SupportOccurrence(PartOccurrence) and Support(Part) and return its property value.
        /// </summary>
        /// <param name="businessObject">Support or SupportOccurence or SupportComponent or SupportComponentOccurence of type bussiness object</param>
        /// <param name="interfaceName">Interface to which the property belongs.</param>
        /// <param name="propertyName">Name of property.</param>
        /// <param name="propertyQueryType">Specifies whether the property is to be queried on occurence or catalog or both, by default its both.</param>
        /// <returns>Returns PropertyValue</returns>
        /// <code>
        /// GetPropertyValueInternal(businessObject, "IJUAhsHeight1","Height1")
        /// </code>
        private static PropertyValue GetPropertyValueInternal(BusinessObject businessObject, string interfaceName, string propertyName, QueryPropertyOn propertyQueryType = QueryPropertyOn.Both)
        {
            PropertyValue propertyValue = null;
            try
            {
                if (propertyQueryType == QueryPropertyOn.CatalogOnly)
                {
                    if (businessObject is Ingr.SP3D.ReferenceData.Middle.Part)
                    {
                        if (businessObject.SupportsInterface(interfaceName))
                            propertyValue = businessObject.GetPropertyValue(interfaceName, propertyName);
                    }
                    else if (businessObject is Ingr.SP3D.Support.Middle.Support)
                    {
                        BusinessObject supportPart = businessObject.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                        if (supportPart.SupportsInterface(interfaceName))
                            propertyValue = supportPart.GetPropertyValue(interfaceName, propertyName);
                    }
                    else if (businessObject is Ingr.SP3D.Support.Middle.SupportComponent)
                    {
                        BusinessObject supportComponentPart = businessObject.GetRelationship("madeFrom", "part").TargetObjects[0];
                        if (supportComponentPart.SupportsInterface(interfaceName))
                            propertyValue = supportComponentPart.GetPropertyValue(interfaceName, propertyName);
                    }
                    else if (businessObject is BusinessObject)
                    {
                        if (businessObject.SupportsInterface(interfaceName))
                            propertyValue = businessObject.GetPropertyValue(interfaceName, propertyName);
                    }
                }
                else if (propertyQueryType == QueryPropertyOn.OccurrenceOnly)
                {
                    if (businessObject is Ingr.SP3D.Support.Middle.Support)
                    {
                        if (businessObject.SupportsInterface(interfaceName))
                            propertyValue = businessObject.GetPropertyValue(interfaceName, propertyName);
                    }
                    else if (businessObject is Ingr.SP3D.Support.Middle.SupportComponent)
                    {
                        if (businessObject.SupportsInterface(interfaceName))
                            propertyValue = businessObject.GetPropertyValue(interfaceName, propertyName);
                    }
                }
                else
                {
                    if (businessObject is Ingr.SP3D.ReferenceData.Middle.Part)
                    {
                        if (businessObject.SupportsInterface(interfaceName))
                            propertyValue = businessObject.GetPropertyValue(interfaceName, propertyName);
                    }
                    else if (businessObject is Ingr.SP3D.Support.Middle.Support)
                    {

                        if (businessObject.SupportsInterface(interfaceName))
                            propertyValue = businessObject.GetPropertyValue(interfaceName, propertyName);
                        else
                        {
                            BusinessObject supportPart = businessObject.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];
                            if (supportPart.SupportsInterface(interfaceName))
                                propertyValue = supportPart.GetPropertyValue(interfaceName, propertyName);
                        }
                    }
                    else if (businessObject is Ingr.SP3D.Support.Middle.SupportComponent)
                    {
                        if (businessObject.SupportsInterface(interfaceName))
                            propertyValue = businessObject.GetPropertyValue(interfaceName, propertyName);
                        else
                        {
                            BusinessObject supportComponentPart = businessObject.GetRelationship("madeFrom", "part").TargetObjects[0];
                            if (supportComponentPart.SupportsInterface(interfaceName))
                                propertyValue = supportComponentPart.GetPropertyValue(interfaceName, propertyName);
                        }
                    }
                    else if (businessObject is BusinessObject)
                    {
                        if (businessObject.SupportsInterface(interfaceName))
                            propertyValue = businessObject.GetPropertyValue(interfaceName, propertyName);
                    }
                }
            }
            catch (Ingr.SP3D.Common.Exceptions.CmnPropertyInfoNotAvailableException ex)
            {
                Type myType = typeof(FrameAssemblyServices);
                CmnException e1 = new CmnException("Invalid Attribute : " + propertyName + "Queried on" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + ex.Message, ex);
                throw e1;
            }
            catch (Ingr.SP3D.Common.Exceptions.CmnInterfaceInfoNotAvailableException ex)
            {
                Type myType = typeof(FrameAssemblyServices);
                CmnException e1 = new CmnException("Invalid Interface : " + propertyName + "Queried on" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + ex.Message, ex);
                throw e1;
            }
            catch (Exception e)
            {
                Type myType = typeof(FrameAssemblyServices);
                CmnException e1 = new CmnException("Unable to get the value of " + propertyName + " property on" + interfaceName + " interface" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }

            return propertyValue;
        }
        /// <summary>
        /// This method will be called to check if property value provided is null in th catalag and assaigns a default value for it
        /// </summary>
        /// <param name="propertyvalue"> takes PropertyValue as parameter</param>
        /// <code>
        /// PropertyEmptyCheck(propertyvalue)
        /// </code>
        public static PropertyValue PropertyEmptyCheck(PropertyValue propertyvalue)
        {
            try
            {
                ValueType propertyValueType = propertyvalue.PropertyInfo.PropertyType;
                double doubleValue = 0;
                int intValue = 0;
                bool boolValue = false;

                if (propertyValueType.ToString() == "PTDouble")
                {
                    try
                    {
                        doubleValue = (double)((PropertyValueDouble)propertyvalue).PropValue;
                    }
                    catch
                    {
                        ((PropertyValueDouble)propertyvalue).PropValue = 0;
                    }
                    return propertyvalue;
                }
                else if (propertyValueType.ToString() == "PTInteger")
                {
                    try
                    {
                        intValue = (int)((PropertyValueInt)propertyvalue).PropValue;
                    }
                    catch
                    {
                        ((PropertyValueInt)propertyvalue).PropValue = 0;
                    }
                    return propertyvalue;
                }
                else if (propertyValueType.ToString() == "PTBool")
                {
                    try
                    {
                        boolValue = (bool)((PropertyValueBoolean)propertyvalue).PropValue;
                    }
                    catch
                    {
                        ((PropertyValueBoolean)propertyvalue).PropValue = false;
                    }
                    return propertyvalue;
                }
                else
                    return propertyvalue;
            }
            catch (Exception e)
            {
                Type myType = typeof(FrameAssemblyServices);
                CmnException e1 = new CmnException("Unable to get the value property on the given interface" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);

                throw e1;
            }
        }
    }
}
