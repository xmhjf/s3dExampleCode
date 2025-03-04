//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   GenericAssyRules.cs
//   Author       :Vijaya
//   Creation Date:1.Oct.2013 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  1.Oct.2013     Vijaya   CR-CP-240910 	Convert HgrAttributesRules to C# .Net   
//  12.aug.2014     PVK     TR-CP-257296 	Resolve coverity issues found in June 2014 report 
//  31.Oct.2014     PVK     TR-CP-260301	Resolve coverity issues found in August 2014 report
//  09.Feb.2015     Siva    TR-CP-261379    When a Conduit support type is is changed, it is not re-placed correctly
//  26.Oct.2015     PVK       		        Resolve coverity issues found in Octpber 2015 report
//  30.Nov.2015     PVK      DI-CP-276798	Replace the use of any HS_Utility parts
//  21.Mar.2016     PVK      TR-CP-288920	Issues found in HS_Assembly_V2
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Linq;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Route.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    //----------------------------------------------------------------------
    //This Rule retuns duct angle offset value.
    //ProgId : AttributesRules,Ingr.SP3D.Content.Support.Rules.GetDuctAngleOffset
    //----------------------------------------------------------------------     
    public class GetDuctAngleOffset : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            GenericHelper genericHelper = new GenericHelper(support);
            Collection<object> valuesByRule = new Collection<object>();
            string primaryCSStandard = string.Empty, primarySectionType = string.Empty, primaryCS = string.Empty;
            PropertyValue legEndOffset, legRoutegap;

            genericHelper.GetDataByRule("HgrSupPrimaryCSSelection", (BusinessObject)support, out valuesByRule);
            if (valuesByRule != null)
            {
                primaryCSStandard = (string)valuesByRule[0];
                primarySectionType = (string)valuesByRule[1];
                primaryCS = (string)valuesByRule[2];
            }

            Dictionary<string, object> parameter = new Dictionary<String, Object>();
            parameter.Add("SectionName", primaryCS.Trim());
            parameter.Add("SectionStandard", primaryCSStandard.Trim());
            parameter.Add("ClassName", primarySectionType.Trim());
            //get offset D and G                                    
            genericHelper.GetDataByRule("HS_HgrServ_OffsetAngSize", "IJUAHngServ_OffsetByAngleSize", "D", parameter, out legEndOffset);
            genericHelper.GetDataByRule("HS_HgrServ_OffsetAngSize", "IJUAHngServ_OffsetByAngleSize", "G", parameter, out legRoutegap);
            double[] attributeValues = new double[2];

            attributeValues[0] = Convert.ToDouble(((PropertyValueDouble)legEndOffset).PropValue);
            attributeValues[1] = Convert.ToDouble(((PropertyValueDouble)legRoutegap).PropValue);
            return attributeValues;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns part based on cross section data.
    //ProgId : AttributesRules,Ingr.SP3D.Content.Support.Rules.GetPrimaryCS
    //---------------------------------------------------------------------- 
    public class GetPrimaryCS : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            SupportedHelper supportedHelper = new SupportedHelper(support);
            SupportHelper supportHelper = new SupportHelper(support);
            BusinessObject supportedObject = supportHelper.SupportedObjects[0];
            RefPortHelper refPortHelper = new RefPortHelper(support);
            IPipePathFeature pipeInfo = supportedObject as IPipePathFeature;
            double width = 0.0, height = 0.0;
            NominalDiameter pipeDiameter = new NominalDiameter();
            if (pipeInfo != null)
            {
                pipeDiameter.Size = pipeInfo.NPD.Size;
                width = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, pipeDiameter.Size, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
            }
            IConduitPathFeature conduitInfo = supportedObject as IConduitPathFeature;
            if (conduitInfo != null)
            {
                pipeDiameter.Size = conduitInfo.NCD.Size;
                width = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, pipeDiameter.Size, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
            }
            IRouteFeatureWithCrossSection routeInfo = supportedObject as IRouteFeatureWithCrossSection;
            if (routeInfo != null)
                width = routeInfo.Width;
          
            if (supportHelper.SupportingObjects.Count > 0)
                height = refPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Direct);
            else
                height = 1.1;

            //some code has to be here
            string[] attributeValues = new string[3];
            IEnumerable<BusinessObject> parts = null;
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("HS_HgrServ_AngByDistance");
            parts = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            parts = parts.Where(part => (((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHngServ_AngleByDist", "NominalWidthFrom")).PropValue) < width && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHngServ_AngleByDist", "NominalWidthTo")).PropValue) >= width && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHngServ_AngleByDist", "NominalHeightFrom")).PropValue) < height && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHngServ_AngleByDist", "NominalHeightTo")).PropValue) >= height));
            if (parts.Count() > 0)
            {
                attributeValues[0] = (string)((PropertyValueString)parts.ElementAt(0).GetPropertyValue("IJUAHngServ_AngleByDist", "SectionStandard")).PropValue;
                attributeValues[1] = (string)((PropertyValueString)parts.ElementAt(0).GetPropertyValue("IJUAHngServ_AngleByDist", "ClassName")).PropValue;
                attributeValues[2] = (string)((PropertyValueString)parts.ElementAt(0).GetPropertyValue("IJUAHngServ_AngleByDist", "SectionName")).PropValue;
            }
            else
                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "GetPrimaryCS" + ": " + "ERROR: " + "Route - Struct distance does not fall in the range of Minimum height and Maximum height.", "", "AttributeRules.cs", 1);

            GC.SuppressFinalize(parts);
            return attributeValues;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns part based on cross section data.
    //ProgId : AttributesRules,Ingr.SP3D.Content.Support.Rules.GetStructCS
    //----------------------------------------------------------------------    
    public class GetStructCS : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            SupportedHelper supportedHelper = new SupportedHelper(support);
            PipeObjectInfo pipe = (PipeObjectInfo)supportedHelper.SupportedObjectInfo(1);
            RefPortHelper refPortHelper = new RefPortHelper(support);
            GenericHelper genericHelper = new GenericHelper(support);
            double pipeDiameter = 0.0;

            pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, pipe.NominalDiameter.Size, UnitName.DISTANCE_METER);
            //Determine the Distance between Route and Structure
            double routeStructDistance = refPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Direct);
                   
            //Since the Part has a relation with a crosssection, the existing PartSelection methods
            //are not suitable to determine a Part with the SectionName, hence Hardcoding the PartNumbers.
            Collection<object> valuesCollection = new Collection<object>();
            string[] attributeValues = new string[3];
            try
            {
                bool value = genericHelper.GetDataByRule("HgrAISCversion", null, out valuesCollection);
            }
            catch { }

            if (valuesCollection != null)
            {
                if (string.IsNullOrEmpty((string)valuesCollection[0]))
                    attributeValues[1] = "AISC-LRFD-3.0";
                else
                    attributeValues[1] = (string)valuesCollection[0];
            }
            if (routeStructDistance >= 2.95 && routeStructDistance <= 3.05)
            {
                //3 inch Standard Shoe Size
                if (pipeDiameter >= 0.75 && pipeDiameter <= 6.5)
                {

                    attributeValues[0] = "G4G_1410_01_P01";
                    attributeValues[2] = "ST";
                }
                else if (pipeDiameter >= 8.00 && pipeDiameter <= 11.00)
                {
                    attributeValues[0] = "G4G_1410_01_P02";
                    attributeValues[2] = "WT";
                }
            }
            else if (routeStructDistance >= 3.95 && routeStructDistance <= 4.05)
            {
                //4 inch Standard Shoe Size
                if (pipeDiameter >= 0.75 && pipeDiameter <= 6.5)
                {
                    attributeValues[0] = "G4G_1410_01_P07";
                    attributeValues[2] = "ST";
                }
                else if (pipeDiameter >= 8.00 && pipeDiameter <= 11.00)
                {

                    attributeValues[0] = "G4G_1410_01_P08";
                    attributeValues[2] = "WT";
                }
            }
            else if (routeStructDistance >= 4.95 && routeStructDistance <= 5.05)
            {
                //5 inch Standard Shoe Size
                if (pipeDiameter >= 0.5 && pipeDiameter <= 6.5)
                {
                    attributeValues[0] = "G4G_1410_01_P13";
                    attributeValues[2] = "ST";
                }
                else if (pipeDiameter >= 8.00 && pipeDiameter <= 11)
                {

                    attributeValues[0] = "G4G_1410_01_P14";
                    attributeValues[2] = "WT";
                }
            }
            return attributeValues;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns part based on cross section data.
    //ProgId : AttributesRules,Ingr.SP3D.Content.Support.Rules.GetStructOffset
    //----------------------------------------------------------------------  
    public class GetStructOffset : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            double[] attribute = new double[] { 0.05 };
            return attribute;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns part based on cross section data.
    //ProgId : AttributesRules,Ingr.SP3D.Content.Support.Rules.GetStructPars
    //----------------------------------------------------------------------     
    public class GetStructPars : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            SupportHelper supportHelper = new SupportHelper(support);
            int suppotedObjectCount = supportHelper.SupportedObjects.Count;
            GenericHelper genericHelper = new GenericHelper(support);
            double pipeRadius = 0.0, width = 0.0, maximumWidth = 0.0;
            Collection<BusinessObject> routeObjectCollection = supportHelper.SupportedObjects;
            for (int index = 0; index < suppotedObjectCount; index++)
            {
                IPipePathFeature pipeInfo = routeObjectCollection[index] as IPipePathFeature;
                NominalDiameter pipeDiameter = new NominalDiameter();
                if (pipeInfo != null)
                {
                    pipeDiameter.Size = pipeInfo.NPD.Size;
                    pipeRadius = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, pipeDiameter.Size, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    width = 2 * pipeRadius;
                }
                IConduitPathFeature conduitInfo = routeObjectCollection[index] as IConduitPathFeature;
                if (conduitInfo != null)
                {
                    pipeDiameter.Size = conduitInfo.NCD.Size;
                    pipeRadius = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, pipeDiameter.Size, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    width = 2 * pipeRadius;
                }
                IRouteFeatureWithCrossSection routeInfo = routeObjectCollection[index] as IRouteFeatureWithCrossSection;
                if (routeInfo != null)
                    width = routeInfo.Width;
                if (width > maximumWidth)
                    maximumWidth = width;
            }
            //Testing data
            Collection<object> valuesCollection = new Collection<object>();
            object[] attributeValues = new object[10];
            try
            {
                bool value = genericHelper.GetDataByRule("HgrAISCversion", null, out valuesCollection);
            }
            catch { }
            if (valuesCollection != null)
            {
                if (string.IsNullOrEmpty((string)valuesCollection[0]))
                    attributeValues[0] = "AISC-LRFD-3.0";
                else
                    attributeValues[0] = (string)valuesCollection[0];
            }
            if (maximumWidth <= 1.00)
            {
                attributeValues[1] = "L";
                attributeValues[2] = "L3x3x5/16";     // SHAPE
                attributeValues[3] = 2.00;               // A
                attributeValues[4] = 27.00;              // B min
                attributeValues[5] = 48.00;              // B max
                attributeValues[6] = 6.00;             // P
                attributeValues[7] = 10.5;           // C
                attributeValues[8] = 1.5;            // D
                attributeValues[9] = 1.75;           // E
            }
            else if (maximumWidth > 1.00 && maximumWidth <= 2.00)
            {
                attributeValues[1] = "L";
                attributeValues[2] = "L3x3x5/16";       // SHAPE
                attributeValues[3] = 2.5;
                attributeValues[4] = 27.00;
                attributeValues[5] = 48.00;
                attributeValues[6] = 6.00;
                attributeValues[7] = 10.5;
                attributeValues[8] = 1.5;
                attributeValues[9] = 1.75;
            }
            else if (maximumWidth > 2.00 && maximumWidth <= 4.00)
            {
                attributeValues[1] = "L";
                attributeValues[2] = "L3x3x5/16";    // SHAPE
                attributeValues[3] = 3.5;
                attributeValues[4] = 26.00;
                attributeValues[5] = 48.00;
                attributeValues[6] = 6.00;
                attributeValues[7] = 10.5;
                attributeValues[8] = 1.5;
                attributeValues[9] = 1.75;
            }
            else
            {
                attributeValues[1] = "L";
                attributeValues[2] = "L4x4x3/8";
                attributeValues[3] = 5.75;
                attributeValues[4] = 24.00;
                attributeValues[5] = 48.00;
                attributeValues[6] = 8.00;
                attributeValues[7] = 12.5;
                attributeValues[8] = 2.00;
                attributeValues[9] = 2.00;
            }
            for (int index = 4; index < attributeValues.Length; index++)
            {
                attributeValues[index] = MiddleServiceProvider.UOMMgr.ConvertUnitToDBU(UnitType.Distance, (double)attributeValues[index], UnitName.DISTANCE_INCH);
            }
                return attributeValues;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns part based on cross section data.
    //ProgId : AttributesRules,Ingr.SP3D.Content.Support.Rules.GetAnglePartByLF
    //----------------------------------------------------------------------     
    public class GetAnglePartByLF : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            SupportHelper supportHelper = new SupportHelper(support);
            int suppotedObjectCount = supportHelper.SupportedObjects.Count;
            double pipeDiameter = 0.0;
            Collection<BusinessObject> routeObjectCollection = supportHelper.SupportedObjects;
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

            IEnumerable<BusinessObject> parts = null;
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("HS_HgrServ_SectLdFactor");
            parts = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            string pipeSpec = "HYDOIL1";
            double loadFactor = 0.0;
            for (int index = 0; index < suppotedObjectCount; index++)
            {
                IPipePathFeature pipeInfo = routeObjectCollection[index] as IPipePathFeature;
                NominalDiameter diameter = new NominalDiameter();
                double result = 0.0;
                if (pipeInfo != null)
                {
                    diameter.Size = pipeInfo.NPD.Size;
                    if (pipeInfo.NPD.Units == "in")
                        pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, diameter.Size, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    else if (pipeInfo.NPD.Units == "mm")
                        pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, diameter.Size, UnitName.DISTANCE_MILLIMETER, UnitName.DISTANCE_METER);
                }
                IConduitPathFeature conduitInfo = routeObjectCollection[index] as IConduitPathFeature;
                
                if (conduitInfo != null)
                {
                    diameter.Size = conduitInfo.NCD.Size;
                    if (conduitInfo.NCD.Units == "in")
                        pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, diameter.Size, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);
                    else if (conduitInfo.NCD.Units == "mm")
                        pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, diameter.Size, UnitName.DISTANCE_MILLIMETER, UnitName.DISTANCE_METER);
                }

                IRouteFeatureWithCrossSection routeInfo = routeObjectCollection[index] as IRouteFeatureWithCrossSection;
                if (routeInfo != null)
                    pipeDiameter = routeInfo.Width;

                //build up the query              
                parts = parts.Where(part => (((string)((PropertyValueString)part.GetPropertyValue("IJUAHngServ_SectionLoadFactor", "PipeSpec")).PropValue) == pipeSpec && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHngServ_SectionLoadFactor", "NominalPipeDiameter")).PropValue) >= pipeDiameter));
                if (parts.Count() > 0)
                    result = (double)((PropertyValueDouble)parts.ElementAt(0).GetPropertyValue("IJUAHngServ_SectionLoadFactor", "LoadFactor")).PropValue;
                loadFactor = result + loadFactor;
            }
            object[] attributeValues = new object[5];
            parts = null;
            parts = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            parts = parts.Where(part => (((string)((PropertyValueString)part.GetPropertyValue("IJUAHngServ_SectionLoadFactor", "PipeSpec")).PropValue) == pipeSpec && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHngServ_SectionLoadFactor", "LoadFactor")).PropValue) >= loadFactor));
            if (parts.Count() > 0)
            {
                attributeValues[0] = (string)((PropertyValueString)parts.ElementAt(0).GetPropertyValue("IJUAHngServ_SectionLoadFactor", "SectionStandard")).PropValue;
                attributeValues[1] = (string)((PropertyValueString)parts.ElementAt(0).GetPropertyValue("IJUAHngServ_SectionLoadFactor", "ClassName")).PropValue;
                attributeValues[2] = (string)((PropertyValueString)parts.ElementAt(0).GetPropertyValue("IJUAHngServ_SectionLoadFactor", "SectionName")).PropValue;
                attributeValues[3] = (double)((PropertyValueDouble)parts.ElementAt(0).GetPropertyValue("IJUAHngServ_SectionLoadFactor", "Beam_W")).PropValue;
                attributeValues[4] = (double)((PropertyValueDouble)parts.ElementAt(0).GetPropertyValue("IJUAHngServ_SectionLoadFactor", "Beam_L")).PropValue;
            }
            GC.SuppressFinalize(parts);
            return attributeValues;
        }
    }

    public class FeatureType : IHgrRule
    {
        public Array AttributeValue(BusinessObject support)
        {
            object[] attributeValues = new object[1];

            Ingr.SP3D.Support.Middle.Support oSupport = (Ingr.SP3D.Support.Middle.Support)support;
            SupportHelper supportHelper = new SupportHelper(oSupport);
            
            attributeValues[0] = "UnKnown";

            Collection<BusinessObject> routeObjectCollection = supportHelper.SupportedObjects;

            if (routeObjectCollection != null && routeObjectCollection.Count > 0)
            {
                if ((routeObjectCollection[0].SupportsInterface("IJRteStraightPathFeat")) || (routeObjectCollection[0].SupportsInterface("IJRteCurvePathFeat")))
                {
                    attributeValues[0] = "STRAIGHT";
                }
                else if (routeObjectCollection[0].SupportsInterface("IJRteTurnPathFeat"))
                {
                    attributeValues[0] = "TURN";
                }
                else if (routeObjectCollection[0].SupportsInterface("IJRteAlongLegPathFeat"))
                {
                    attributeValues[0] = "PART";
                }
                else if (routeObjectCollection[0].SupportsInterface("IJRteEndPathFeat"))
                {
                    attributeValues[0] = "END";
                }
                else if (routeObjectCollection[0].SupportsInterface("IJRteSurfMountPathFeat"))
                {
                    attributeValues[0] = "SURFACE";
                }
                else
                {
                    attributeValues[0] = "UnKnown";
                }
            }
            
            return attributeValues;
        }
    }

}
