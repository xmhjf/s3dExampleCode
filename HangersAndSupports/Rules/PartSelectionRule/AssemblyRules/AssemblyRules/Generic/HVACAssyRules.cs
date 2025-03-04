//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014, Intergraph Corporation. All rights reserved.
//
//   GenericAssyRules.cs
//   Author       :Vijaya
//   Creation Date:30.July.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  18.Sep.2013     Vijaya   CR-CP-233075  Convert HgrHVACAssyRules to C# .Net    
//  12.aug.2014     PVK      TR-CP-257296 	Resolve coverity issues found in June 2014 report 
//  31.Oct.2014     PVK      TR-CP-260301	Resolve coverity issues found in August 2014 report
//  26.Oct.2015     PVK      Resolve coverity issues found in Octpber 2015 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Linq;
using Ingr.SP3D.Support.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Route.Middle;

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
    //This Rule retuns array of width and height offset values.
    //ProgId : AssemblyRules,Ingr.SP3D.Content.Support.Rules.BoltSize
    //----------------------------------------------------------------------
    public class BoltSize : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;
            SupportedHelper supportedHelper = new SupportedHelper(support);
            SupportHelper supportHelper = new SupportHelper(support);
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();          
            double maximumDuctSize = 0.0, ductSize = 0.0;  
            BusinessObject supportedObject = supportHelper.SupportedObjects[0];
            for (int index = 1; index <= supportHelper.SupportedObjects.Count; index++)
            {
                 supportedObject = supportHelper.SupportedObjects[index-1];
                 IRouteFeatureWithCrossSection ductInfo = supportedObject as IRouteFeatureWithCrossSection;
                 if (ductInfo != null)
                 {
                     if (ductInfo.CrossSectionShape == 1 )
                         ductSize = ductInfo.Width;
                     else if (ductInfo.CrossSectionShape == 4 )
                         ductSize = ductInfo.Width / 2;
                 }
                if (ductSize > maximumDuctSize)
                    maximumDuctSize = ductSize;
            }        
            maximumDuctSize = maximumDuctSize * 1000;
            //get Section Type
            IEnumerable<BusinessObject> hvacParts = null;
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("HVACAssy_HgrBoltSize");
            hvacParts = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            string[] boltSize = new string[1];

            hvacParts = hvacParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAHgrHVACSrvBoltSize", "UnitType")).PropValue).ToLower() == "mm" && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrHVACSrvBoltSize", "DuctSizeMin")).PropValue) < maximumDuctSize && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrHVACSrvBoltSize", "DuctSizeMax")).PropValue) > maximumDuctSize);
              if (hvacParts.Count() > 0)
                  boltSize[0] = (string)((PropertyValueString)hvacParts.ElementAt(0).GetPropertyValue("IJUAHgrHVACSrvBoltSize", "BoltSize")).PropValue;
              GC.SuppressFinalize(hvacParts);
            return boltSize;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns array of width and height offset values.
    //ProgId : AssemblyRules,Ingr.SP3D.Content.Support.Rules.HangerSectionSize
    //----------------------------------------------------------------------
    public class HangerSectionSize : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;            
            SupportHelper supportHelper = new SupportHelper(support);
            BoundingBoxHelper boundingBoxHelper = new BoundingBoxHelper(support);
            GenericHelper genericHelper = new GenericHelper(support);
            RefPortHelper refPortHelper = new RefPortHelper(support);
            boundingBoxHelper.CreateStandardBoundingBoxes(false);
            BoundingBox boundingBox;
            if (supportHelper.PlacementType == PlacementType.PlaceByStruct)
                boundingBox = boundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
            else
                boundingBox = boundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
            double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height, routeStructAngle = 0.0, routeAngle = 0.0, routeStructDistance = 0.0;
            BusinessObject supportedObject = supportHelper.SupportedObjects[0];
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();           
            string hvacShape = string.Empty, sectionSize = string.Empty, steelStadard = string.Empty;
            IRouteFeatureWithCrossSection ductInfo = supportedObject as IRouteFeatureWithCrossSection; ;

            if (ductInfo!=null)
            {
                if (ductInfo.CrossSectionShape == 1)
                    hvacShape = "Rectangular";
                else if (ductInfo.CrossSectionShape == 4)
                    hvacShape = "Round";
            }
            routeStructAngle = refPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);
            routeAngle = refPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
            if ((Math.Abs(routeAngle) < (0 + 0.001) && Math.Abs(routeAngle) > (0 - 0.001)) || (Math.Abs(routeAngle) < (Math.Round(Math.PI,7) + 0.001) && Math.Abs(routeAngle) > (Math.Round(Math.PI,7) - 0.001)))
                routeStructDistance = RouteStructDistance(supportHelper, refPortHelper, PortDistanceType.Horizontal);
            else
                if ((Math.Abs(routeStructAngle) < (0 + 0.001) && Math.Abs(routeStructAngle) > (0 - 0.001)) || (Math.Abs(routeStructAngle) < (Math.Round(Math.PI, 7) + 0.001) && Math.Abs(routeStructAngle) > (Math.Round(Math.PI, 7) - 0.001)))
                    routeStructDistance = RouteStructDistance(supportHelper, refPortHelper, PortDistanceType.Horizontal);
                else
                    routeStructDistance = RouteStructDistance(supportHelper, refPortHelper, PortDistanceType.Vertical);
            if (hvacShape.Equals("Round") && ductInfo != null)
                routeStructDistance = routeStructDistance - (ductInfo.Width/2);
            else if (hvacShape.Equals("Rectangular"))
                routeStructDistance = routeStructDistance - boundingBoxHeight / 2;
            routeStructDistance = routeStructDistance * 1000;
            //get Section Type
            IEnumerable<BusinessObject> hvacParts = null;
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("HVACAssy_HgrSectSize");
            hvacParts = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

            hvacParts = hvacParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAHgrHVACSrvSectSize", "UnitType")).PropValue).ToLower() == "mm" && (((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrHVACSrvSectSize", "Hmin")).PropValue) < routeStructDistance && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrHVACSrvSectSize", "Hmax")).PropValue) > routeStructDistance) && (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrHVACSrvSectSize", "HVACShape")).PropValue.ToUpper() == hvacShape.ToUpper());
              if (hvacParts.Count() > 0)
                  sectionSize = (string)((PropertyValueString)hvacParts.ElementAt(0).GetPropertyValue("IJUAHgrHVACSrvSectSize", "SectionSize")).PropValue;
          
            bool value = genericHelper.GetDataByRule("HgrHVACSteelStandardName", (BusinessObject)support, out steelStadard);

            partClass = (PartClass)catalogBaseHelper.GetPartClass("HgrHVACStCorrespond");
            hvacParts = null;
            hvacParts = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            string[] hangerSectionSize = new string[1];
            hvacParts = hvacParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAHgrHVACStCorrespon", "Size")).PropValue == sectionSize) && ((string)((PropertyValueString)part.GetPropertyValue("IJUAHgrHVACStCorrespon", "StdName")).PropValue.ToUpper() == steelStadard.ToUpper()));
                if (hvacParts.Count() > 0)
                    hangerSectionSize[0] = (string)((PropertyValueString)hvacParts.ElementAt(0).GetPropertyValue("IJUAHgrHVACStCorrespon", "SectionSize")).PropValue;
                GC.SuppressFinalize(hvacParts);
            return hangerSectionSize;
        }
        /// <summary>
        /// This method Gets the distance between the route and structure.
        /// </summary>
        /// <param name="supportHelper">A class, that provides methods and properties of a support object. </param>
        ///  <param name="refPortHelper">A class, that provides methods and properties of support reference ports.</param>
        ///  <param name="distanceType">Hanger port distance type</param>
        /// <returns>double</returns>
        /// <code> 
        /// routeStructDistance = RouteStructDistance(supportHelper, refPortHelper, PortDistanceType.Vertical);
        /// </code>
        public double RouteStructDistance(SupportHelper supportHelper, RefPortHelper refPortHelper, PortDistanceType distanceType)
        {

            int ductCount = supportHelper.SupportedObjects.Count;
            double distance = 0.0, maximumHorizontalDistance = 0.0;
            string routePortName = string.Empty;
            for (int routeIndex = 1; routeIndex <= ductCount; routeIndex++)
            {
                if (routeIndex == 1)
                    routePortName = "Route";
                else
                    routePortName = "Route_" + routeIndex;
                distance = refPortHelper.DistanceBetweenPorts("Structure", routePortName, distanceType);
                if (distance > maximumHorizontalDistance)
                    maximumHorizontalDistance = distance;
            }
            return maximumHorizontalDistance;
        }
    }
    //----------------------------------------------------------------------
    //This Rule retuns array of width and height offset values.
    //ProgId : AssemblyRules,Ingr.SP3D.Content.Support.Rules.HVACSectionSize
    //----------------------------------------------------------------------                                                 
    public class HVACSectionSize : IHgrRule
    {
        public Array AttributeValue(BusinessObject SupportOrComponent)
        {
            Ingr.SP3D.Support.Middle.Support support = (Ingr.SP3D.Support.Middle.Support)SupportOrComponent;          
            SupportHelper supportHelper = new SupportHelper(support);
            BoundingBoxHelper boundingBoxHelper = new BoundingBoxHelper(support);
            GenericHelper genericHelper = new GenericHelper(support);
            RefPortHelper refPortHelper = new RefPortHelper(support);
           
            int pipeCount = supportHelper.SupportedObjects.Count;
            double distance = 0.0, routeStructDistance = 0.0;
            string routePortName = string.Empty,steelStandard = string.Empty;
            for (int routeIndex = 1; routeIndex <= pipeCount; routeIndex++)
            {
                if (routeIndex == 1)
                    routePortName = "Route";
                else
                    routePortName = "Route_" + routeIndex;
                distance = refPortHelper.DistanceBetweenPorts(routePortName, "Structure", PortDistanceType.Direct);
                if (distance > routeStructDistance)
                    routeStructDistance = distance;
            }
            boundingBoxHelper.CreateStandardBoundingBoxes(false);
            BoundingBox boundingBox;
            if (supportHelper.PlacementType == PlacementType.PlaceByStruct)
                boundingBox = boundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
            else
                boundingBox = boundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
            double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height;
            if (routeStructDistance + (boundingBoxHeight / 2) > boundingBoxWidth)
                distance = routeStructDistance + (boundingBoxHeight / 2);
            else
                distance = boundingBoxWidth;
            distance = distance * 1000;     
            string[] size = new string[2];
           
            CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
            IEnumerable<BusinessObject> hvacParts = null;
            PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("HVACAssy_SecByRule");
            hvacParts = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

            hvacParts = hvacParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAHgrSecByRule", "UnitType")).PropValue).ToLower() == "mm" && (((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrSecByRule", "LorAmin")).PropValue) < distance && ((double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrSecByRule", "LorAmax")).PropValue) > distance)).OrderBy<BusinessObject, string>(x => (string)((PropertyValueString)x.GetPropertyValue("IJUAHgrSecByRule", "SectionSize")).PropValue);
            if (hvacParts.Count() > 0)
            {                
                size[0] = (string)((PropertyValueString)hvacParts.ElementAt(0).GetPropertyValue("IJUAHgrSecByRule", "SectionSize")).PropValue;
                size[1] = (string)((PropertyValueString)hvacParts.ElementAt(1).GetPropertyValue("IJUAHgrSecByRule", "SectionSize")).PropValue;
            }           
            bool value = genericHelper.GetDataByRule("HgrHVACSSteelStandardName", (BusinessObject)support, out steelStandard);
            partClass = (PartClass)catalogBaseHelper.GetPartClass("HgrHVACStCorrespond");

            hvacParts = null;
            hvacParts = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            string[] SectionSize = new string[2];
            hvacParts = hvacParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAHgrHVACStCorrespon", "Size")).PropValue == size[0]) && ((string)((PropertyValueString)part.GetPropertyValue("IJUAHgrHVACStCorrespon", "StdName")).PropValue == steelStandard.ToUpper()));
            if (hvacParts.Count() > 0)           
                SectionSize[0] = (string)((PropertyValueString)hvacParts.ElementAt(0).GetPropertyValue("IJUAHgrHVACStCorrespon", "SectionSize")).PropValue;

            hvacParts = null;
            hvacParts = partClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
            hvacParts = hvacParts.Where(part => ((string)((PropertyValueString)part.GetPropertyValue("IJUAHgrHVACStCorrespon", "Size")).PropValue == size[1]) && ((string)((PropertyValueString)part.GetPropertyValue("IJUAHgrHVACStCorrespon", "StdName")).PropValue == steelStandard.ToUpper()));
            if (hvacParts.Count() > 0)   
            SectionSize[1] = (string)((PropertyValueString)hvacParts.ElementAt(0).GetPropertyValue("IJUAHgrHVACStCorrespon", "SectionSize")).PropValue;
            GC.SuppressFinalize(hvacParts); 
            return SectionSize;
        }
    }
}