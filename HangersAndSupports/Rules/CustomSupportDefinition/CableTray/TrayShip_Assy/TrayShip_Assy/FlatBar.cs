//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   FlatBar.cs
//   TrayShip_Assy,Ingr.SP3D.Content.Support.Rules.FlatBar
//   Author       :  Vijay
//   Creation Date:  12/07/2013   
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   12/07/2013     Vijay    CR-CP-224487  Convert HS_TrayShip_Assy to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;

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
    public class FlatBar : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Part Index's
        private const string FLATBAR = "FLATBAR";
        private double length, flatBarThickness, flatBarWidth;
        private string sectionSize;

        public override Collection<PartInfo> Parts
        {
            get
            {
                IEnumerable<BusinessObject> traySrvFBDimPart = null;
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmFlatBarL", "Length")).PropValue;
                    PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAhsTrayAsmFBSecSize", "SectionSize");
                    sectionSize = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodelist.PropValue).ShortDisplayName;

                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    PartClass traySrvFBDim = (PartClass)catalogBaseHelper.GetPartClass("TrayShipSrv_FBSecDim");
                    traySrvFBDimPart = traySrvFBDim.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;

                    traySrvFBDimPart = traySrvFBDimPart.Where(part => ((int)((PropertyValueCodelist)part.GetPropertyValue("IJUAhsTraySrvFBDim", "SectionSize")).PropValue == sectionSizeCodelist.PropValue));
                    if (traySrvFBDimPart.Count() > 0)
                    {
                        flatBarWidth = (double)((PropertyValueDouble)traySrvFBDimPart.ElementAt(0).GetPropertyValue("IJUAhsTraySrvFBDim", "Width")).PropValue;
                        flatBarThickness = (double)((PropertyValueDouble)traySrvFBDimPart.ElementAt(0).GetPropertyValue("IJUAhsTraySrvFBDim", "Thickness")).PropValue;
                    }
                    //Create the list of parts
                    parts.Add(new PartInfo(FLATBAR, "TrayShip_RecStrap_1"));

                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
                finally
                {
                    if (traySrvFBDimPart is IDisposable)
                    {
                        ((IDisposable)traySrvFBDimPart).Dispose(); // This line will be executed
                    }
                }
            }
        }


        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------

        public override int ConfigurationCount
        {
            get
            {
                return 2;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);

                double routeAngle = RefPortHelper.AngleBetweenPorts("BBR_High", PortAxisType.X, OrientationAlong.Global_Z);
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("BBR_High", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);

                double distRouteStruct;

                if ((Math.Abs(routeAngle) < (0 + 0.01) && Math.Abs(routeAngle) > (0 - 0.01)) || (Math.Abs(routeAngle) < (Math.PI + 0.01) && Math.Abs(routeAngle) > (Math.PI - 0.01)))
                    distRouteStruct = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_High", PortDistanceType.Horizontal_Perpendicular);
                else
                {
                    if ((Math.Abs(routeStructAngle) < (0 + 0.01) && Math.Abs(routeStructAngle) > (0 - 0.01)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.01) && Math.Abs(routeStructAngle) > (Math.PI - 0.01)))
                        distRouteStruct = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_High", PortDistanceType.Horizontal);
                    else if (Math.Abs(routeStructAngle) < (Math.PI / 2 + 0.01) && Math.Abs(routeStructAngle) > (Math.PI / 2 - 0.01))
                        distRouteStruct = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_High", PortDistanceType.Vertical);
                    else
                        distRouteStruct = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_High", PortDistanceType.Horizontal_Perpendicular);
                }

                componentDictionary[FLATBAR].SetPropertyValue(length, "IJOAHgrTrayWidth", "Width");
                componentDictionary[FLATBAR].SetPropertyValue(flatBarWidth, "IJOAHgrTrayDepth", "Depth");
                componentDictionary[FLATBAR].SetPropertyValue(distRouteStruct - flatBarThickness, "IJOAHgrTrayHeight", "Height");
                componentDictionary[FLATBAR].SetPropertyValue(flatBarThickness, "IJOAHgrTrayThickness", "Thickness");

                JointHelper.CreateRigidJoint(FLATBAR, "Route", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.Y, -flatBarThickness, 0, 0);
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();

                    routeConnections.Add(new ConnectionInfo(FLATBAR, 1));       //partindex, routeindex

                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(FLATBAR, 1));      //partindex, routeindex

                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }

        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                double length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmFlatBarL", "Length")).PropValue;
                PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAhsTrayAsmFBSecSize", "SectionSize");
                sectionSize = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodelist.PropValue).ShortDisplayName;
                                                   
                bomString = sectionSize + "-" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_MILLIMETER);

                return bomString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        #endregion
    }
}
