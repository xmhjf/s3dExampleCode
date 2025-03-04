//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   AngleBarSupp.cs
//   TrayShip_Assy,Ingr.SP3D.Content.Support.Rules.AngleBarSupp
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
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
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

    public class AngleBarSupp : CustomSupportDefinition 
    {
        //Part Index's
        private const string VERTSECTION1 = "VERTSECTION1";
        private const string VERTSECTION2 = "VERTSECTION2";

        private double underLength;
        private string sectionSize;
        public double length1, length2;
        public static double Length11 { get; set; }
        public static double Length22 { get; set; }
      
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    underLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAngleBar", "UnderLength")).PropValue;
                    PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAHgrSecSize", "SectionSize");
                    sectionSize = sectionSizeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(sectionSizeCodelist.PropValue).ShortDisplayName;

                    //Create the list of parts
                    parts.Add(new PartInfo(VERTSECTION1, sectionSize));
                    parts.Add(new PartInfo(VERTSECTION2, sectionSize));

                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
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
                //Load standard bounding box definition
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;

                //==========================
                //1. Load standard bounding box definition
                //==========================

                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                //=========================
                //2. Get bounding box boundary objects dimension information
                //=========================

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                //====== ======
                //3. retrieve dimension of the bounding box
                //====== ======
                // Get route box geometry

                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | dHeight
                // |____________________|
                //        dWidth

                double width = boundingBox.Width;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                BusinessObject horizontalSectionPart = componentDictionary[VERTSECTION1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double SectionWidth = crosssection.Width;

                PropertyValueCodelist topbeginMiterCodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (topbeginMiterCodelist.PropValue == -1)
                    topbeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist topendMiterCodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION1].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (topendMiterCodelist.PropValue == -1)
                    topendMiterCodelist.PropValue = 1;

                PropertyValueCodelist anglebeginMiterCodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION2].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (anglebeginMiterCodelist.PropValue == -1)
                    anglebeginMiterCodelist.PropValue = 1;
                PropertyValueCodelist angleendMiterCodelist = (PropertyValueCodelist)componentDictionary[VERTSECTION2].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (angleendMiterCodelist.PropValue == -1)
                    angleendMiterCodelist.PropValue = 1;

                //====== ======
                // Set Values of Part Occurance Attributes
                //====== ======

                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION1].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTSECTION1].SetPropertyValue(topbeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTSECTION1].SetPropertyValue(topendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[VERTSECTION2].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[VERTSECTION2].SetPropertyValue(anglebeginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[VERTSECTION2].SetPropertyValue(angleendMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                //=======================================
                //Do Something if more than one Structure
                //=======================================
                //get structure count

                Boolean[] isOffsetApplied = TrayShipAseemblyServices.GetIsLugEndOffsetApplied(this);
                string[] structPort = TrayShipAseemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0];
                string rightStructPort = structPort[1];

                double routeAngle = RefPortHelper.AngleBetweenPorts("BBR_Low", PortAxisType.X, OrientationAlong.Global_Z);
                double routeStructAngle = RefPortHelper.AngleBetweenPorts("BBR_Low", PortAxisType.Y, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);

                double distRouteStruct1, distRouteStruct2;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if ((Math.Abs(routeAngle) < (0 + 0.01) && Math.Abs(routeAngle) > (0 - 0.01)) || (Math.Abs(routeAngle) < (Math.PI + 0.01) && Math.Abs(routeAngle) > (Math.PI - 0.01)))
                    {
                        distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_Low", PortDistanceType.Horizontal);
                        distRouteStruct2 = distRouteStruct1;
                    }
                    else
                    {
                        if ((Math.Abs(routeStructAngle) < (0 + 0.01) && Math.Abs(routeStructAngle) > (0 - 0.01)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.01) && Math.Abs(routeStructAngle) > (Math.PI - 0.01)))
                        {
                            distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_Low", PortDistanceType.Horizontal);
                            distRouteStruct2 = distRouteStruct1;
                        }
                        else if (Math.Abs(routeStructAngle) < (Math.PI / 2 + 0.01) && Math.Abs(routeStructAngle) > (Math.PI / 2 - 0.01))
                        {
                            distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_Low", PortDistanceType.Vertical);
                            distRouteStruct2 = distRouteStruct1;
                        }
                        else
                        {
                            distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_Low", PortDistanceType.Horizontal);
                            distRouteStruct2 = distRouteStruct1;
                        }
                    }
                }
                else
                {
                    if ((Math.Abs(routeAngle) < (0 + 0.01) && Math.Abs(routeAngle) > (0 - 0.01)) || (Math.Abs(routeAngle) < (Math.PI + 0.01) && Math.Abs(routeAngle) > (Math.PI - 0.01)))
                    {
                        if (leftStructPort != rightStructPort)
                        {
                            distRouteStruct1 = RefPortHelper.DistanceBetweenPorts(leftStructPort, "BBR_Low", PortDistanceType.Horizontal_Perpendicular);
                            distRouteStruct2 = RefPortHelper.DistanceBetweenPorts(rightStructPort, "BBR_Low", PortDistanceType.Horizontal_Perpendicular);
                        }
                        else
                        {
                            distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_Low", PortDistanceType.Horizontal_Perpendicular);
                            distRouteStruct2 = distRouteStruct1;
                        }
                    }
                    else
                    {
                        if ((Math.Abs(routeStructAngle) < (0 + 0.01) && Math.Abs(routeStructAngle) > (0 - 0.01)) || (Math.Abs(routeStructAngle) < (Math.PI + 0.01) && Math.Abs(routeStructAngle) > (Math.PI - 0.01)))
                        {
                            if (leftStructPort != rightStructPort)
                            {
                                distRouteStruct1 = RefPortHelper.DistanceBetweenPorts(leftStructPort, "BBR_Low", PortDistanceType.Horizontal);
                                distRouteStruct2 = RefPortHelper.DistanceBetweenPorts(rightStructPort, "BBR_Low", PortDistanceType.Horizontal);
                            }
                            else
                            {
                                distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_Low", PortDistanceType.Horizontal);
                                distRouteStruct2 = distRouteStruct1;
                            }
                        }
                        else if (Math.Abs(routeStructAngle) < (Math.PI / 2 + 0.01) && Math.Abs(routeStructAngle) > (Math.PI / 2 - 0.01))
                        {
                            if (leftStructPort != rightStructPort)
                            {
                                distRouteStruct1 = RefPortHelper.DistanceBetweenPorts(leftStructPort, "BBR_Low", PortDistanceType.Vertical);
                                distRouteStruct2 = RefPortHelper.DistanceBetweenPorts(rightStructPort, "BBR_Low", PortDistanceType.Vertical);
                            }
                            else
                            {
                                distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_Low", PortDistanceType.Vertical);
                                distRouteStruct2 = distRouteStruct1;
                            }
                        }
                        else
                        {
                            if (leftStructPort != rightStructPort)
                            {
                                distRouteStruct1 = RefPortHelper.DistanceBetweenPorts(leftStructPort, "BBR_Low", PortDistanceType.Horizontal_Perpendicular);
                                distRouteStruct2 = RefPortHelper.DistanceBetweenPorts(rightStructPort, "BBR_Low", PortDistanceType.Horizontal_Perpendicular);
                            }
                            else
                            {
                                distRouteStruct1 = RefPortHelper.DistanceBetweenPorts("Structure", "BBR_Low", PortDistanceType.Horizontal_Perpendicular);
                                distRouteStruct2 = distRouteStruct1;
                            }
                        }
                    }
                }                

                string sectionPort1, sectionPort2, overLengthPort1, overLengthPort2;
                TrayShipAseemblyServices.ConfigIndex verticalSection1, verticalSection2 = new TrayShipAseemblyServices.ConfigIndex();
                if (Configuration == 1)
                {
                    verticalSection1 = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                    verticalSection2 = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX);
                    sectionPort1 = "BeginCap";
                    sectionPort2 = "EndCap";
                    overLengthPort1 = "BeginOverLength";
                    overLengthPort2 = "EndOverLength";
                }
                else
                {
                    verticalSection1 = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                    verticalSection2 = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                    sectionPort1 = "EndCap";
                    sectionPort2 = "BeginCap";
                    overLengthPort1 = "EndOverLength";
                    overLengthPort2 = "BeginOverLength";
                }

                if (underLength != 0)
                {
                    componentDictionary[VERTSECTION1].SetPropertyValue(-underLength, "IJUAHgrOccOverLength", overLengthPort1);
                    componentDictionary[VERTSECTION2].SetPropertyValue(-underLength, "IJUAHgrOccOverLength", overLengthPort2);
                }

                componentDictionary[VERTSECTION1].SetPropertyValue(distRouteStruct1, "IJUAHgrOccLength", "Length");
                componentDictionary[VERTSECTION2].SetPropertyValue(distRouteStruct2, "IJUAHgrOccLength", "Length");

                length1 = distRouteStruct1 - underLength;
                Length11 = length1;
                length2 = distRouteStruct2 - underLength;
                Length22 = length2;
                //=============
                //Create Joints
                //=============
                //Create a collection to hold the joints

                //Add joint between Vertical Section 1 and Route
                JointHelper.CreateRigidJoint(VERTSECTION1, sectionPort1, "-1", "BBR_Low", verticalSection1.A, verticalSection1.B, verticalSection1.C, verticalSection1.D, 0, 0, SectionWidth / 2);

                //Add joint between Vertical Section 2 and Route
                JointHelper.CreateRigidJoint(VERTSECTION2, sectionPort2, "-1", "BBR_Low", verticalSection2.A, verticalSection2.B, verticalSection2.C, verticalSection2.D, 0, -width, SectionWidth / 2);

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

                    routeConnections.Add(new ConnectionInfo(VERTSECTION1, 1));       //partindex, routeindex

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

                    structConnections.Add(new ConnectionInfo(VERTSECTION1, 1));      //partindex, routeindex

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

        
    }
}
