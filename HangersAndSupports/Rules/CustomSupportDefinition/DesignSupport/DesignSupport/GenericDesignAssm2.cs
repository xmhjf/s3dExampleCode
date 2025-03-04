//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   GenericDesignAssm2.cs
//   DesignSupport,Ingr.SP3D.Support.Content.Rules.GenericDesignAssm2
//   Author       :  Pavan
//   Creation Date:  28.Sep.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   08.Nov.2012    Pavan    DI-CP-220480  Check items in from Shelfsets
//   25.Nov.2013    BS       DI-CP-241803  Checked back .Net progid to DesignSupport part
//   23-2-03-2015   Chethan  TR-CP-268570  Namespace inconsistency in .NET content for few H&S project  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;

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
    public class GenericDesignAssm2 : CustomSupportDefinition
    {
        //Constants
        private const string DESIGN_SUPPORT = "DESIGN_SUPPORT";
        private double insulationThick;
        BoundingBox BBX;
        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                    BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                    

                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    parts.Add(new PartInfo(DESIGN_SUPPORT, "HgrGenericDesignBBX"));
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
                return 0;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;


            if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
            {
                BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
            }
            else
            {
                BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
            }

            Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

            double width = BBX.Width;
            double height = BBX.Height;

            int supEdObjCount = SupportHelper.SupportedObjects.Count;
            insulationThick = GetInsulationThickness(supEdObjCount);



            componentDictionary[DESIGN_SUPPORT].SetPropertyValue(width, "IJUAHgrOccGeometry", "Width");
            componentDictionary[DESIGN_SUPPORT].SetPropertyValue(height, "IJUAHgrOccGeometry", "Height");
            if (componentDictionary[DESIGN_SUPPORT].SupportsInterface("IJOAhsDesignSupportType"))
            {
                PropertyValueCodelist designSupType = (PropertyValueCodelist)support.GetPropertyValue("IJOAhsDesignSupportType", "DesignSupportType");
                int designSupType1;
                if (designSupType.PropValue <= 0)
                    designSupType1 = 1;
                else
                    designSupType1 = designSupType.PropValue;
                componentDictionary[DESIGN_SUPPORT].SetPropertyValue(designSupType1, "IJOAhsDesignSupportType", "DesignSupportType");
            }
            if (componentDictionary[DESIGN_SUPPORT].SupportsInterface("IJUAhsInsulationThick"))
                componentDictionary[DESIGN_SUPPORT].SetPropertyValue(insulationThick, "IJUAhsInsulationThick", "InsulationThick");

            //====== ======
            //Create Joints
            //====== ======
            //Create the Joint between the Bounding Box Low and Struct Reference Ports
            if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
            {
                JointHelper.CreateRigidJoint("-1", "BBSR_Low", DESIGN_SUPPORT, "Structure", Plane.YZ, Plane.XY, Axis.Y, Axis.X, 0, 0, 0);
            }
            else
            {
                JointHelper.CreateRigidJoint("-1", "BBR_Low", DESIGN_SUPPORT, "Structure", Plane.YZ, Plane.XY, Axis.Y, Axis.X, 0, 0, 0);
            }

        }
        //-----------------------------------------------------------------------------------
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                Collection<ConnectionInfo> routeConns = new Collection<ConnectionInfo>();
                for (int iIndex = 1; iIndex <= SupportHelper.SupportedObjects.Count; iIndex++)
                {
                    routeConns.Add(new ConnectionInfo(DESIGN_SUPPORT, iIndex));
                }
                return routeConns;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                Collection<ConnectionInfo> structConns = new Collection<ConnectionInfo>();
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    structConns.Add(new ConnectionInfo(DESIGN_SUPPORT, 1));
                }
                return structConns;
            }
        }
        private double GetInsulationThickness(int iSupEdObjCount)
        {
            try
            {
                double insulationThickness = 0;
                double tempInsulationThickness = 0;

                for (int iIndex = 1; iIndex <= iSupEdObjCount; iIndex++)
                {
                    if (SupportedHelper.SupportedObjectInfo(iIndex).SupportedObjectType == SupportedObjectType.Pipe)
                    {
                        // Get maximum thickness among all route objects
                        PipeObjectInfo routeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(iIndex);
                        tempInsulationThickness = routeInfo.InsulationThickness;
                    }
                    else if (SupportedHelper.SupportedObjectInfo(iIndex).SupportedObjectType == SupportedObjectType.HVAC)
                    {
                        // Get maximum thickness among all duct objects
                        DuctObjectInfo ductInfo = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(iIndex);
                        tempInsulationThickness = ductInfo.InsulationThickness;
                    }
                    if (tempInsulationThickness > insulationThickness)
                    {
                        insulationThickness = tempInsulationThickness;
                    }
                }
                return insulationThickness;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetInsulationThickness method." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }

        }
    }
}


