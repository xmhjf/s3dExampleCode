//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   TypeU.cs
//   SupportStructAssemblyInfoRules,Ingr.SP3D.Content.Support.Rules.StructSup
//   Author       :Vijaya
//   Creation Date:5.Aug.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  5.Aug.2013     Vijaya   CR-CP-224488  Convert HgrSupStructAssmInfoRules to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class StructSup : CustomSupportDefinition
    {
        //Constants
        private const string STRUCTPROFILE1 = "STRUCTPROFILE1";
        private const string STRUCTPROFILE2 = "STRUCTPROFILE2";
        private const string STRUCTPROFILE3 = "STRUCTPROFILE3";       
        double d, gap;        
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
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    object[] attributeCollection = SupportStructAssemblyServices.GetPipeStructuralASMAttributes(this);

                    //Get the attributes from assembly
                    d = (double)attributeCollection[0];
                    gap = (double)attributeCollection[1];

                    //Use the default selection rule to get a catalog part for each part class               
                    parts.Add(new PartInfo(STRUCTPROFILE1, "StructProfile", "HgrPipePartSelRule.CStructProfilePart"));
                    parts.Add(new PartInfo(STRUCTPROFILE2, "StructProfile", "HgrPipePartSelRule.CStructProfilePart"));
                    parts.Add(new PartInfo(STRUCTPROFILE3, "StructProfile", "HgrPipePartSelRule.CStructProfilePart"));

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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                //load standard bounding box.
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
               
                //Create Joints
                BusinessObject horizontalSectionPart = componentDictionary[STRUCTPROFILE2].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                //Get SteelWidth and SteelDepth 
                double steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                string bottomStructPort = string.Empty;
                double lugOffset = d;

                //Create the Joint between the RteLow Reference Port and the Left Bottom Symbol
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", STRUCTPROFILE2, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -steelDepth / 2 - gap, -steelWidth / 2 - lugOffset);
                else
                    JointHelper.CreateRigidJoint("-1", "BBR_Low", STRUCTPROFILE2, "BeginCap", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -steelDepth / 2 - gap, -steelWidth / 2 - lugOffset, 0);

                //Create the Plane Joint between the RteHigh Reference Port
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePlanarJoint("-1", "BBSR_High", STRUCTPROFILE2, "EndCap", Plane.ZX, Plane.XY, lugOffset + steelWidth / 2);
                else
                    JointHelper.CreatePlanarJoint("-1", "BBR_High", STRUCTPROFILE2, "EndCap", Plane.ZX, Plane.XY, lugOffset + steelWidth / 2);

                //Add a Prismatic Joint defining the flexible bottom member
                JointHelper.CreatePrismaticJoint(STRUCTPROFILE2, "BeginCap", STRUCTPROFILE2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add rigid joint between beam 1 and beam 2
                JointHelper.CreateRigidJoint(STRUCTPROFILE1, "EndCap", STRUCTPROFILE2, "BeginCap", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, 0, -steelWidth, 0);

                //Add rigid joint between beam 2 and beam 3
                JointHelper.CreateRigidJoint(STRUCTPROFILE3, "EndCap", STRUCTPROFILE2, "EndCap", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, 0, -steelWidth, 0);

                //Add prismatic joint for the flexible memeber
                JointHelper.CreatePrismaticJoint(STRUCTPROFILE1, "BeginCap", STRUCTPROFILE1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add prismatic joint for the flexible memeber
                JointHelper.CreatePrismaticJoint(STRUCTPROFILE3, "BeginCap", STRUCTPROFILE3, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add prismatic joint between beam 1 and structure
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePrismaticJoint("-1", "Structure", STRUCTPROFILE1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0);
                else
                    JointHelper.CreatePlanarJoint("-1", "Structure", STRUCTPROFILE1, "BeginCap", Plane.XY, Plane.XY, 0);

                //Add planar joint between structure and beam 3
                JointHelper.CreatePlanarJoint("-1", "Structure", STRUCTPROFILE3, "BeginCap", Plane.XY, Plane.XY, 0);

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

                    routeConnections.Add(new ConnectionInfo(STRUCTPROFILE2, 1)); // partindex, routeindex

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
                    //Create a collection to hold the ALL Structure connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(STRUCTPROFILE1, 1)); // partindex, Structureindex
                    structConnections.Add(new ConnectionInfo(STRUCTPROFILE3, 1)); // partindex, Structureindex

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




