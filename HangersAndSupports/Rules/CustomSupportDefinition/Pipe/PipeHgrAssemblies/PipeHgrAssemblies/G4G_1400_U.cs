
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   G4G_1400_U.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.G4G_1400_U
//   Author       : Hema
//   Creation Date: 04.April.2013
//   Description:   Converted HS_Assembly VB Project to C# .Net 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   04.April.2013   Hema     Converted HS_Assembly VB Project to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle.Services;
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

    public class G4G_1400_U : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.G4G_1400_U"
        //----------------------------------------------------------------------------------
        //Constants
        private const string HGRSUPUBOLT = "HgrSupUBolt";

        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    //Get SupportHelper
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                    //Create a new collection to hold the caltalog parts
                    PartClass hgrSupUBoltPartClass = (PartClass)catalogBaseHelper.GetPartClass("HgrSupUBolt");

                    //Use the default selection rule to get a catalog part for each part class
                    string partselection = hgrSupUBoltPartClass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                    parts.Add(new PartInfo(HGRSUPUBOLT, "HgrSupUBolt", partselection));

                    // Return the collection of Catalog Parts
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
                //Don't support the Route Connection Button
                return 0;
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                //Get IJHgrInputConfig Hlpr Interface off of passed Helper
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========

                BoundingBoxHelper.CreateStandardBoundingBoxes(false);

                //====== ======
                //Create Joints
                //====== ======

                //----------------------------------------------------
                //Add a Prismatic Joint defining the flexible U-Bolt
                JointHelper.CreatePrismaticJoint(HGRSUPUBOLT, "Route", HGRSUPUBOLT, "Structure", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //----------------------------------------------------
                //Add a Joint between U-Bolt Structure Port and the Reference Structure Port
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    //Beam Structure... Planar Slot Joint
                    JointHelper.CreatePlanarSlotJoint("-1", "Structure", HGRSUPUBOLT, "Structure", Plane.XY, Plane.XY, Axis.X, 0, 0);
                else
                    JointHelper.CreatePointOnPlaneJoint(HGRSUPUBOLT, "Structure", "-1", "Structure", Plane.XY);
                //----------------------------------------------------

                //Add a Joint between U-Bolt and Pipe
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    //Beam Structure... Add a Cylindrical Joint on the Reference Route Port
                    JointHelper.CreateCylindricalJoint(HGRSUPUBOLT, "Route", "-1", "Route", Axis.X, Axis.X, 0);
                else
                    //Plate Structure... Add a Spherical Joint
                    JointHelper.CreateRigidJoint("-1", "Route", HGRSUPUBOLT, "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
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

                    //First Part return by GetAssemblyCatalogParts() is HgrSupUBolt
                    //Connects to First Route Input (i.e. the pipe)
                    routeConnections.Add(new ConnectionInfo(HGRSUPUBOLT, 1));

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

                    //First Part return by GetAssemblyCatalogParts() is HgrSupUBolt
                    //Connects to First Structure Input (i.e. the beam or plate)
                    structConnections.Add(new ConnectionInfo(HGRSUPUBOLT, 1));

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

