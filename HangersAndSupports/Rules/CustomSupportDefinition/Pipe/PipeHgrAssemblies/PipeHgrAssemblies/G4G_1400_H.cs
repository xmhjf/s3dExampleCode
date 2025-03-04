
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   G4G_1400_U.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.G4G_1400_H
//   Author       : Hema
//   Creation Date: 04.April.2013
//   Description:   Converted HS_Assembly VB Project to C# .Net 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   04.April.2013   Hema     Converted HS_Assembly VB Project to C# .Net 
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
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

    public class G4G_1400_H : CustomSupportDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.G4G_1400_H"
        //----------------------------------------------------------------------------------
        //Constants
        //private const string G4G_1400_H = "G4G_1400_H";
        private const string G4G_1461_01 = "G4G_1461_01_P01";
        private const string G4G_1460_01 = "G4G_1460_01_P02";
        private const string G4G_1451_06 = "G4G_1451_06_P03";

        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    //Gets SupportHelper
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                    //Create a new collection to hold the caltalog parts
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                    PartClass g4g_1461_01PartClass, g4g_1460_01PartClass, g4g_1451_06PartClass;
                    string partselection, partselection1, partselection2;

                    //Create the list of part classes required by the type
                    g4g_1461_01PartClass = (PartClass)catalogBaseHelper.GetPartClass("G4G_1461_01");
                    g4g_1460_01PartClass = (PartClass)catalogBaseHelper.GetPartClass("G4G_1460_01");
                    g4g_1451_06PartClass = (PartClass)catalogBaseHelper.GetPartClass("G4G_1451_06");

                    //Use the default selection rule to get a catalog part for each part class
                    partselection = g4g_1461_01PartClass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                    partselection1 = g4g_1460_01PartClass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                    partselection2 = g4g_1451_06PartClass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();

                    //HngSupHgrAssemblySymbols.ThreadTop  (Beam Clamp)
                    parts.Add(new PartInfo(G4G_1461_01, "G4G_1461_01", partselection));

                    //HngSupHgrAssemblySymbols.ThreadFlex (Flexible Rod)
                    parts.Add(new PartInfo(G4G_1460_01, "G4G_1460_01", partselection1));

                    //HngSupHgrAssemblySymbols.BottomClamp  (Pipe Clamp)
                    parts.Add(new PartInfo(G4G_1451_06, "G4G_1451_06", partselection2));

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
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);

                //=== ========== == ==== ===========
                //Set Attributes on Part Occurrences
                //=== ========== == ==== ===========

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                //Get the Diameter of the Primary Pipe
                ConduitObjectInfo routeInfo = (ConduitObjectInfo)SupportedHelper.SupportedObjectInfo(1);

                double pipeRadius, BeamWidth = 0.0, byPointAngle1;

                //Get the Diameter of the Primary Pipe
                pipeRadius = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, routeInfo.NominalDiameter.Size / 2.0, UnitName.DISTANCE_INCH, UnitName.DISTANCE_METER);

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    BeamWidth = 0.5;
                else
                {
                    if ((SupportHelper.SupportingObjects.Count != 0))
                        BeamWidth = SupportingHelper.SupportingObjectInfo(1).Width;
					else
						BeamWidth = 0.5;
                }
                //If unable to retrieve BeamWidth, Structure could be a slab.
                //Use the pipe radius to set the clamp dimensions.
                if (BeamWidth <= 0.0)
                    BeamWidth = 2.0 * pipeRadius;

                //Set the width on the Beam Clamp
                //Get the Beam Clamp
                (componentDictionary[G4G_1461_01]).SetPropertyValue((0.75 * BeamWidth), "IJUAHgrOccGeometry", "Width");

                byPointAngle1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Direct);

                //======================================================
                //Get the Current Location in the Route Connection Cycle
                //======================================================

                //====== ======
                //Create Joints
                //====== ======

                //Create the Flexible (Cylindrical) Joint between the ports of the rod
                JointHelper.CreateCylindricalJoint(G4G_1460_01, "RodTop", G4G_1460_01, "RodBottom", Axis.Z, Axis.Z, 0);

                //----------------------------------------------------
                //Add a Joint between Pipe and Pipe Clamp
                if (!((Math.Abs(byPointAngle1) >= -0.0000001) && (Math.Abs(byPointAngle1) <= 0.0000001)))
                    JointHelper.CreateRevoluteJoint(G4G_1451_06, "Route", "-1", "Route", Axis.X, Axis.X);
                else
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint(G4G_1451_06, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                    else
                        JointHelper.CreateRigidJoint(G4G_1451_06, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);



                //----------------------------------------------------------------------------
                //Add a Joint between the Supporting Object (Beam or Plate) and the Beam Clamp
                if (!((Math.Abs(byPointAngle1) >= -0.0000001) && (Math.Abs(byPointAngle1) <= 0.0000001)))

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        //Beam Structure ... Add a Prismatic Joint between Beam and Beam Clamp
                        if (Configuration == 1)
                            JointHelper.CreateCylindricalJoint(G4G_1461_01, "Structure", "-1", "Structure", Axis.X, Axis.X, 0);
                        else
                            JointHelper.CreatePlanarJoint(G4G_1461_01, "Structure", "-1", "Structure", Plane.XY, Plane.XY, 0);
                    else
                        //Plate Structure ... Add a Translation Joint Between Plate and Beam Clamp
                        if (Configuration == 1)
                            JointHelper.CreateTranslationalJoint(G4G_1461_01, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0);
                        else
                            JointHelper.CreateTranslationalJoint(G4G_1461_01, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0);
                else
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        //Beam Structure ... Add a Prismatic Joint between Beam and Beam Clamp
                        if (Configuration == 1)
                            JointHelper.CreatePrismaticJoint(G4G_1461_01, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                        else
                            JointHelper.CreatePrismaticJoint(G4G_1461_01, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0);
                    else
                        //Plate Structure ... Add a Translation Joint Between Plate and Beam Clamp
                        if (Configuration == 1)
                            JointHelper.CreateTranslationalJoint(G4G_1461_01, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0);
                        else
                            JointHelper.CreateTranslationalJoint(G4G_1461_01, "Structure", "-1", "Structure", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0);

                //----------------------------------------------------
                //Add a Spherical Joint between Beam Clamp and Top of Rod
                if (!((Math.Abs(byPointAngle1) >= -0.0000001) && (Math.Abs(byPointAngle1) <= 0.0000001)))
                    if (Configuration == 1)
                        JointHelper.CreateRevoluteJoint(G4G_1461_01, "Rod", G4G_1460_01, "RodTop", Axis.Y, Axis.X);
                    else
                        JointHelper.CreateRevoluteJoint(G4G_1461_01, "Rod", G4G_1460_01, "RodTop", Axis.X, Axis.Y);
                else
                    JointHelper.CreateSphericalJoint(G4G_1461_01, "Rod", G4G_1460_01, "RodTop");

                //----------------------------------------------------
                //Add a Spherical Joint between Pipe Clamp and Bottom of Rod
                if (!((Math.Abs(byPointAngle1) >= -0.0000001) && (Math.Abs(byPointAngle1) <= 0.0000001)))
                    JointHelper.CreateRigidJoint(G4G_1451_06, "Rod", G4G_1460_01, "RodBottom", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    JointHelper.CreateSphericalJoint(G4G_1451_06, "Rod", G4G_1460_01, "RodBottom");


                //----------------------------------------------------
                //Add a Vertical Joint to the Rod Z axis
                JointHelper.CreateGlobalAxesAlignedJoint(G4G_1460_01, "RodBottom", Axis.Z, Axis.Z);
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

                    //"G4G_1451_06" (i.e. the pipe clamp).  It is the third part
                    //Add the PipeClampConnections to the Route Collection of Connections.
                    routeConnections.Add(new ConnectionInfo(G4G_1451_06, 1));

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

                    //Add the PipeClampConnections to the Structure Collection of Connections.
                    structConnections.Add(new ConnectionInfo(G4G_1461_01, 1));

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

