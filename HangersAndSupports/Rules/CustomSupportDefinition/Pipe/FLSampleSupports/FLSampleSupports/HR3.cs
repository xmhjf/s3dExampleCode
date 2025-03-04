//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//  HR3.cs
//   FLSampleSupports,Ingr.SP3D.Content.Support.Rules.HR3
//   Author       : Hema
//   Creation Date: 29-Aug-013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//    29-Aug-2013  Hema CR-CP-224478 Convert FlSample_Supports to C# .Net 
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
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class HR3 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string TOPCLAMP = "TOPCLAMP";
        private const string BOTTOMCLAMP = "BOTTOMCLAMP";
        private const string BOTTOMROD = "BOTTOMROD";
        private const string TOPEYENUT = "TOPEYENUT";
        private const string BOTTOMEYENUT = "BOTTOMEYENUT";
        private const string TOPHEXNUT = "TOPHEXNUT";
        private const string BOTTOMHEXNUT = "BOTTOMHEXNUT";
        private const string ROD = "ROD";

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
                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(2);
                    Double topPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);

                    pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    Double bottomPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);

                    //checking min&max bottom&top pipes size
                    if (bottomPipeDiameter > 6)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Error: Maximum Bottom Clamp exceed 6in.", "", "HR3", 66);
                        return null;
                    }
                    else if (bottomPipeDiameter > 3)
                    {
                        if (topPipeDiameter < 6)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: Minimun Top Clamp should be at least 6in..", "", "HR3", 73);
                            return null;
                        }
                    }
                    else if (bottomPipeDiameter > 1)
                    {
                        if (topPipeDiameter < 4)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: Minimun Top Clamp should be at least 4in..", "", "HR3", 81);
                            return null;
                        }
                    }
                    else
                        if (topPipeDiameter < 3)
                        {
                            MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: Minimun Top Clamp should be at least 3in..", "", "HR3", 88);
                            return null;
                        }

                    string rod = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Rod", "Rod")).PropValue;
                    string eyeNut = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_EyeNut", "EyeNut")).PropValue;
                    string hexNut = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_HexNut", "HexNut")).PropValue;

                    Double rodSize = FLSampleSupportServices.GetDataByCondition("Rule_PipeRod", "IJUAHgrRule_PipeRod", "RodSize", "IJUAHgrRule_PipeRod", "PipeDiaIn", topPipeDiameter.ToString());
                    PropertyValue topClampValue, bottomClampValue;
                     
                    Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                    parameter.Add("PipeDiaIn", topPipeDiameter.ToString());
                    GenericHelper.GetDataByRule("Rule_PipeClamp", "IJUAHgrRule_PipeClamp", "Clamp", parameter, out topClampValue);

                    parameter = new Dictionary<String, Object>();
                    parameter.Add("PipeDiaIn", bottomPipeDiameter.ToString());
                    GenericHelper.GetDataByRule("Rule_PipeClamp", "IJUAHgrRule_PipeClamp", "Clamp", parameter, out bottomClampValue);
                    String topClamp=topClampValue.ToString(), bottomClamp=bottomClampValue.ToString();
                    rod = FLSampleSupportServices.GetStringDataByCondition("Anv09_RodET", "IJDPart", "PartNumber", "IJUAhsRodDiameter", "RodDiameter", rodSize - 0.000001, rodSize + 0.000001);
                    eyeNut = FLSampleSupportServices.GetStringDataByCondition("Anv09_EyeNut", "IJDPart", "PartNumber", "IJUAhsRodDiameter", "RodDiameter", rodSize - 0.000001, rodSize + 0.000001);
                    hexNut = FLSampleSupportServices.GetStringDataByCondition("Anv09_HexNut", "IJDPart", "PartNumber", "IJUAhsRodDiameter", "RodDiameter", rodSize - 0.000001, rodSize + 0.000001);

                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    PartClass bottomClampPartClass = (PartClass)catalogBaseHelper.GetPartClass(bottomClamp);

                    //Use the default selection rule to get a catalog part for each part class
                    string partselection = bottomClampPartClass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();

                    PartClass topClampPartClass = (PartClass)catalogBaseHelper.GetPartClass(topClamp);

                    string partselection1 = topClampPartClass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();

                    parts.Add(new PartInfo(TOPCLAMP, topClamp, partselection1));
                    parts.Add(new PartInfo(BOTTOMCLAMP, bottomClamp, partselection));
                    parts.Add(new PartInfo(TOPEYENUT, eyeNut));
                    parts.Add(new PartInfo(BOTTOMEYENUT, eyeNut));
                    parts.Add(new PartInfo(TOPHEXNUT, hexNut));
                    parts.Add(new PartInfo(BOTTOMHEXNUT, hexNut));
                    parts.Add(new PartInfo(ROD, rod));
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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)     //It is placed by point only
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: Can not place-by-structure please place-by-point.", "", "HR3.cs", 151);
                    return;
                }
                //Top Joints
                Double angle = RefPortHelper.AngleBetweenPorts("Route_2", PortAxisType.X, "Route", PortAxisType.X, OrientationAlong.Direct);
                if ((angle >= 179.9999 && angle <= 180.0001) || (angle >= -0.0001 && angle <= 0.0001))
                    JointHelper.CreateRigidJoint("-1", "Route_2", TOPCLAMP, "Route", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                else
                    JointHelper.CreatePrismaticJoint(TOPCLAMP, "Route", "-1", "Route_2", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0);

                JointHelper.CreateRevoluteJoint(TOPCLAMP, "Wing", TOPEYENUT, "Eye", Axis.Y, Axis.Y);
                JointHelper.CreateRigidJoint(TOPHEXNUT, "Bottom", TOPEYENUT, "Surface", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(TOPEYENUT, "RodEnd", ROD, "RodEnd1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreatePrismaticJoint(ROD, "RodEnd1", ROD, "RodEnd2", Plane.YZ, Plane.YZ, Axis.Z, Axis.NegativeZ, 0, 0);
                JointHelper.CreateGlobalAxesAlignedJoint(ROD, "RodEnd2", Axis.Z, Axis.Z);

                //Bottom Joints
                JointHelper.CreatePrismaticJoint(BOTTOMCLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                JointHelper.CreateRevoluteJoint(BOTTOMCLAMP, "Wing", BOTTOMEYENUT, "Eye", Axis.Y, Axis.Y);
                JointHelper.CreateRigidJoint(BOTTOMHEXNUT, "Bottom", BOTTOMEYENUT, "Surface", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                //Middle Joint
                JointHelper.CreateRevoluteJoint(BOTTOMEYENUT, "RodEnd", ROD, "RodEnd2", Axis.Z, Axis.Z);
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 1;
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

                    routeConnections.Add(new ConnectionInfo(BOTTOMCLAMP, 1));// partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(TOPCLAMP, 2));// partindex, routeindex
                    
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
        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = string.Empty;
            try
            {
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(2);
                Double topPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);

                pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(2);
                Double bottomPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);

                bomDescription = "HR3 - " + Convert.ToString(bottomPipeDiameter) + "in - " + Convert.ToString(topPipeDiameter) + "in";
                return bomDescription;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOMDescription." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}
