//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   GU6.cs
//   FLSampleSupports,Ingr.SP3D.Content.Support.Rules.GU6
//   Author       : Hema
//   Creation Date: 29-Aug-013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//    29-Aug-2013  Hema CR-CP-224478 Convert FlSample_Supports to C# .Net 
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;

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

    public class GU6 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string PLATE1 = "PLATE1";
        private const string PLATE2 = "PLATE2";
        private const string PLATE11 = "PLATE11";
        private const string PLATE22 = "PLATE22";
        private const string GUSSET1 = "GUSSET1";
        private const string GUSSET2 = "GUSSET2";

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

                    string plate1 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Plate", "Plate_1")).PropValue;
                    string plate2 = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Plate", "Plate_2")).PropValue;
                    string gusset = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrSPL_Plate", "Plate_3")).PropValue;

                    parts.Add(new PartInfo(PLATE1, plate1));
                    parts.Add(new PartInfo(PLATE2, plate2));
                    parts.Add(new PartInfo(PLATE11, plate1));
                    parts.Add(new PartInfo(PLATE22, plate2));
                    parts.Add(new PartInfo(GUSSET1, gusset));
                    parts.Add(new PartInfo(GUSSET2, gusset));

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

                Double plate1Length, plate1Height, plate1Thickness, plate2Length, plate2Height, plate2Thickness, plate3Length, plate3Height, plate3Thickness;

                plate1Length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateL_1")).PropValue;
                plate1Height = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateH_1")).PropValue;
                plate1Thickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateT_1")).PropValue;
                plate2Length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateL_2")).PropValue;
                plate2Height = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateH_2")).PropValue;
                plate2Thickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateT_2")).PropValue;
                plate3Length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateL_3")).PropValue;
                plate3Height = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateH_3")).PropValue;
                plate3Thickness = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrSPL_Plate", "PlateT_3")).PropValue;

                componentDictionary[PLATE1].SetPropertyValue(plate1Length, "IJOAHgrUtility_PLATE", "DEPTH");
                componentDictionary[PLATE1].SetPropertyValue(plate1Height, "IJOAHgrUtility_PLATE", "WIDTH");
                componentDictionary[PLATE1].SetPropertyValue(plate1Thickness, "IJOAHgrUtility_PLATE", "THICKNESS");
                componentDictionary[PLATE1].SetPropertyValue("280x100x12 THK CARBON STEEL PLATE.", "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                componentDictionary[PLATE11].SetPropertyValue(plate1Length, "IJOAHgrUtility_PLATE", "DEPTH");
                componentDictionary[PLATE11].SetPropertyValue(plate1Height, "IJOAHgrUtility_PLATE", "WIDTH");
                componentDictionary[PLATE11].SetPropertyValue(plate1Thickness, "IJOAHgrUtility_PLATE", "THICKNESS");
                componentDictionary[PLATE11].SetPropertyValue("280x100x12 THK CARBON STEEL PLATE.", "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                componentDictionary[PLATE2].SetPropertyValue(plate2Length, "IJOAHgrUtility_PLATE", "DEPTH");
                componentDictionary[PLATE2].SetPropertyValue(plate2Height, "IJOAHgrUtility_PLATE", "WIDTH");
                componentDictionary[PLATE2].SetPropertyValue(plate2Thickness, "IJOAHgrUtility_PLATE", "THICKNESS");
                componentDictionary[PLATE2].SetPropertyValue("280x100x12 THK CARBON STEEL PLATE.", "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                componentDictionary[PLATE22].SetPropertyValue(plate2Length, "IJOAHgrUtility_PLATE", "DEPTH");
                componentDictionary[PLATE22].SetPropertyValue(plate2Height, "IJOAHgrUtility_PLATE", "WIDTH");
                componentDictionary[PLATE22].SetPropertyValue(plate2Thickness, "IJOAHgrUtility_PLATE", "THICKNESS");
                componentDictionary[PLATE22].SetPropertyValue("280x100x12 THK CARBON STEEL PLATE.", "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                componentDictionary[GUSSET1].SetPropertyValue(plate3Length, "IJOAHgrGUSSET", "W");
                componentDictionary[GUSSET1].SetPropertyValue(plate3Height, "IJOAHgrGUSSET", "H");
                componentDictionary[GUSSET1].SetPropertyValue(plate3Thickness, "IJOAHgrGUSSET", "THICKNESS");
                componentDictionary[GUSSET1].SetPropertyValue("268x268x12 THK CARBON STEEL PLATE.", "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                componentDictionary[GUSSET2].SetPropertyValue(plate3Length, "IJOAHgrGUSSET", "W");
                componentDictionary[GUSSET2].SetPropertyValue(plate3Height, "IJOAHgrGUSSET", "H");
                componentDictionary[GUSSET2].SetPropertyValue(plate3Thickness, "IJOAHgrGUSSET", "THICKNESS");
                componentDictionary[GUSSET2].SetPropertyValue("268x268x12 THK CARBON STEEL PLATE.", "IJOAHgrUtilMMBomDesc", "InputBomDesc");

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                Double ndpPipeDiameterInch = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);
                if (ndpPipeDiameterInch < 41.99999 || ndpPipeDiameterInch > 48.000001)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: GU6 can only be placed from 42in to 48in pipe.", "", "GU6", 131);

                Double ndpPipeDiameter = pipeInfo.OutsideDiameter;
                ndpPipeDiameter = ndpPipeDiameter / 2;
                if (pipeInfo.InsulationThickness > 0)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: Not placable for Insulated pipe.", "", "GU6", 136);

                Double verticalDIS = ndpPipeDiameter;
                Double x = ((ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 * (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 + (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 * (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 + (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 * (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 + (verticalDIS - (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 - plate1Thickness * Math.Sqrt(2) / 2) * (verticalDIS - (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2 - plate1Thickness * Math.Sqrt(2) / 2) - ((plate1Thickness + plate1Height) * Math.Sqrt(2) / 2 - verticalDIS) * ((plate1Thickness + plate1Height) * Math.Sqrt(2) / 2 - verticalDIS)) / (2 * (ndpPipeDiameter + 0.003) * Math.Sqrt(2) / 2);

                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                if (supportingType == "Slab")
                {
                    x = x - (0.134 + 0.012) * Math.Sqrt(2) / 2;
                    JointHelper.CreateRigidJoint("-1", "Structure", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.14 * Math.Sqrt(2) / 2, x, 0);
                    JointHelper.CreateRigidJoint(PLATE1, "TopStructure", PLATE2, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, -0.006 * Math.Sqrt(2) / 2, (0.134 * Math.Sqrt(2) - 0.006 * Math.Sqrt(2) / 2), 0);
                    JointHelper.CreateRigidJoint(PLATE2, "TopStructure", GUSSET1, "Other", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0.134 * Math.Sqrt(2) / 2, 0.134 * Math.Sqrt(2) / 2, -0.006);

                    JointHelper.CreateRigidJoint("-1", "Structure", PLATE11, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0.14 * Math.Sqrt(2) / 2, -x, 0);
                    JointHelper.CreateRigidJoint(PLATE11, "TopStructure", PLATE22, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, -0.006 * Math.Sqrt(2) / 2, (0.134 * Math.Sqrt(2) - 0.006 * Math.Sqrt(2) / 2), 0);
                    JointHelper.CreateRigidJoint(PLATE22, "TopStructure", GUSSET2, "Other", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0.134 * Math.Sqrt(2) / 2, 0.134 * Math.Sqrt(2) / 2, -0.006);
                }
                else
                {
                    x = x + (0.134 - 2 * 0.012) * Math.Sqrt(2) / 2;
                    JointHelper.CreateRigidJoint("-1", "Structure", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, (0.14 * Math.Sqrt(2) / 2), 0, x);
                    JointHelper.CreateRigidJoint(PLATE1, "TopStructure", PLATE2, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, -0.006 * Math.Sqrt(2) / 2, (0.134 * Math.Sqrt(2) - 0.006 * Math.Sqrt(2) / 2), 0);
                    JointHelper.CreateRigidJoint(PLATE2, "TopStructure", GUSSET1, "Other", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0.134 * Math.Sqrt(2) / 2, 0.134 * Math.Sqrt(2) / 2, -0.006);

                    JointHelper.CreateRigidJoint("-1", "Structure", PLATE11, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0.14 * Math.Sqrt(2) / 2, 0, -x);
                    JointHelper.CreateRigidJoint(PLATE11, "TopStructure", PLATE22, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, -0.006 * Math.Sqrt(2) / 2, (0.134 * Math.Sqrt(2) - 0.006 * Math.Sqrt(2) / 2), 0);
                    JointHelper.CreateRigidJoint(PLATE22, "TopStructure", GUSSET2, "Other", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0.134 * Math.Sqrt(2) / 2, 0.134 * Math.Sqrt(2) / 2, -0.006);
                }
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

                    routeConnections.Add(new ConnectionInfo(PLATE2, 1)); // partindex, routeindex

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

                    structConnections.Add(new ConnectionInfo(PLATE2, 1)); // partindex, Structureindex    

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
            return "GU6";
        }
        #endregion
    }
}
