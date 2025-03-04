//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5365_Guide_H.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5365_Guide_H
//   Author       :  BS
//   Creation Date:  04/07/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  08/07/2013    BS  CR-CP-224485- Converted HS_FINL_Assy to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using System.Collections.Generic;
namespace Ingr.SP3D.Content.Support.Rules
{
    public class SFS5365_Guide_H : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string LEFTGUIDE = "LeftGuide";
        private const string RIGHTGUIDE = "RightGuide";
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    parts.Add(new PartInfo(LEFTGUIDE, "Utility_USER_FIXED_BOX_1"));
                    parts.Add(new PartInfo(RIGHTGUIDE, "Utility_USER_FIXED_BOX_1"));

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
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(LEFTGUIDE, 1));      //partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(RIGHTGUIDE, 1));      //partindex, routeindex
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
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    structConnections.Add(new ConnectionInfo(LEFTGUIDE, 1));      //partindex, structindex
                    structConnections.Add(new ConnectionInfo(RIGHTGUIDE, 1));      //partindex, structindex
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

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                double npdMetric;
                //to change NomPipeDia to metric unit
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                if (!pipeInfo.NominalDiameter.Units.Equals("mm"))
                    npdMetric = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
                else
                    npdMetric = pipeInfo.NominalDiameter.Size / 1000.0;
                //check valid pipe size
                NominalDiameter minNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                NominalDiameter maxNominalDiameter = new NominalDiameter();
                maxNominalDiameter.Size = 150;
                maxNominalDiameter.Units = "mm";
                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 32, 90 }, "mm");

                //check valid pipe size
                if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5365_Anchor_C.cs", 130);

                double width = 5 / 1000.0, depth = 50 / 1000.0, length = pipeInfo.OutsideDiameter / 2.0 + 30 / 1000.0;
                string bomDesc = "Flat bar " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, depth, UnitName.DISTANCE_MILLIMETER) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_MILLIMETER) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, width, UnitName.DISTANCE_MILLIMETER);

                (componentDictionary[LEFTGUIDE]).SetPropertyValue(length, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                (componentDictionary[LEFTGUIDE]).SetPropertyValue(width, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                (componentDictionary[LEFTGUIDE]).SetPropertyValue(depth, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                (componentDictionary[LEFTGUIDE]).SetPropertyValue(bomDesc, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");

                (componentDictionary[RIGHTGUIDE]).SetPropertyValue(length, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                (componentDictionary[RIGHTGUIDE]).SetPropertyValue(width, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                (componentDictionary[RIGHTGUIDE]).SetPropertyValue(depth, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");
                (componentDictionary[RIGHTGUIDE]).SetPropertyValue(bomDesc, "IJOAHgrUtility_USER_FIXED_BOX", "BOM_DESC");


                JointHelper.CreateRigidJoint("-1", "Route", RIGHTGUIDE, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -30 / 1000.0, pipeInfo.OutsideDiameter / 2.0 + 2.5 / 1000.0, 0);
                JointHelper.CreateRigidJoint("-1", "Route", LEFTGUIDE, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -30 / 1000.0, -pipeInfo.OutsideDiameter / 2.0 - 2.5 / 1000.0, 0);
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of SFS5365_Guide_H." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string BOMString = "";
            try
            {
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                return BOMString = "Guide H SFS 5365 DN " + pipeInfo.NominalDiameter.Size;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5365_Guide_H" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion


    }

}