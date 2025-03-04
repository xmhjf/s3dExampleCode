//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5397_E1.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5397_E1
//   Author       :  Manikanth
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     Manikanth   CR-CP-224492- Convert HS_FINL_Assy to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
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
    [SymbolVersion("1.0.0.0")]
    public class SFS5397_E1 : CustomSupportDefinition, ICustomHgrBOMDescription
    {

        string size3;
        double L1, L2, shoeH, gap, overHang;
        private const string HORSTEEL1 = "HorSteel1";
        private const string HORSTEEL2 = "HorSteel2";
        private const string HORSTEEL3 = "HorSteel3";
        private const string SPINEL1 = "SpineL1";
        private const string SPINEL2 = "SpineL2";
        private const int steelDensityKGperM = 7900;
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PropertyValueCodelist sizeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize3", "Size3");
                    size3 = sizeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeList.PropValue).DisplayName;
                    gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLGap2", "Gap2")).PropValue;
                    L1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLL1", "L1")).PropValue;
                    L2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLL2", "L2")).PropValue;
                    shoeH = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLOverHang", "OverHang")).PropValue;

                    parts.Add(new PartInfo(HORSTEEL1, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(HORSTEEL2, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(HORSTEEL3, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(SPINEL1, "Utility_GENERIC_L_1"));
                    parts.Add(new PartInfo(SPINEL2, "Utility_GENERIC_L_1"));
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

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                //To get Pipe Nom Dia

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                NominalDiameter currentDiameter = new NominalDiameter();
                currentDiameter = pipeInfo.NominalDiameter;
                double pipeDia = pipeInfo.OutsideDiameter;
                NominalDiameter minNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                NominalDiameter maxNominalDiameter = new NominalDiameter();
                maxNominalDiameter.Size = 1200;
                maxNominalDiameter.Units = "mm";
                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 10, 1200, 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                //check valid pipe size
                if (IsPipeSizeValid(currentDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5397_E1.cs", 104);
                    return;
                }

                CatalogStructHelper catalogHelper = new CatalogStructHelper();
                PropertyValue tsSize, lSize;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size3", size3);

                GenericHelper.GetDataByRule("FINLSrv_SFS5397_E1_Dim", "IJUAFINLSrv_SFS5397_E1_Dim", "TS_Size", parameter, out tsSize);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_E1_Dim", "IJUAFINLSrv_SFS5397_E1_Dim", "Brace_Size", parameter, out lSize);
                double maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_E1_Dim", "IJUAFINLSrv_SFS5397_E1_Dim", "Max_Len", "IJUAFINLSrv_SFS5397_E1_Dim", "Size3", size3);
                double ma = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_E1_Dim", "IJUAFINLSrv_SFS5397_E1_Dim", "MA", "IJUAFINLSrv_SFS5397_E1_Dim", "Size3", size3);
                double mp = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_E1_Dim", "IJUAFINLSrv_SFS5397_E1_Dim", "MP", "IJUAFINLSrv_SFS5397_E1_Dim", "Size3", size3);
                double fU = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_E1_Dim", "IJUAFINLSrv_SFS5397_E1_Dim", "FU", "IJUAFINLSrv_SFS5397_E1_Dim", "Size3", size3);
                double fQ = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_E1_Dim", "IJUAFINLSrv_SFS5397_E1_Dim", "FQ", "IJUAFINLSrv_SFS5397_E1_Dim", "Size3", size3);
                support.SetPropertyValue(fU, "IJUAFINLFQFU", "FU");
                support.SetPropertyValue(fQ, "IJUAFINLFQFU", "FQ");
                support.SetPropertyValue(ma, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(mp, "IJUAFINLMPMA", "MP");

                double widthTS, flangeThickness, depthTS, webThickness;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize), out widthTS, out flangeThickness, out webThickness, out depthTS);
                double WidthL, depthL, TL;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(lSize), out WidthL, out TL, out webThickness, out depthL);

                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5397_E1.cs", 132);
                    return;
                }


                double horDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                double horLength = horDistance + overHang;

                if (((horDistance + overHang) > maxLength) || (L1 > maxLength) || (L2 > maxLength))
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + maxLength * 1000 + "mm exceeded.", "", "SFS5397_E1.cs", 141);

                if (overHang < pipeDia / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Overhang must exceed pipe outside dia:" + pipeDia / 2 * 1000 + " mm..", "", "SFS5397_E1.cs", 144);

                double horSteelWeight, horSteelWeight1, horSteelWeight2;

                if (size3 == "3")
                {
                    horSteelWeight = (widthTS * depthTS * horLength - ((widthTS - 2 * 5.0 / 1000) * (depthTS - 2 * 5.0 / 1000) * horLength)) * steelDensityKGperM;
                    horSteelWeight1 = (widthTS * depthTS * L2 - ((widthTS - 2 * 5.0 / 1000) * (depthTS - 2 * 5.0 / 1000) * L2)) * steelDensityKGperM;
                    horSteelWeight2 = (widthTS * depthTS * L1 - ((widthTS - 2 * 5.0 / 1000) * (depthTS - 2 * 5.0 / 1000) * L1)) * steelDensityKGperM;
                }
                else
                {
                    horSteelWeight = (widthTS * depthTS * horLength - ((widthTS - 2 * 4.0 / 1000) * (depthTS - 2 * 4.0 / 1000) * horLength)) * steelDensityKGperM;
                    horSteelWeight2 = (widthTS * depthTS * L2 - ((widthTS - 2 * 4.0 / 1000) * (depthTS - 2 * 4.0 / 1000) * L2)) * steelDensityKGperM;
                    horSteelWeight1 = (widthTS * depthTS * L1 - ((widthTS - 2 * 4.0 / 1000) * (depthTS - 2 * 4.0 / 1000) * L1)) * steelDensityKGperM;
                }


                double LL = 30.0 / 1000 + gap + gap + widthTS + 10.0 / 1000;

                string hor1BOM = tsSize + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L1, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                //horizontal HSS section
                (componentDictionary[HORSTEEL1]).SetPropertyValue(widthTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(depthTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(L1, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(hor1BOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(horSteelWeight1, "IJOAHgrUtilMetricInWt", "InputWeight");

                string hor2BOM = tsSize + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, L2, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[HORSTEEL2]).SetPropertyValue(widthTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(depthTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(L2, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(hor2BOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(horSteelWeight2, "IJOAHgrUtilMetricInWt", "InputWeight");

                string hor3BOM = tsSize + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[HORSTEEL3]).SetPropertyValue(widthTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[HORSTEEL3]).SetPropertyValue(depthTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[HORSTEEL3]).SetPropertyValue(horLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[HORSTEEL3]).SetPropertyValue(hor3BOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[HORSTEEL3]).SetPropertyValue(horSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                string hor4BOM = tsSize + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                //verical spine L
                (componentDictionary[SPINEL1]).SetPropertyValue(WidthL, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                (componentDictionary[SPINEL1]).SetPropertyValue(depthL, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                (componentDictionary[SPINEL1]).SetPropertyValue(LL, "IJOAHgrUtility_GENERIC_L", "L");
                (componentDictionary[SPINEL1]).SetPropertyValue(hor3BOM, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");
                (componentDictionary[SPINEL1]).SetPropertyValue(TL, "IJOAHgrUtility_GENERIC_L", "THICKNESS");

                string hor5BOM = tsSize + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[SPINEL2]).SetPropertyValue(WidthL, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                (componentDictionary[SPINEL2]).SetPropertyValue(depthL, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                (componentDictionary[SPINEL2]).SetPropertyValue(LL, "IJOAHgrUtility_GENERIC_L", "L");
                (componentDictionary[SPINEL2]).SetPropertyValue(hor4BOM, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");
                (componentDictionary[SPINEL2]).SetPropertyValue(TL, "IJOAHgrUtility_GENERIC_L", "THICKNESS");

                double theta = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);

                if (theta < Math.PI / 2)
                {
                    //this is the one on the pipe
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL3, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -overHang, -(widthTS / 2 + pipeDia / 2 + shoeH), 0);
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL2, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, horDistance, gap - (depthTS / 2 + pipeDia / 2 + shoeH), 0);
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, horDistance, 2 * gap - (depthTS / 2 + pipeDia / 2 + shoeH), 0);
                    //spineL
                    JointHelper.CreateRigidJoint(HORSTEEL3, "EndOther", SPINEL1, "Structure", Plane.XY, Plane.ZX, Axis.X, Axis.Z, 0, -widthTS / 2 - 10.0 / 1000, depthTS / 2);
                    JointHelper.CreateRigidJoint(HORSTEEL3, "EndOther", SPINEL2, "Structure", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, 0, LL - widthTS / 2 - 10.0 / 1000, -depthTS / 2);
                }
                else
                {
                    //this is the one on the pipe
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL3, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, -overHang, (widthTS / 2 + pipeDia / 2 + shoeH), 0);
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL2, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, horDistance, -gap + (depthTS / 2 + pipeDia / 2 + shoeH), 0);
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, horDistance, 2 * -gap + (depthTS / 2 + pipeDia / 2 + shoeH), 0);
                    //spineL
                    JointHelper.CreateRigidJoint(HORSTEEL3, "EndOther", SPINEL1, "Structure", Plane.XY, Plane.ZX, Axis.X, Axis.Z, 0, -widthTS / 2 - 10.0 / 1000, depthTS / 2);
                    JointHelper.CreateRigidJoint(HORSTEEL3, "EndOther", SPINEL2, "Structure", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, 0, LL - widthTS / 2 - 10.0 / 1000, -depthTS / 2);
                }


            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of FINL_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();

                    routeConnections.Add(new ConnectionInfo(HORSTEEL1, 1)); // partindex, routeindex


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

                    structConnections.Add(new ConnectionInfo(HORSTEEL1, 1)); // partindex, routeindex


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

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            PropertyValueCodelist sizeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAFINLSize3", "Size3");
            size3 = sizeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeList.PropValue).DisplayName;
            string BOMdesc = "Console support E1" + size3 + " SFS 5397";

            return BOMdesc;
        }
    }
}