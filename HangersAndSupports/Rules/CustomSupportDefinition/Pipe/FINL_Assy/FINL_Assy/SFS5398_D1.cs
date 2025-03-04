//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5398_D1.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5398_D1
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
    public class SFS5398_D1 : CustomSupportDefinition, ICustomHgrBOMDescription
    {

        string size4;
        private const string HORSTEEL = "HorSteel";
        private const string VERSTEEL = "VerSteel";
        private const string PLATE1 = "Plate1";
        private const string PLATE2 = "Plate2";
        private const int steelDensityKGperM = 7900;
        double X1, X2, H, shoeH;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PropertyValueCodelist sizeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize4", "Size4");
                    size4 = sizeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeList.PropValue).DisplayName;
                    shoeH = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    X1 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLX1", "X1")).PropValue;
                    X2 = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLX2", "X2")).PropValue;
                    H = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLHeight", "H")).PropValue;

                    parts.Add(new PartInfo(VERSTEEL, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(HORSTEEL, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(PLATE1, "Util_FourHolePl_Metric_1"));   //from_utility_metric
                    parts.Add(new PartInfo(PLATE2, "Util_FourHolePl_Metric_1"));   //from_utility_metric
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
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5398_D1.cs", 100);
                    return;
                }

                CatalogStructHelper catalogHelper = new CatalogStructHelper();


                PropertyValue tsSize1, tsSize2;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size4", size4);

                GenericHelper.GetDataByRule("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "TS_Size1", parameter, out tsSize1);
                GenericHelper.GetDataByRule("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "TS_Size2", parameter, out tsSize2);
                double maxSpan = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "Max_Span", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "Max_Len", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double ma = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "MA", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double mp = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "MP", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double fMax = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "Fmax", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double plateT = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "Plate_T", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double plateD = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "Plate_D", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double plateW = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "Plate_W", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double inset = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "Inset", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double C = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "C", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double holeDia = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "D", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double fU = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "MA", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);
                double fQ = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_D1_Dim", "IJUAFINLSrv_SFS5398_D1_Dim", "MP", "IJUAFINLSrv_SFS5398_D1_Dim", "Size4", size4);

                support.SetPropertyValue(fU, "IJUAFINLFQFU", "FU");
                support.SetPropertyValue(fQ, "IJUAFINLFQFU", "FQ");
                support.SetPropertyValue(ma, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(mp, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(fMax, "IJUAFINLFmax", "Fmax");

                double widthHorTS, flangeThickness, depthhorTS, webThickness;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize1), out widthHorTS, out flangeThickness, out webThickness, out depthhorTS);
                double widthVertTS, depthVertTS, TL;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize2), out widthVertTS, out TL, out webThickness, out depthVertTS);

                double horLength = X1 + X2;


                double vertLength = H + C / 2;

                // Checking for max H
                if (H > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + maxLength * 1000 + "mm exceeded.", "", "SFS5398_D1.cs", 224);
                //If the span length entered is greater than allowed
                if (X1 + X2 > maxSpan)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Maximum allowable length of " + maxLength * 1000 + "mm exceeded.", "", "SFS5398_D1.cs", 149);
                    return;
                }
                if ((pipeDia / 2 > X1) || (pipeDia / 2 > X2))
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "X1 and X2 must exceed Pipe outside radius.", "", "SFS5398_D1.cs", 154);
                    return;
                }
                double vertsSteelWeight = (widthVertTS * depthVertTS * vertLength - ((widthVertTS - 2 * 5.0 / 1000) * (depthVertTS - 2 * 5.0 / 1000) * vertLength)) * steelDensityKGperM;
                double horSteelweight;

                if (size4 == "3")
                    horSteelweight = (widthHorTS * depthhorTS * horLength - ((widthHorTS - 2 * 5.0 / 1000) * (depthhorTS - 2 * 5.0 / 1000) * horLength)) * steelDensityKGperM;
                else
                    horSteelweight = (widthHorTS * depthhorTS * horLength - ((widthHorTS - 2 * 4.0 / 1000) * (depthhorTS - 2 * 4.0 / 1000) * horLength)) * steelDensityKGperM;

                string horBOM = tsSize1 + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[HORSTEEL]).SetPropertyValue(widthHorTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[HORSTEEL]).SetPropertyValue(depthhorTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[HORSTEEL]).SetPropertyValue(horLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[HORSTEEL]).SetPropertyValue(horBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[HORSTEEL]).SetPropertyValue(horSteelweight, "IJOAHgrUtilMetricInWt", "InputWeight");

                string vertBOM = tsSize2 + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, vertLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[VERSTEEL]).SetPropertyValue(widthVertTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[VERSTEEL]).SetPropertyValue(depthVertTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[VERSTEEL]).SetPropertyValue(vertLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[VERSTEEL]).SetPropertyValue(vertBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[VERSTEEL]).SetPropertyValue(vertsSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                string plateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateT, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Steel Plate " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateW, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateD, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + ", 4 eq spaced " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, holeDia, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " dia. holes " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, inset, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " from edges";

                (componentDictionary[PLATE1]).SetPropertyValue(plateD, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[PLATE1]).SetPropertyValue(plateW, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[PLATE1]).SetPropertyValue(plateT, "IJOAHgrUtilMetricT", "T");
                (componentDictionary[PLATE1]).SetPropertyValue(holeDia, "IJOAHgrUtilMetricHoleSize", "HoleSize");
                (componentDictionary[PLATE1]).SetPropertyValue(inset, "IJOAHgrUtilMetricC", "C");
                (componentDictionary[PLATE1]).SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                (componentDictionary[PLATE2]).SetPropertyValue(plateD, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[PLATE2]).SetPropertyValue(plateW, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[PLATE2]).SetPropertyValue(plateT, "IJOAHgrUtilMetricT", "T");
                (componentDictionary[PLATE2]).SetPropertyValue(holeDia, "IJOAHgrUtilMetricHoleSize", "HoleSize");
                (componentDictionary[PLATE2]).SetPropertyValue(inset, "IJOAHgrUtilMetricC", "C");
                (componentDictionary[PLATE2]).SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                double theta = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "", PortAxisType.X, OrientationAlong.Global_Z);

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference || SupportHelper.PlacementType == PlacementType.PlaceByPoint)
                {
                    if (Math.Abs(theta) < Math.PI / 2)
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL, "StartOther", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, widthHorTS / 2 + plateT, pipeDia / 2 + shoeH + depthhorTS / 2, horLength - X1);
                    else
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL, "StartOther", Plane.XY, Plane.ZX, Axis.X, Axis.Z, widthHorTS / 2 + plateT, -(pipeDia / 2 + shoeH + depthhorTS / 2), -X2);

                    JointHelper.CreateRigidJoint(HORSTEEL, "StartOther", VERSTEEL, "StartOther", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, -depthVertTS / 2, (widthVertTS / 2 - widthHorTS / 2), -depthhorTS / 2);
                    JointHelper.CreateRigidJoint(VERSTEEL, "EndOther", PLATE1, "TopStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, -widthVertTS / 2, 0);
                    JointHelper.CreateRigidJoint(HORSTEEL, "EndOther", PLATE2, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                else
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Support can only be placed on Wall By-Point.", "", "SFS5398_D1.cs", 212);
                    return;
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

                    routeConnections.Add(new ConnectionInfo(HORSTEEL, 1)); // partindex, routeindex


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

                    structConnections.Add(new ConnectionInfo(PLATE1, 1)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(PLATE2, 1)); // partindex, routeindex


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
            string BOMdesc = "Gate support D2 -" + size4 + " SFS 5398";

            return BOMdesc;
        }
    }
}