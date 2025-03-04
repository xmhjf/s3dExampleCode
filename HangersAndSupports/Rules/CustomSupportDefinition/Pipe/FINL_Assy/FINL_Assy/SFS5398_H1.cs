//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5398_H1.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5398_H1
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
    public class SFS5398_H1 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        string size3;
        double span, shoeH, overHang;
        private const string VERSTEEL = "VerSteel";
        private const string HORSTEEL1 = "HorSteel1";
        private const string HORSTEEL2 = "HorSteel2";
        private const string PLATE1 = "Plate1";
        private const string PLATE2 = "Plate2";
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

                    span = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLSpan", "Span")).PropValue;
                    shoeH = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLOverHang", "OverHang")).PropValue;

                    parts.Add(new PartInfo(VERSTEEL, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(HORSTEEL1, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(HORSTEEL2, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(PLATE1, "Util_FourHolePl_Metric_1"));
                    parts.Add(new PartInfo(PLATE2, "Util_FourHolePl_Metric_1"));
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
                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                //check valid pipe size
                if (IsPipeSizeValid(currentDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5398_H1.cs", 104);
                    return;
                }
                PropertyValue tsSize1, tsSize2;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size3", size3);
                CatalogStructHelper catalogHelper = new CatalogStructHelper();
                GenericHelper.GetDataByRule("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "TS_Size", parameter, out tsSize1);
                GenericHelper.GetDataByRule("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "TS_Size", parameter, out tsSize2);
                double maxSpan = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "Max_Span", "IJUAFINLSrv_SFS5398_H1_Dim", "Size3", size3);
                double maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "Max_Len", "IJUAFINLSrv_SFS5398_H1_Dim", "Size3", size3);
                double ma = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "MA", "IJUAFINLSrv_SFS5398_H1_Dim", "Size3", size3);
                double mp = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "MP", "IJUAFINLSrv_SFS5398_H1_Dim", "Size3", size3);
                double fMax = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "Fmax", "IJUAFINLSrv_SFS5398_H1_Dim", "Size3", size3);
                double plateT = 10.0 / 1000;
                double plateW = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "Plate_W", "IJUAFINLSrv_SFS5398_H1_Dim", "Size3", size3);
                double C = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "C", "IJUAFINLSrv_SFS5398_H1_Dim", "Size3", size3);
                double holeDia = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "D", "IJUAFINLSrv_SFS5398_H1_Dim", "Size3", size3);
                double fU = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "MA", "IJUAFINLSrv_SFS5398_H1_Dim", "Size3", size3);
                double fQ = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5398_H1_Dim", "IJUAFINLSrv_SFS5398_H1_Dim", "MP", "IJUAFINLSrv_SFS5398_H1_Dim", "Size3", size3);
                support.SetPropertyValue(fU, "IJUAFINLFQFU", "FU");
                support.SetPropertyValue(fQ, "IJUAFINLFQFU", "FQ");
                support.SetPropertyValue(ma, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(mp, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(fMax, "IJUAFINLFmax", "Fmax");

                double widthHorTS, flangeThickness, depthhorTS, webThickness;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize1), out widthHorTS, out flangeThickness, out webThickness, out depthhorTS);
                double widthVertTS, depthVertTS, TL;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", Convert.ToString(tsSize2), out widthVertTS, out TL, out webThickness, out depthVertTS);

                double verticalLength = span;

                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5398_H1.cs", 139);
                    return;
                }

                double horDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);

                double horLength = horDistance + overHang + depthhorTS - plateT;


                if (horLength + plateT > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + maxLength * 1000 + "mm exceeded.", "", "SFS5398_H1.cs", 149);

                if (span > maxSpan)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Maximum allowable length of " + maxSpan * 1000 + "mm exceeded.", "", "SFS5398_H1.cs", 153);
                    return;
                }
                if (span < pipeDia + shoeH)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Given Shoe Height, span must exceed pipe O.D. plus shoe height: " + pipeDia * 1000 + shoeH * 1000 + " mm.", "", "SFS5398_H1.cs", 157);

                if (overHang < pipeDia / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Overhang must exceed pipe outside dia:" + pipeDia / 2 * 1000 + " mm.", "", "SFS5398_H1.cs", 160);


                double horSteelweight = (widthHorTS * depthhorTS * horLength - ((widthHorTS - 2 * 4.0 / 1000) * (depthhorTS - 2 * 4.0 / 1000) * horLength)) * steelDensityKGperM;

                double vertsSteelWeight = (widthVertTS * depthVertTS * verticalLength - ((widthVertTS - 2 * 4.0 / 1000) * (depthVertTS - 2 * 4.0 / 1000) * verticalLength)) * steelDensityKGperM;

                string vertBOM = tsSize1 + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[VERSTEEL]).SetPropertyValue(widthHorTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[VERSTEEL]).SetPropertyValue(depthhorTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[VERSTEEL]).SetPropertyValue(verticalLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[VERSTEEL]).SetPropertyValue(vertBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[VERSTEEL]).SetPropertyValue(vertsSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                string hor1BOM = tsSize2 + "Length" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                (componentDictionary[HORSTEEL1]).SetPropertyValue(widthVertTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(depthVertTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(horLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(hor1BOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[HORSTEEL1]).SetPropertyValue(horSteelweight, "IJOAHgrUtilMetricInWt", "InputWeight");

                (componentDictionary[HORSTEEL2]).SetPropertyValue(widthVertTS, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(depthVertTS, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(horLength, "IJOAHgrUtilMetricL", "L");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(hor1BOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                (componentDictionary[HORSTEEL2]).SetPropertyValue(horSteelweight, "IJOAHgrUtilMetricInWt", "InputWeight");

                string plateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateT, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Steel Plate " + (MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateW, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0)) + "x" + (MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateW, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0)) + ", 4 eq spaced " + (MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, holeDia, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0)) + " dia. holes " + (MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, ((plateW - C) / 2), UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0)) + " from edges";

                (componentDictionary[PLATE1]).SetPropertyValue(plateW, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[PLATE1]).SetPropertyValue(plateW, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[PLATE1]).SetPropertyValue(plateT, "IJOAHgrUtilMetricT", "T");
                (componentDictionary[PLATE1]).SetPropertyValue(holeDia, "IJOAHgrUtilMetricHoleSize", "HoleSize");
                (componentDictionary[PLATE1]).SetPropertyValue(((plateW - C) / 2), "IJOAHgrUtilMetricC", "C");
                (componentDictionary[PLATE1]).SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                (componentDictionary[PLATE2]).SetPropertyValue(plateW, "IJOAHgrUtilMetricWidth", "Width");
                (componentDictionary[PLATE2]).SetPropertyValue(plateW, "IJOAHgrUtilMetricDepth", "Depth");
                (componentDictionary[PLATE2]).SetPropertyValue(plateT, "IJOAHgrUtilMetricT", "T");
                (componentDictionary[PLATE2]).SetPropertyValue(holeDia, "IJOAHgrUtilMetricHoleSize", "HoleSize");
                (componentDictionary[PLATE2]).SetPropertyValue(((plateW - C) / 2), "IJOAHgrUtilMetricC", "C");
                (componentDictionary[PLATE2]).SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");


                double theta = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "", PortAxisType.X, OrientationAlong.Global_Z);

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference || SupportHelper.PlacementType == PlacementType.PlaceByPoint)
                {
                    if (Math.Abs(theta) < Math.PI / 2)
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, -overHang - depthhorTS, -pipeDia / 2 - depthhorTS / 2 - shoeH, 0);
                    else
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, -overHang - depthhorTS, pipeDia / 2 + depthhorTS / 2 + shoeH, 0);

                    JointHelper.CreateRigidJoint(HORSTEEL1, "StartOther", VERSTEEL, "StartOther", Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, depthhorTS / 2, depthhorTS / 2, 0);

                    JointHelper.CreateRigidJoint(VERSTEEL, "EndOther", HORSTEEL2, "StartOther", Plane.XY, Plane.ZX, Axis.X, Axis.X, depthhorTS / 2, depthhorTS / 2, 0);

                    JointHelper.CreateRigidJoint(HORSTEEL1, "EndOther", PLATE1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                    JointHelper.CreateRigidJoint(HORSTEEL2, "EndOther", PLATE2, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                }
                else
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Support can only be placed on Wall By-Point.", "", "SFS5398_H1.cs", 225);
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
            PropertyValueCodelist sizeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJUAFINLSize3", "Size3");
            size3 = sizeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeList.PropValue).DisplayName;
            string BOMdesc = "Gate support H1 -" + size3 + " SFS 5398";

            return BOMdesc;
        }
    }
}