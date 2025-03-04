//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5397_D3.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5397_D3
//   Author       :  Vijaya
//   Creation Date: 1-Jul-2013 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   1-Jul-2013     Vijaya   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
//  11-Dec-2014      PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Content.Support.Symbols;

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
    public class SFS5397_D3 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HORSTEEL5397D3 = "HORSTEEL5397_D2";
        private const string VERTSTEEL5397D3 = "VERTSTEEL5397_D2";

        private string size3;
        double pipeDiameter = 0.0, shoeHeight = 0.0, overHang = 0.0, horizontalLength;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PropertyValueCodelist sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize3", "Size3");
                    size3 = sizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeCodeList.PropValue).DisplayName;

                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLOverHang", "OverHang")).PropValue;

                    horizontalLength = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal) + overHang;

                    parts.Add(new PartInfo(HORSTEEL5397D3, "Utility_GENERIC_W_1"));
                    if (horizontalLength >= 500.00 / 1000.00)// with clamps
                        parts.Add(new PartInfo(VERTSTEEL5397D3, "Utility_CUTBACK_W2_1"));


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

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                // To get Pipe Nominal Diameter
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                pipeDiameter = pipeInfo.OutsideDiameter + pipeInfo.InsulationThickness * 2;

                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                NominalDiameter minNominalDiameter = new NominalDiameter(), maxNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                maxNominalDiameter.Size = 1200;
                maxNominalDiameter.Units = "mm";

                if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5397_D3.cs", 110);
                    return;
                }

                double maxLen = 0.0, MA = 0.0, MP = 0.0;
                PropertyValue WSize1, WSize2;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size3", size3);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_D3_Dim", "IJUAFINLSrv_SFS5397_D3_Dim", "W_Size1", parameter, out WSize1);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_D3_Dim", "IJUAFINLSrv_SFS5397_D3_Dim", "W_Size2", parameter, out WSize2);
                string wSize1 = WSize1.ToString(), wSize2 = WSize2.ToString();


                maxLen = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_D3_Dim", "IJUAFINLSrv_SFS5397_D3_Dim", "Max_Len", "IJUAFINLSrv_SFS5397_D3_Dim", "Size3", size3);
                MA = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_D3_Dim", "IJUAFINLSrv_SFS5397_D3_Dim", "MA", "IJUAFINLSrv_SFS5397_D3_Dim", "Size3", size3);
                MP = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_D3_Dim", "IJUAFINLSrv_SFS5397_D3_Dim", "MP", "IJUAFINLSrv_SFS5397_D3_Dim", "Size3", size3);


                support.SetPropertyValue(MA, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(MP, "IJUAFINLMPMA", "MP");
                double w1Width, w1FlangeThickness, w1WebThickness, w1Depth, w2Width, w2FlangeThickness, w2WebThickness, w2Depth;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "W", wSize1, out  w1Width, out   w1FlangeThickness, out  w1WebThickness, out  w1Depth);
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "W", wSize2, out  w2Width, out   w2FlangeThickness, out  w2WebThickness, out  w2Depth);


                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5397_D3.cs", 137);
                    return;
                }

                double theta = Math.Abs(Math.PI / 2 - RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "", PortAxisType.Y, OrientationAlong.Global_Z));
                double angleOffset = Math.Tan(theta) * w1Width;


                double structDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                horizontalLength = structDistance + overHang;


                if (structDistance + overHang > maxLen)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + (maxLen * 1000).ToString() + "mm exceeded.", "", "SFS5397_D3.cs", 150);


                if (overHang < pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Overhang must exceed pipe outside dia:" + (pipeDiameter / 2 * 1000).ToString() + "mm.", "", "SFS5397_D3.cs", 154);


                double verticalLength = (550.00 / 1000.00 - 1.0 / 2.0 * (w2Depth / Math.Tan(Math.PI / 3) + (w2Depth / Math.Tan(Math.PI / 6))));
                double verticalOffset = 550.00 / 1000.00 * Math.Cos(Math.PI / 6) - horizontalLength;


                string inputBOM = wSize1 + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[HORSTEEL5397D3].SetPropertyValue(w1Width, "IJOAHgrUtility_GENERIC_W", "WIDTH");
                componentDictionary[HORSTEEL5397D3].SetPropertyValue(w1Depth, "IJOAHgrUtility_GENERIC_W", "DEPTH");
                componentDictionary[HORSTEEL5397D3].SetPropertyValue(horizontalLength, "IJOAHgrUtility_GENERIC_W", "L");
                componentDictionary[HORSTEEL5397D3].SetPropertyValue(w1FlangeThickness, "IJOAHgrUtility_GENERIC_W", "T_FLANGE");
                componentDictionary[HORSTEEL5397D3].SetPropertyValue(w1WebThickness, "IJOAHgrUtility_GENERIC_W", "T_WEB");
                componentDictionary[HORSTEEL5397D3].SetPropertyValue(inputBOM, "IJOAHgrUtility_GENERIC_W", "BOM_DESC");


                if (horizontalLength >= 500.00 / 1000.00)
                {

                    inputBOM = wSize2 + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, Math.Round(verticalLength, 3), UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                    componentDictionary[VERTSTEEL5397D3].SetPropertyValue(w2Width, "IJOAHgrUtility_GENERIC_W", "WIDTH");
                    componentDictionary[VERTSTEEL5397D3].SetPropertyValue(w2Depth, "IJOAHgrUtility_GENERIC_W", "DEPTH");
                    componentDictionary[VERTSTEEL5397D3].SetPropertyValue(verticalLength, "IJOAHgrUtility_GENERIC_W", "L");
                    componentDictionary[VERTSTEEL5397D3].SetPropertyValue(w2FlangeThickness, "IJOAHgrUtility_GENERIC_W", "T_FLANGE");
                    componentDictionary[VERTSTEEL5397D3].SetPropertyValue(w2FlangeThickness, "IJOAHgrUtility_GENERIC_W", "T_WEB");
                    componentDictionary[VERTSTEEL5397D3].SetPropertyValue(Math.PI / 3, "IJOAHgrUtility_CUTBACK", "ANGLE");//rotae about y-axis
                    componentDictionary[VERTSTEEL5397D3].SetPropertyValue(Math.PI / 6, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                    componentDictionary[VERTSTEEL5397D3].SetPropertyValue(inputBOM, "IJOAHgrUtility_GENERIC_W", "BOM_DESC");

                }

                double flipAngle = 0.0;

                if (Configuration == 2)
                    flipAngle = Math.PI;

                double verticalDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                double theta1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference || SupportHelper.PlacementType == PlacementType.PlaceByPoint)
                {
                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D3, "Structure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, (structDistance + overHang), (pipeDiameter / 2 + w1Depth + shoeHeight - angleOffset / 2 - w1Depth / 2 + verticalDistance), 0);
                    else
                        if (Symbols.HgrCompareDoubleService.cmpdbl(flipAngle, Math.PI) == true)
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D3, "Structure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeY, (structDistance + overHang), (-pipeDiameter / 2 - shoeHeight - angleOffset / 2 - w1Depth / 2 + verticalDistance), 0);

                    if (horizontalLength >= 500.00 / 1000.00)
                        JointHelper.CreateRigidJoint(HORSTEEL5397D3, "BotStructure", VERTSTEEL5397D3, "EndStructure", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0, 0, -verticalOffset);

                }
                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (Math.Abs(theta1) < Math.PI / 2)
                        {
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D3, "Structure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, structDistance + overHang, 0, -pipeDiameter / 2 - shoeHeight - angleOffset / 2 - w1Depth / 2);

                            if (horizontalLength >= 500.00 / 1000.00)
                                JointHelper.CreateRigidJoint(HORSTEEL5397D3, "BotStructure", VERTSTEEL5397D3, "EndStructure", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0, 0, -verticalOffset);

                        }
                        else
                        {//flip
                            JointHelper.CreateRigidJoint("-1", "Structure", HORSTEEL5397D3, "Structure", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeZ, structDistance + overHang, 0, (pipeDiameter / 2 + shoeHeight + angleOffset / 2 + w1Depth / 2));

                            if (horizontalLength >= 500.00 / 1000.00)
                                JointHelper.CreateRigidJoint(HORSTEEL5397D3, "BotStructure", VERTSTEEL5397D3, "EndStructure", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0, 0, -verticalOffset);

                        }
                    }
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
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();

                    routeConnections.Add(new ConnectionInfo(HORSTEEL5397D3, 1));

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
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(HORSTEEL5397D3, 1));

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
        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomString = "";
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                PropertyValueCodelist sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize3", "Size3");

                bomString = "Console support D3 -" + sizeCodeList.PropValue + " SFS 5397";

                return bomString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - RectHorTypeB" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}


