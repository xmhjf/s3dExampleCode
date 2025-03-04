//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5365_Guide_A.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5365_Guide_A
//   Author       :  BS
//   Creation Date:  04/07/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  04/07/2013       BS    CR-CP-224485- Converted HS_FINL_Assy to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using System.Linq;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class SFS5365_Guide_A : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        CustomSupportDefinition Shoe, Guide, l1, l2;

        string slideType = string.Empty, sectionSize = string.Empty;
        int guideSide;
        double shoeWidthFromDB, npdMetric, shoeHeightFromDB, clampInnerDiameter, gap, widthL, thicknessL, depthL, sectionL, webThicknessL;//shoeHeight,shoeWidth,
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    slideType = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSlideType", "SlideType")).PropValue;
                    PropertyValueCodelist guideSub = ((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLGuideType", "GuideType"));
                    string sGuideSub = guideSub.PropertyInfo.CodeListInfo.GetCodelistItem((int)guideSub.PropValue).DisplayName;

                    Shoe = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + slideType.Trim(), support);
                    Guide = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + sGuideSub.Trim(), support);


                    PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    npdMetric = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
                    shoeWidthFromDB = FINLAssemblyServices.GetDataByCondition("FINLSrv_" + slideType + "_Dim", "IJUAFINLSrv_" + slideType + "_Dim", "Shoe_W", "IJUAFINLSrv_" + slideType + "_Dim", "Pipe_Nom_Dia_m", (npdMetric * 1000).ToString());
                    shoeHeightFromDB = FINLAssemblyServices.GetDataByCondition("FINLSrv_" + slideType + "_Dim", "IJUAFINLSrv_" + slideType + "_Dim", "Shoe_H", "IJUAFINLSrv_" + slideType + "_Dim", "Pipe_Nom_Dia_m", (npdMetric * 1000).ToString());
                    if (slideType.Equals("SFS5859") || slideType.Equals("SFS5860"))
                    {
                        sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSectionSize2", "SectionSize2")).PropValue;
                        sectionL = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLSectionL", "SectionL")).PropValue;

                        l1 = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + "Utility_L_Wrapper", support);
                        FINLAssemblyServices.SetValueOnPropertyType(l1, "Index", 1);
                        FINLAssemblyServices.SetValueOnPropertyType(l1, "Override", true);
                        FINLAssemblyServices.SetValueOnPropertyType(l1, "SectionSize", sectionSize);
                        FINLAssemblyServices.SetValueOnPropertyType(l1, "SectionL", sectionL);
                        parts = l1.Parts;

                        l2 = FINLAssemblyServices.GetAssembly("FINL_Assy,Ingr.SP3D.Content.Support.Rules." + "Utility_L_Wrapper", support);
                        FINLAssemblyServices.SetValueOnPropertyType(l2, "Index", 2);
                        FINLAssemblyServices.SetValueOnPropertyType(l2, "Override", true);
                        FINLAssemblyServices.SetValueOnPropertyType(l2, "SectionSize", sectionSize);
                        FINLAssemblyServices.SetValueOnPropertyType(l2, "SectionL", sectionL);
                        parts = new Collection<PartInfo>(parts.Concat(l2.Parts).ToList());

                        FINLAssemblyServices.GetCrossSectionDimensions("Euro", "L", sectionSize, out widthL, out thicknessL, out webThicknessL, out depthL);
                        if (slideType.Equals("SFS5859"))
                            clampInnerDiameter = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5370", "IJUAHgrPipe_Dia", "PIPE_DIA", "IJUAFINL_PipeND_mm", "PipeND", npdMetric - 0.003, npdMetric + 0.003);
                        else
                            clampInnerDiameter = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5856", "IJUAHgrPipe_Dia", "PIPE_DIA", "IJUAFINL_PipeND_mm", "PipeND", npdMetric - 0.003, npdMetric + 0.003);
                    }
                    
                    guideSide = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLGuideSide", "GuideSide")).PropValue;
                    gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLGap2", "Gap2")).PropValue;

                    FINLAssemblyServices.SetValueOnPropertyType(Guide, "Override", true);
                    FINLAssemblyServices.SetValueOnPropertyType(Guide, "ShoeHeight", shoeHeightFromDB);
                    FINLAssemblyServices.SetValueOnPropertyType(Guide, "ShoeWidth", (shoeWidthFromDB + 2 * depthL));
                    FINLAssemblyServices.SetValueOnPropertyType(Guide, "GuideSide", guideSide);
                    FINLAssemblyServices.SetValueOnPropertyType(Guide, "Gap", gap);

                    parts = new Collection<PartInfo>(parts.Concat(Guide.Parts).Concat(Shoe.Parts).ToList());
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
                    routeConnections = Shoe.SupportedConnections;
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
                    structConnections = Shoe.SupportingConnections;
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

                ReadOnlyCollection<object> shoeJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, Shoe, oSupCompColl);
                ReadOnlyCollection<object> guideJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, Guide, oSupCompColl);
                JointHelper.m_oCollOfJoints = new Collection<object>(shoeJoints.Concat(guideJoints).ToList());

                Collection<ConnectionInfo> iSubAssyRouteConShoe = Shoe.SupportedConnections;
                Collection<ConnectionInfo> iSubAssyRoutConGuide = Guide.SupportedConnections;
                JointHelper.CreateRigidJoint("-1", "Route", iSubAssyRoutConGuide[0].PartKey, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                if (slideType.Equals("SFS5859") || slideType.Equals("SFS5860"))
                {
                    ReadOnlyCollection<object> l1Joints = FINLAssemblyServices.GetSubAssemblyJoints(this, l1, oSupCompColl);
                    ReadOnlyCollection<object> l2Joints = FINLAssemblyServices.GetSubAssemblyJoints(this, l2, oSupCompColl);
                    JointHelper.m_oCollOfJoints = new Collection<object>(JointHelper.m_oCollOfJoints.Concat(l1Joints).Concat(l2Joints).ToList());

                    Collection<ConnectionInfo> iSubAssyRouteConL1 = l1.SupportedConnections;
                    Collection<ConnectionInfo> iSubAssyRouteConL2 = l2.SupportedConnections;

                    JointHelper.CreateRigidJoint("-1", "Route", iSubAssyRouteConL1[0].PartKey, "Connection", Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeZ, -sectionL / 2, clampInnerDiameter / 2 + shoeHeightFromDB, -shoeWidthFromDB / 2);
                    JointHelper.CreateRigidJoint("-1", "Route", iSubAssyRouteConL2[0].PartKey, "Connection", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX, clampInnerDiameter / 2 + shoeHeightFromDB, shoeWidthFromDB / 2, sectionL / 2);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of SFS5365_Guide_A." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                return BOMString = "Guide A SFS 5365 DN  " + pipeInfo.NominalDiameter.Size;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5365_Guide_A" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion


    }

}