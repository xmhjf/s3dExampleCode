//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_353.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_353
//   Author       :  Hema
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013     Hema    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//                          WeightCG implementation is commented ,as in VB "WEIGHT_PER_LENGTH" property is not available in refdata.
//   30-Dec-2014    PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//   11-Jun-2015    PVK      TR-CP-239160	Issues observed in .Net PSL Parts
//   16-Jul-2015    PVK      Resolve coverity issues found in July 2015 report
//   26-Oct-2015    PVK      Resolve coverity issues found in October 2015 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_353 : HangerComponentSymbolDefinition, ICustomWeightCG, ICustomHgrBOMDescription
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_353"

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "STANCHION_DIA", "STANCHION_DIA", 1)]
        public InputDouble STANCHION_DIA;
        [InputDouble(3, "ELBOW_OPT", "ELBOW_OPT", 1)]
        public InputDouble ELBOW_OPT;
        [InputDouble(4, "Length", "Length", 0.999999)]
        public InputDouble Length;
        [InputString(5, "PIPE_NOM_DIA", "PIPE_NOM_DIA", "No Value")]
        public InputString PIPE_NOM_DIA;
        [InputDouble(6, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(7, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(8, "B", "B", 0.999999)]
        public InputDouble B;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("PIPE", "PIPE")]
        [SymbolOutput("BASE", "BASE")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                //-----------------------------------------------------------------------
                // Construction of Physical Aspect 
                //----------------------------------------------------------------------
                Part part = (Part)PartInput.Value;
                const Double constant1 = 0.1143, constant2 = 0.015;
                Double length = Length.Value;
                Double pipeDiameter = PIPE_DIA.Value;
                Double d = D.Value;
                Double b = B.Value;
                String pipeNominalDiameter = PIPE_NOM_DIA.Value;
                Double stanchionDiameter = STANCHION_DIA.Value;
                Double elbowOption = ELBOW_OPT.Value;
                Double elbowRadius = Convert.ToDouble(pipeNominalDiameter) * 3;
                Double actualStanchionDiameter, lAdj = constant2, Y, outerRadius, L, baseWidth;
                String actualElbowOption;

                if (stanchionDiameter < 1 || stanchionDiameter > 28)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidStanctionDiameterCodelist, "STANCHION_DIA codelist values should be between 1 and 28"));
                    return;
                }
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    actualStanchionDiameter = Convert.ToDouble(metadataManager.GetCodelistInfo("PSL_353_STANCHION", "UDP").GetCodelistItem((int)stanchionDiameter).ShortDisplayName.Trim()) / 1000;
                else
                    actualStanchionDiameter = constant1;

                if (metadataManager != null)
                    actualElbowOption = metadataManager.GetCodelistInfo("PSL_ELBOW", "UDP").GetCodelistItem((int)elbowOption).ShortDisplayName.Trim();
                else
                    actualElbowOption = "For Straight Pipe";

                if (actualElbowOption.Equals("For 1.5D Elbow"))
                    elbowRadius = Convert.ToDouble(pipeNominalDiameter) * 1.5;

                if (actualElbowOption.Equals("For 5D Elbow"))
                    elbowRadius = Convert.ToDouble(pipeNominalDiameter) * 5;

                Y = elbowRadius + actualStanchionDiameter / 2;
                outerRadius = elbowRadius + pipeDiameter / 2;

                if (!actualElbowOption.Equals("For Straight Pipe") && pipeDiameter > actualStanchionDiameter)
                    lAdj = Math.Sqrt(outerRadius * outerRadius - Y * Y);

                if (actualElbowOption.Equals("For Straight Pipe") && pipeDiameter > actualStanchionDiameter)
                {
                    lAdj = Math.Sqrt((pipeDiameter / 2) * (pipeDiameter / 2) - (actualStanchionDiameter / 2) * (actualStanchionDiameter / 2));
                    elbowRadius = 0;
                }

                L = length - elbowRadius;
                baseWidth = b;

                //Intializing SymbolGeomHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (HgrCompareDoubleService.cmpdbl(length - d - lAdj , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDLAdj, "( Length - Thickness of Base - lAdj) value cannot be zero"));
                    return;
                }
                if (actualStanchionDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidactualStanchionDiameter, "Stanchion Size (No greater than Pipe) should be greater than zero"));
                    return;
                }
                if (baseWidth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBaseWidth, "Width of Base Plate should be greater than zero"));
                    return;
                }
                if (d <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidD, "Thickness of Base Plate should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d pipe = symbolGeometryHelper.CreateCylinder(null, actualStanchionDiameter / 2, length - d - lAdj);
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, 0, length - d));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                pipe.Transform(matrix);
                m_Symbolic.Outputs["PIPE"] = pipe;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-baseWidth / 2, -baseWidth / 2, length - d);
                Projection3d baseBox = symbolGeometryHelper.CreateBox(null, baseWidth, baseWidth, d, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                baseBox.Transform(matrix);
                m_Symbolic.Outputs["BASE"] = baseBox;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_353"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try //WeightCG implementation is commented ,as in VB "WEIGHT_PER_LENGTH" property is not available in refdata.
            {
                double weight = 0, cogX=0, cogY=0, cogZ=0;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                double materialDensity = 0;
                string materialType = "", materialGrade = "";

                materialType = "Steel - Carbon"; //Only information given about the material is that it should be compatable with pipe. So using the most common type
                materialGrade = "A36"; //Only information given about the material is that it should be compatable with pipe. So using the most common type
                Material material1 = catalogStructHelper.GetMaterial(materialType, materialGrade);
                if(material1 != null)
                    materialDensity = material1.Density;

                VolumeCG baseVolumeCG = null;
                if (supportComponent != null)
                    baseVolumeCG = supportComponent.GetVolumeAndCOG();

                if (baseVolumeCG != null)
                {
                    weight = baseVolumeCG.Volume * materialDensity;
                    cogX = baseVolumeCG.COGX;
                    cogY = baseVolumeCG.COGY;
                    cogZ = baseVolumeCG.COGZ;
                }
                if (supportComponent != null)
                    supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_353."));
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                String pipeNominalDiameter = (String)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_353", "PIPE_NOM_DIA")).PropValue;
                double lengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccLength_mm", "Length")).PropValue;
                string length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER);

                int material = (int)((PropertyValueCodelist)part.GetPropertyValue("IJOAHgrPSL_MATERIAL1", "MATERIAL")).PropValue;
                PropertyValueCodelist stanctionDiameterCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrPSL_353", "STANCHION_DIA");
                if (stanctionDiameterCodelist.PropValue == 0)
                    stanctionDiameterCodelist.PropValue = 1;

                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                string stanctionDiameter = metadataManager.GetCodelistInfo("PSL_353_STANCHION", "UDP").GetCodelistItem(stanctionDiameterCodelist.PropValue).ShortDisplayName;

                PropertyValueCodelist elbowOptionCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrPSL_353", "ELBOW_OPT");
                if (elbowOptionCodelist.PropValue == 0)
                    elbowOptionCodelist.PropValue = 1;

                String elbowOption = metadataManager.GetCodelistInfo("PSL_ELBOW", "UDP").GetCodelistItem(elbowOptionCodelist.PropValue).ShortDisplayName;

                bomDescription = "PSL 353A Tubular Base, Pipe Size: " + pipeNominalDiameter + ", Stanchion Size: " + stanctionDiameter + ", L=" + length + " " + elbowOption + ", Material " + material;

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_353"));
                return "";
            }
        }
        #endregion
    }
}
