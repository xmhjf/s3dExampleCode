//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG436.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG436
//   Author       :  Hema
//   Creation Date:  03-05-2013
//   Description:    Converted Anvil_Parts VB Project to C#.Net Project

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   03-05-2013     Hema     CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
//   11-Dec-2014     PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
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
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]

    public class Anvil_FIG436 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG436"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "TYP", "TYP", 1)]
        public InputDouble m_oTYP;
        [InputDouble(3, "BASE_TYP", "BASE_TYP", 1)]
        public InputDouble m_oBASE_TYP;
        [InputDouble(4, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(5, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        [InputDouble(6, "BASE_TEMP", "BASE_TEMP", 0.999999)]
        public InputDouble m_dBASE_TEMP;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("WEB", "WEB")]
        [SymbolOutput("T_BASE", "T_BASE")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
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
                Part part = (Part)m_PartInput.Value;

                Double H = 0, W = 0, BL = 0, tHeight = 0.1016, thickness = 0.00635, TL = 0.3048, TW = 0.104775;
                Double pipeDiameter = m_dPIPE_DIA.Value;
                Double baseTemp = m_dBASE_TEMP.Value;
                Double typ = m_oTYP.Value;
                Double baseTyp = m_oBASE_TYP.Value;

                string actualTyp = string.Empty, actualBaseTyp = string.Empty,realFigName = "FIG436_BOLT",auxilaryclass="Anvil_FIG436_BOLT";

                //Initializing SymbolGeometryHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    actualTyp = metadataManager.GetCodelistInfo("Anvil_Shoe_Type", "UDP").GetCodelistItem((int)typ).ShortDisplayName.Trim();
                else
                    actualTyp = "Type 1";
               
                if (metadataManager != null)
                    actualBaseTyp = metadataManager.GetCodelistInfo("Anvil_Shoe_Base", "UDP").GetCodelistItem((int)baseTyp).ShortDisplayName.Trim();
                else
                    actualBaseTyp = "Welded";

                if (actualBaseTyp == "Welded")
                {
                    realFigName = "FIG436_WELD";
                    auxilaryclass = "Anvil_FIG436_WELD";
                }
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass anvilFIG436Class = (PartClass)catalogBaseHelper.GetPartClass(auxilaryclass);
                ReadOnlyCollection<BusinessObject> fig436Classes = anvilFIG436Class.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects; 
                foreach (BusinessObject classItem in fig436Classes)
                {
                    if (((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAnvil_" + realFigName, "TYPE")).PropValue == actualTyp))
                    {
                        H = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_" + realFigName, "H")).PropValue;
                        W = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_" + realFigName, "W")).PropValue;
                        BL = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_" + realFigName, "BL")).PropValue;
                        break;
                    }
                }
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -(H + pipeDiameter / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (typ < 1 && typ > 6)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTypCodelist, "TYP should be between 1 to 6"));
                    return;
                }
                if (baseTyp < 1 && baseTyp > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidbaseTypCodelist, "BASE_TYP should be between 1 to 2"));
                    return;
                }
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidWGTZero, "W value should be greater than zero"));
                    return;
                }
                if (BL <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidBLGTZero, "BL value should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(H, 0) == true && HgrCompareDoubleService.cmpdbl(tHeight, 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHAndTHeight, "H value and T_height cannot be zero"));
                    return;
                }
                if (TW <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTWGTZero, "TW value should be greater than zero"));
                    return;
                }
                if (TL <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTLGTZero, "TL value should be greater than zero"));
                    return;
                }
                if (thickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidThickness, "Thickness should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(-W / 2, -BL / 2, pipeDiameter / 2 + H - (H - tHeight));
                Projection3d Base = symbolGeometryHelper.CreateBox(null, W, BL, H - tHeight, 9);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                Base.Transform(matrix);
                m_Symbolic.Outputs["BASE"] = Base;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-thickness / 2, -TL / 2, pipeDiameter / 2);
                Projection3d web = symbolGeometryHelper.CreateBox(null, thickness, TL, tHeight - thickness, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                web.Transform(matrix);
                m_Symbolic.Outputs["WEB"] = web;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-TW / 2, -TL / 2, pipeDiameter / 2 + tHeight - thickness);
                Projection3d tBase = symbolGeometryHelper.CreateBox(null, TW, TL, thickness, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                tBase.Transform(matrix);
                m_Symbolic.Outputs["T_BASE"] = tBase;

            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG436"));
                    return;
                }
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

                //To get FINISH
                PropertyValueCodelist finishCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrShoeFinish", "FINISH");
                if (finishCodeList.PropValue == 0)
                    finishCodeList.PropValue = 1;

                string finish = finishCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodeList.PropValue).DisplayName;

                //To get TYP
                PropertyValueCodelist typCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG436", "TYP");
                if (typCodeList.PropValue < 1 && typCodeList.PropValue > 6)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTypCodelist, "Typ should be between 1 to 6"));
                    return "";
                }
                string typ = typCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(typCodeList.PropValue).DisplayName;
                
               
                //To get BASE_TYP
                PropertyValueCodelist baseTypeCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG436", "BASE_TYP");
                if (baseTypeCodeList.PropValue < 1 && baseTypeCodeList.PropValue > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidbaseTypCodelist, "baseTyp should be between 1 to 2"));
                    return "";
                }
                string baseType = baseTypeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(baseTypeCodeList.PropValue).DisplayName;

                bomDescription = part.PartDescription + ", Base Type: " + typ + ", Base Connection: " + baseType;

                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG436"));
                return ""; 
            }
        }
        #endregion

    }

}
