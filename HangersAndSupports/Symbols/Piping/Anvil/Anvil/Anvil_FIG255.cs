//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG255.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG255
//   Author       : Vijaya 
//   Creation Date: 1-May-2013 
//   Description: Initial Creation-CR-CP-222292 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   1-May-2013    Vijaya    CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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
    public class Anvil_FIG255 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG255"
        //----------------------------------------------------------------------------------
       

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "INSUL", "INSUL", 1)]
        public InputDouble m_dINSUL;
        [InputDouble(3, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(4, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(5, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("CYL", "CYL")]
        [SymbolOutput("LEG1", "LEG1")]
        [SymbolOutput("FOOT1", "FOOT1")]
        [SymbolOutput("LEG2", "LEG2")]
        [SymbolOutput("FOOT2", "FOOT2")]
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
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();

                Double insulation = m_dINSUL.Value, pipeDiameter = m_dPIPE_DIA.Value, L = m_dL.Value;
                double actualInsulation = 0.0,B = 0, D = 0, E = 0, G = 0, T = 0;                

                //=================================================
                //Construction of Physical Aspect 
                //=================================================                
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    actualInsulation = double.Parse(metadataManager.GetCodelistInfo("Anvil_FIG255_Insulat", "UDP").GetCodelistItem((int)insulation).ShortDisplayName) * 25.4 / 1000;
                else
                    actualInsulation = 0.0254;             
                
                string size = string.Empty;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();

                PartClass anvilFIG2553 = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG255_INSUL"), anvilFIG2554 = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG255_SIZE");
                ReadOnlyCollection<BusinessObject> anvilFIG2553Classes = anvilFIG2553.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects, anvilFIG2554Classes = anvilFIG2554.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in anvilFIG2553Classes)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_INSUL", "INSUL_T")).PropValue) >= actualInsulation - 0.001 && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_INSUL", "INSUL_T")).PropValue) <= actualInsulation + 0.001 && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_INSUL", "PIPE_DIA")).PropValue) >= pipeDiameter - 0.001 && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_INSUL", "PIPE_DIA")).PropValue) <= pipeDiameter + 0.001)
                    {
                        size = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_INSUL", "SIZE")).PropValue;
                        break;
                    }
                }
                foreach (BusinessObject classItem in anvilFIG2554Classes)
                {
                    if (((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_SIZE", "SIZE")).PropValue == size))
                    {
                        B = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_SIZE", "B")).PropValue;
                        D = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_SIZE", "D")).PropValue;
                        E = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_SIZE", "E")).PropValue;
                        G = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_SIZE", "G")).PropValue;
                        T = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_SIZE", "T")).PropValue;
                        break;
                    }
                }

                double calc = Math.Sqrt(((B / 2) * (B / 2)) - ((E / 3) * (E / 3)));
                double legHeight = (B / 2 - calc) + (D - B / 2);               
               
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -D), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (T <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTGTZero, "T value should be greater than zero"));
                    return;
                }
                if (G <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidGGTZero, "G value should be greater than zero"));
                    return;
                }
                if (E <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidEGTZero, "E value should be greater than zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidBGTZero, "B value should be greater than zero"));
                    return;
                }
                if (legHeight <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidlegHeightGTZero, "Leg height should be greater than zero"));
                    return;
                }
                
                Vector normal = new Position(G / 2, 0, 0).Subtract(new Position(-G / 2, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(-G / 2, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d cylinder = symbolGeometryHelper.CreateCylinder(null, B / 2.0, normal.Length);
                m_Symbolic.Outputs["CYL"] = cylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-E / 3 - T, -G / 2, -D);
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d leg1 = symbolGeometryHelper.CreateBox(null, T, G, legHeight, 9);
                leg1.Transform(rotateMatrix);
                m_Symbolic.Outputs["LEG1"] = leg1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-E / 3 - T - E / 3, -G / 2, -D);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d foot1 = symbolGeometryHelper.CreateBox(null, E / 3, G, T, 9);
                foot1.Transform(rotateMatrix);
                m_Symbolic.Outputs["FOOT1"] = foot1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(E / 3, -G / 2, -D);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d leg2 = symbolGeometryHelper.CreateBox(null, T, G, legHeight, 9);
                leg2.Transform(rotateMatrix);
                m_Symbolic.Outputs["LEG2"] = leg2;
                
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(E / 3 + T, -G / 2, -D);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d foot2 = symbolGeometryHelper.CreateBox(null, E / 3, G, T, 9);
                foot2.Transform(rotateMatrix);
                m_Symbolic.Outputs["FOOT2"] = foot2;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG255"));
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
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                double pipeDiameter = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrPipe_Dia", "PIPE_DIA")).PropValue;
                PropertyValueCodelist insulationCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG255", "INSUL");
                string insulation = insulationCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(insulationCodeList.PropValue).ShortDisplayName;
                double actualInsulation = 0.0;
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    actualInsulation = double.Parse(metadataManager.GetCodelistInfo("Anvil_FIG255_Insulat", "UDP").GetCodelistItem((int)insulationCodeList.PropValue).ShortDisplayName) * 25.4 / 1000;
                else
                    actualInsulation = 0.0254;
                string size = string.Empty;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass anvilFIG2553 = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG255_INSUL");
                ReadOnlyCollection<BusinessObject> anvilFIG2553Classes = anvilFIG2553.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in anvilFIG2553Classes)
                {
                    if (((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_INSUL", "INSUL_T")).PropValue) >= actualInsulation - 0.001 && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_INSUL", "INSUL_T")).PropValue) <= actualInsulation + 0.001 && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_INSUL", "PIPE_DIA")).PropValue) >= pipeDiameter - 0.001 && ((double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_INSUL", "PIPE_DIA")).PropValue) <= pipeDiameter + 0.001)
                    {
                        size = (string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAnvil_FIG255_INSUL", "SIZE")).PropValue;
                        break;
                    }
                }

                bomDescription = catalogPart.PartDescription + " Size No. " + size + ", Insulation Thickness: " + insulation + " in";
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG255S"));
                return ""; 
            }
        }
        #endregion

    }

}
