//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG167.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG167
//   Author       :  Rajeswari
//   Creation Date:  09-May-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   09-May-2013  Rajeswari CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
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
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class Anvil_FIG167 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG167"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "INSULAT", "INSULAT", 1)]
        public InputDouble m_oINSULAT;
        [InputString(3, "COPPER", "COPPER", "No Value")]
        public InputString m_oCOPPER;
        [InputDouble(4, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY", "BODY")]
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

                Double pipeDiameter = m_dPIPE_DIA.Value;
                String copper = m_oCOPPER.Value;
                Double insult = m_oINSULAT.Value;

                Double actualInsult, radius = 0, L = 0;
                String heading, size = null;

                heading = "E";

                //Validating Inputs
                if (insult < 1 || insult > 5)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidInsultCodelist, "INSULAT codelist value should be between 1 & 5."));
                    return;
                }
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    actualInsult = Convert.ToDouble(metadataManager.GetCodelistInfo("Anvil_FIG167_Insulat", "UDP").GetCodelistItem((int)insult).ShortDisplayName) * 25.4 / 1000;
                else
                    actualInsult = 0.0127;

                if (copper == "Yes")
                {
                    if (actualInsult <= 0.0127)
                        heading = "A2";
                    else
                    {
                        if (actualInsult <= 0.01905)
                            heading = "B2";
                        else
                        {
                            if (actualInsult <= 0.0254)
                                heading = "C2";
                            else
                            {
                                if (actualInsult <= 0.0508)
                                    heading = "D2";
                            }
                        }
                    }
                }
                else
                {
                    if (actualInsult <= 0.0127)
                        heading = "A";
                    else
                    {
                        if (actualInsult <= 0.01905)
                            heading = "B";
                        else
                        {
                            if (actualInsult <= 0.0254)
                                heading = "C";
                            else
                            {
                                if (actualInsult <= 0.0508)
                                    heading = "D";
                            }
                        }
                    }
                }

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                ReadOnlyCollection<BusinessObject> fig167SizesParts, fig167ParamsParts;
                PartClass anvilfig167SizesPartClass = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG167_SIZES");
                fig167SizesParts = anvilfig167SizesPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject part1 in fig167SizesParts)
                {
                    if (((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG167_SIZES", "PIPE_DIA")).PropValue > pipeDiameter - 0.001) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG167_SIZES", "PIPE_DIA")).PropValue < (pipeDiameter + 0.001)))
                    {
                        size = (string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_FIG167_SIZES", heading)).PropValue;
                        break;
                    }
                }

                PartClass anvilfig167ParamsPartClass = (PartClass)catalogBaseHelper.GetPartClass("FIG167_PARAMS");
                fig167ParamsParts = anvilfig167ParamsPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject part1 in fig167ParamsParts)
                {
                    if (((string)((PropertyValueString)part1.GetPropertyValue("IJUAHgrAnvil_FIG167_PARAMS", "SIZE")).PropValue == size))
                    {
                        L = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG167_PARAMS", "L")).PropValue;
                        radius = (double)((PropertyValueDouble)part1.GetPropertyValue("IJUAHgrAnvil_FIG167_PARAMS", "RADIUS")).PropValue;
                        break;
                    }
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -radius), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (radius == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidRadiusNZero, "Radius cannot be zero"));
                    return;
                }
                if (L == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidLNZero, "L value should not be lessthan zero"));
                    return;
                }

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d outerArc = symbolGeometryHelper.CreateArc(null, radius, Math.PI);
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -L / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                matrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                outerArc.Transform(matrix);
                curveCollection.Add(outerArc);

                Projection3d body = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, -0), L, false);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch//General Unhandled exception 
            {

                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG167"));
                return;
            }
        }
        #endregion

    }

}
