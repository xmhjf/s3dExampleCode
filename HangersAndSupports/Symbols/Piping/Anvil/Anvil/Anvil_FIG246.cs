//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG246.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG246
//   Author       : Vijaya
//   Creation Date: 30-April-2013
//   Description: Initial Creation-CR-CP-222292

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30-April-2013  Vijaya  CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

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
    public class Anvil_FIG246 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG246"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(3, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(4, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(5, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(6, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(7, "M", "M", 0.999999)]
        public InputDouble m_dM;
        [InputDouble(8, "K", "K", 0.999999)]
        public InputDouble m_dK;
        [InputDouble(9, "EXACT_OD", "EXACT_OD", 0.999999)]
        public InputDouble m_dEXACT_OD;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("TOP_BOLT", "TOP_BOLT")]
        [SymbolOutput("BODY", "BODY")]
        [SymbolOutput("FILLER", "FILLER")]
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
                
                const double CONST_LENGTH = 0.0381;
                const double CONST_1 = -0.01905;
                Double pipeDiameter = m_dPIPE_DIA.Value, C = m_dC.Value, E = m_dE.Value, F = m_dF.Value, H = m_dH.Value, M = m_dM.Value, K = m_dK.Value;

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, (E - F / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHGTZero, "H value should be greater than zero"));
                    return;
                }
                if (M <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidMGTZero, "M value should be greater than zero"));
                    return;
                }
                if (F <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidFGTZero, "F value should be greater than zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================   

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                symbolGeometryHelper.ActivePosition = new Position(-(C / 2 + H), -M / 2, pipeDiameter / 2 + H / 2);
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d top1 = symbolGeometryHelper.CreateBox(null, H, M, E - pipeDiameter / 2 + M / 2 - H / 2, 9);
                top1.Transform(rotateMatrix);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(C / 2, -M / 2, pipeDiameter / 2 + H / 2);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d top2 = symbolGeometryHelper.CreateBox(null, H, M, E - pipeDiameter / 2 + M / 2 - H / 2, 9);
                top2.Transform(rotateMatrix);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, C / 2 + 2 * H, E).Subtract(new Position(0, -(C / 2 + 2 * H), E));
                symbolGeometryHelper.ActivePosition = new Position(0, -(C / 2 + 2 * H), E);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBolt = symbolGeometryHelper.CreateCylinder(null, F / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = topBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + H, H);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(-Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, -H / 2, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                body.Transform(rotateMatrix);
                m_Symbolic.Outputs["BODY"] = body;

                curveCollection.Add(new Line3d(new Position(CONST_1, -K / 2, 0), new Position(CONST_1, -pipeDiameter / 2, 0)));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d arc = symbolGeometryHelper.CreateArc(null, pipeDiameter / 2, Math.PI);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(Math.PI, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                rotateMatrix.Translate(new Vector(0, CONST_1, 0));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                arc.Transform(rotateMatrix);
                curveCollection.Add(arc);

                curveCollection.Add(new Line3d(new Position(CONST_1, pipeDiameter / 2, 0), new Position(CONST_1, K / 2, 0)));

                Projection3d filter = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), CONST_LENGTH, false);
                m_Symbolic.Outputs["FILLER"] = filter;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG246"));
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
                double exactOD = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG246", "EXACT_OD")).PropValue;
                if (exactOD == 0)
                    exactOD = pipeDiameter;
                string OD = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, exactOD, UnitName.DISTANCE_INCH).Trim();

                bomDescription = catalogPart.PartDescription + ", Exact O.D." + OD;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG246"));
                return ""; 
            }
        }
        #endregion
    }

}
