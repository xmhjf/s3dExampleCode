//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_PC2.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_PC2
//   Author       :  Vijaya
//   Creation Date:  22-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   22-08-2013     Vijaya    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_PC2 : HangerComponentSymbolDefinition
    {
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_PC2"

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(3, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(4, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble PIPE_DIA;
        [InputDouble(5, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(6, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(7, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(8, "D", "D", 0.999999)]
        public InputDouble D;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("TOP1", "TOP1")]
        [SymbolOutput("TOP2", "TOP2")]
        [SymbolOutput("BOT1", "BOT1")]
        [SymbolOutput("BOT2", "BOT2")]
        [SymbolOutput("TOP_BOLT", "TOP_BOLT")]
        [SymbolOutput("BOT_BOLT", "BOT_BOLT")]
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
                Part part = (Part)PartInput.Value;
                Double pipeDiameter = PIPE_DIA.Value, a = A.Value, b = B.Value, c = C.Value, d = D.Value, e = E.Value, f = F.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, c - a / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs               
                if (f <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFGTZ, "F value should be greater than zero"));
                    return;
                }
                if (e <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidEGTZ, "E value should be greater than zero"));
                    return;
                }
                if ((c + d - pipeDiameter / 2) <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidCDPipeDiaGTZ, "( C + D - PIPE_DIA/2 ) value should be greater than zero"));
                    return;
                }
                if (a <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(-(b / 2 + f), -e / 2, pipeDiameter / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d top1 = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2, 9);
                top1.Transform(matrix);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(b / 2, -e / 2, pipeDiameter / 2);
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d top2 = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2, 9);
                top2.Transform(matrix);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(-(b / 2 + f), -e / 2, -(c + d));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d bottom1 = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2, 9);
                bottom1.Transform(matrix);
                m_Symbolic.Outputs["BOT1"] = bottom1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                matrix.SetIdentity();
                symbolGeometryHelper.ActivePosition = new Position(b / 2, -e / 2, -(c + d));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d bottom2 = symbolGeometryHelper.CreateBox(null, f, e, c + d - pipeDiameter / 2, 9);
                bottom2.Transform(matrix);
                m_Symbolic.Outputs["BOT2"] = bottom2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, b / 2 + 2 * f, -c).Subtract(new Position(0, -(b / 2 + 2 * f), -c));
                symbolGeometryHelper.ActivePosition = new Position(0, -(b / 2 + 2 * f), -c);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBolt = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = topBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                normal = new Position(0, b / 2 + 2 * f, c).Subtract(new Position(0, -(b / 2 + 2 * f), c));
                symbolGeometryHelper.ActivePosition = new Position(0, -(b / 2 + 2 * f), c);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d botomBolt = symbolGeometryHelper.CreateCylinder(null, a / 2, normal.Length);
                m_Symbolic.Outputs["BOT_BOLT"] = botomBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, pipeDiameter / 2 + f, e);
                matrix.SetIdentity();
                matrix.Rotate(3 * Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -e / 2, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_PC2.cs"));
                    return;
                }
            }
        }
        #endregion
    }
}
