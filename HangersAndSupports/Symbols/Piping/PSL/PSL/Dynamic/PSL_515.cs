//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_515.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_515
//   Author       :  Rajeswari
//   Creation Date:  22-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   22-08-2013   Rajeswari  CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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
    public class PSL_515 : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_515"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "PIN_L", "PIN_L", 0.999999)]
        public InputDouble PIN_L;
        [InputDouble(3, "BOLT_L", "BOLT_L", 0.999999)]
        public InputDouble BOLT_L;
        [InputDouble(4, "A", "A", 0.999999)]
        public InputDouble A;
        [InputDouble(5, "C", "C", 0.999999)]
        public InputDouble C;
        [InputDouble(6, "N", "N", 0.999999)]
        public InputDouble N;
        [InputDouble(7, "B", "B", 0.999999)]
        public InputDouble B;
        [InputDouble(8, "X", "X", 0.999999)]
        public InputDouble X;
        [InputDouble(9, "D", "D", 0.999999)]
        public InputDouble D;
        [InputDouble(10, "B1", "B1", 0.999999)]
        public InputDouble B1;
        [InputDouble(11, "F", "F", 0.999999)]
        public InputDouble F;
        [InputDouble(12, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(13, "T", "T", 0.999999)]
        public InputDouble T;
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
        [SymbolOutput("TOP_BOLT1", "TOP_BOLT1")]
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
                Double a = A.Value;
                Double b = B.Value;
                Double c = C.Value;
                Double d = D.Value;
                Double e = E.Value;
                Double f = F.Value;
                Double b1 = B1.Value;
                Double t = T.Value;
                Double x = X.Value;
                Double n = N.Value;
                Double pinLength = PIN_L.Value;
                Double boltLength = BOLT_L.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, c - a / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (t <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidTGTZ, "T value should be greater than zero"));
                    return;
                }
                if (d <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidDGTZ, "D value should be greater than zero"));
                    return;
                }
                if (a <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidAGTZ, "A value should be greater than zero"));
                    return;
                }
                if (n <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidNGTZ, "N value should be greater than zero"));
                    return;
                }
                if (x <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidXGTZ, "X value should be greater than zero"));
                    return;
                }

                Projection3d top1 = symbolGeometryHelper.CreateBox(null, t, d, c + f - x / 2.0, 9);
                matrix.Translate(new Vector(-(b / 2.0 + t), -d / 2.0, x / 2.0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                top1.Transform(matrix);
                m_Symbolic.Outputs["TOP1"] = top1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Projection3d top2 = symbolGeometryHelper.CreateBox(null, t, d, c + f - x / 2.0, 9);
                matrix.SetIdentity();
                matrix.Translate(new Vector(b / 2.0, -d / 2.0, x / 2.0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                top2.Transform(matrix);
                m_Symbolic.Outputs["TOP2"] = top2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Projection3d bot1 = symbolGeometryHelper.CreateBox(null, t, d, (b1 + e) - x / 2.0, 9);
                matrix.SetIdentity();
                matrix.Translate(new Vector(-(b / 2.0 + t), -d / 2.0, -b1 - e));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                bot1.Transform(matrix);
                m_Symbolic.Outputs["BOT1"] = bot1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Projection3d bot2 = symbolGeometryHelper.CreateBox(null, t, d, (b1 + e) - x / 2.0, 9);
                matrix.SetIdentity();
                matrix.Translate(new Vector(b / 2.0, -d / 2.0, -b1 - e));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                bot2.Transform(matrix);
                m_Symbolic.Outputs["BOT2"] = bot2;

                Vector normal = new Position(0, pinLength / 2.0, c).Subtract(new Position(0, -pinLength / 2.0, c));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -pinLength / 2.0, c);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBolt = symbolGeometryHelper.CreateCylinder(null, a / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT"] = topBolt;

                normal = new Position(0, boltLength / 2.0, b1).Subtract(new Position(0, -boltLength / 2.0, b1));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -boltLength / 2.0, b1);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d topBolt1 = symbolGeometryHelper.CreateCylinder(null, n / 2.0, normal.Length);
                m_Symbolic.Outputs["TOP_BOLT1"] = topBolt1;

                normal = new Position(0, boltLength / 2.0, -b1).Subtract(new Position(0, -boltLength / 2.0, -b1));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -boltLength / 2.0, -b1);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d botBolt = symbolGeometryHelper.CreateCylinder(null, n / 2.0, normal.Length);
                m_Symbolic.Outputs["BOT_BOLT"] = botBolt;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d body = symbolGeometryHelper.CreateCylinder(null, x / 2.0 + t, d);
                matrix.SetIdentity();
                matrix.Rotate(3 * Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -d / 2.0, 0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                body.Transform(matrix);
                m_Symbolic.Outputs["BODY"] = body;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_515.cs"));
                    return;
                }
            }
        }

        #endregion

    }
}
