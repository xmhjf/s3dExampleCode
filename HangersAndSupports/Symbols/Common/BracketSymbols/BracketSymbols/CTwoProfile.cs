
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//   File:
//  CTwoProfile.cs
//   BracketSymbols,Ingr.SP3D.Content.Support.Symbols.CTwoProfile
//   Author       :  Mahanth
//   Creation Date:  25.07.2013
//  Description:
//      This class module is the implementation of the CTwoProfile C# Symbol. 
//      This class inherits from CustomSymbolDefinition.

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//25.07.2013     Mahanth   CR-CP-222299 Convert HgrBracketSymbols VB Project to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    [CacheOption(CacheOptionType.Cached)]
    public class CTwoProfile : CustomSymbolDefinition
    {
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "AngleHeight", "Angle Height", 0.08)]
        public InputDouble AngleHeight;
        [InputDouble(3, "AngleWidth", "Angle Width", 0.08)]
        public InputDouble AngleWidth;
        [InputDouble(4, "BottomHeight", "Hight from edge of bottom bracket", 0.01)]
        public InputDouble BottomHeight;
        [InputDouble(5, "TopWidth", "Width from edge of top bracket", 0.01)]
        public InputDouble TopWidth;
        [InputDouble(6, "CornerRadius", "Origin corner Radius", 0.01)]
        public InputDouble CornerRadius;
        [InputDouble(7, "BracketThickness", "Thickness of Bracket", 0.01)]
        public InputDouble BracketThickness;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("SimplePhysical", "SimplePhysical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Bracket", "Bracket")]
        [SymbolOutput("BracketPort", "BracketPort")]
        public AspectDefinition SimplePhysicalAspect;

        #endregion

        #region "Construct Outputs"

        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)PartInput.Value;
                Double angleHeight = AngleHeight.Value;
                Double angleWidth = AngleWidth.Value;
                Double bottomHeight = BottomHeight.Value;
                Double topWidth = TopWidth.Value;
                Double cornerRadius = CornerRadius.Value;
                Double thickness = BracketThickness.Value;
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Collection<ICurve> colOfGeom = new Collection<ICurve>();
                //-----------------------------------------------------------------------
                // Construction of Physical Aspect 
                //----------------------------------------------------------------------     
                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, BracketSymbolsLocalizer.GetString(BracketSymbolsResourceIDs.ErrInvalidTEZero, "Thickess cannot be zero"));
                    return;

                }
                colOfGeom.Add(new Line3d(new Position(cornerRadius, thickness / 2, 0), new Position(angleWidth, thickness / 2, 0)));
                colOfGeom.Add(new Line3d(new Position(angleWidth, thickness / 2, 0), new Position(angleWidth, thickness / 2, bottomHeight)));
                colOfGeom.Add(new Line3d(new Position(angleWidth, thickness / 2, bottomHeight), new Position(topWidth, thickness / 2, angleHeight)));
                colOfGeom.Add(new Line3d(new Position(topWidth, thickness / 2, angleHeight), new Position(0, thickness / 2, angleHeight)));
                colOfGeom.Add(new Line3d(new Position(0, thickness / 2, angleHeight), new Position(0, thickness / 2, cornerRadius)));
                colOfGeom.Add(new Arc3d(new Position(0, thickness / 2, 0), null, new Position(0, thickness / 2, cornerRadius), new Position(cornerRadius, thickness / 2, 0)));

                //CREATE HANGER PORTS
                Port port = new Port(OccurrenceConnection, part, "HgrPort_1", new Position(0, -thickness / 2.0, 0), new Vector(1.0, 0, 0.0), new Vector(0, -1.0, 0.0));
                SimplePhysicalAspect.Outputs["BracketPort"] = port;
                Projection3d bracket = new Projection3d(OccurrenceConnection, new ComplexString3d(colOfGeom), new Vector(0, -1, 0), thickness, true);
                SimplePhysicalAspect.Outputs["Bracket"] = bracket;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, BracketSymbolsLocalizer.GetString(BracketSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CTwoProile.cs."));
                    return;
                }
            }
        }

        #endregion

    }
}
