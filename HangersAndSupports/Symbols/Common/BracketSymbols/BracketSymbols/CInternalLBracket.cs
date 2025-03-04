
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//   File:
//  CInternalLBracket.cs
//   BracketSymbols,Ingr.SP3D.Content.Support.Symbols.CInternalLBracket
//   Author       :  Mahanth
//   Creation Date:  25.07.2013
//  Description:
//      This class module is the implementation of the CInternalLBracket C# Symbol. 
//      This class inherits from CustomSymbolDefinition.

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//25.07.2013     Mahanth     CR-CP-222299 Convert HgrBracketSymbols VB Project to C# .Net 
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
    public class CInternalLBracket : CustomSymbolDefinition
    {
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "AngleHeight", "Angle Height", 0.08)]
        public InputDouble AngleHeight;
        [InputDouble(3, "AngleWidth", "Angle Width", 0.08)]
        public InputDouble AngleWidth;
        [InputDouble(4, "AngleThickness", "Angle Thickness", 0.005)]
        public InputDouble AngleThickness;
        [InputDouble(5, "GapEAHeight", "Gap from edge of bottom  Angle", 0.01)]
        public InputDouble GapEAHeight;
        [InputDouble(6, "GapEAWidth", "Gap from edge top angle", 0.01)]
        public InputDouble GapEAWidth;
        [InputDouble(7, "BottomHeight", "Hight from edge of bottom bracket", 0.01)]
        public InputDouble BottomHeight;
        [InputDouble(8, "TopWidth", "Width from edge of top bracket", 0.01)]
        public InputDouble TopWidth;
        [InputDouble(9, "CornerRadius", "Origin corner Radius", 0.01)]
        public InputDouble CornerRadius;
        [InputDouble(10, "BracketThickness", "Thickness of Bracket", 0.01)]
        public InputDouble BracketThickness;
        [InputDouble(11, "GapBTWBrackets", "Gap between Brackets", 0.1)]
        public InputDouble GapBTWBrackets;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("SimplePhysical", "SimplePhysical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Bracket_1", "Bracket_1")]
        [SymbolOutput("Bracket_2", "Bracket_2")]
        [SymbolOutput("Bracket_3", "Bracket_3")]
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
                Double angleThickness = AngleThickness.Value;
                Double gapEAHeight = GapEAHeight.Value;
                Double gapEAWidth = GapEAWidth.Value;
                Double bottomHeight = BottomHeight.Value;
                Double topWidth = TopWidth.Value;
                Double cornerRadius = CornerRadius.Value;
                Double thickness = BracketThickness.Value;
                Double gapBTWBrackets = GapBTWBrackets.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                Collection<ICurve> collectionOfGeometry = new Collection<ICurve>();
                //-----------------------------------------------------------------------
                // Construction of Physical Aspect 
                //----------------------------------------------------------------------       

                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, BracketSymbolsLocalizer.GetString(BracketSymbolsResourceIDs.ErrInvalidTEZero, "Thickess cannot be zero "));
                    return;

                }

                // Creation of Bracket_1
                collectionOfGeometry.Add(new Line3d(new Position(cornerRadius, thickness / 2, 0), new Position(angleWidth - gapEAWidth, thickness / 2, 0)));
                collectionOfGeometry.Add(new Line3d(new Position(angleWidth - gapEAWidth, thickness / 2, 0), new Position(angleWidth - gapEAWidth, thickness / 2, bottomHeight)));
                collectionOfGeometry.Add(new Line3d(new Position(angleWidth - gapEAWidth, thickness / 2, bottomHeight), new Position(topWidth, thickness / 2, angleHeight - gapEAHeight)));
                collectionOfGeometry.Add(new Line3d(new Position(topWidth, thickness / 2, angleHeight - gapEAHeight), new Position(0, thickness / 2, angleHeight - gapEAHeight)));
                collectionOfGeometry.Add(new Line3d(new Position(0, thickness / 2, angleHeight - gapEAHeight), new Position(0, thickness / 2, cornerRadius)));
                collectionOfGeometry.Add(new Arc3d(new Position(0, thickness / 2, 0), null, new Position(0, thickness / 2, cornerRadius), new Position(cornerRadius, thickness / 2, 0)));
                Projection3d bracket = new Projection3d(OccurrenceConnection, new ComplexString3d(collectionOfGeometry), new Vector(0, -1, 0), thickness, true);
                SimplePhysicalAspect.Outputs["Bracket_1"] = bracket;

                // Creation of Bracket_2
                collectionOfGeometry.Clear();
                collectionOfGeometry.Add(new Line3d(new Position(cornerRadius, thickness / 2 + gapBTWBrackets, 0), new Position(angleWidth - gapEAWidth, thickness / 2 + gapBTWBrackets, 0)));
                collectionOfGeometry.Add(new Line3d(new Position(angleWidth - gapEAWidth, thickness / 2 + gapBTWBrackets, 0), new Position(angleWidth - gapEAWidth, thickness / 2 + gapBTWBrackets, bottomHeight)));
                collectionOfGeometry.Add(new Line3d(new Position(angleWidth - gapEAWidth, thickness / 2 + gapBTWBrackets, bottomHeight), new Position(topWidth, thickness / 2 + gapBTWBrackets, angleHeight - gapEAHeight)));
                collectionOfGeometry.Add(new Line3d(new Position(topWidth, thickness / 2 + gapBTWBrackets, angleHeight - gapEAHeight), new Position(0, thickness / 2 + gapBTWBrackets, angleHeight - gapEAHeight)));
                collectionOfGeometry.Add(new Line3d(new Position(0, thickness / 2 + gapBTWBrackets, angleHeight - gapEAHeight), new Position(0, thickness / 2 + gapBTWBrackets, cornerRadius)));
                collectionOfGeometry.Add(new Arc3d(new Position(0, thickness / 2 + gapBTWBrackets, 0), null, new Position(0, thickness / 2 + gapBTWBrackets, cornerRadius), new Position(cornerRadius, thickness / 2 + gapBTWBrackets, 0)));
                bracket = new Projection3d(OccurrenceConnection, new ComplexString3d(collectionOfGeometry), new Vector(0, -1, 0), thickness, true);
                SimplePhysicalAspect.Outputs["Bracket_2"] = bracket;

                // Creation of Bracket_3
                collectionOfGeometry.Clear();
                collectionOfGeometry.Add(new Line3d(new Position(cornerRadius, thickness / 2 - gapBTWBrackets, 0), new Position(angleWidth - gapEAWidth, thickness / 2 - gapBTWBrackets, 0)));
                collectionOfGeometry.Add(new Line3d(new Position(angleWidth - gapEAWidth, thickness / 2 - gapBTWBrackets, 0), new Position(angleWidth - gapEAWidth, thickness / 2 - gapBTWBrackets, bottomHeight)));
                collectionOfGeometry.Add(new Line3d(new Position(angleWidth - gapEAWidth, thickness / 2 - gapBTWBrackets, bottomHeight), new Position(topWidth, thickness / 2 - gapBTWBrackets, angleHeight - gapEAHeight)));
                collectionOfGeometry.Add(new Line3d(new Position(topWidth, thickness / 2 - gapBTWBrackets, angleHeight - gapEAHeight), new Position(0, thickness / 2 - gapBTWBrackets, angleHeight - gapEAHeight)));
                collectionOfGeometry.Add(new Line3d(new Position(0, thickness / 2 - gapBTWBrackets, angleHeight - gapEAHeight), new Position(0, thickness / 2 - gapBTWBrackets, cornerRadius)));
                collectionOfGeometry.Add(new Arc3d(new Position(0, thickness / 2 - gapBTWBrackets, 0), null, new Position(0, thickness / 2 - gapBTWBrackets, cornerRadius), new Position(cornerRadius, thickness / 2 - gapBTWBrackets, 0)));
                bracket = new Projection3d(OccurrenceConnection, new ComplexString3d(collectionOfGeometry), new Vector(0, -thickness, 0), thickness, true);
                SimplePhysicalAspect.Outputs["Bracket_3"] = bracket;

                //CREATE HANGER PORTS
                Port port = new Port(OccurrenceConnection, part, "HgrPort_1", new Position(0, 0, 0), new Vector(1.0, 0, 0.0), new Vector(0, -1.0, 0));
                SimplePhysicalAspect.Outputs["BracketPort"] = port;

            }
            catch (Exception)
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, BracketSymbolsLocalizer.GetString(BracketSymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of CInternalLBracket.cs."));
                    return;
                }
            }
        }

        #endregion

    }
}
