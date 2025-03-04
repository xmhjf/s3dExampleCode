//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   Ubolt.cs
//   PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.Ubolt
//   Author       :  BS
//   Creation Date:  03-OCT-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   03-OCT-2013     BS      CR-CP-222280--Convert HS_Lisega2010 VB Project to C# .Net
//   05-May-2016     PVK	 TR-CP-293853	Copy/Pasting Cable Tray Deprecated supports results record exceptions
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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
    [CacheOption(CacheOptionType.NonCached)]
    public class Ubolt : HangerComponentSymbolDefinition
    {//----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.Ubolt"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "PipeRadius", "PipeRadius", 1)]
        public InputDouble PipeRadius;
        [InputDouble(3, "BoltDia", "Bolt Diameter", 0.125)]
        public InputDouble BoltDia;
        [InputDouble(4, "Overhang", "Overhang", 2.5)]
        public InputDouble Overhang;
        [InputDouble(5, "Length", "Length", 1)]
        public InputDouble Length;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("SimplePhysical", "SimplePhysical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("SemiCircle", "SemiCircle")]
        [SymbolOutput("LeftLeg", "Left Leg")]
        [SymbolOutput("RightLeg", "Right Leg")]
        [SymbolOutput("PipePort", "Pipe Port")]
        [SymbolOutput("BasePort", "Base Port")]
        public AspectDefinition simplePhysicalAspect;

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
                Part sectionPart = (Part)PartInput.Value;
                double pipeRadius = PipeRadius.Value, boltDiameter = BoltDia.Value, overhang = Overhang.Value, length = Length.Value;
                try
                {
                    RelationCollection hgrRelation = Occurrence.GetRelationship("SupportHasComponents", "Support");
                    BusinessObject businessObject = hgrRelation.TargetObjects[0];
                    SupportedHelper supportedHelper = new SupportedHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                    SupportHelper supportHelper = new SupportHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                    PipeObjectInfo pipeInfo = null;
                    if (supportHelper.SupportedObjects.Count > 0 && supportedHelper.SupportedObjectInfo(1) != null)
                    {
                        pipeInfo = ((PipeObjectInfo)supportedHelper.SupportedObjectInfo(1));
                    }
                    if (pipeInfo != null)
                    {
                        pipeRadius = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, pipeInfo.NominalDiameter.Size / 2.0, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Distance, pipeInfo.NominalDiameter.Units), UnitName.DISTANCE_METER);
                    }
                    else
                    {
                        pipeRadius = PipeRadius.Value; 
                    }
                    Occurrence.SetPropertyValue(pipeRadius, "IJUAHgrRadius", "PipeRadius");
                    pipeRadius = (double)((PropertyValueDouble)Occurrence.GetPropertyValue("IJUAHgrRadius", "PipeRadius")).PropValue;
                }
                catch { pipeRadius = PipeRadius.Value; }
                //-----------------------------------------------------------------------
                // Construction of different outputs in each aspect need to be done here
                //----------------------------------------------------------------------
                if (boltDiameter == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrInvalidBoltDiameterNEZ, "Bolt Diameter should not be equals to zero."));
                    return;
                }

                //Error Check and Adject input values
                //Correct Length if necessary
                if (length < 0)
                    length = 0.1 * pipeRadius;
                //Correct BaseDepth if necessary
                if (boltDiameter < 0)
                    boltDiameter = 0.15 * pipeRadius;
                // Correct Theta if Necessary
                if (overhang < 0)
                    overhang = 0;
                //========================================
                // Ubolt contour And projection
                //========================================
                //Construct the Revolution forming the curved part of the "U" shape
                double boltRadius = boltDiameter / 2;
                Circle3d boltXSection = new Circle3d(new Position(pipeRadius + boltRadius, 0.0, 0.0), new Vector(0.0, 1.0, 0.0), boltRadius);
                //Create the Bolt Cross section curve to revolve.
                Revolution3d semiCircle = new Revolution3d(boltXSection, new Vector(0.0, 0.0, 1.0), new Position(0.0, 0.0, 0.0), Math.PI, true);
                //Create Right Leg
                Projection3d rightLeg = new Projection3d(boltXSection, new Vector(0.0, -1.0, 0.0), length + overhang, true);
                //Create the Bolt Cross section curve to for the left leg.
                Circle3d leftBoltXSection = new Circle3d(new Position(-pipeRadius - boltRadius, 0.0, 0.0), new Vector(0.0, 1.0, 0.0), boltRadius);
                //Create Left Leg
                Projection3d leftLeg = new Projection3d(leftBoltXSection, new Vector(0.0, -1.0, 0.0), length + overhang, true);
                //Add Ubolt sections to the output
                simplePhysicalAspect.Outputs["SemiCircle"] = semiCircle;
                simplePhysicalAspect.Outputs["LeftLeg"] = rightLeg;
                simplePhysicalAspect.Outputs["RightLeg"] = leftLeg;
                //==========================================
                // Hanger Ports
                //==========================================
                Port port1 = new Port(OccurrenceConnection, sectionPart, "Route", new Position(0, 0, 0), new Vector(0, 0, 1.0), new Vector(0, 1.0, 0));
                simplePhysicalAspect.Outputs["PipePort"] = port1;
                Port port2 = new Port(OccurrenceConnection, sectionPart, "Structure", new Position(0.0, -length, 0), new Vector(0, 0, 1.0), new Vector(0, 1.0, 0));
                simplePhysicalAspect.Outputs["BasePort"] = port2;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Ubolt.cs"));
                    return;
                }
            }
        }

        #endregion

    }
}
