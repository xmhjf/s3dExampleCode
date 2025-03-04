//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   ThreadTop.cs
//   PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.ThreadTop
//   Author       :  BS
//   Creation Date:  03-OCT-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who        change description
//   -----------     ---        ------------------
//   03-OCT-2013     BS         CR-CP-222280--Convert HS_Lisega2010 VB Project to C# .Net
//   03-OCT-2013     Vinay      TR-CP-284319 	Multiple Record exception dumps are created on copy pasting supports 
//   12-APR-2016     Adi Soumitra Mondal TR-CP-289118 48 minidump(s) at 'DistributionConnectivity!CRteDistribConnection::GetDefaultGas & TR-CP-289119 30 minidump(s) at 'DistributionConnectivity!CRteDistribConnection::GetConnectedP
//   05-May-2016     PVK	     TR-CP-293853	Copy/Pasting Cable Tray Deprecated supports results record exceptions
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using System.Collections.ObjectModel;

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
    public class ThreadTop : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PipeHgrAssemblySymbols,Ingr.SP3D.Content.Support.Symbols.ThreadTop"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "Height", "Height", 0.5)]
        public InputDouble Height;
        [InputDouble(3, "Width", "Width", 0.5)]
        public InputDouble Width;
        [InputDouble(4, "Depth", "Depth", 0.5)]
        public InputDouble Depth;
        [InputDouble(5, "PinRadius", "PinRadius", 0.05)]
        public InputDouble PinRadius;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("XSProjection", "3d project contour")]
        [SymbolOutput("Pin", "3d pin")]
        [SymbolOutput("EndPin1", "end pin 1")]
        [SymbolOutput("EndPin2", "end pin 2")]
        [SymbolOutput("EyeNut", "eye nut")]
        [SymbolOutput("EyeNutBolt", "eye nut bolt")]
        public AspectDefinition symbolicAspect;

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
                double height = Height.Value, width = Width.Value, depth = Depth.Value, pinRadius = PinRadius.Value;
                try
                {
                    RelationCollection hgrRelation = Occurrence.GetRelationship("SupportHasComponents", "Support");
                    BusinessObject businessObject = hgrRelation.TargetObjects[0];
                    SupportedHelper supportedHelper = new SupportedHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                    SupportingHelper supportingHelper = new SupportingHelper((Ingr.SP3D.Support.Middle.Support)businessObject);
                    SupportHelper supportHelper = new SupportHelper((Ingr.SP3D.Support.Middle.Support)businessObject);

                    PipeObjectInfo pipeInfo = null;
                    CableTrayObjectInfo cableInfo = null;
                    if (supportHelper.SupportedObjects.Count > 0 && supportedHelper.SupportedObjectInfo(1)!=null)
                    {
                        if (supportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.Pipe)
                        {
                            pipeInfo = (PipeObjectInfo)supportedHelper.SupportedObjectInfo(1);
                        }
                    
                        if (supportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.CableTray)
                        {
                            cableInfo = (CableTrayObjectInfo)supportedHelper.SupportedObjectInfo(1);
                        }
                    }

                    double structBase=0;

                    if (supportHelper.SupportingObjects.Count > 0
                        && supportingHelper.SupportingObjectInfo(1)!= null)
                    {
                        structBase = supportingHelper.SupportingObjectInfo(1).Width;
                    }
                    else
                    {
                        structBase = 0;
                    }

                    if (structBase == 0)
                    {
                        if (pipeInfo !=null)
                             width = 0.75 * pipeInfo.NominalDiameter.Size;
                        else if (cableInfo!=null)
                            width = 0.75 * cableInfo.Width;
                    }
                    else
                        width = 0.75 * structBase;

                    Occurrence.SetPropertyValue(width, "IJUAHgrOccGeometry", "Width");
                    width = (double)((PropertyValueDouble)Occurrence.GetPropertyValue("IJUAHgrOccGeometry", "Width")).PropValue;
                }
                catch { width = Width.Value; }

                if (width <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrInvalidWidthGTZ, "Width should be greater than zero."));
                    return;
                }

                double thickness = width / 10, pinOffset = width / 5;
                //Set all input based on width
                height = width;
                depth = width;
                pinRadius = width / 10;

                //-----------------------------------------------------------------------
                // Construction of different outputs in each aspect need to be done here
                //----------------------------------------------------------------------
                //construct the contour
                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, -width / 2, height));
                pointCollection.Add(new Position(0, -width / 2, 0));
                pointCollection.Add(new Position(0, width / 2, 0));
                pointCollection.Add(new Position(0, width / 2, height));
                pointCollection.Add(new Position(0, ((width / 2) - thickness), height));
                pointCollection.Add(new Position(0, ((width / 2) - thickness), thickness));
                pointCollection.Add(new Position(0, -((width / 2) - thickness), thickness));
                pointCollection.Add(new Position(0, -((width / 2) - thickness), height));
                pointCollection.Add(new Position(0, -(width / 2), height));
                Projection3d projection = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), depth, true);

                Circle3d circle = new Circle3d(new Position(depth / 2, (width / 2 + pinOffset), (2 * height / 3)), new Vector(0, 1, 0), pinRadius);

                double pinLength = width + 2 * pinOffset;
                Projection3d pin = new Projection3d(circle, new Vector(0, -1, 0), pinLength, true);

                double circleZVal = (2 * height / 3) - (2 * pinRadius);
                double endPinRadius = pinRadius / 5;
                circle = new Circle3d(new Position(depth / 2, -(width / 2 + pinOffset / 2), circleZVal), new Vector(0, 0, 1), endPinRadius);
                circle = new Circle3d(new Position(depth / 2, (width / 2 + pinOffset / 2), circleZVal), new Vector(0, 0, 1), endPinRadius);

                Projection3d oEndPin1 = new Projection3d(circle, new Vector(0, 0, 1), 4 * pinRadius, true);
                Projection3d oEndPin2 = new Projection3d(circle, new Vector(0, 0, 1), 4 * pinRadius, true);

                double minAxis = pinRadius / 2, majAxis = height / 3;
                double axisCenterZ = 2 * height / 3 + (majAxis - minAxis - pinRadius);

                Torus3d eyeNut = new Torus3d(new Position(depth / 2, 0, axisCenterZ), new Vector(0, 1, 0), new Vector(0, 0, 1), majAxis, minAxis, false);
                double bContX = depth / 2 - 2 * minAxis, bContY = 2 * minAxis, bContZ = axisCenterZ + majAxis;

                pointCollection.Clear();
                pointCollection.Add(new Position(bContX, -bContY, bContZ));
                pointCollection.Add(new Position(bContX, bContY, bContZ));
                pointCollection.Add(new Position(bContX, bContY, bContZ + 4 * minAxis));
                pointCollection.Add(new Position(bContX, -bContY, bContZ + 4 * minAxis));
                pointCollection.Add(new Position(bContX, -bContY, bContZ));

                Projection3d eyeNutBolt = new Projection3d(new LineString3d(pointCollection), new Vector(1, 0, 0), 4 * minAxis, true);
                //==========================================
                // create hgrports as part of the output
                //==========================================
                Port port1 = new Port(OccurrenceConnection, sectionPart, "Structure", new Position(depth / 2, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                symbolicAspect.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, sectionPart, "Rod", new Position(width / 2, 0, bContZ + 4 * minAxis), new Vector(1, 0, 0), new Vector(0, 0, 1));
                symbolicAspect.Outputs["Port2"] = port2;

                symbolicAspect.Outputs["XSProjection"] = projection;
                symbolicAspect.Outputs["Pin"] = pin;
                symbolicAspect.Outputs["EndPin1"] = oEndPin1;
                symbolicAspect.Outputs["EndPin2"] = oEndPin2;
                symbolicAspect.Outputs["EyeNut"] = eyeNut;
                symbolicAspect.Outputs["EyeNutBolt"] = eyeNutBolt;

            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PipeHgrAssemblySymbolsLocalizer.GetString(PipeHgrAssemblySymbolsResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of ThreadTop.cs"));
                    return;
                }
            }
        }

        #endregion

    }
}
