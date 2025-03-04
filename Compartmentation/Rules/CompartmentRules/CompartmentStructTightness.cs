using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Compartmentation.Middle;
using Ingr.SP3D.Structure.Middle;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Compartmentation
{
    public class CompartmentStructTightness : CompartmentComputeRuleBase
    {
        public override void Evaluate(Compartment compartment)
        {
            #region Input Error Handling
            if (compartment == null)
            {
                throw new CmnArgumentNullException("input argument is null");
            }
            #endregion
           
            try
            {
                Tightness tightness;              
                tightness = GetMinimumTightness(compartment);

                if ((tightness >= Tightness.UnSpecified) && (tightness <= Tightness.AirTight))
                {
                    compartment.SetPropertyValue((int)tightness, "IJCompartTightness", "StructTightness");
                }
             
                if (tightness == Tightness.UnSpecified)
                {
                    PropertyValue propertyValue = compartment.GetPropertyValue("IJCompartTightness", "CompartTightness");
                    List<CodelistItem> codeListMembers = propertyValue.PropertyInfo.CodeListInfo.CodelistMembers;

                    foreach (var codeLstItem in codeListMembers)
                    {
                        if (codeLstItem.Name == "Undefined")
                        {
                            compartment.SetPropertyValue(codeLstItem, "IJCompartTightness", "CompartTightness");
                            break;
                        }
                    }
                }
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        /// <summary>
        ///  returns the tightness of the given compartment
        /// </summary>
        /// <param name="compartEntity"></param>
        /// <returns></returns>
        private Tightness GetMinimumTightness(Compartment compartment)
        {           
            Tightness tightness = Tightness.UnSpecified;
            try
            {
                if (compartment.ConstructionType == SP3D.Compartmentation.Middle.ConstructionType.ByFaces)
                {
                    Collection<BoundaryDefinition> boundaries = null;
                    compartment.GetInputs(out boundaries);
                    bool notProperlyBounded = false;

                    if (boundaries != null)
                    {
                        int count = 1;

                        foreach (BoundaryDefinition boundaryDefinition in boundaries)
                        {
                            if (boundaryDefinition.BoundaryObject is PlateSystemBase)
                            {
                                //Get the Tightness
                                PlateSystemBase plateSystemBase = (PlateSystemBase)boundaryDefinition.BoundaryObject;

                                if (plateSystemBase != null)
                                {
                                    if (count == 1)
                                    {
                                        tightness = plateSystemBase.Tightness;

                                    }
                                    else if (tightness > plateSystemBase.Tightness)
                                    {
                                        tightness = plateSystemBase.Tightness;
                                    }
                                    count += 1;
                                }
                            }
                            else
                            {
                                notProperlyBounded = true;
                                break;
                            }
                        }

                        if (notProperlyBounded == true)
                        {
                            tightness = Tightness.UnSpecified;
                        }
                    }
                }
            }
            catch (CmnException e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
            return tightness;
        }
    }
}
