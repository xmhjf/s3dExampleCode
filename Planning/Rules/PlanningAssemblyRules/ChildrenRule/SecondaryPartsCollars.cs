using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Planning;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Planning
{
    /// <summary>
    ///  SecondaryPartsCollars rule class Implements PlnAssemblyRangeRuleBase
    /// </summary>
    public class SecondaryPartsCollars : PlnAssemblyChildrenRuleBase
    {
        #region Methods

        /// <summary>
        /// Compute seconday children automatically based on input data 
        /// </summary>
        /// <param name="childernData"></param>        
        public override void Evaluate(ChildernData childernData)
        {
            #region Input Error Handling

            if (childernData == null)
            {
                throw new ArgumentNullException("childernData Information is null");
            }

            #endregion //Input Error Handling

            try
            {
                ReadOnlyCollection<BusinessObject> primaryChildren = childernData.PrimaryChildren;
                Collection<BusinessObject> secondaryChildrenCollection = new Collection<BusinessObject>();

                foreach (BusinessObject primaryChild in primaryChildren)
                {
                    if (primaryChild is PlatePart && !(primaryChild is CollarPart))
                    {
                        PlatePart platePart = (PlatePart)primaryChild;
                        PlateSystem plateSys = platePart.RootPlateSystem as PlateSystem;
                        ReadOnlyCollection<BusinessObject> collarParts = GetCollarParts((BusinessObject)plateSys);

                        if ((collarParts != null))
                        {
                            foreach (BusinessObject item in collarParts)
                            {
                                CollarPart collarPart = (CollarPart)item;
                                PlatePart penetratedPart = (PlatePart)collarPart.PenetratedObject;

                                if (object.Equals(penetratedPart, platePart))
                                {
                                    secondaryChildrenCollection.Add(item);
                                }
                            }
                        }
                    }
                }
                childernData.SecondaryChildren = new ReadOnlyCollection<BusinessObject>(secondaryChildrenCollection);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }

        # endregion
    }
}
