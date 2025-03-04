using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using System.Runtime.InteropServices;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Planning
{
    /// <summary>
    /// SecondaryPartsProfiles rule class Implements PlnAssemblyRangeRuleBase
    /// </summary>
    public class SecondaryPartsProfiles : PlnAssemblyChildrenRuleBase
    {
        //Profile Length Tolerance
        private const double m_LengthTol = 0.5;

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
                    if ((primaryChild is PlatePart) && !(primaryChild is CollarPart))
                    {
                        IAssemblyChild assyChild = (IAssemblyChild)primaryChild;
                        IAssembly assembly = assyChild.AssemblyParent;
                        ReadOnlyCollection<BusinessObject> profColl = GetValidProfilesOnPlate(primaryChild, assembly);

                        foreach (BusinessObject profile in profColl)
                        {
                            double profileLength = GetProfileLengthOnThePlate(profile, primaryChild);
                            StiffenerPart stiffenerPart = (StiffenerPart)profile;
                            Curve3d stiffenerPartAxis = stiffenerPart.Axis;
                            double stiffenerLength = stiffenerPartAxis.Length;

                            if ((profileLength / stiffenerLength) > m_LengthTol)
                            {
                                secondaryChildrenCollection.Add((BusinessObject)profile);
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
