
/*'***************************************************************************
'  Copyright (C) 2015, Intergraph Corporation.  All rights reserved.
'
'  Project: NamingRules
'
'  Abstract: The file contains an implementation of the default naming rule
'
'  History:
'  Sravanya            10th March 2015               Creation
'****************************************************************************/
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System;
using Ingr.SP3D.Compartmentation.Middle;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Ingr.SP3D.Content.Compartmentation
{
    public class CompSpaceRule : NameRuleBase
    {
        private const string strFormat = "{0:0000}";
        /// <summary>
        /// Creates a name for the object passed in. The name is based on the parents name and object name. It is assumed that all Naming Parents and the Object implement JNamedItem.
        /// The Naming Parents are added in AddNamingParents() of the same interface. Both these methods are called from the naming rule semantic. 
        /// </summary>
        /// <param name="entity">The business object whose name is being computed. </param>
        /// <param name="parents">Naming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject entity, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            try
            {
                if (entity == null)
                {
                    throw new ArgumentNullException("entity");
                }
                if (activeEntity == null)
                {
                    throw new ArgumentNullException("activeEntity");
                }

                string strName = string.Empty, strLocation = string.Empty;
                long count;
                BusinessObject parent;

                ISpaceChild spaceChild = (ISpaceChild)entity;
                strName = Convert.ToString(entity.GetPropertyValue("IJNamedItem", "Name"));

                if (!string.IsNullOrEmpty(strName))
                {
                    return;
                }

                //'Get the parent's name
                parent = spaceChild.SpaceParent as BusinessObject;                

                if(parent != null)
                    strName = Convert.ToString(parent.GetPropertyValue("IJNamedItem", "Name"));

                //Returns the number of occurrence of a string in addtion to the LocationID
                GetCountAndLocationID(strName, out count, out strLocation);

                //'Add LocationID, if available
                if (strLocation != string.Empty)
                {
                    strName = strName + "-" + strLocation + "-" + string.Format(strFormat, count);
                }
                else
                {
                    strName = strName + "-" + string.Format(strFormat, count);
                }

                entity.SetPropertyValue(strName, "IJNamedItem", "Name");
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("FrameRule.ComputeName: Error encountered (" + e.Message + ")");
            }
        }

        /// <summary>
        ///   All the Naming Parents that need to participate in an objects naming are added here to the collection. Dummy function which does nothing
        /// </summary>
        /// <param name="entity">The business object whose name is being computed. </param>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            Collection<BusinessObject> oRetColl = new Collection<BusinessObject>();
            try
            {
                if (entity == null)
                {
                    throw new ArgumentNullException("entity");
                }

                ISpaceChild spaceChild = (ISpaceChild)entity;
                oRetColl.Add((BusinessObject)spaceChild.SpaceParent);
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("FrameRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }

            return oRetColl;
        }
    }
}
