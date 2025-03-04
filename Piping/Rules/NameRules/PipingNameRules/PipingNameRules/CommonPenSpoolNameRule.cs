//*******************************************************************

//'  Copyright (C) 2006 Intergraph Corporation.  All rights reserved.
//    '
//    '  Project: PipingNameRules
//    '  Class:   CommonPenSpoolNameRule
//    '
//    '  Abstract: The file contains the Implementation for naming rule interface for Pen Spools

//    '14/05/2015   chandrakala   CR-CP-269177  Route Pipe VB6 Rules need to be replaced  

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Route.Middle;


namespace Ingr.SP3D.Content.Piping.Rules
{
    public class CommonPenSpoolNameRule : NameRuleBase
    {
        /// <summary>
        /// Computes the name for the given entity. 
        /// </summary>
        /// <param name="entity">The business object whose name is being computed.</param>
        /// <param name="parents">The parentsNaming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject entity, ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {

            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }
            if (parents == null)
            {
                throw new ArgumentNullException("The naming parents of the entity to be named are null");
            }
            if (activeEntity == null)
            {
                throw new ArgumentNullException("activeEntity");
            }
            try
            {
                BusinessObject plate = null;
                if (parents.Count > 0) plate = parents[0];
                if (plate != null)
                {
                    // Gets the Name  of the plate then append with PEN.                
                    string name = base.GetName(plate) + "_" + "PEN";
                    // Set name on entity
                    base.SetName(entity, name);
                }
                plate = null;
            }
            catch (Exception)
            {
                throw new Exception("Unexpected error:  PipingNameRules.CommonPenSpoolNameRule.ComputeName");
            }
        }

        /// <summary>
        /// Gets the naming parents from naming rule.
        /// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection. 
        /// The parents added here are used in computing the name of the object in ComputeName() of the same interface. 
        /// Both these methods are called from naming rule semantic. 
        /// </summary>
        /// <param name="entity">BusinessObject for which naming parents are required.</param>
        /// <returns>Collection of BusinessObjects.</returns>
        public override Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }
            try
            {
                Collection<BusinessObject> parentsColl = new Collection<BusinessObject>();
                PenetrationSpool penspool = (PenetrationSpool)entity;
                //Get child collection of "IsChildOf" and  "IsSpoolChildOf"    
                ReadOnlyCollection<IAssemblyChild> assembly = penspool.AssemblyChildren;
                BusinessObject tmpchild = null;
                // Get the penetration plate (naming parent) from the collection
                foreach (BusinessObject child in assembly)
                {
                    tmpchild = child;
                    //if child object is not of spool then exit for loop                   
                    Spool spool = child as Spool;
                    if (spool == null) { break; }
                }
                parentsColl.Add(tmpchild);
                return parentsColl;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
