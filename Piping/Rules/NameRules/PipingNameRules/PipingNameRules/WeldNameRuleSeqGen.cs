/*******************************************************************
  Copyright (C) 2004 Intergraph Corporation.  All rights reserved.

Project:   PipingNameRules
Class:     WeldNameRuleSeqGen

  Abstract: The file contains the Implementation for naming rule  for Welds

14/05/2015   Swetha   CR-CP-269177  Route Pipe VB6 Rules need to be replaced  

******************************************************************/
using Ingr.SP3D.Common.Middle;
using System;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Piping.Rules
{
    /// <summary>
    /// class implementing the naming rule  for Welds
    /// </summary>
    public class WeldNameRuleSeqGen : NameRuleBase
    {
        /// <summary>
        /// Creates a name for the given entity. The name is based on the Pipeline
        /// to which the Weld belongs.  It is something like this: "Pipeline Name" + Index.
        /// The Pipeline is found in AddNamingParents() and is returned as the first
        /// item in business object collection. The naming rule semantic first calls AddNamingParents()
        /// and then ComputeName().
        ///          
        /// Other naming rules may base the name on other parents and may construct a
        /// different form for the name.
        /// </summary>
        /// <param name="entity">The business object whose name is being computed.</param>
        /// <param name="parents">The parentsNaming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject entity, System.Collections.ObjectModel.ReadOnlyCollection<BusinessObject> parents, BusinessObject activeEntity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }

            if (parents == null || parents.Count <= 0)
            {
                throw new ArgumentNullException("The naming parents of the entity to be named are null");
            }

            if (activeEntity == null)
            {
                throw new ArgumentNullException("activeEntity");
            }
            try
            {
                INamedItem pipelineNamedItem;
                INamedItem weldNamedItem;
                string pipeLineName = string.Empty;

                //GetNamingParents returns business object collection.
                //GetNamingParents placed this Weld's Pipeline as the first item in the collection.            

                //Get the INamedItem inteface for the Pipeline
                pipelineNamedItem = (INamedItem)parents[0];

                //This is the interface on which we change the Weld's name
                weldNamedItem = (INamedItem)entity;

                if (pipelineNamedItem != null)
                {
                    pipeLineName = pipelineNamedItem.Name;
                }

                if (string.IsNullOrEmpty(pipeLineName))
                {
                    pipeLineName = "Pipeline";
                }

                //If base part of name hasn't changed, don't compute a new name.
                PropertyValueString sequenceProperty = (PropertyValueString)entity.GetPropertyValue("IJSequence", "Id");
                string sequenceId = sequenceProperty.PropValue;

                //Build the Weld's name and set it.
                string weldName = pipeLineName + sequenceId;

                if (weldNamedItem.Name != weldName)
                {
                    SetName(entity, weldName);
                }
            }
            catch (Exception)
            {

                throw new Exception("Unexpected error:  PipingNameRules.WeldNameRuleSeqGen.ComputeName");
            }

        }

        /// <summary>
        /// Gets the naming parents from naming rule.
        /// All the Naming Parents that need to participate in an objects naming are added here to the
        /// business object collection. The parents added here are used in computing the name of the object in
        /// ComputeName() of the same class. Both these methods are called from naming rule semantic.
        /// In the case of Welds, we are using the Pipeline as our naming parent. Pipeline->Run->Part->Weld relation
        /// is used here
        /// The naming rule semantic will take the parents returned in the business object  collection and                   
        /// create a relationship so that whenever their name changes, the rule is again fired and
        /// the name is updated by this rule.
        /// </summary>
        /// <param name="entity">BusinessObject for which naming parents are required.</param>
        /// <returns>ReadOnlyCollection of BusinessObjects</returns>
        public override System.Collections.ObjectModel.Collection<BusinessObject> GetNamingParents(BusinessObject entity)
        {
            if (entity == null)
            {
                throw new ArgumentNullException("entity");
            }

            //Create business object collection for parents of the given entity. In this case, only the Pipeline is needed.
            Collection<BusinessObject> parentsColl = new Collection<BusinessObject>();
            try
            {
                //Get Part, which is parent of Weld
                RelationCollection relColl = entity.GetRelationship("OwnsImpliedItems", "Owner");
                ReadOnlyCollection<BusinessObject> boColl = relColl.TargetObjects;

                BusinessObject partOccurenceBO = null;
                if (boColl != null && boColl.Count > 0)
                {
                    partOccurenceBO = boColl[0];
                }

                if (partOccurenceBO != null)
                {
                    //Get Run, which is parent of Part
                    SystemChildHelper child = new SystemChildHelper(partOccurenceBO);
                    BusinessObject parent = (BusinessObject)child.SystemParent;

                    child = null;

                    //Get Pipeline, which is parent of Run
                    child = new SystemChildHelper(parent);
                    parent = (BusinessObject)child.SystemParent;

                    //Add Pipeline to parents collection
                    parentsColl.Add(parent);

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return parentsColl;
        }
    }
}
