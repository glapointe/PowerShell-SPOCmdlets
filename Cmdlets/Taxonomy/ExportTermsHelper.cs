using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Taxonomy
{
    public class ExportTermsHelper
    {
        private XmlDocument _xml = new XmlDocument();
        private ClientContext _ctx;
        private IEnumerable<int> _languages = null;

        public ExportTermsHelper(ClientContext ctx)
        {
            _ctx = ctx;
        }
        
        private XmlElement AddTermStoreElement(XmlElement parent, TermStore termStore)
        {
            XmlElement element = _xml.CreateElement("TermStore");
            if (parent == null)
                _xml.AppendChild(element);
            else
                parent.AppendChild(element);

            element.SetAttribute("Name", termStore.Name);
            element.SetAttribute("Id", termStore.Id.ToString());
            element.SetAttribute("IsOnline", termStore.IsOnline.ToString());
            element.SetAttribute("WorkingLanguage", termStore.WorkingLanguage.ToString());
            element.SetAttribute("DefaultLanguage", termStore.DefaultLanguage.ToString());
            if (!termStore.SystemGroup.IsPropertyAvailable("Id") || !termStore.IsPropertyAvailable("ContentTypePublishingHub"))
            {
                _ctx.Load(termStore, ts => ts.ContentTypePublishingHub);
                _ctx.Load(termStore.SystemGroup, s => s.Id);
                _ctx.ExecuteQuery();
            }
            element.SetAttribute("SystemGroup", termStore.SystemGroup.Id.ToString());
            element.SetAttribute("ContentTypePublishingHub", termStore.ContentTypePublishingHub);

            return element;
        }

        private XmlElement GetTermStoreElement(XmlElement parent, TermStore termStore)
        {
            if (parent == null)
                return AddTermStoreElement(parent, termStore);

            XmlElement termStoreElement = parent.SelectSingleNode("./TermStore[@Id='" + termStore.Id + "']") as XmlElement;
            if (termStoreElement == null)
                termStoreElement = AddTermStoreElement(parent, termStore);
            return termStoreElement;
        }

        private XmlElement GetGroupElement(XmlElement parent, TermGroup group)
        {
            if (parent == null)
                return AddGroupElement(parent, group);

            XmlElement groupElement = parent.SelectSingleNode("./Groups/Group[@Id='" + group.Id + "']") as XmlElement;
            if (groupElement == null)
                groupElement = AddGroupElement(parent, group);

            return groupElement;
        }

        private XmlElement AddGroupElement(XmlElement parent, TermGroup group)
        {
            XmlElement element = _xml.CreateElement("Group");

            if (parent == null)
                _xml.AppendChild(element);
            else
            {
                if (parent.Name == "Groups")
                    parent.AppendChild(element);
                else
                {
                    XmlElement groupsElement = parent.SelectSingleNode("./Groups") as XmlElement;
                    if (groupsElement == null)
                    {
                        groupsElement = _xml.CreateElement("Groups");
                        parent.AppendChild(groupsElement);
                    }
                    groupsElement.AppendChild(element);
                }
            }

            element.SetAttribute("Id", group.Id.ToString());
            element.SetAttribute("Name", group.Name);
            element.SetAttribute("Description", group.Description);
            element.SetAttribute("CreatedDate", group.CreatedDate.ToString());
            element.SetAttribute("LastModifiedDate", group.LastModifiedDate.ToString());
            element.SetAttribute("IsSystemGroup", group.IsSystemGroup.ToString());
            element.SetAttribute("IsSiteCollectionGroup", group.IsSiteCollectionGroup.ToString());

            return element;
        }


        private XmlElement GetTermSetElement(XmlElement parent, TermSet termSet)
        {
            if (parent == null)
                return AddTermSetElement(parent, termSet);

            XmlElement termSetElement = parent.SelectSingleNode("./TermSets/TermSet[@Id='" + termSet.Id + "']") as XmlElement;
            if (termSetElement == null)
                termSetElement = AddTermSetElement(parent, termSet);

            return termSetElement;
        }


        private XmlElement AddTermSetElement(XmlElement parent, TermSet termSet)
        {
            XmlElement element = _xml.CreateElement("TermSet");
            if (parent == null)
                _xml.AppendChild(element);
            else
            {
                if (parent.Name == "TermSets")
                    parent.AppendChild(element);
                else
                {
                    XmlElement termSetsElement = parent.SelectSingleNode("./TermSets") as XmlElement;
                    if (termSetsElement == null)
                    {
                        termSetsElement = _xml.CreateElement("TermSets");
                        parent.AppendChild(termSetsElement);
                    }
                    termSetsElement.AppendChild(element);
                }
            }

            element.SetAttribute("Id", termSet.Id.ToString());
            element.SetAttribute("Name", termSet.Name);
            element.SetAttribute("Description", termSet.Description);
            element.SetAttribute("CreatedDate", termSet.CreatedDate.ToString());
            element.SetAttribute("LastModifiedDate", termSet.LastModifiedDate.ToString());
            element.SetAttribute("Contact", termSet.Contact);
            element.SetAttribute("Owner", termSet.Owner);
            element.SetAttribute("IsAvailableForTagging", termSet.IsAvailableForTagging.ToString());
            element.SetAttribute("IsOpenForTermCreation", termSet.IsOpenForTermCreation.ToString());
            element.SetAttribute("CustomSortOrder", termSet.CustomSortOrder);

            XmlElement propertiesElement = _xml.CreateElement("CustomProperties");
            element.AppendChild(propertiesElement);
            foreach (string key in termSet.CustomProperties.Keys)
            {
                XmlElement propertyElement = _xml.CreateElement("CustomProperty");
                propertiesElement.AppendChild(propertyElement);
                propertyElement.SetAttribute("Name", key);
                propertyElement.SetAttribute("Value", termSet.CustomProperties[key]);
            }

            XmlElement stakeholdersElement = _xml.CreateElement("Stakeholders");
            element.AppendChild(stakeholdersElement);
            foreach (string stakeholder in termSet.Stakeholders)
            {
                XmlElement stakeholderElement = _xml.CreateElement("Stakeholder");
                stakeholderElement.SetAttribute("Value", stakeholder);
                stakeholdersElement.AppendChild(stakeholderElement);
            }

            return element;
        }



        private XmlElement GetTermElement(XmlElement parent, Term term)
        {
            if (parent == null)
                return AddTermElement(parent, term);

            XmlElement termElement = parent.SelectSingleNode("./Terms/Term[@Id='" + term.Id + "']") as XmlElement;
            if (termElement == null)
                termElement = AddTermElement(parent, term);

            return termElement;
        }


        private XmlElement AddTermElement(XmlElement parent, Term term)
        {
            XmlElement element = _xml.CreateElement("Term");
            if (parent == null)
                _xml.AppendChild(element);
            else
            {
                if (parent.Name == "TermSets")
                    parent.AppendChild(element);
                else
                {
                    XmlElement termsElement = parent.SelectSingleNode("./Terms") as XmlElement;
                    if (termsElement == null)
                    {
                        termsElement = _xml.CreateElement("Terms");
                        parent.AppendChild(termsElement);
                    }
                    termsElement.AppendChild(element);
                }
            }
            
            element.SetAttribute("Id", term.Id.ToString());
            element.SetAttribute("Name", term.Name);
            element.SetAttribute("CreatedDate", term.CreatedDate.ToString());
            element.SetAttribute("LastModifiedDate", term.LastModifiedDate.ToString());
            element.SetAttribute("Owner", term.Owner);
            element.SetAttribute("IsDeprecated", term.IsDeprecated.ToString());
            element.SetAttribute("IsAvailableForTagging", term.IsAvailableForTagging.ToString());
            element.SetAttribute("IsKeyword", term.IsKeyword.ToString());
            element.SetAttribute("IsReused", term.IsReused.ToString());
            element.SetAttribute("IsRoot", term.IsRoot.ToString());
            element.SetAttribute("IsSourceTerm", term.IsSourceTerm.ToString());
            element.SetAttribute("CustomSortOrder", term.CustomSortOrder);
            element.SetAttribute("IsPinned", term.IsPinned.ToString());
            element.SetAttribute("IsPinnedRoot", term.IsPinnedRoot.ToString());
            element.SetAttribute("PathOfTerm", term.PathOfTerm);

            if (!term.SourceTerm.IsPropertyAvailable("Id") || !term.SourceTerm.IsPropertyAvailable("Name"))
            {
                _ctx.Load(term.SourceTerm, st => st.Id, st => st.Name);
                _ctx.ExecuteQuery();
            }
            element.SetAttribute("SourceTerm", term.SourceTerm.Id + "|" + term.SourceTerm.Name);

            if (!term.PinSourceTermSet.IsPropertyAvailable("Id") || !term.PinSourceTermSet.IsPropertyAvailable("Name"))
            {
                _ctx.Load(term.PinSourceTermSet, st => st.Id, st => st.Name);
                _ctx.ExecuteQuery();
            }
            if (!term.PinSourceTermSet.ServerObjectIsNull.Value)
                element.SetAttribute("PinSourceTermSet", term.PinSourceTermSet.Id + "|" + term.PinSourceTermSet.Name);

            XmlElement descriptionsElement = _xml.CreateElement("Descriptions");
            element.AppendChild(descriptionsElement);
            LoadLanguages(term.TermStore);
            foreach (int lcid in _languages)
            {
                XmlElement descriptionElement = _xml.CreateElement("Description");
                descriptionsElement.AppendChild(descriptionElement);
                descriptionElement.SetAttribute("Language", lcid.ToString());
                var desc = term.GetDescription(lcid);
                _ctx.ExecuteQuery();
                descriptionElement.SetAttribute("Value", desc.Value);
            }

            XmlElement propertiesElement = _xml.CreateElement("CustomProperties");
            element.AppendChild(propertiesElement);
            foreach (string key in term.CustomProperties.Keys)
            {
                XmlElement propertyElement = _xml.CreateElement("CustomProperty");
                propertiesElement.AppendChild(propertyElement);
                propertyElement.SetAttribute("Name", key);
                propertyElement.SetAttribute("Value", term.CustomProperties[key]);
            }

            XmlElement localPropertiesElement = _xml.CreateElement("LocalCustomProperties");
            element.AppendChild(localPropertiesElement);
            foreach (string key in term.LocalCustomProperties.Keys)
            {
                XmlElement propertyElement = _xml.CreateElement("LocalCustomProperty");
                localPropertiesElement.AppendChild(propertyElement);
                propertyElement.SetAttribute("Name", key);
                propertyElement.SetAttribute("Value", term.LocalCustomProperties[key]);
            }

            XmlElement mergedTermIdsElement = _xml.CreateElement("MergedTermIds");
            element.AppendChild(mergedTermIdsElement);
            foreach (Guid mergedTermId in term.MergedTermIds)
            {
                XmlElement mergedTermIdElement = _xml.CreateElement("MergedTermId");
                mergedTermIdsElement.AppendChild(mergedTermIdElement);
                mergedTermIdElement.SetAttribute("Value", mergedTermId.ToString());
            }

            return element;
        }


        private XmlElement AddLabelElement(XmlElement termElement, Label label)
        {
            XmlElement labelsElement = termElement.SelectSingleNode("./Labels") as XmlElement;
            if (labelsElement == null)
            {
                labelsElement = _xml.CreateElement("Labels");
                termElement.AppendChild(labelsElement);
            }

            XmlElement element = _xml.CreateElement("Label");
            labelsElement.AppendChild(element);
            element.SetAttribute("Value", label.Value);
            element.SetAttribute("Language", label.Language.ToString());
            element.SetAttribute("IsDefaultForLanguage", label.IsDefaultForLanguage.ToString());

            return element;
        }

        private void LoadLanguages(TermStore termStore)
        {
            if (_languages != null) return;
            if (!termStore.IsPropertyAvailable("Languages"))
            {
                _ctx.Load(termStore, l => l.Languages);
                _ctx.ExecuteQuery();
            }
            _languages = termStore.Languages;
        }

        public XmlDocument Export()
        {
            TaxonomySession ts = TaxonomySession.GetTaxonomySession(_ctx);
            return Export(ts);
        }

        public XmlDocument Export(TaxonomySession ts)
        {
            if (ts == null)
                throw new ArgumentNullException("ts", "The TaxonomySession object is null.");

            XmlElement element = _xml.CreateElement("TermStores");
            _xml.AppendChild(element);

            var tsc = ts.TermStores;
            if (!ts.IsObjectPropertyInstantiated("TermStores") || !ts.TermStores.AreItemsAvailable)
            {
                _ctx.Load(tsc);
                _ctx.ExecuteQuery();
            }
            foreach (TermStore termStore in tsc)
            {
                LoadLanguages(termStore);
                Export(element, termStore);
            }

            return _xml;
        }

        public XmlDocument Export(TermStore termStore)
        {
            if (termStore == null)
                throw new ArgumentNullException("termStore", "The TermStore object is null.");

            LoadLanguages(termStore);
            XmlElement termStoreElement = AddTermStoreElement(null, termStore);
            var groups = termStore.Groups;
            if (!termStore.IsObjectPropertyInstantiated("Groups") || !termStore.Groups.AreItemsAvailable)
            {
                _ctx.Load(groups);
                _ctx.ExecuteQuery();
            }
            foreach (TermGroup group in groups)
            {
                Export(termStoreElement, group);
            }
            return _xml;
        }

        
        public XmlDocument Export(TermGroup group)
        {
            if (group == null)
                throw new ArgumentNullException("group", "The Group object is null.");

            LoadLanguages(group.TermStore);
            XmlElement groupElement = AddGroupElement(null, group);
            var termSets = group.TermSets;
            if (!group.IsObjectPropertyInstantiated("TermSets") || !group.TermSets.AreItemsAvailable)
            {
                _ctx.Load(termSets);
                _ctx.ExecuteQuery();
            }
            foreach (TermSet termSet in termSets)
            {
                Export(groupElement, termSet);
            }
            return _xml;
        }

        public XmlDocument Export(TermSet termSet)
        {
            if (termSet == null)
                throw new ArgumentNullException("termSet", "The TermSet object is null.");

            LoadLanguages(termSet.TermStore);
            XmlElement termSetElement = AddTermSetElement(null, termSet);
            var terms = termSet.Terms;
            if (!termSet.IsObjectPropertyInstantiated("Terms") || !termSet.Terms.AreItemsAvailable)
            {
                _ctx.Load(terms);
                _ctx.ExecuteQuery();
            }
            foreach (Term term in terms)
            {
                Export(termSetElement, term);
            }
            return _xml;
        }

        public XmlDocument Export(Term term)
        {
            if (term == null)
                throw new ArgumentNullException("term", "The Term object is null.");

            LoadLanguages(term.TermStore);
            XmlElement termElement = AddTermElement(null, term);
            var labels = term.Labels;
            if (!term.IsObjectPropertyInstantiated("Labels") || !term.Labels.AreItemsAvailable)
            {
                _ctx.Load(labels);
                _ctx.ExecuteQuery();
            }
            foreach (Label label in term.Labels)
            {
                AddLabelElement(termElement, label);
            }

            var terms = term.Terms;
            if (!term.IsObjectPropertyInstantiated("Terms") || !term.Terms.AreItemsAvailable)
            {
                _ctx.Load(terms);
                _ctx.ExecuteQuery();
            }
            foreach (Term childTerm in term.Terms)
            {
                Export(termElement, childTerm);
            }
            return _xml;
        }

        private void Export(XmlElement parentElement, TermStore termStore)
        {
            if (termStore == null)
                throw new ArgumentNullException("termStore", "The TermStore object is null.");

            XmlElement termStoreElement = GetTermStoreElement(parentElement, termStore);
            var groups = termStore.Groups;
            if (!termStore.IsObjectPropertyInstantiated("Groups") || !termStore.Groups.AreItemsAvailable)
            {
                _ctx.Load(groups);
                _ctx.ExecuteQuery();
            }
            foreach (TermGroup termSet in groups)
            {
                Export(termStoreElement, termSet);
            }

        }
        
        private void Export(XmlElement parentElement, TermGroup group)
        {
            XmlElement groupElement = GetGroupElement(parentElement, group);
            var termSets = group.TermSets;
            if (!group.IsObjectPropertyInstantiated("TermSets") || !group.TermSets.AreItemsAvailable)
            {
                _ctx.Load(termSets);
                _ctx.ExecuteQuery();
            }
            foreach (TermSet termSet in termSets)
            {
                Export(groupElement, termSet);
            }

        }

        private void Export(XmlElement parentElement, TermSet termSet)
        {
            XmlElement termSetElement = GetTermSetElement(parentElement, termSet);
            var terms = termSet.Terms;
            if (!termSet.IsObjectPropertyInstantiated("Terms") || !termSet.Terms.AreItemsAvailable)
            {
                _ctx.Load(terms);
                _ctx.ExecuteQuery();
            }
            foreach (Term term in terms)
            {
                Export(termSetElement, term);
            }
        }

        private void Export(XmlElement parentElement, Term term)
        {
            XmlElement termElement = GetTermElement(parentElement, term);
            var labels = term.Labels;
            if (!term.IsObjectPropertyInstantiated("Labels") || !term.Labels.AreItemsAvailable)
            {
                _ctx.Load(labels);
                _ctx.ExecuteQuery();
            }
            foreach (Label label in labels)
            {
                AddLabelElement(termElement, label);
            }

            var terms = term.Terms;
            if (!term.IsObjectPropertyInstantiated("Terms") || !term.Terms.AreItemsAvailable)
            {
                _ctx.Load(terms);
                _ctx.ExecuteQuery();
            }
            foreach (Term childTerm in terms)
            {
                Export(termElement, childTerm);
            }
        }

        
    }
}
