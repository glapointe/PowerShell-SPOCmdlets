using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Taxonomy
{
    public class ImportTermsHelper
    {
        private XmlDocument _xml;
        private ClientContext _ctx;
        private BaseSPOCmdlet _cmdlet = null;
        private int _workingLanguage = 0;

        public ImportTermsHelper(BaseSPOCmdlet cmdlet, ClientContext ctx, XmlDocument xml)
        {
            _cmdlet = cmdlet;
            _ctx = ctx;
            _xml = xml;
            if (xml.DocumentElement == null)
                throw new Exception("The XML provided does not include a root element.");
        }

        public void Import()
        {
            TaxonomySession ts = TaxonomySession.GetTaxonomySession(_ctx);
            Import(ts);
        }

        public void Import(TaxonomySession ts)
        {
            if (ts == null)
                throw new ArgumentNullException("ts", "The TaxonomySession object is null.");

            XmlNodeList termStoreNodes = _xml.SelectNodes("//TermStore");
            if (termStoreNodes == null || termStoreNodes.Count == 0)
                return;

            var termStores = ts.TermStores;
            _ctx.Load(termStores);
            _ctx.ExecuteQuery();

            foreach (XmlElement termStoreElement in termStoreNodes)
            {
                TermStore termStore = null;
                if (termStoreNodes.Count == 1 && termStores.Count == 1)
                {
                    termStore = termStores[0];
                }
                else
                {
                    string termStoreName = termStoreElement.GetAttribute("Name");

                    termStore = ts.TermStores.GetByName(termStoreName);
                    _ctx.ExecuteQuery();
                    if (termStore == null || termStore.ServerObjectIsNull.Value)
                    {
                        _cmdlet.WriteWarning(string.Format("Unable to locate target Term Store: {0}", termStoreName));
                        continue;
                    }
                }
                LoadWorkingLanguage(termStore);
                _cmdlet.WriteVerbose(string.Format("Importing into Term Store: {0}", termStore.Name));

                XmlNodeList groupNodes = termStoreElement.SelectNodes("./Groups/Group");
                if (groupNodes == null || groupNodes.Count == 0)
                {
                    _cmdlet.WriteWarning("No Group elements were defined in the import file for the Term Store.");
                    continue;
                }
                foreach (XmlElement groupElement in groupNodes)
                {
                    Import(groupElement, termStore);
                }
                termStore.CommitAll();
                termStore.Context.ExecuteQuery();
            }
        }

        public void Import(TermStore parentTermStore)
        {
            if (parentTermStore == null || parentTermStore.ServerObjectIsNull.Value)
                throw new ArgumentNullException("parentTermStore", "The parent TermStore object is null.");

            LoadWorkingLanguage(parentTermStore);
            if (_xml.DocumentElement.Name == "Groups")
            {
                XmlNodeList groupNodes = _xml.SelectNodes("./Groups/Group");
                if (groupNodes == null || groupNodes.Count == 0)
                    return;

                foreach (XmlElement groupElement in groupNodes)
                {
                    Import(groupElement, parentTermStore);
                }
            }
            else if (_xml.DocumentElement.Name == "Group")
            {
                Import(_xml.DocumentElement, parentTermStore);
            }
            parentTermStore.CommitAll();
            parentTermStore.Context.ExecuteQuery();
        }

        public void Import(TermGroup parentGroup)
        {
            if (parentGroup == null || parentGroup.ServerObjectIsNull.Value)
                throw new ArgumentNullException("parentGroup", "The parent Group object is null.");

            LoadWorkingLanguage(parentGroup.TermStore);
            if (_xml.DocumentElement.Name == "TermSets")
            {
                XmlNodeList termSetNodes = _xml.SelectNodes("./TermSets/TermSet");
                if (termSetNodes == null || termSetNodes.Count == 0)
                {
                    _cmdlet.WriteWarning("No Term Set elements were defined in the import file for the Group.");
                    return;
                }

                foreach (XmlElement termSetElement in termSetNodes)
                {
                    Import(termSetElement, parentGroup);
                }
            }
            else if (_xml.DocumentElement.Name == "TermSet")
            {
                Import(_xml.DocumentElement, parentGroup);
            }
            parentGroup.TermStore.CommitAll();
            parentGroup.Context.ExecuteQuery();
        }

        public void Import(TermSet parentTermSet)
        {
            if (parentTermSet == null || parentTermSet.ServerObjectIsNull.Value)
                throw new ArgumentNullException("parentTermSet", "The parent TermSet object is null.");

            LoadWorkingLanguage(parentTermSet.TermStore);
            XmlNodeList termNodes;
            if (_xml.DocumentElement.Name == "Terms")
            {
                termNodes = _xml.SelectNodes("./Terms/Term");
                if (termNodes == null || termNodes.Count == 0)
                {
                    _cmdlet.WriteWarning("No Term elements were defined in the import file for the Term Set.");
                    return;
                }
                foreach (XmlElement termElement in termNodes)
                {
                    Import(termElement, parentTermSet);
                }
            }
            else if (_xml.DocumentElement.Name == "Term")
            {
                Import(_xml.DocumentElement, parentTermSet);
            }
            parentTermSet.TermStore.CommitAll();
        }


        public void Import(Term parentTerm)
        {
            if (parentTerm == null || parentTerm.ServerObjectIsNull.Value)
                throw new ArgumentNullException("parentTerm", "The parent Term object is null.");

            LoadWorkingLanguage(parentTerm.TermStore);
            XmlNodeList termNodes;
            if (_xml.DocumentElement.Name == "Terms")
            {
                termNodes = _xml.SelectNodes("./Terms/Term");
                if (termNodes == null || termNodes.Count == 0)
                {
                    _cmdlet.WriteWarning("No Term elements were defined in the import file for the Term.");
                    return;
                }
                foreach (XmlElement termElement in termNodes)
                {
                    Import(termElement, parentTerm);
                }
            }
            else if (_xml.DocumentElement.Name == "Term")
            {
                Import(_xml.DocumentElement, parentTerm);
            }
            parentTerm.TermStore.CommitAll();
            parentTerm.Context.ExecuteQuery();
        }

        private void Import(XmlElement groupElement, TermStore parentTermStore)
        {
            string groupName = groupElement.GetAttribute("Name");
            Guid groupId = new Guid(groupElement.GetAttribute("Id"));
            LoadWorkingLanguage(parentTermStore);

            TermGroup group = null;
            ExceptionHandlingScope scope = new ExceptionHandlingScope(_ctx);
            using (scope.StartScope())
            {
                using (scope.StartTry())
                {
                    group = parentTermStore.Groups.GetByName(groupName);
                    _ctx.Load(group);
                }
                using (scope.StartCatch())
                {
                }
            }
            _ctx.ExecuteQuery();

            if (group == null || group.ServerObjectIsNull == null || group.ServerObjectIsNull.Value)
            {
                _cmdlet.WriteVerbose(string.Format("Creating Group: {0}", groupName));

                group = parentTermStore.CreateGroup(groupName, groupId);
                group.Description = groupElement.GetAttribute("Description");
                group.Context.ExecuteQuery();
                parentTermStore.CommitAll();
                parentTermStore.Context.ExecuteQuery();
            }
            XmlNodeList termSetNodes = groupElement.SelectNodes("./TermSets/TermSet");
            if (termSetNodes != null && termSetNodes.Count > 0)
            {
                foreach (XmlElement termSetElement in termSetNodes)
                {
                    Import(termSetElement, group);
                }
            }
        }

        private void Import(XmlElement termSetElement, TermGroup parentGroup)
        {
            string termSetName = termSetElement.GetAttribute("Name");
            LoadWorkingLanguage(parentGroup.TermStore);

            TermSet termSet = null;
            ExceptionHandlingScope scope = new ExceptionHandlingScope(_ctx);
            using (scope.StartScope())
            {
                using (scope.StartTry())
                {
                    termSet = parentGroup.TermSets.GetByName(termSetName);
                    _ctx.Load(termSet);
                }
                using (scope.StartCatch())
                {
                }
            }
            _ctx.ExecuteQuery();

            if (termSet == null || termSet.ServerObjectIsNull == null || termSet.ServerObjectIsNull.Value)
            {
                _cmdlet.WriteVerbose(string.Format("Creating Term Set: {0}", termSetName));

                int lcid = _workingLanguage;
                Guid id = Guid.NewGuid();
                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("Id")))
                    id = new Guid(termSetElement.GetAttribute("Id"));
                termSet = parentGroup.CreateTermSet(termSetName, id, lcid);

                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("Contact")))
                    termSet.Contact = termSetElement.GetAttribute("Contact");
                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("Description")))
                    termSet.Description = termSetElement.GetAttribute("Description");
                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("CustomSortOrder")))
                    termSet.CustomSortOrder = termSetElement.GetAttribute("CustomSortOrder");
                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("IsAvailableForTagging")))
                    termSet.IsAvailableForTagging = bool.Parse(termSetElement.GetAttribute("IsAvailableForTagging"));
                if (!string.IsNullOrEmpty(termSetElement.GetAttribute("Owner")))
                    termSet.Owner = termSetElement.GetAttribute("Owner");

                termSet.IsOpenForTermCreation = true;
                parentGroup.TermStore.CommitAll();
                _ctx.Load(termSet);
                _ctx.ExecuteQuery();
            }
            
            XmlNodeList stakeholderNodes = termSetElement.SelectNodes("./Stakeholders/Stakeholder");
            if (stakeholderNodes != null && stakeholderNodes.Count > 0)
            {
                bool modified = false;
                foreach (XmlElement stakeholderElement in stakeholderNodes)
                {
                    string stakeholder = stakeholderElement.GetAttribute("Value");
                    if (!termSet.Stakeholders.Contains(stakeholder))
                    {
                        termSet.AddStakeholder(stakeholder);
                    }
                }
                if (modified)
                    _ctx.ExecuteQuery();
            }

            XmlNodeList termsNodes = termSetElement.SelectNodes("./Terms/Term");
            if (termsNodes != null && termsNodes.Count > 0)
            {
                foreach (XmlElement termElement in termsNodes)
                {
                    Import(termElement, termSet);
                }
            }
            

            if (!string.IsNullOrEmpty(termSetElement.GetAttribute("IsOpenForTermCreation")))
            {
                bool isOpenForTermCreation = bool.Parse(termSetElement.GetAttribute("IsOpenForTermCreation"));
                if (isOpenForTermCreation != termSet.IsOpenForTermCreation && termSet.Name != "Hashtags")
                {
                    var orphanedTermsTermSet = termSet.TermStore.OrphanedTermsTermSet;
                    termSet.Context.Load(orphanedTermsTermSet, o => o.Id);
                    termSet.Context.ExecuteQuery();

                    if (orphanedTermsTermSet.Id != termSet.Id)
                    {
                        termSet.IsOpenForTermCreation = isOpenForTermCreation;
                        termSet.Context.ExecuteQuery();
                    }
                }
            }
        }

        private void Import(XmlElement termElement, TermSetItem parentTermSetItem)
        {
            string termName = termElement.GetAttribute("Name");
            LoadWorkingLanguage(parentTermSetItem.TermStore);

            Term term = null;
            ExceptionHandlingScope scope = new ExceptionHandlingScope(_ctx);
            using (scope.StartScope())
            {
                using (scope.StartTry())
                {
                    term = parentTermSetItem.Terms.GetByName(termName);
                    _ctx.Load(term);
                }
                using (scope.StartCatch())
                {
                }
            }
            _ctx.ExecuteQuery();


            if (term == null || term.ServerObjectIsNull == null || term.ServerObjectIsNull.Value)
            {
                if (!string.IsNullOrEmpty(termElement.GetAttribute("IsSourceTerm")) &&
                    !bool.Parse(termElement.GetAttribute("IsSourceTerm")))
                {
                    string[] sourceTermInfo = termElement.GetAttribute("SourceTerm").Split('|');

                    Term sourceTerm = null;
                    ExceptionHandlingScope scope1 = new ExceptionHandlingScope(_ctx);
                    using (scope1.StartScope())
                    {
                        using (scope1.StartTry())
                        {
                            sourceTerm = parentTermSetItem.TermStore.GetTerm(new Guid(sourceTermInfo[0]));
                            _ctx.Load(sourceTerm);
                        }
                        using (scope1.StartCatch())
                        {
                        }
                    }
                    _ctx.ExecuteQuery();

                    if (sourceTerm == null || sourceTerm.ServerObjectIsNull == null || sourceTerm.ServerObjectIsNull.Value)
                    {
                        LabelMatchInformation lmi = new LabelMatchInformation(parentTermSetItem.Context);
                        lmi.StringMatchOption = StringMatchOption.ExactMatch;
                        lmi.DefaultLabelOnly = true;
                        lmi.ResultCollectionSize = 1;
                        lmi.TermLabel = sourceTermInfo[1];
                        lmi.TrimUnavailable = false;
                        TermCollection sourceTerms = parentTermSetItem.TermStore.GetTerms(lmi);
                        _ctx.ExecuteQuery();
                        if (sourceTerms != null && !sourceTerms.ServerObjectIsNull.Value && sourceTerms.Count > 0)
                            sourceTerm = sourceTerms[0];
                    }
                    if (sourceTerm != null)
                    {
                        _cmdlet.WriteVerbose(string.Format("Creating Reference Term: {0}", termName)); 
                        term = parentTermSetItem.ReuseTerm(sourceTerm, false);
                        _ctx.ExecuteQuery();
                    }
                    else
                        _cmdlet.WriteWarning(string.Format("The Source Term, {0}, was not found. {1} will be created without linking.", sourceTermInfo[1], termName));
                }
                if (term == null || term.ServerObjectIsNull == null || term.ServerObjectIsNull.Value)
                {
                    _cmdlet.WriteVerbose(string.Format("Creating Term: {0}", termName));
                    
                    int lcid = _workingLanguage;
                    Guid id = Guid.NewGuid();
                    if (!string.IsNullOrEmpty(termElement.GetAttribute("Id")))
                        id = new Guid(termElement.GetAttribute("Id"));
                    term = parentTermSetItem.CreateTerm(termName, lcid, id);

                    if (!string.IsNullOrEmpty(termElement.GetAttribute("CustomSortOrder")))
                        term.CustomSortOrder = termElement.GetAttribute("CustomSortOrder");
                    if (!string.IsNullOrEmpty(termElement.GetAttribute("IsAvailableForTagging")))
                        term.IsAvailableForTagging = bool.Parse(termElement.GetAttribute("IsAvailableForTagging"));
                    if (!string.IsNullOrEmpty(termElement.GetAttribute("Owner")))
                        term.Owner = termElement.GetAttribute("Owner");

                    if (!string.IsNullOrEmpty(termElement.GetAttribute("IsDeprecated")) &&
                        bool.Parse(termElement.GetAttribute("IsDeprecated")))
                        term.Deprecate(true);

                    _ctx.ExecuteQuery();
                }
            }

            XmlNodeList descriptionNodes = termElement.SelectNodes("./Descriptions/Description");
            if (descriptionNodes != null && descriptionNodes.Count > 0)
            {
                foreach (XmlElement descriptionElement in descriptionNodes)
                {
                    term.SetDescription(descriptionElement.GetAttribute("Value"),
                        int.Parse(descriptionElement.GetAttribute("Language")));
                }
                _ctx.ExecuteQuery();
            }

            XmlNodeList propertyNodes = termElement.SelectNodes("./CustomProperties/CustomProperty");
            if (propertyNodes != null && propertyNodes.Count > 0)
            {
                foreach (XmlElement propertyElement in propertyNodes)
                {
                    term.SetCustomProperty(propertyElement.GetAttribute("Name"), 
                        propertyElement.GetAttribute("Value"));
                }
                _ctx.ExecuteQuery();
            }

            XmlNodeList localPropertyNodes = termElement.SelectNodes("./LocalCustomProperties/LocalCustomProperty");
            if (localPropertyNodes != null && localPropertyNodes.Count > 0)
            {
                foreach (XmlElement propertyElement in localPropertyNodes)
                {
                    term.SetLocalCustomProperty(propertyElement.GetAttribute("Name"),
                        propertyElement.GetAttribute("Value"));
                }
                _ctx.ExecuteQuery();
            }

            XmlNodeList labelNodes = termElement.SelectNodes("./Labels/Label");
            if (labelNodes != null && labelNodes.Count > 0)
            {
                foreach (XmlElement labelElement in labelNodes)
                {
                    string labelValue = labelElement.GetAttribute("Value");
                    int lcid = int.Parse(labelElement.GetAttribute("Language"));
                    bool isDefault = bool.Parse(labelElement.GetAttribute("IsDefaultForLanguage"));
                    var labels = term.GetAllLabels(lcid);
                    parentTermSetItem.Context.Load(labels);
                    parentTermSetItem.Context.ExecuteQuery();

                    Label label = labels.FirstOrDefault(currentLabel => currentLabel.Value == labelValue);
                    if (label == null || label.ServerObjectIsNull.Value)
                    {
                        term.CreateLabel(labelValue, lcid, isDefault);
                        parentTermSetItem.Context.ExecuteQuery();
                    }
                    else
                    {
                        if (isDefault && !label.IsDefaultForLanguage)
                        {
                            label.SetAsDefaultForLanguage();
                            parentTermSetItem.Context.ExecuteQuery();
                        }
                    }
                }
            }
            parentTermSetItem.TermStore.CommitAll();
            _ctx.ExecuteQuery();

            XmlNodeList termsNodes = termElement.SelectNodes("./Terms/Term");
            if (termsNodes != null && termsNodes.Count > 0)
            {
                foreach (XmlElement childTermElement in termsNodes)
                {
                    Import(childTermElement, term);
                }
            }
        }

        private void LoadWorkingLanguage(TermStore termStore)
        {
            if (_workingLanguage > 0) return;
            if (!termStore.IsPropertyAvailable("WorkingLanguage"))
            {
                _ctx.Load(termStore, l => l.WorkingLanguage);
                _ctx.ExecuteQuery();
            }
            _workingLanguage = termStore.WorkingLanguage;
        }

    }
}
