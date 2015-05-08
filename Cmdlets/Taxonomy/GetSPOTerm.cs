using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.Data.Taxonomy;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Taxonomy
{
    [Cmdlet("Get", "SPOTerm", SupportsShouldProcess = false)]
    [CmdletGroup("Managed Metadata")]
    [CmdletDescription("Retrieves a Term or Terms matching the specified criteria.")]
    [RelatedCmdlets(typeof(ExportSPOTaxonomy))]
    [Example(Code = "PS C:\\> Get-SPOTerm -Identity \"My Term\" -StringMatchOption \"ExactMatch\"",
        Remarks = "This example retrieves one or more Terms matching the value My Term.")]
    public class GetSPOTerm : BaseSPOCmdlet
    {
        public GetSPOTerm()
        {
            StringMatchOption = Microsoft.SharePoint.Client.Taxonomy.StringMatchOption.ExactMatch;
            ResultCollectionSize = 1;
            //LCID = 0;
        }

        [Alias("Term", "Value", "TermLabel")]
        [Parameter(Mandatory = true,
            ValueFromPipeline = true,
            Position = 0,
            HelpMessage = "The Term to retrieve.")]
        public string Identity { get; set; }

        [Parameter(Mandatory = false,
            Position = 1,
            HelpMessage = "Retrieve terms that match the value exactly or just start with the value. The default value is ExactMatch.")]
        public StringMatchOption StringMatchOption { get; set; }

        [Parameter(Mandatory = false,
            Position = 2,
            HelpMessage = "If specified,only default labels should be matched against in the search.")]
        public SwitchParameter DefaultLabelOnly { get; set; }

        //[Parameter(Mandatory = false,
        //    Position = 3,
        //    HelpMessage = "If specified, exclude Term objects that are in the keyword’s TermSet.")]
        //public SwitchParameter ExcludeKeyword { get; set; }

        [Parameter(Mandatory = false,
            Position = 4,
            HelpMessage = "If specified, trim Terms that are unavailable.")]
        public SwitchParameter TrimUnavailable { get; set; }

        [Parameter(Mandatory = false,
            Position = 5,
            HelpMessage = "If specified, trim out Term objects that have Term.IsDeprecated set to true.")]
        public SwitchParameter TrimDeprecated { get; set; }

        [Parameter(Mandatory = false,
            Position = 6,
            HelpMessage = "Indicates the maximum number of objects to include in the returned collection. The default value is 1.")]
        public int ResultCollectionSize { get; set; }

        //[Alias("Locale", "Language")]
        //[Parameter(Mandatory = false,
        //    Position = 7,
        //    HelpMessage = "The language code identifier (LCID) of the Term.Labels to be matched.")]
        //public int LCID { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            TaxonomySession ts = TaxonomySession.GetTaxonomySession(ctx);
            //if (LCID == 0)
            //    LCID = GetWorkingLanguage(ts);

            //LabelMatchInformation lmi = new LabelMatchInformation(ctx);
            //ctx.Load(lmi);
            //ctx.ExecuteQuery();
            //lmi.StringMatchOption = StringMatchOption;
            //lmi.DefaultLabelOnly = DefaultLabelOnly.ToBool();
            //lmi.ExcludeKeyword = ExcludeKeyword.ToBool();
            //lmi.TrimUnavailable = TrimUnavailable.ToBool();
            //lmi.TrimDeprecated = TrimDeprecated.ToBool();
            //lmi.ResultCollectionSize = ResultCollectionSize;
            //lmi.Lcid = LCID;
            //lmi.TermLabel = Identity;

            //var terms = ts.GetTerms(lmi);

            var terms = ts.GetTermsInDefaultLanguage(Identity, DefaultLabelOnly, StringMatchOption, ResultCollectionSize, TrimUnavailable, TrimDeprecated);
            ctx.Load(terms);
            ctx.ExecuteQuery();
            foreach (var term in terms)
            {
                WriteObject(new SPOTerm(term));
            }
        }
        //private int GetWorkingLanguage(TaxonomySession ts)
        //{
        //    var ctx = base.Context;
        //    ctx.Load(ts.TermStores);
        //    ctx.ExecuteQuery();

        //    TermStore termStore = ts.TermStores[0];
        //    ctx.Load(termStore, t => t.WorkingLanguage);
        //    ctx.ExecuteQuery();
        //    return termStore.WorkingLanguage;
        //}

    }
}
