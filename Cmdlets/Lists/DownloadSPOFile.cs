using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using Microsoft.SharePoint.Client;
using Lapointe.SharePointOnline.PowerShell.Data.Sites;
using Lapointe.SharePointOnline.PowerShell.Attributes;
using Lapointe.SharePointOnline.PowerShell.Data.Lists;
using System.IO;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets.Lists
{
    [Cmdlet("Download", "SPOFile", SupportsShouldProcess = false, DefaultParameterSetName = "File")]
    [CmdletGroup("Lists")]
    [CmdletDescription("Downloads a specific File from a given Site.",
        "Download and save a specific file from a Site given the server relative URL to the file. The Site provided by the -Web parameter must correspond to the Site where the file is located or else an argument out of range exception will be thrown.")]
    [RelatedCmdlets(typeof(ConnectSPOSite))]
    [Example(Code = "PS C:\\> Download-SPOFile -Web \"/\" -File \"/Shared Documents/MyFile.docx\" -TargetFolder \"c:\\temp\" -Overwrite",
        Remarks = "Downloads the file MyFile.docx from the Documents library within the root Site of the current Site Collection and saves to the c:\\temp folder.")]
    public class DownloadSPOFile : BaseSPOCmdlet
    {
        [ValidateNotNullOrEmpty,
        Parameter(
            Position = 0,
            ValueFromPipeline = true,
            Mandatory = true,
            HelpMessage = "Specifies the identity of the Site containing the File(s) to retrieve.\r\n\r\nThe type must be a valid server relative URL, in the form /site_name, or an SPOWeb object, Microsoft.SharePoint.Client.Web object, or GUID representing the Site ID.")]
        public SPOWebPipeBind Web { get; set; }

        [ValidateNotNullOrEmpty,
        Parameter(
            ParameterSetName = "File",
            Position = 1,
            Mandatory = true,
            HelpMessage = "The server relative URL(s) to the file(s) to retrieve.")]
        public SPOFilePipeBind[] File { get; set; }


        [Parameter(Mandatory = true,
            ParameterSetName = "Folder",
            Position = 1,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "The path to the folder to save.")]
        [ValidateNotNull]
        public SPOFolderPipeBind Folder { get; set; }

        [Parameter(ParameterSetName = "Folder",
            Position = 2,
            HelpMessage = "Save all child folders and their files.")]
        public SwitchParameter Recursive { get; set; }

        [Parameter(Mandatory = true,
            Position = 3,
            HelpMessage = "The path to the folder to save the files to.")]
        [ValidateDirectoryExists]
        [Alias("Path")]
        public string TargetFolder { get; set; }

        [Parameter(Position = 4, HelpMessage = "Overwrite existing files.")]
        public SwitchParameter Overwrite { get; set; }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            var ctx = base.Context;

            SPOWeb web = new SPOWeb(ctx.Site.OpenWeb(Web.Read()));

            if (ParameterSetName == "File")
            {
                foreach (var filePipe in File)
                {
                    var file = filePipe.Read(web);
                    WriteFile(TargetFolder, file);
                }
            }
            else if (ParameterSetName == "Folder")
            {
                var folder = Folder.Read(web);
                WriteFiles(TargetFolder, folder);
            }

        }

        private void WriteFiles(string targetFolder, SPOFolder folder)
        {
            if (!Directory.Exists(targetFolder))
            {
                WriteVerbose(string.Format("Creating folder {0}...", targetFolder));
                Directory.CreateDirectory(targetFolder);
            }

            foreach (SPOFile file in folder.GetFiles())
            {
                WriteFile(targetFolder, file);
            }
            if (Recursive)
            {
                foreach (SPOFolder childFolder in folder.GetFolders(false))
                {
                    WriteFiles(Path.Combine(targetFolder, childFolder.Name), childFolder);
                }
            }
        }

        private void WriteFile(string targetFolder, SPOFile file)
        {
            string fullName = Path.Combine(targetFolder, file.Name);
            if (System.IO.File.Exists(fullName) && !Overwrite)
            {
                WriteWarning(string.Format("Unable to save \"{0}\". File already exists. Use the -Overwrite parameter to overwrite.", fullName));
                return;
            }
            WriteVerbose(string.Format("Saving {0} to {1}...", file.Name, targetFolder));
            file.SaveFileLocally(targetFolder, Overwrite);
        }
    }
}
