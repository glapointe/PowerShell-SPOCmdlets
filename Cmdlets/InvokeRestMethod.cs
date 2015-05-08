using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management.Automation;
using Lapointe.PowerShell.MamlGenerator.Attributes;
using Lapointe.SharePointOnline.PowerShell.PipeBind;
using System.Net;
using Lapointe.SharePointOnline.PowerShell.Enums;
using System.IO;
using Microsoft.PowerShell.Commands;
using System.Collections;
using System.Xml;
using System.Globalization;
using Lapointe.SharePointOnline.PowerShell.Utilities;
using System.Web;
using Lapointe.SharePointOnline.PowerShell.Validators;

namespace Lapointe.SharePointOnline.PowerShell.Cmdlets
{
    [Cmdlet("Invoke", "SPORestMethod", SupportsShouldProcess = false)]
    [CmdletGroup("General")]
    [CmdletDescription("Sends an HTTP or HTTPS request to a SharePoint Online REST-compliant web service.",
        "This function sends an HTTP or HTTPS request to a Representational State Transfer (REST)-compliant (\"RESTful\") SharePoint Online web service. The Connect-SPOSite cmdlet must be run first in order to set the credentials for the environment.\r\n\r\nSome source code for this cmdlet was referenced from http://blog.vgrem.com/2014/02/27/working-with-the-sharepoint-online-rest-service-via-powershell/.")]
    [RelatedCmdlets(typeof(ConnectSPOSite), typeof(DisconnectSPOSite))]
    [Example(Code = "PS C:\\> Connect-SPOSite -Url \"https://contoso.sharepoint.com/\" -Credential \"admin@contoso.com\"\r\n" +
                    "PS C:\\> Invoke-SPORestMethod \"https://contoso.sharepoint.com/_api/web\"",
        Remarks = "This example invokes the REST call at https://contoso.sharepoint.com/_api/web.")]
    public class InvokeRestMethod : PSCmdlet
    {
        public InvokeRestMethod() { 
            Method = WebRequestMethod.Get;
            JSONVerbosity = "Verbose";
            XHTTPMethod = WebRequestMethod.Default;
        }

        [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, HelpMessage = "The REST endpoint URL to call.")]
        [ValidateNotNullOrEmpty]
        public UrlPipeBind Url { get; set; }

        [Parameter(Mandatory = false, Position = 1, HelpMessage = "Specifies the method used for the web request. The default value is \"Get\".")]
        [ValidateNotNullOrEmpty]
        public WebRequestMethod Method { get; set; }

        [Parameter(Mandatory = false, Position = 2, ValueFromPipeline = true, HelpMessage = "Specifies the body of the request. The body is the content of the request that follows the headers. You can also pipe a body value to Invoke-SPORestMethod.")]
        [ValidateNotNullOrEmpty]
        [Alias("Metadata")]
        public object Body { get; set; }

        [Parameter(Mandatory = false, Position = 3, HelpMessage = "The \"X-RequestDigest\" header to set. This is most commonly used to provide the form digest variable. Use \"(Invoke-SPORestMethod -Url \"https://contoso.sharepoint.com/_api/contextinfo\" -Method \"Post\").GetContextWebInformation.FormDigestValue\" to get the Form Digest value.")]
        [ValidateNotNullOrEmpty]
        public string RequestDigest { get; set; }
        
        [Parameter(Mandatory = false, Position = 4, HelpMessage = "The \"If-Match\" header to set. Provide this to make sure you are not overwritting an item that has changed since you retrieved it.")]
        [ValidateNotNullOrEmpty]
        public string ETag { get; set; }

        [Parameter(Mandatory = false, Position = 5, HelpMessage = "To work around the fact that many firewalls and other network intermediaries block HTTP verbs other than GET and POST, specify PUT, DELETE, or MERGE requests for -XHTTPMethod with a POST value for -Method.")]
        [ValidateNotNullOrEmpty]
        public WebRequestMethod XHTTPMethod { get; set; }

        [Parameter(Mandatory = false, Position = 6, HelpMessage = "The OData verbosity of the data to return. Valid values are \"Verbose\", \"MinimalMetadata\", and \"NoMetadata\".")]
        [ValidateNotNullOrEmpty]
        [ValidateSet("Verbose", "MinimalMetadata", "NoMetadata")]
        public string JSONVerbosity { get; set; }

        [Parameter(Mandatory = false, Position = 7, HelpMessage = "If the returned data is a binary data object such as a file from a SharePoint site specify the output file name to save the data to.")]
        [ValidateNotNullOrEmpty]
        [ValidateDirectoryExistsAndValidFileName]
        public string OutFile { get; set; }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            if (SPOSiteContext.CurrentSiteContext == null)
                throw new Exception("You must call Connect-SPOSite before you can run the Invoke-SPORestMethod cmdlet.");
        }
        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            Uri url = this.Url.Read();


            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Credentials = SPOSiteContext.CurrentSiteContext.Context.Credentials;
            string odata = ";odata=" + JSONVerbosity.ToLower();
            request.Accept = "application/json" + odata;
            request.ContentType = "application/json" + odata;
            request.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
            request.Method = Method.ToString().ToUpperInvariant();

            if (!string.IsNullOrEmpty(RequestDigest)) {
                request.Headers.Add("X-RequestDigest", RequestDigest);
            }
            if (!string.IsNullOrEmpty(ETag)) {
                request.Headers.Add("If-Match", ETag);
            }
            if (XHTTPMethod != WebRequestMethod.Default) {
                request.Headers.Add("X-HTTP-Method", XHTTPMethod.ToString().ToUpperInvariant());
            }

            AddBody(request, Body);

            //if (!string.IsNullOrEmpty(Metadata)) {
            //    byte[] body = Encoding.UTF8.GetBytes(Metadata);
            //    request.ContentLength = body.Length;
            //    var stream = request.GetRequestStream();
            //    stream.Write(body, 0, body.Length);
            //} else {
            //    request.ContentLength = 0;
            //}

            using (var response = request.GetResponse())
            using (var streamReader = new StreamReader(response.GetResponseStream()))
            {
                if (response.ContentType.Contains("application/octet-stream"))
                {
                    if (!string.IsNullOrEmpty(OutFile))
                    {
                        using (System.IO.FileStream fs = System.IO.File.Create(OutFile))
                        {
                            streamReader.BaseStream.CopyTo(fs);
                        }
                        return;
                    }
                    WriteObject(streamReader.ReadToEnd());
                    return;
                }

                string data = streamReader.ReadToEnd();
                ErrorRecord errRecord;
                if (!response.ContentType.Contains("application/json"))
                {
                    WriteObject(data);
                    return;
                }
                if (data.Contains("\"ID\":") && data.Contains("\"Id\":"))
                    data = data.Replace("\"ID\":", "\"ID-dup\":");
                object results = JsonObject.ConvertFromJson(data, out errRecord);
                if (errRecord != null)
                {
                    WriteWarning("Unable to convert JSON data to PSObject.");
                    WriteObject(data);
                    ThrowTerminatingError(errRecord);
                }
                if (results is PSObject)
                {
                    if (JSONVerbosity.ToLower() != "verbose")
                        WriteObject(results);
                    else
                        if (((PSObject)results).Properties["d"] != null && ((PSObject)results).Properties["d"].Value != null)
                            WriteObject(((PSObject)results).Properties["d"].Value);
                        else
                            WriteObject(results);
                }
                else
                    WriteObject(results);
            }
        }

        private void AddBody(WebRequest request, object body)
        {
            if (body == null) return;

            if (body is PSObject)
            {
                body = ((PSObject)body).BaseObject;
            }
            if (!(body is HtmlWebResponseObject))
            {
                if (body is FormObject)
                {
                    FormObject formObject = body as FormObject;
                    SetRequestContent(request, formObject.Fields);
                }
                else if ((body is IDictionary) && (request.Method != "GET"))
                {
                    IDictionary content = body as IDictionary;
                    SetRequestContent(request, content);
                }
                else if (body is XmlNode)
                {
                    XmlNode xmlNode = body as XmlNode;
                    SetRequestContent(request, xmlNode);
                }
                else if (body is Stream)
                {
                    Stream contentStream = body as Stream;
                    SetRequestContent(request, contentStream);
                }
                else if (body is byte[])
                {
                    byte[] buffer = body as byte[];
                    SetRequestContent(request, buffer);
                }
                else
                {
                    SetRequestContent(request, (string)LanguagePrimitives.ConvertTo(body, typeof(string), CultureInfo.InvariantCulture));
                }
            }
            else
            {
                HtmlWebResponseObject htmlWebResponse = body as HtmlWebResponseObject;
                if (htmlWebResponse.Forms.Count == 1)
                {
                    SetRequestContent(request, htmlWebResponse.Forms[0].Fields);
                }
            }
        }
        internal long SetRequestContent(WebRequest request, byte[] content)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            if (content != null)
            {
                request.ContentLength = content.Length;
                StreamHelper.WriteToStream(content, request.GetRequestStream());
            }
            else
            {
                request.ContentLength = 0L;
            }
            return request.ContentLength;
        }

        internal long SetRequestContent(WebRequest request, IDictionary content)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            if (content == null)
            {
                throw new ArgumentNullException("content");
            }
            string str = this.FormatDictionary(content);
            return this.SetRequestContent(request, str);
        }


        internal long SetRequestContent(WebRequest request, Stream contentStream)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            if (contentStream == null)
            {
                throw new ArgumentNullException("contentStream");
            }
            request.ContentLength = contentStream.Length;
            StreamHelper.WriteToStream(contentStream, request.GetRequestStream(), this);
            return request.ContentLength;
        }

        internal long SetRequestContent(WebRequest request, string content)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            if (content != null)
            {
                byte[] input = StreamHelper.EncodeToBytes(content);
                request.ContentLength = input.Length;
                StreamHelper.WriteToStream(input, request.GetRequestStream());
            }
            else
            {
                request.ContentLength = 0L;
            }
            return request.ContentLength;
        }


        internal long SetRequestContent(WebRequest request, XmlNode xmlNode)
        {
            if (request == null)
            {
                throw new ArgumentNullException("request");
            }
            if (xmlNode != null)
            {
                byte[] input = null;
                XmlDocument document = xmlNode as XmlDocument;
                if ((document != null) && (document.FirstChild is XmlDeclaration))
                {
                    XmlDeclaration firstChild = document.FirstChild as XmlDeclaration;
                    Encoding encoding = Encoding.GetEncoding(firstChild.Encoding);
                    input = StreamHelper.EncodeToBytes(document.OuterXml, encoding);
                }
                else
                {
                    input = StreamHelper.EncodeToBytes(xmlNode.OuterXml);
                }
                request.ContentLength = input.Length;
                StreamHelper.WriteToStream(input, request.GetRequestStream());
            }
            else
            {
                request.ContentLength = 0L;
            }
            return request.ContentLength;
        }

        private string FormatDictionary(IDictionary content)
        {
            if (content == null)
            {
                throw new ArgumentNullException("content");
            }
            StringBuilder builder = new StringBuilder();
            foreach (string str in content.Keys)
            {
                if (0 < builder.Length)
                {
                    builder.Append("&");
                }
                object obj2 = content[str];
                string str2 = HttpUtility.UrlEncode(str);
                string str3 = string.Empty;
                if (obj2 != null)
                {
                    str3 = HttpUtility.UrlEncode(obj2.ToString());
                }
                builder.AppendFormat("{0}={1}", str2, str3);
            }
            return builder.ToString();
        }


    }
}
