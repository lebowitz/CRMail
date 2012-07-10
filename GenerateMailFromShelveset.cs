

namespace crmail
{
    using System;
    using System.IO;
    using System.Text;
    using Microsoft.TeamFoundation.VersionControl.Client;
    using Outlook;
    using System.Globalization;

    /// <summary>
    /// Create outlook email from shelveset
    /// </summary>
    public class MailFromShelveset
    {
        #region Ctor/Init
        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="vcs">Version Control server.</param>
        /// <param name="shelvesetName">Shelveset name.</param>
        /// <param name="reviewAlias">The email alias of the code review panel. Null if none is present</param>
        /// <param name="webViewPort">Port number on Team Foundation server where web-view is hosted. 
        /// This is typically 8090.</param>
        public MailFromShelveset(VersionControlServer vcs, string shelvesetName, 
            string reviewAlias, int webViewPort)
        {
            this.versionControl = vcs;
            this.crAlias = reviewAlias;
            // convert the port number to a :<port> format
            this.webViewPortStr = string.Format(CultureInfo.InvariantCulture, ":{0}", webViewPort);

            // if the shelvesetname is of the form foo;bar then foo is the shelveset
            // name and bar is the alias of the owner.
            string[] textArray = shelvesetName.Split(new char[] { ';' });
            if (textArray.Length > 1)
            {
                this.shelveName = textArray[0];
                this.shelveOwner = textArray[1];
            }
            else
            {
                this.shelveName = textArray[0];
                this.shelveOwner = "."; // Dot means current user
            }
        }
        #endregion 

        #region Public Methods
        /// <summary>
        /// Generate the email.
        /// </summary>
        /// <remarks>This is the driver method which calls other methods to create 
        /// parts of the email.</remarks>
        public void GenerateMail()
        {
            Shelveset shelveset = this.GetShelveset(this.shelveName, this.shelveOwner);

            MailItem item = GenerateMailItem(shelveset);
            item.HTMLBody = this.GenerateMailBody(shelveset);
            item.Display(this);
        }

        #endregion 

        #region Helpers

        /// <summary>
        /// Get the shelveset metadata from the server.
        /// </summary>
        /// <param name="name">Shelveset name</param>
        /// <param name="owner">Shelveset owner.</param>
        /// <returns>Shelveset details.</returns>
        private Shelveset GetShelveset(string name, string owner)
        {
            Console.WriteLine(Resources.InfoGettingShelveset);
            Shelveset[] shelvesetArray = this.versionControl.QueryShelvesets(name, owner);
            if (shelvesetArray.Length == 0)
            {
                throw new System.Exception(string.Format(CultureInfo.InvariantCulture,
                                           Resources.ErrFailToGetShelveset, name));
            }

            Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                                           Resources.InfoGotShelveset, name));
            return shelvesetArray[0];
        }
        
        /// <summary>
        /// Create the email item and fill up to/cc/subject fields.
        /// </summary>
        /// <param name="shelveset">Shelveset.</param>
        /// <returns>MailITem with Subject, To and Cc filled</returns>
        private MailItem GenerateMailItem(Shelveset shelveset)
        {
            Console.WriteLine(Resources.InfoGeneratingEmail);
            Application application = new ApplicationClass();
            MailItem mailItem  = (MailItem) new Outlook.Application().CreateItem(OlItemType.olMailItem);
            mailItem.Subject = string.Format(CultureInfo.InvariantCulture,
                "CR: <area> {0};{1}", shelveset.Name, shelveset.OwnerName);
            mailItem.BodyFormat = OlBodyFormat.olFormatHTML;
            mailItem.To = this.GetReviewer(shelveset);
            mailItem.CC = this.crAlias;

            return mailItem;
        }

        /// <summary>
        /// Generate the email body.
        /// </summary>
        /// <param name="shelve">Shelveset</param>
        /// <returns>Formatted HTML email body.</returns>
        private string GenerateMailBody(Shelveset shelveset)
        {
            // HACKHACK: We are converting tfs server url http://foo:8080 to http://foo:8090 blindly
            string tfsWebUIServer = shelveset.VersionControlServer.TeamFoundationServer.Uri.ToString().Replace(":8080", 
                this.webViewPortStr);
            
            StringBuilder strBuilder = new StringBuilder(0x400);

            // Start the table
            strBuilder.AppendLine("<p");
            strBuilder.AppendLine("<table style=\"font-family:verdana;font-size:10pt;\" border=1 cellspacing=0 cellpadding=0>");

            // Reviewers row
            strBuilder.AppendFormat("<tr><td valign=top><b>{0}</b></td>\n", Resources.EmailReviewer);
            string textReviewer = this.GetReviewer(shelveset);
            if (string.IsNullOrEmpty(textReviewer))
                textReviewer = Resources.EmailReviewerNotPresent;

            if (!string.IsNullOrEmpty(textReviewer))
            {
                strBuilder.AppendFormat("<td valign=top>{0}</td></tr>\n", textReviewer);
            }

            // TFS Server details Row
            strBuilder.AppendFormat("<tr><td valign=top><b>TFS Server</b></td><td valign=top>{0}</td></tr>",
                                    shelveset.VersionControlServer.TeamFoundationServer.Uri);

            strBuilder.AppendFormat("<tr><td valign=top><b>Shelveset</b></td><td valign=top><a href=\"{0}/ss.aspx?ss={1};{2}\">{1};{2}</a></td></tr>", 
                tfsWebUIServer, 
                shelveset.Name, 
                shelveset.OwnerName
            );

            // Description/Comment Row
            string textDescription = this.HtmlEncode(shelveset.Comment);
            strBuilder.AppendFormat("<tr><td valign=top><b>{0}</b></td><td valign=top>{1}</td></tr>\n", 
                                        Resources.EmailDescription, this.HtmlEncode(shelveset.Comment));

            // Bugs Row
            strBuilder.AppendFormat("<tr><td valign=top><b>{0}</b></td><td valign=top>\n", 
                                        Resources.EmailBugs);
            foreach (WorkItemCheckinInfo info in shelveset.WorkItemInfo)
            {
                strBuilder.AppendFormat("<a href={0}WorkItemTracking/Workitem.aspx?artifactMoniker={1}>{1}</a> {2}<br>\n",
                    shelveset.VersionControlServer.TeamFoundationServer.Uri, info.WorkItem.Id, info.WorkItem.Title);
            }

            strBuilder.AppendLine("</td></tr>");

            // Files row
            strBuilder.AppendLine("<tr><td valign=top><b>File(s)</b></td>");
            strBuilder.AppendLine("<td  style=\"font-family:courier new;font-size:10pt;\" valign=top>");
            PendingSet[] pendingSets = this.versionControl.QueryShelvedChanges(shelveset);
            foreach (PendingSet pendingSet in pendingSets)
            {
                foreach (PendingChange changes in pendingSet.PendingChanges)
                {
                    string changeName = changes.ChangeTypeName.Substring(0, 3);
                    strBuilder.Append(changeName);
                    if (changes.ChangeType == ChangeType.Edit)
                    {
                        strBuilder.AppendFormat(" <a href={0}/history.aspx?item={1}>H</a> ", tfsWebUIServer, changes.ItemId);
                        strBuilder.AppendFormat(" <a href={0}/ann.aspx?item={1}>B</a> ", tfsWebUIServer, changes.ItemId);
                        string diffUrl = string.Format("{0}/UI/Pages/Scc/Difference.aspx?oitem={1}&ocs=-1&mpcid={2}",
                            tfsWebUIServer, changes.ItemId, changes.PendingChangeId);

                        strBuilder.AppendLine(string.Format("<a href={0}>{1}</a> <br>",
                            diffUrl, changes.ServerItem));

                    }
                    else if ((changes.ChangeType & ChangeType.Add) == ChangeType.Add)
                    {
                        strBuilder.AppendFormat(" H B ");
                        string newUrl = string.Format("{0}/UI/Pages/Scc/ViewSource.aspx?pcid={1}",
                            tfsWebUIServer, changes.PendingChangeId);

                        strBuilder.AppendLine(string.Format("<a href={0}>{1}</a> <br>",
                            newUrl, changes.ServerItem));
                    }   
                    else if (changes.ChangeType == ChangeType.Delete)
                    {
                        strBuilder.AppendFormat(" <a href={0}/history.aspx?item={1}>H</a> ", tfsWebUIServer, changes.ItemId);
                        strBuilder.AppendFormat(" <a href={0}/ann.aspx?item={1}>B</a> ", tfsWebUIServer, changes.ItemId);
                        string newUrl = string.Format("{0}/UI/Pages/Scc/ViewSource.aspx?pcid={1}",
                            tfsWebUIServer, changes.PendingChangeId);

                        strBuilder.AppendLine(string.Format("<a href={0}>{1}</a> <br>",
                            newUrl, changes.ServerItem));
                    }
                    else
                    {
                        strBuilder.AppendFormat(" H B ");
                        strBuilder.AppendLine(string.Format("{1} <br>",
                               changes.ServerItem));
                    }
                }
            }

            strBuilder.AppendLine("</td></tr>");
            
            // Test rows to be hand filled by the user
            strBuilder.AppendLine("<tr><td valign=top><b>Tests Run</b></td><td valign=top></td></tr>");
            strBuilder.AppendLine("<tr><td valign=top><b>Tests Added/Fixed</b></td><td valign=top></td></tr>");

            // close the table
            strBuilder.AppendLine("</table>");
            return strBuilder.ToString();
        }

        /// <summary>
        /// Get email reviewer name.
        /// </summary>
        private string GetReviewer(Shelveset shelveset)
        {
            foreach (CheckinNoteFieldValue value1 in shelveset.CheckinNote.Values)
            {
                if (value1.Name == "Code Reviewer")
                {
                    return value1.Value;
                }
            }
            return null;
        }

        /// <summary>
        /// Do some html conversion.
        /// </summary>
        private string HtmlEncode(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            // TODO: do propert convertion like it's done in System.Web.HttpServerUtility.HtmlEncode.
            text = text.Replace("\n", "<BR>");
            return text;
        }
        #endregion // Helpers

        #region Private
        private string shelveName;
        private string shelveOwner;
        private VersionControlServer versionControl;
        private string crAlias;
        private string webViewPortStr;
        #endregion
    }
}
