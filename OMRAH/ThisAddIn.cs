using System;
using System.Collections.Specialized;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace OMRAH
{
    public partial class ThisAddIn
    {
        Outlook.Folder inbox;
        Outlook.Items inboxItems;
        string userAddress;
        string userName;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.userAddress = Application.GetNamespace("MAPI").CurrentUser.AddressEntry.GetExchangeUser()?.PrimarySmtpAddress;
            this.userName = Application.GetNamespace("MAPI").CurrentUser.Name;
            Debug.WriteLine("User address: " + userAddress);
            Debug.WriteLine("User name: " + userName);

            this.inbox = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;

            // Define the event handler for new emails
            this.inboxItems = this.inbox.Items;
            this.inboxItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(AutoHandleRequest);

            // Check all meeting requests already in the inbox
            Outlook.Items meetingRequests = this.inboxItems
                //.Restrict("[Unread]=true")
                .Restrict("[MessageClass] = 'IPM.Schedule.Meeting.Request'");
            foreach (Outlook.MeetingItem meetingRequest in meetingRequests) AutoHandleRequest(meetingRequest);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }

        private void AutoHandleRequest(object item)
        {
            if (!(item is Outlook.MeetingItem)) return;
            Outlook.MeetingItem meetingRequest = item as Outlook.MeetingItem;
            Outlook.AppointmentItem appointment = meetingRequest.GetAssociatedAppointment(true);
            if (appointment == null) return;
            Debug.WriteLine("Request subject: " + appointment.Subject);
            StringCollection subjectFilters = Properties.Settings.Default.SubjectFilters;
            if (!Array.Exists(subjectFilters.Cast<string>().ToArray(), filter => appointment.Subject.ToLower().Contains(filter))) return;
            StringCollection recipientFilters = Properties.Settings.Default.RecipientFilters;
            bool unlessDirect = Properties.Settings.Default.UnlessDirect;
            bool filterMatch = false;
            foreach (Outlook.Recipient recipient in appointment.Recipients)
            {
                string recipientAddress = recipient.AddressEntry.GetExchangeUser()?.PrimarySmtpAddress;
                string recipientName = recipient.Name;
                Debug.WriteLine("Recipient address: " + recipientAddress);
                Debug.WriteLine("Recipient name: " + recipientName);
                if (unlessDirect && (recipientAddress == this.userAddress || recipientName == this.userName))
                {
                    filterMatch = false;
                    break;
                }
                else if (
                    Array.Exists(
                        recipientFilters.Cast<string>().ToArray(), filter =>
                            (recipientAddress is string && recipientAddress.Contains(filter))
                            || recipientName.Contains(filter)
                    )
                )
                {
                    filterMatch = true;
                    if (!unlessDirect) break;
                }
            }
            Debug.WriteLine($"Filter match: {filterMatch}");
            if (filterMatch)
            {
                appointment.Categories = ResolveCategory();
                Outlook.MeetingItem response = appointment.Respond(Outlook.OlMeetingResponse.olMeetingTentative, false);
                //response.Send();
                response.Close(Outlook.OlInspectorClose.olDiscard);
                meetingRequest.Delete();
                appointment.Save();
            }
        }

        private string ResolveCategory()
        {
            string categoryName = Properties.Settings.Default.CategoryName;
            Outlook.OlCategoryColor categoryColor = Properties.Settings.Default.CategoryColor;
            Outlook.Categories categories = Application.Session.Categories;
            bool categoryExists;
            try
            {
                Outlook.Category category = categories[categoryName];
                if (category != null)
                {
                    category.Color = categoryColor;
                    categoryExists = true;
                }
                else categoryExists = false;
            }
            catch { categoryExists = false; }
            if (!categoryExists) categories.Add(categoryName, categoryColor);
            return categoryName;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
