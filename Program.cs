using System;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using Microsoft.Lync.Model;
using System.Collections.Generic;

namespace Lync_set_note
{
    class Program
    {
        private static LyncClient lc;

        static int Main(string[] args)
        {
            String userName;

            if (args.Length == 0)
            {
                // getting current user from AD
                userName = Environment.UserName;
            } else
                userName = args[0];

            Console.WriteLine("Setting note for Lync");
            Console.WriteLine("=====================");
            Console.WriteLine("Getting current AD settings...");

            Domain currentDomain = Domain.GetCurrentDomain();
            Console.WriteLine("Current domain: {0}", currentDomain.Name);

            DomainController dc = currentDomain.FindDomainController();
            Console.WriteLine("Domain controller: {0}", dc.Name);

            DirectoryEntry de = new DirectoryEntry("LDAP://" + currentDomain.Name);
            DirectorySearcher deSearch = new DirectorySearcher(de, "(&(objectClass=user)(sAMAccountName="+ userName+"))");
            deSearch.PropertiesToLoad.Add("cn");
            deSearch.PropertiesToLoad.Add("displayName");
            deSearch.PropertiesToLoad.Add("telephoneNumber");

            System.DirectoryServices.SearchResult searchUser= deSearch.FindOne();
            if( searchUser == null) { 
                Console.WriteLine("Not found info about user: {0}", userName);
                return 1;
            }
            DirectoryEntry currentUser = searchUser.GetDirectoryEntry();
            string phoneNo = currentUser.Properties["telephoneNumber"].Value.ToString();

            Console.WriteLine("Found user: {0}, phone: {1}", currentUser.Name, phoneNo);

            try
            {
                lc = LyncClient.GetClient();
            }
            catch (LyncClientException ex)
            {
                Console.WriteLine("Error 0x{0:X}: {1}", ex.InternalCode, ex.Message);
                return 1;
            }
            Console.WriteLine("Current personal note: {0}", getLyncNote());
            setLyncNote("Тел. " + phoneNo);

            return 0;
        }

        private static void setLyncNote(string newNote)
        {
    
            try
            {
                Dictionary<PublishableContactInformationType, object> publishData = new Dictionary<PublishableContactInformationType, object>();

                publishData.Add(PublishableContactInformationType.PersonalNote, newNote);

                object[] asyncState = { lc.Self };

                lc.Self.BeginPublishContactInformation(publishData, PublishContactInformationCallback, asyncState);
            }
            catch
            { }
        }

        private static string getLyncNote()
        {
            string text = string.Empty;
            try
            {
                text = lc.Self.Contact.GetContactInformation(ContactInformationType.PersonalNote).ToString();
                return text;
            }
            catch
            { return text; }
        }

        #region Callbacks
        // <summary> 
        // Callback invoked when Self.BeginPublishContactInformation is completed 
        // </summary> 
        // <param name="result">The status of the asynchronous operation</param> 
        private static void PublishContactInformationCallback(IAsyncResult result)
        {
            lc.Self.EndPublishContactInformation(result);
            Console.WriteLine("Phone sets as personal note succeccfully");
        }
        #endregion
    }
}
