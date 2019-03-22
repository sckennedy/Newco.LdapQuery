using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;

namespace LdapQuery
{
    class Program
    {
        private static string _outputFolder = @"C:\code\LdapInfo";
        static void Main(string[] args)
        {
            if (!string.IsNullOrEmpty(args[1]))
                _outputFolder = args[1];

            var directoryEntry = new DirectoryEntry("LDAP://cisco.com/DC=cisco,DC=com", @"simkenne@cisco.com", "C0ntr4ct0r!!");

            if (!Directory.Exists(_outputFolder))
                Directory.CreateDirectory(_outputFolder);

            var lines = new List<string>();
            var notFound = new List<string>();

            lines.Add("Username|FirstName|LastName|PhoneNumber|Mobile|Title|Department|Manager|Email|DeskNumber|Building|Street|City|State|PostCode|Country");

            DirectorySearcher searcher = CreateSearcher(directoryEntry);
            SearchResultCollection results = searcher.FindAll();
            
            foreach (SearchResult result in results)
            {
                //if we didn't find anyone then add to the notFound list
                if (result != null)
                {
                    var user = User.FromDirectorySearchResult(result);
                    if (user.FirstName != "Not Found" 
                        && user.FirstName != "Generic" 
                        && user.FirstName != "Admin" 
                        && !user.Username.StartsWith("CONF") 
                        && !user.Username.StartsWith("MCU")
                        && !user.Username.StartsWith("CTS_"))
                    {
                        lines.Add(user.ToSingleLineString("|"));
                        Console.WriteLine($"{user.Username} done.");
                    }
                }
            }

            Console.WriteLine("Completed");
            if (lines.Count > 0)
            {
                var fileName = WriteOutputToFile(lines, "LdapResults");
                Console.WriteLine($"File generated at {fileName}");
            }

            if (notFound.Count > 0)
            {
                var notFoundFilename = WriteOutputToFile(notFound, "NotFound");
                Console.WriteLine($"List of users not found at {notFoundFilename}");
            }
            Console.ReadLine();
        }

        private static string WriteOutputToFile(List<string> lines, string name)
        {
            //create filename based on timestamp of the day
            var fileName = $@"{_outputFolder}\{name}_{DateTime.Now.ToShortDateString().Replace("/", "_")}-{DateTime.Now.ToShortTimeString().Replace(":", "-")}.txt";
            File.WriteAllLines(fileName, lines.ToArray());
            return fileName;
        }

        private static DirectorySearcher CreateSearcher(DirectoryEntry directoryEntry)
        {
            var searcher = new DirectorySearcher(directoryEntry)
            {
                PageSize = int.MaxValue,
                Filter = $"(&(objectCategory=person)(objectClass=user))"
            };

            searcher.PropertiesToLoad.Add("sAMAccountName");
            searcher.PropertiesToLoad.Add("givenName");
            searcher.PropertiesToLoad.Add("sn");
            searcher.PropertiesToLoad.Add("telephoneNumber");
            searcher.PropertiesToLoad.Add("mobile");
            searcher.PropertiesToLoad.Add("physicalDeliveryOfficeName");
            searcher.PropertiesToLoad.Add("title");
            searcher.PropertiesToLoad.Add("department");
            searcher.PropertiesToLoad.Add("manager");
            searcher.PropertiesToLoad.Add("mail");
            searcher.PropertiesToLoad.Add("displayName");
            searcher.PropertiesToLoad.Add("name");
            searcher.PropertiesToLoad.Add("streetAddress");
            searcher.PropertiesToLoad.Add("postOfficeBox");
            searcher.PropertiesToLoad.Add("l");
            searcher.PropertiesToLoad.Add("st");
            searcher.PropertiesToLoad.Add("postalCode");
            searcher.PropertiesToLoad.Add("co");
            return searcher;
        }
    }
}
