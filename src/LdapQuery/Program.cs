using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;

namespace LdapQuery
{
    class Program
    {
        private static string _usernamesFile = @"C:\code\usernames.txt";
        private static string _outputFolder = @"C:\code\LdapInfo";
        static void Main(string[] args)
        {
            if (!string.IsNullOrEmpty(args[0]))
                _usernamesFile = args[0];

            if (!string.IsNullOrEmpty(args[1]))
                _outputFolder = args[1];

            var directoryEntry = new DirectoryEntry("LDAP://cisco.com");

            if (!File.Exists(_usernamesFile))
            {
                Console.WriteLine($"{_usernamesFile} does not exist.  Terminating.");
                Console.ReadLine();
                return;
            }

            if (!Directory.Exists(_outputFolder))
                Directory.CreateDirectory(_outputFolder);

            var usernames = File.ReadAllLines(_usernamesFile);
            var lines = new List<string>();
            var notFound = new List<string>();

            lines.Add("Username|FirstName|LastName|PhoneNumber|Mobile|Title|Department|Manager|Email|DeskNumber|Building|Street|City|State|PostCode|Country");

            foreach (var username in usernames)
            {
                DirectorySearcher searcher = CreateSearcher(directoryEntry, username);

                var result = searcher.FindOne();

                //if we didn't find anyone then add to the notFound list
                if (result == null)
                {
                    Console.WriteLine($"{username} not found.");
                    notFound.Add(username);
                }
                else
                {
                    var user = User.FromDirectorySearchResult(result, username);
                    lines.Add(user.ToSingleLineString("|"));
                    Console.WriteLine($"{username} done.");
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

        private static DirectorySearcher CreateSearcher(DirectoryEntry directoryEntry, string username)
        {
            var searcher = new DirectorySearcher(directoryEntry)
            {
                PageSize = int.MaxValue,
                Filter = $"(&(objectCategory=person)(objectClass=user)(sAMAccountName={username}))"
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
