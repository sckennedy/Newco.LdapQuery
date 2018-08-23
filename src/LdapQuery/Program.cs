using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.IO;

namespace LdapQuery
{
    class Program
    {
        private static string usernamesFile = @"C:\code\usernames.txt";
        private static string outputFolder = @"C:\code\LdapInfo";
        static void Main(string[] args)
        {
            if (!string.IsNullOrEmpty(args[0]))
                usernamesFile = args[0];

            if (!string.IsNullOrEmpty(args[1]))
                outputFolder = args[1];

            DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://cisco.com");

            if (!File.Exists(usernamesFile))
            {
                Console.WriteLine($"{usernamesFile} does not exist.  Terminating.");
                Console.ReadLine();
                return;
            }

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            string[] usernames = File.ReadAllLines(usernamesFile);
            var lines = new List<string>();
            var notFound = new List<string>();

            lines.Add("Username|FirstName|Initials|LastName|PhoneNumber|Mobile|Title|Department|Manager|Email|DeskNumber|Building|Street|City|State|Country");

            foreach (var username in usernames)
            {
                string[] parts = username.Split(new char[] { '|' });

                DirectorySearcher searcher = CreateSearcher(directoryEntry, parts[0], parts[1]);

                var results = searcher.FindAll();

                //if we didn't find anyone then add to the notFound list
                if (results == null || results.Count == 0)
                {
                    Console.WriteLine($"{username} not found.");
                    notFound.Add(username);
                }
                else
                {
                    if (results.Count == 1)
                    {
                        //if we just have one result then go with that
                        var user = User.FromDirectorySearchResult(results[0], username);
                        lines.Add(user.ToSingleLineString("|"));
                    }
                    else
                    {
                        var found = false;
                        foreach (SearchResult result in results)
                        {
                            //otherwise loop through each result and try and find one that has a matching first name
                            var givenName = result.Properties["givenName"][0].ToString();
                            if (parts[0].ToLower().Contains(givenName.ToLower()) || givenName.ToLower().Contains(parts[0].ToLower()))
                            {
                                var user = User.FromDirectorySearchResult(result, username);
                                lines.Add(user.ToSingleLineString("|"));
                                Console.WriteLine($"{username} done.");
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                        {
                            Console.WriteLine($"{username} not found.");
                            notFound.Add(username);
                        }
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
            var fileName = $@"{outputFolder}\{name}_{DateTime.Now.Year}_{DateTime.Now.Month}_{DateTime.Now.Day}-{DateTime.Now.ToShortTimeString().Replace(":", "-")}.txt";
            File.WriteAllLines(fileName, lines.ToArray());
            return fileName;
        }

        private static DirectorySearcher CreateSearcher(DirectoryEntry directoryEntry, string firstName, string surname)
        {
            var searcher = new DirectorySearcher(directoryEntry)
            {
                PageSize = int.MaxValue,
                Filter = $"(&(objectCategory=person)(objectClass=user)(sn={surname}))"
            };

            searcher.PropertiesToLoad.Add("sAMAccountName");
            searcher.PropertiesToLoad.Add("givenName");
            searcher.PropertiesToLoad.Add("initials");
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
