using System;
using System.DirectoryServices;

namespace LdapQuery
{
    internal class User
    {
        public string Username { get; set; }
        public string FirstName { get; set; }
        public string Initials { get; set; }
        public string LastName { get; set; }
        public string TelephoneNumber { get; set; }
        public string Mobile { get; set; }
        public string Office { get; set; }
        public string Title { get; set; }
        public string Department { get; set; }
        public string Manager { get; set; }
        public string Email { get; set; }
        public string SearchedText { get; set; }
        public string Name { get; set; }
        public string Street { get; set; }
        public string PoBox { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string PostCode { get; set;}
        public string Country { get; set; }

        public static User FromDirectorySearchResult(SearchResult result, string username)
        {
            return new User
            {
                Username = GetProperty(result, "sAMAccountName"),
                FirstName = GetProperty(result, "givenName"),
                LastName = GetProperty(result, "sn"),
                TelephoneNumber = GetProperty(result, "telephoneNumber"),
                Mobile = GetProperty(result, "mobile"),
                Office = GetProperty(result, "physicalDeliveryOfficeName"),
                Title = GetProperty(result, "title"),
                Department = GetProperty(result, "department"),
                Manager = GetProperty(result, "manager"),
                Email = GetProperty(result, "mail"),
                SearchedText = username,
                Name = GetProperty(result, "name"),
                Street = GetProperty(result, "streetAddress"),
                PoBox = GetProperty(result, "postOfficeBox"),
                City = GetProperty(result, "l"),
                State = GetProperty(result, "st"),
                PostCode = GetProperty(result, "postalCode"),
                Country = GetProperty(result, "co")
            };
        }

        public string ToSingleLineString(string delimeter)
        {
            var d = "|";

            if (!string.IsNullOrEmpty(delimeter))
                d = delimeter;

            return $"{Username}{d}{FirstName}{d}{LastName}{d}{TelephoneNumber}{d}" +
                $"{Mobile}{d}{Title}{d}{Department}{d}{Manager}" +
                $"{d}{Email}{d}{Office}{d}{PoBox}{d}{Street}{d}{City}{d}{State}{d}{PostCode}{d}{Country}";
        }

        public static string GetProperty(SearchResult result, string property)
        {
            if (property == "manager" && result.Properties.Contains(property))
            {
                var mgr = result.Properties[property][0].ToString();
                var endpoint = mgr.IndexOf(",", StringComparison.Ordinal) - 3;
                return mgr.Substring(3, endpoint);
            }

            if (property == "initials" && result.Properties.Contains(property))
            {
                return result.Properties[property][0].ToString();
            }

            if (result.Properties.Contains(property))
            {
                var value = result.Properties[property][0].ToString();
                return string.IsNullOrEmpty(value) ? "Not Found" : value;
            }

            return "Not Found";
        }
    }
}
