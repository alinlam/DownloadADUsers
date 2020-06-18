using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Dynamic;
using System.Linq;
using System.Text;

namespace ADUsers
{
    public class ADUser
    {
        public string UserPrincipalName { get; }

        public Dictionary<string, object> Properties { get; }

        public ADUser(string userPrincipalName, Dictionary<string, object> properties)
        {
            this.Properties = properties;
            this.UserPrincipalName = userPrincipalName;
        }
    }

    public class Services
    {
        private string _ldap;
        private string _username;
        private string _password;
        public Services(string ldap, string username, string password)
        {
            this._ldap = ldap;
            this._username = username;
            this._password = password;
        }
        public IEnumerable<ADUser> GetUsers()
        {
            int pageNumber = 1;
            int pageSize = 100;
            var results = new List<ADUser>();

            while (true)
            {
                Console.WriteLine($"Getting {pageSize} users of page #{pageNumber}");

                var skipRows = (pageNumber - 1) * pageSize;
                var users = this.GetUsersInOU(this._ldap, this._username, this._password, pageSize, skipRows, (skipRows + pageSize));
                if (users.Count == 0)
                {
                    break;
                }

                results.AddRange(users);
                Console.WriteLine($"Got {users.Count} users of page #{pageNumber}, up the total to {results.Count} records");

                if (users.Count < pageSize)
                {
                    Console.WriteLine($"Finished downloading process");
                    break;
                }

                // Continue next page
                pageNumber++;
            }

            return results;
        }

        private List<ADUser> GetUsersInOU(string ldap, string username, string password, int pageSize, int fromIndex, int toIndex)
        {
            Console.WriteLine($"Accessing {ldap}");

            string filter = $"(&(objectCategory=person)(objectClass=user)(|(givenName=*)(sn=*)))";
            var usersFound = new List<ADUser>();
            

            using (DirectoryEntry dirEntry = new DirectoryEntry(ldap, username, password))
            {
                using (DirectorySearcher dirSearcher = new DirectorySearcher(dirEntry, filter, null))
                {
                    dirSearcher.Sort = new SortOption("cn", SortDirection.Ascending);
                    dirSearcher.PageSize = pageSize;
                    dirSearcher.SearchScope = SearchScope.Subtree;
                    dirSearcher.PropertiesToLoad.Add(string.Empty);                    

                    using (var searchResults = dirSearcher.FindAll())
                    {
                        for (int i = fromIndex; i < toIndex; i++)
                        {
                            SearchResult result = null;
                            try
                            {
                                result = searchResults[i];
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex);
                                break;
                            }

                            using (var de = result.GetDirectoryEntry())
                            {
                                var properties = new Dictionary<string, object>();
                                foreach (string key in de.Properties.PropertyNames)
                                {
                                    properties.Add(key, de.Properties[key]);
                                }

                                var userPrincipalName = (de.Properties["userPrincipalName"] as PropertyValueCollection).Value as string;

                                usersFound.Add(new ADUser(userPrincipalName, properties));

                                Console.WriteLine($"Got Directory Entry for {userPrincipalName}, found: {properties.Count} properties");
                            }
                        }                        
                    }

                    dirEntry.Close();
                    dirEntry.Dispose();
                }
            }

            return usersFound;
        }
    }
}
