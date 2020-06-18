using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;

namespace ADUsers
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("This console application will download all AD accounts and the properties");
            Console.WriteLine("Please input following information");
            
            Console.Write("LDAP path to the OU: ");
            var ldap = Console.ReadLine();

            Console.Write("Username: ");
            var username = Console.ReadLine();

            Console.Write("Password: ");
            var password= Console.ReadLine();

            Console.WriteLine("Start downloading process...");

            var service = new Services(ldap, username, password);
            var users = service.GetUsers();

            var now = DateTime.Now;
            var fileName = $"allUsers_{now.Year}{now.Month}{now.Day}{now.Hour}{now.Minute}{now.Second}.json";
            //FileStream fs = new FileStream(fileName, FileMode.Create);
            
            using (StreamWriter sw = new StreamWriter(@$"{fileName}", true))
            {
                //Console.SetOut(sw);

                //Console.WriteLine("[");
                sw.WriteLine("[");

                var lastUser = users.Last();
                foreach (var user in users)
                {
                    //Console.WriteLine("\t{");
                    //Console.WriteLine($"\t\"UserPrincipalName\": \"{user.UserPrincipalName}\",");
                    //Console.WriteLine($"\t\"Properties\":");
                    //Console.WriteLine("\t\t[");

                    sw.WriteLine("\t{");
                    sw.WriteLine($"\t\"UserPrincipalName\": \"{user.UserPrincipalName}\",");
                    sw.WriteLine($"\t\"Properties\":");
                    sw.WriteLine("\t\t[");

                    var lastProp = user.Properties.Last();
                    foreach (var prop in user.Properties)
                    {
                        if (prop.Key == lastProp.Key)
                        {
                            //Console.WriteLine($"\t\t\t{JsonConvert.SerializeObject(prop)}");
                            sw.WriteLine($"\t\t\t{JsonConvert.SerializeObject(prop)}");
                        }
                        else
                        {
                            //Console.WriteLine($"\t\t\t{JsonConvert.SerializeObject(prop)},");
                            sw.WriteLine($"\t\t\t{JsonConvert.SerializeObject(prop)},");
                        }
                    }

                    //Console.WriteLine("\t\t]");
                    sw.WriteLine("\t\t]");

                    if (user.UserPrincipalName == lastUser.UserPrincipalName)
                    {
                        //Console.WriteLine("\t}");
                        sw.WriteLine("\t}");
                    }
                    else
                    {
                        //Console.WriteLine("\t},");
                        sw.WriteLine("\t},");
                    }

                }
                //Console.WriteLine("]");
                sw.WriteLine("]");
            }           

            //TextWriter tmp = Console.Out;
            //Console.SetOut(tmp);
            Console.WriteLine($"Completed downloaded {users.Count()} users, data is saved in file {fileName}");

            Console.WriteLine("Press any key to exit the program");
            Console.ReadKey();
        }
    }
}
