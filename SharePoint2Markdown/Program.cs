using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using Figgle;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;

class Program
{
    // For data validation
    static string GetValidFolderName(string folderName)
    {
        string ret = Regex.Replace(folderName.Trim(), "[^A-Za-z0-9_ ]+", "");
        return ret.Replace(" ", string.Empty);
    }

    // For downloading assets from the Sharepoint site
    static void DownloadContent(string url, string path, string name)
    {
        using (WebClient client = new WebClient())
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Green;
                // Enable these two lines instead of client.Credentials = CredentialCache.DefaultCredentials if you wish to hard code credentials!
                //client.UseDefaultCredentials = true;
                //client.Credentials = new NetworkCredential("USER", "PASSWORD", "DOMAIN");
                client.Credentials = CredentialCache.DefaultCredentials;
                client.DownloadFile(new Uri(url), $@"{path}\{name}");
                Console.WriteLine($"Successfully downloaded {name} to {path}");

            }
            catch (WebException e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Web execption error! {e}");
            }
        }
    }

    static void Main(string[] args)
    {
        // Pointless flashy title, because...
        Console.WriteLine(FiggleFonts.Ogre.Render("SharePoint2Markdown"));

        Console.ForegroundColor = ConsoleColor.Magenta;

        // User defined input
        string exportPath = "exports";
        Console.Write("Site URL (eg. https://sharepointserver.your.domain): ");
        string siteURL = Console.ReadLine();

        try
        {
            // Starting with ClientContext, the constructor requires a URL to the
            // server running SharePoint.
            ClientContext context = new ClientContext(siteURL);

            // The SharePoint web at the URL.
            Web web = context.Web;

            // What list to search, this should search all posts on your site, however you may need to change depending on your site layout
            SP.List list = web.Lists.GetByTitle("Posts");

            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/>" +
                "<Value Type='Number'>0</Value></Geq></Where></Query></View>";
            ListItemCollection collListItem = list.GetItems(query);

            context.Load(collListItem);
            context.ExecuteQuery();

            // Create instance of converter
            var convertor = new ReverseMarkdown.Converter();

            foreach (ListItem item in collListItem)
            {
                string title = convertor.Convert(item["Title"].ToString());
                string body = convertor.Convert(item["Body"].ToString());
                FieldUserValue authorFUV = (FieldUserValue)item["Author"];
                string author = convertor.Convert(authorFUV.LookupValue);
                FieldLookupValue[] categoryFLV = (FieldLookupValue[])item["PostCategory"];
                List<string> categories = new List<string>();
                foreach (FieldLookupValue category in categoryFLV)
                {
                    categories.Add(convertor.Convert(category.LookupValue));
                }
                string created = convertor.Convert(item["Created"].ToString());
                string[] lines = body.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                List<string> urls = new List<string>();

                // Pattern for capturing URLS
                Regex regex = new Regex(@"\!\[([^\]]*)\]\(([^)]*)\)");

                // Loop through and add URLS to list
                foreach (Match match in regex.Matches(body))
                {
                    string value = match.Groups[2].Value;
                    if (value.StartsWith("/"))
                    {
                        urls.Add(value);
                        int pos = value.LastIndexOf("/") + 1;
                        string newURL = value.Substring(pos, value.Length - pos);
                        body = body.Replace(value, "/" + newURL);
                    }
                }

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Got post with title \"{0}\"", title);

                string titleStrip = GetValidFolderName(title);

                if (!Directory.Exists(exportPath))
                {
                    Directory.CreateDirectory(exportPath);
                }

                // Write content to markdown file
                using (StreamWriter outputFile = new StreamWriter(Path.Combine(exportPath, titleStrip + ".md"), true))
                {
                    // Header for markdown properties, this is tailored for Wiki JS so you may change as needed
                    string header = $"---\ntitle: {title}\ndescription:\npublished: true\ndate: {created}\ntags: ";
                    int iteration = 0;
                    foreach (string category in categories)
                    {
                        iteration++;
                        header += category;
                        if (iteration < categories.Count)
                        {
                            header += ", ";
                        }
                    }
                    header += $"\neditor: {author}\ndateCreated: {created}\n---";
                    outputFile.WriteLine(header);
                    outputFile.WriteLine(body);
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Successfully wrote post body to the markdown file.");
                }

                Console.WriteLine("Checking through post content...");

                // Download assets using URLS captured above
                if (urls.Count() != 0)
                {
                    foreach (string url in urls)
                    {
                        string downloadURL = siteURL + url;
                        string fileName = url.Split('/').Last();
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("Content with name {0} found, downloading...", fileName);
                        DownloadContent(downloadURL, exportPath, fileName);
                    }
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("No content found!");
                }
            }
            Console.ResetColor();
            Console.WriteLine("\nEnd.");
            Console.ReadLine();
        }
        catch (DirectoryNotFoundException e)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Directory not found error! {e}");
            Console.ReadLine();
        }
        catch (WebException e)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error with site URL! {e}");
            Console.ReadLine();
        }
        catch (ArgumentException e)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error with connection! Is your site URL correct? {e}");
            Console.ReadLine();
        }
    }
}