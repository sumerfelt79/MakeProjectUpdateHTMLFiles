using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Windows.Forms;

namespace MakeProjectUpdateHTMLFiles
{
    class Program
    {
        static void Main(string[] args)
        {
            List<string> jobsList = new List<string>();
            using (ClientContext projectUpdatesList = new ClientContext("https://sharepoint.wilsonconst.com/PWA"))
            {
                List jobs = projectUpdatesList.Web.Lists.GetByTitle("Project Updates");
                CamlQuery query = new CamlQuery();
                query.ViewXml = "<View><Query><OrderBy><FieldRef Name='Title' Ascending='True' /></OrderBy></Query></View>";
                ListItemCollection job = jobs.GetItems(query);
                projectUpdatesList.Load(job);
                projectUpdatesList.ExecuteQuery();
                foreach (ListItem j in job)
                {
                    jobsList.Add(j["Title"].ToString());
                }
            }
            makeHTMLFiles(jobsList);
        }
        static void makeHTMLFiles(List<string> jobsList)
        {
            var directories = Directory.GetDirectories("P:");
            List<string> jobNumbers = new List<string>();
            foreach (string f in directories)
            {
                string dirName = f.Replace("P:", "");
                jobNumbers.Add(dirName);
            }
            foreach (string job in jobsList)
            {
                string jn = jobNumbers.Find(j => j == job);
                if (jn != null)
                {
                    string fileName = string.Concat(job, " - Project Updates.html");
                    string puList = string.Concat(@"P:\", job, "\\Lists\\Project Updates\\", job, " - Project Update.html");
                    FileInfo fi = new FileInfo(puList);
                    if (!fi.Exists)
                    {
                        string meta = string.Concat("<meta http-equiv=\"refresh\" content=\"0; url=https://sharepoint.wilsonconst.com/PWA/", job, "/Lists/Project%20Updates/Top%20View.aspx\" />");
                        string lines = string.Concat("<html>\r\n<head>\r\n", meta, "\r\n</head>\r\n...\r\n</html>");
                        System.IO.StreamWriter file = new System.IO.StreamWriter(puList, true);
                        file.WriteLine(lines);
                        file.Close();
                    }
                }
            }
        }
    }
}
