using CommonResources.Models;
using EnvDTE;
using System.Collections.Generic;
using System.Linq;

namespace CommonResources
{
    public static class SharedWindow
    {
        const string SolutionFolder = "{66A26720-8FB5-11D2-AA7E-00C04F688DDE}";

        public static void OpenCrmPage(string url, CrmConn selectedConnection, DTE dte)
        {
            if (selectedConnection == null) return;
            string connString = selectedConnection.ConnectionString;
            if (string.IsNullOrEmpty(connString)) return;

            string[] connParts = connString.Split(';');
            string urlPart = connParts.FirstOrDefault(s => s.ToUpper().StartsWith("URL="));
            if (!string.IsNullOrEmpty(urlPart))
            {
                string[] urlParts = urlPart.Split('=');
                string baseUrl = (urlParts[1].EndsWith("/")) ? urlParts[1] : urlParts[1] + "/";

                var props = dte.Properties["CRM Developer Extensions", "General"];
                bool useDefaultWebBrowser = (bool)props.Item("UseDefaultWebBrowser").Value;

                if (useDefaultWebBrowser) //User's default browser
                    System.Diagnostics.Process.Start(baseUrl + url);
                else //Internal VS browser
                    dte.ItemOperations.Navigate(baseUrl + url);
            }
        }

        public static IEnumerable<Project> GetProjects(Projects projects)
        {
            var list = new List<Project>();
            var item = projects.GetEnumerator();

            while (item.MoveNext())
            {
                var project = item.Current as Project;

                if (project == null) continue;

                if (project.Kind.ToUpper() == SolutionFolder)
                    list.AddRange(GetFolderProjects(project));
                else
                    list.Add(project);
            }
            
            return list.OrderBy(p => p.Name);
        }

        private static IEnumerable<Project> GetFolderProjects(Project folder)
        {
            var list = new List<Project>();

            foreach (ProjectItem item in folder.ProjectItems)
            {
                var subProject = item.SubProject;

                if (subProject == null) continue;

                if (subProject.Kind.ToUpper() == SolutionFolder)
                    list.AddRange(GetFolderProjects(subProject));
                else
                    list.Add(subProject);
            }

            return list;
        }
    }
}
