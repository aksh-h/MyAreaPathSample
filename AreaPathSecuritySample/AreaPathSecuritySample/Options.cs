using CommandLine;

namespace AddUserToAreaPath
{
    public class Options
    {
        [Option('a', "account", HelpText = "The url of the Azure Devops account like https://dev.azure.com/fabrikam", Required = true)]
        public string AccountUrl { get; set; }

        [Option('p', "project", HelpText = "The name of the project", Required = true)]
        public string ProjectName { get; set; }

        [Option('c', "area", HelpText = "The name of the area path. Do not enter path like parent/child, just the name like child", Required = true)]
        public string AreaPath { get; set; }

        [Option('g', "group", HelpText = "The name of the group that you want to add to the area path.", Required = true)]
        public string GroupName { get; set; }
    }
}
