# AreaPathSecuritySample

Sample C# console app for adding a group to a area path and giving the group read/write perms to the work items in the area path.

```bash
.\AreaPathSecuritySample.exe --help

  -a, --account    Required. The url of the Azure Devops account like https://dev.azure.com/fabrikam

  -p, --project    Required. The name of the project

  -c, --area       Required. The name of the area path. Do not enter path like parent/child, just the name like child

  -g, --group      Required. The name of the group that you want to add to the area path.

  --help           Display this help screen.

  --version        Display version information.

Usage: AreaPathSecuritySample.exe -a yourAccountUrl -p yourPojectName -c yourAreaPathName -g yourGroupName
```
