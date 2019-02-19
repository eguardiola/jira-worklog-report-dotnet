using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Atlassian.Jira;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace JiraWorkLogReport
{
    class JiraWorkLogReport
    {
        static void Main(string[] args)
        {
            BuildJiraWorkLogReport().GetAwaiter().GetResult();
        }

        private static async Task BuildJiraWorkLogReport()
        {
            DateTime fromDate = new DateTime(2019, 2, 1);
            DateTime toDate = new DateTime(2019, 2, 19);

            var jira = Jira.CreateRestClient("https://jira.example.com", "username", "*****");
            var issues = await jira.Issues.GetIssuesFromJqlAsync($"worklogDate >= {fromDate:yyyy-M-d} and worklogdate <= {toDate:yyyy-M-d}", int.MaxValue, 0);

            var usersByUserNameMap = new Dictionary<string, JiraUser>();
            var worklogDtos = new List<WorklogDto>();

            foreach (var issue in issues)
            {
                var worklogs = await issue.GetWorklogsAsync();
                var component = issue.Components.FirstOrDefault()?.Name;
                var department = issue.CustomFields["Departamento"]?.Values.FirstOrDefault();

                foreach (var worklog in worklogs)
                {
                    if (worklog.StartDate?.Date >= fromDate && worklog.StartDate?.Date <= toDate)
                    {
                        if (!usersByUserNameMap.ContainsKey(worklog.Author))
                            usersByUserNameMap.Add(worklog.Author, await jira.Users.GetUserAsync(worklog.Author));

                        worklogDtos.Add(new WorklogDto
                        {
                            IssueProject = issue.Project,
                            IssueKey = issue.Key.Value,
                            IssueType = issue.Type.Name,
                            IssueSummary = issue.Summary,
                            Component = component,
                            Department = department,
                            Author = usersByUserNameMap[worklog.Author].DisplayName,
                            StartDate = worklog.StartDate,
                            MinutesSpent = worklog.TimeSpentInSeconds / 60,
                            HoursSpent = worklog.TimeSpentInSeconds / (double)(60 * 60),
                            Year = worklog.StartDate?.Year,
                            Month = worklog.StartDate?.Month
                        });

                    }
                }
            }

            var pck = new ExcelPackage();
            ExcelWorksheet dataSheet = pck.Workbook.Worksheets.Add("Data");
            dataSheet.Cells["A1"].LoadFromCollection(worklogDtos.OrderBy(c => c.StartDate), true, TableStyles.Light10);
            dataSheet.Cells[dataSheet.Dimension.Address].AutoFitColumns();
            dataSheet.Column(8).Style.Numberformat.Format = "yyyy-mm-dd h:mm";
            dataSheet.Column(9).Style.Numberformat.Format = "0.00";
            dataSheet.Column(10).Style.Numberformat.Format = "0.00";

            File.WriteAllBytes($"c:\\JiraWorklogReport {fromDate:yyyy-M-d} - {toDate:yyyy-M-d}.xlsx", pck.GetAsByteArray());
        }
    }
}
