using System;

namespace JiraWorkLogReport
{
    internal class WorklogDto
    {
        public string IssueProject { get; set; }
        public string IssueKey { get; set; }
        public string IssueType { get; set; }
        public string IssueSummary { get; set; }
        public string Component { get; set; }
        public string Department { get; set; }
        public string Author { get; set; }
        public DateTime? StartDate { get; set; }
        public long MinutesSpent { get; set; }
        public double HoursSpent { get; internal set; }
        public int? Year { get; set; }
        public int? Month { get; set; }
        
    }
}