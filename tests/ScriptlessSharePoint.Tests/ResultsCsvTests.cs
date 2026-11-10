using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.History;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class ResultsCsvTests
{
    private static readonly SpNode Hr = SpNode.FromSite(new SiteSummary("https://contoso.sharepoint.com/sites/hr", "Human Resources, Team", "GROUP#0", null, null));
    private static readonly SpNode Legal = SpNode.FromSite(new SiteSummary("https://contoso.sharepoint.com/sites/legal", "Legal", "GROUP#0", null, null));

    [Fact]
    public void A_report_gets_a_row_per_line_with_its_parts_in_columns()
    {
        var csv = ResultsCsv.Build(
        [
            new(null, Hr, RunOutcome.Done, "2 people.", "Ada Lovelace  ·  ada@contoso.com  ·  Owner\r\nAlan Turing  ·  alan@contoso.com"),
            new(null, Legal, RunOutcome.Failed, "Access denied.", null),
        ]);
        var lines = csv.Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal("Name,Address,Result,Message,Detail 1,Detail 2,Detail 3", lines[0]);
        Assert.Equal("\"Human Resources, Team\",https://contoso.sharepoint.com/sites/hr,Done,2 people.,Ada Lovelace,ada@contoso.com,Owner", lines[1]);
        Assert.Equal("\"Human Resources, Team\",https://contoso.sharepoint.com/sites/hr,Done,2 people.,Alan Turing,alan@contoso.com,", lines[2]);
        Assert.Equal("Legal,https://contoso.sharepoint.com/sites/legal,Failed,Access denied.,,,", lines[3]);
    }

    [Fact]
    public void A_reports_table_names_its_columns_and_reads_as_lines()
    {
        var table = ReportTable.Of(["Name", "Email or sign-in", "Administrator", "Guest"],
            [["Ada Lovelace", "ada@contoso.com", "administrator", ""], ["Bob", "bob_x#ext#", null, "guest"]]);
        var result = CommandResult.Report("2 people.", table);
        Assert.Equal($"Ada Lovelace  ·  ada@contoso.com  ·  administrator{Environment.NewLine}Bob  ·  bob_x#ext#  ·  guest", result.Detail);

        var lines = ResultsCsv.Build([new(null, Hr, RunOutcome.Done, result.Message, result.Detail, result.Table)]).Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal("Name,Address,Result,Message,Name,Email or sign-in,Administrator,Guest", lines[0]);
        Assert.EndsWith(",Bob,bob_x#ext#,,guest", lines[2]);
    }

    [Fact]
    public void Details_of_several_targets_line_up_by_column_name()
    {
        var hr = ReportTable.Of(["Title", "Owner", "Storage used"], [["HR", "ada@contoso.com", "4 GB of 1 TB"]]);
        var legal = ReportTable.Of(["Title", "Storage used"], [["Legal", "1 GB of 1 TB"]]);
        Assert.Equal(["Title", "Owner", "Storage used"], ReportTable.Union([hr, legal]));
        Assert.Equal(["Legal", "", "1 GB of 1 TB"], legal.Cells(legal.Rows[0], ReportTable.Union([hr, legal])));

        var lines = ResultsCsv.Build([new(null, Hr, RunOutcome.Done, "Details below.", null, hr), new(null, Legal, RunOutcome.Done, "Details below.", null, legal)])
            .Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal("Name,Address,Result,Message,Title,Owner,Storage used", lines[0]);
        Assert.Equal("Legal,https://contoso.sharepoint.com/sites/legal,Done,Details below.,Legal,,1 GB of 1 TB", lines[2]);
    }

    [Fact]
    public void Lines_without_columns_become_details()
    {
        var table = ReportTable.FromLines("a  ·  b\r\n\r\nc");
        Assert.Equal(["Detail 1", "Detail 2"], table.Columns);
        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(["c"], table.Rows[1]);
    }

    [Fact]
    public void A_plans_results_have_the_step_and_formulas_are_kept_as_text()
    {
        var csv = ResultsCsv.Build([new(2, Legal with { Title = "=HYPERLINK(\"x\")" }, RunOutcome.Unchanged, "Already on.", null)]);
        var lines = csv.Split("\r\n", StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal("Step,Name,Address,Result,Message", lines[0]);
        Assert.Equal("2,\"'=HYPERLINK(\"\"x\"\")\",https://contoso.sharepoint.com/sites/legal,Unchanged,Already on.", lines[1]);
    }
}
