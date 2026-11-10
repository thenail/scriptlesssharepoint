using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class CommandCatalogTests
{
    [Fact]
    public void Ids_are_unique()
    {
        var duplicates = CommandCatalog.All.GroupBy(c => c.Id).Where(g => g.Count() > 1).Select(g => g.Key);
        Assert.Empty(duplicates);
    }

    [Fact]
    public void Every_kind_of_row_has_commands_beyond_reading()
    {
        foreach (var kind in Enum.GetValues<NodeKind>())
        {
            var commands = CommandCatalog.For(kind);
            Assert.True(commands.Count(c => !c.ReadOnly) >= 3, $"{kind} has too few commands");
        }
    }

    [Fact]
    public void Commands_are_described_and_grouped()
    {
        foreach (var c in CommandCatalog.All)
        {
            Assert.False(string.IsNullOrWhiteSpace(c.Title), c.Id);
            Assert.False(string.IsNullOrWhiteSpace(c.Description), c.Id);
            Assert.Contains(c.Group, CommandCatalog.GroupOrder);
            Assert.NotEmpty(c.Targets);
            Assert.Equal(c.Parameters.Count, c.Parameters.Select(p => p.Name).Distinct().Count());
            Assert.True(c.Glyph.Length == 1 && c.Glyph[0] >= '', $"{c.Id}: the glyph is one icon character");
        }
    }

    [Fact]
    public void Choices_have_options_and_their_defaults_are_among_them()
    {
        foreach (var p in CommandCatalog.All.SelectMany(c => c.Parameters))
        {
            if (p.Kind == ParameterKind.Choice)
            {
                Assert.NotEmpty(p.Choices);
                Assert.True(p.Default is null || p.Choices.Any(c => c.Value == p.Default), $"{p.Name}: default {p.Default}");
            }
            if (p.Default is { } d)
            {
                Assert.Null(p.Check(d));
            }
        }
    }

    [Fact]
    public void Things_that_delete_ask_first()
    {
        foreach (var c in CommandCatalog.All.Where(c => !c.ReadOnly && (c.Id.Contains("delete") || c.Id.Contains("purge") || c.After == After.Remove && c.Id != "content.move")))
        {
            Assert.True(c.Destructive, c.Id);
        }
        Assert.All(CommandCatalog.All.Where(c => c.ReadOnly), c => Assert.False(c.Destructive, c.Id));
    }

    [Fact]
    public void The_pane_lists_commands_in_group_order()
    {
        var groups = CommandCatalog.For(NodeKind.SiteCollection).Select(c => c.Group).Distinct().ToList();
        var order = groups.Select(g => CommandCatalog.GroupOrder.ToList().IndexOf(g)).ToList();
        Assert.Equal(order.OrderBy(i => i), order);
        Assert.Equal("Open", groups[0]);
        Assert.Equal("Delete", groups[^1]);
    }

    [Theory]
    [InlineData(ParameterKind.Number, "12", true)]
    [InlineData(ParameterKind.Number, "twelve", false)]
    [InlineData(ParameterKind.Number, "0", false)]
    [InlineData(ParameterKind.Url, "https://contoso.sharepoint.com/sites/a", true)]
    [InlineData(ParameterKind.Url, "http://contoso.sharepoint.com", false)]
    [InlineData(ParameterKind.Toggle, "true", true)]
    [InlineData(ParameterKind.Toggle, "maybe", false)]
    [InlineData(ParameterKind.Text, "", false)]
    public void Parameters_check_their_values(ParameterKind kind, string value, bool ok)
    {
        var p = new CommandParameter("x", "X", kind) { Min = 1, Max = 100 };
        Assert.Equal(ok, p.Check(value) is null);
    }

    [Fact]
    public void Optional_parameters_may_be_empty()
    {
        Assert.Null(new CommandParameter("x", "X", ParameterKind.Text) { Required = false }.Check(" "));
        Assert.Null(new CommandParameter("x", "X", ParameterKind.Toggle).Check(null));
    }

    [Fact]
    public async Task The_runner_checks_the_inputs_first_and_runs_only_on_kinds_it_applies_to()
    {
        var calls = new List<string>();
        var command = new Command
        {
            Id = "test",
            Title = "Test",
            Group = "Settings",
            Description = "A test.",
            Targets = [NodeKind.Library],
            Parameters = [new("name", "Name", ParameterKind.Text)],
            Run = (c, n) =>
            {
                lock (calls)
                {
                    calls.Add($"{n.Title}={c.Args.Text("name")}");
                }
                return n.Title == "Bad" ? throw new InvalidOperationException("nope") : Task.FromResult(CommandResult.Done("ok"));
            },
        };
        var nodes = new[]
        {
            Node(NodeKind.Library, "Docs"),
            Node(NodeKind.Library, "Bad"),
            Node(NodeKind.List, "Tasks"),
            Node(NodeKind.Library, "Docs"), // the same one twice: once
        };
        var client = new SharePointClient(new Connections.ConnectionProfile { TenantUrl = "https://contoso.sharepoint.com" }, new NoTokens(), false);

        await Assert.ThrowsAsync<ArgumentException>(() => CommandRunner.RunAsync(command, nodes, new CommandArgs(), client, new NoHost()));

        var args = new CommandArgs { ["name"] = "x" };
        var results = await CommandRunner.RunAsync(command, nodes, args, client, new NoHost());
        Assert.Equal(2, results.Count);
        Assert.True(results.Single(r => r.Target.Title == "Docs").Result.Ok);
        Assert.Equal("nope", results.Single(r => r.Target.Title == "Bad").Result.Message);
        Assert.Equal(["Bad=x", "Docs=x"], calls.Order());
    }

    [Fact]
    public void Destinations_take_a_folder_address_as_the_browser_shows_it()
    {
        Assert.Equal("https://contoso.sharepoint.com/sites/a/Shared Documents/2024/a.docx",
            Csom.Destination("https://contoso.sharepoint.com/sites/a/Shared%20Documents/2024/", "a.docx"));
        Assert.Equal("https://contoso.sharepoint.com/sites/a/Shared Documents/a.docx",
            Csom.Destination("https://contoso.sharepoint.com/sites/a/Shared%20Documents/Forms/AllItems.aspx", "a.docx"));
        Assert.Equal("/sites/a/Docs/new.txt", Csom.WithLeaf("/sites/a/Docs/old.txt", "new.txt"));
    }

    [Fact]
    public void Downloads_never_overwrite()
    {
        var dir = Directory.CreateTempSubdirectory("ssp-download-").FullName;
        try
        {
            Assert.Equal(Path.Combine(dir, "a.txt"), CommandCatalog.FreeName(dir, "a.txt"));
            File.WriteAllText(Path.Combine(dir, "a.txt"), "");
            Assert.Equal(Path.Combine(dir, "a (2).txt"), CommandCatalog.FreeName(dir, "a.txt"));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact]
    public void Nodes_know_their_names_and_keys()
    {
        var file = new SpNode
        {
            Kind = NodeKind.File,
            Title = "Plan.docx",
            Url = "https://contoso.sharepoint.com/sites/a/Docs/Plan.docx",
            WebUrl = "https://contoso.sharepoint.com/sites/a",
            ServerRelativeUrl = "/sites/a/Docs/Plan.docx",
        };
        Assert.Equal("Plan.docx", file.LeafName);
        Assert.Equal("/sites/a/Docs", file.ParentFolder);
        Assert.Equal(file.Key, (file with { Url = file.Url.ToUpperInvariant() }).Key);
        Assert.Equal("3 libraries", NodeKinds.Count(NodeKind.Library, 3));
        Assert.Equal("1 site collection", NodeKinds.Count(NodeKind.SiteCollection, 1));
    }

    private static SpNode Node(NodeKind kind, string title) => new()
    {
        Kind = kind,
        Title = title,
        Url = $"https://contoso.sharepoint.com/sites/a/{title}",
        WebUrl = "https://contoso.sharepoint.com/sites/a",
        ListId = Guid.NewGuid(),
    };

    private sealed class NoHost : ICommandHost
    {
        public void OpenUrl(string url) { }
        public void CopyText(string text) { }
    }

    private sealed class NoTokens : Auth.ITokenProvider
    {
        public Task<string> GetTokenAsync(string resource, bool interactive, CancellationToken cancellationToken = default) => Task.FromResult("x");
        public Task SignOutAsync() => Task.CompletedTask;
    }
}
