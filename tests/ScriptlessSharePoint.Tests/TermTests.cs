using System.Xml.Linq;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class TermTests
{
    [Theory]
    [InlineData("Europe/Denmark", new[] { "Europe", "Denmark" })]
    [InlineData(" Europe / Denmark / Copenhagen ", new[] { "Europe", "Denmark", "Copenhagen" })]
    [InlineData("Europe;Denmark", new[] { "Europe", "Denmark" })]
    [InlineData("Asia", new[] { "Asia" })]
    [InlineData(" / ", new string[0])]
    public void Term_paths_split_at_slashes(string typed, string[] parts) =>
        Assert.Equal(parts, CommandCatalog.TermPathOf(typed));

    [Theory]
    [InlineData("Research & Development", "research ＆ development", true)]
    [InlineData("Sales  team", "Sales team", true)]
    [InlineData("\"Quoted\"", "＂quoted＂", true)]
    [InlineData("Sales", "Marketing", false)]
    public void Term_names_match_as_the_term_store_keeps_them(string a, string b, bool same) =>
        Assert.Equal(same, CommandCatalog.SameTerm(a, b));

    [Fact]
    public void Managed_metadata_columns_come_with_their_note_column()
    {
        var (noteXml, fieldXml) = CommandCatalog.TaxonomyFieldXml("Country & region", multi: true, required: true, "Contoso Columns");
        var note = XElement.Parse(noteXml);
        var field = XElement.Parse(fieldXml);
        Assert.Equal("Note", (string?)note.Attribute("Type"));
        Assert.Equal("Countryregion" + "TaxHTField0", (string?)note.Attribute("StaticName"));
        Assert.Equal("TRUE", (string?)note.Attribute("Hidden"));
        Assert.Equal("TaxonomyFieldTypeMulti", (string?)field.Attribute("Type"));
        Assert.Equal("TRUE", (string?)field.Attribute("Mult"));
        Assert.Equal("Country & region", (string?)field.Attribute("DisplayName"));
        Assert.Equal("Contoso Columns", (string?)field.Attribute("Group"));
        // the column's text goes in the note column
        var textField = field.Descendants("Property").Single(p => (string?)p.Element("Name") == "TextField").Element("Value")!.Value;
        Assert.Equal((string?)note.Attribute("ID"), textField);

        var single = XElement.Parse(CommandCatalog.TaxonomyFieldXml("Topic", false, false, null).Field);
        Assert.Equal("TaxonomyFieldType", (string?)single.Attribute("Type"));
        Assert.Null(single.Attribute("Mult"));
        Assert.Null(single.Attribute("Group"));
    }

    [Fact]
    public void Term_store_commands_are_the_tenants_and_deleting_asks_first()
    {
        var terms = CommandCatalog.All.Where(c => c.Group == "Term store").ToList();
        Assert.True(terms.Count >= 10);
        Assert.All(terms, c => Assert.True(c.AppliesTo(NodeKind.Tenant), c.Id));
        Assert.All(terms.Where(c => c.Id.Contains("delete")), c => Assert.True(c.Destructive, c.Id));
        Assert.True(CommandCatalog.Find("list.add-term-column")!.AppliesTo(NodeKind.Library));
        Assert.True(CommandCatalog.Find("content.set-term")!.AppliesTo(NodeKind.Item));
    }
}
