using System.Xml.Linq;

using ScriptlessSharePoint.Commands;
using ScriptlessSharePoint.SharePoint;

namespace ScriptlessSharePoint.Tests;

public sealed class ColumnTests
{
    [Fact]
    public void Column_xml_has_the_name_type_and_group()
    {
        var x = XElement.Parse(CommandCatalog.FieldXml("Project code & phase", "Text", null, true, "Contoso Columns"));
        Assert.Equal("Text", (string?)x.Attribute("Type"));
        Assert.Equal("Project code & phase", (string?)x.Attribute("DisplayName"));
        Assert.Equal("Projectcodephase", (string?)x.Attribute("Name"));
        Assert.Equal("Projectcodephase", (string?)x.Attribute("StaticName"));
        Assert.Equal("TRUE", (string?)x.Attribute("Required"));
        Assert.Equal("Contoso Columns", (string?)x.Attribute("Group"));
        Assert.True(Guid.TryParse((string?)x.Attribute("ID"), out _));
    }

    [Fact]
    public void Choice_columns_list_their_choices_and_list_columns_have_no_group()
    {
        var x = XElement.Parse(CommandCatalog.FieldXml("Status", "Choice", "Open\r\n Closed \n\n<Other>", false));
        Assert.Equal(["Open", "Closed", "<Other>"], x.Descendants("CHOICE").Select(c => c.Value));
        Assert.Null(x.Attribute("Group"));
        Assert.Equal("FALSE", (string?)x.Attribute("Required"));
    }

    [Theory]
    [InlineData("2024 budget", "Column2024budget")]
    [InlineData("!!!", "Column")]
    public void Internal_names_start_with_a_letter(string name, string internalName) =>
        Assert.Equal(internalName, (string?)XElement.Parse(CommandCatalog.FieldXml(name, "Text", null, false)).Attribute("Name"));

    [Fact]
    public void New_content_types_go_under_their_parent()
    {
        var id = CommandCatalog.ChildContentTypeId("0x0101");
        Assert.StartsWith("0x010100", id);
        Assert.Equal(4 + 2 + 32, id.Length - 2);
        Assert.Matches("^0x[0-9A-F]+$", id);
        Assert.NotEqual(id, CommandCatalog.ChildContentTypeId("0x0101"));
    }

    [Fact]
    public void Column_and_content_type_commands_share_a_group()
    {
        var group = CommandCatalog.All.Where(c => c.Group == "Columns and content types").Select(c => c.Id).ToList();
        Assert.Contains("list.add-column", group);
        Assert.Contains("web.add-column", group);
        Assert.Contains("web.new-content-type", group);
        Assert.Contains("list.content-types", group);
        Assert.True(CommandCatalog.For(NodeKind.Library).Any(c => c.Id == "list.add-site-column"));
    }
}
