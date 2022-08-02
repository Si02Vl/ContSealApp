using AngleSharp.Dom;
using AngleSharp.Html.Parser;

namespace ContSealApp_Parse;

public class Class1
{
    public IEnumerable<string> AngleSharp()
    {
        List<string> hrefTags = new List<string>();

        var parser = new HtmlParser();
        var document = parser.Parse();
        foreach (IElement element in document.QuerySelectorAll("a"))
        {
            hrefTags.Add(element.GetAttribute("href"));
        }

        Console.WriteLine(hrefTags);
    }
}