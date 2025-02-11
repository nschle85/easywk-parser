namespace easywk_parser_dotnet;

public static class StarterLineHtmlFormatter
{
    public static string ToHtmlRow(StarterLine starterLine)
    {
        return $"""
                  <tr height=17 style='height:12.75pt'>
                   <td height=17 class=xlGeneral style='height:12.75pt'>{starterLine.Wettkampf}</td>
                   <td height=17 class=xlGeneral style='height:12.75pt'>{starterLine.Lauf}</td>
                   <td height=17 class=xlGeneral style='height:12.75pt'>{starterLine.Bahn}</td>
                   <td height=17 class=xlGeneral style='height:12.75pt'>{starterLine.Teilnehmer}</td>
                   <td height=17 class=xlGeneral style='height:12.75pt'>{starterLine.Verein}</td>
                   <td height=17 class=xlGeneral style='height:12.75pt'>{starterLine.Meldezeit}</td>
                 </tr>
                """;
    }
}