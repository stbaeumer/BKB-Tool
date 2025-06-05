using CookComputing.XmlRpc;
using System.Text;
using Microsoft.Extensions.Configuration;

#pragma warning disable CS8603 // Mögliche Null-Verweis-Rückgabe
#pragma warning disable CS8602 // Dereferenzierung eines möglicherweise null-Objekts.
#pragma warning disable CS8604 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8620 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8600 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8618 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8619 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0219 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8625 // Möglicher Null-Verweis-Argument
#pragma warning disable CS8601 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0168 // Möglicher Null-Verweis-Argument
#pragma warning disable CS0618 // Möglicher Null-Verweis-Argument
#pragma warning disable NU1903 // Möglicher Null-Verweis-Argument
#pragma warning disable NU1902 // Möglicher Null-Verweis-Argument

public struct StructSchema
{
    public string name;
    public StructField[] fields;
}

public struct StructField
{
    public string name;
    public string label;
    public string type;
    public string multi;
    
}

public interface IDokuWikiApi : IXmlRpcProxy
{
    [XmlRpcMethod("struct.getData")]
    XmlRpcStruct GetData(string schemaName, string pageId);
    
    [XmlRpcMethod("struct.getSchema")]
    StructSchema GetSchema(string schemaName);
    
    [XmlRpcMethod("dokuwiki.getVersion")]
    string GetVersion();

    [XmlRpcMethod("wiki.getPage")]
    string GetPage(string page);

    [XmlRpcMethod("wiki.putPage")]
    bool PutPage(string page, string content, XmlRpcStruct options);
}

public class DokuwikiZugriff
{
    public DokuwikiZugriff(IConfiguration configuration)
    {
        Global.Konfig("WikiUrl", Global.Modus.Update, configuration, "URL zum dokuwiki xmlrpc.","",Global.Datentyp.Url);
        Global.Konfig("WikiJsonUser", Global.Modus.Update, configuration, "Benutzer, mit dem auf Json zugegriffen wird.");
        Global.Konfig("WikiJsonUserKennwort", Global.Modus.Update, configuration, "Kennwort");
        
        // Proxy erstellen
        Proxy = XmlRpcProxyGen.Create<IDokuWikiApi>();
        ((XmlRpcClientProtocol)Proxy).Url = Global.WikiUrl;

        // Manuelle HTTP-Header setzen
        var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{Global.WikiJsonUser}:{Global.WikiJsonUserKennwort}"));
        ((XmlRpcClientProtocol)Proxy).Headers.Add("Authorization", "Basic " + credentials);
    }

    public IDokuWikiApi Proxy { get; set; }
    public XmlRpcStruct Options { get; set; }

    public void GetVersion()
    {
        var version = Proxy.GetVersion();
        Console.WriteLine($"DokuWiki Version: {version}");
    }

    public string GetPage(string page)
    {
        var pageContent = Proxy.GetPage("start");
        //Console.WriteLine($"Seiteninhalt: {pageContent}");
        return pageContent;
    }
    
    public void PutPage(string page, string content)
    {
        Proxy.PutPage(page, content, new XmlRpcStruct());
        Console.WriteLine("Seite aktualisiert!");
    }
}