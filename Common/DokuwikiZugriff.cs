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

    // NEU: Holt die Tabellendaten (Aggregationen) des Struct-Plugins
    [XmlRpcMethod("plugin.struct.getAggregationData")]
    
    object[] GetAggregationData(string[] schemas, string[] columns, object[] filters, string sortBy);

    [XmlRpcMethod("plugin.struct.saveData")]
    bool SaveStructData(string page, object data, string summary = "Updated via API");
    
    [XmlRpcMethod("plugin.struct.addGlobalRow")]
    object AddGlobalRow(XmlRpcStruct data, string summary = "");

    // NEU: Ermöglicht das Abfragen aller registrierten API-Endpunkte des Servers
    [XmlRpcMethod("system.listMethods")]
    string[] ListMethods();

    // Testweise das Interface anpassen, falls die API ein Array/Objekt-Wrapper benötigt:
[XmlRpcMethod("plugin.struct.saveData")]
bool SaveStructData(string pageId, string schemaName, object data);

}

public class DokuwikiZugriff
{
    public IDokuWikiApi Proxy { get; set; }
    public XmlRpcStruct Options { get; set; } // Hier ist es definiert!

    public DokuwikiZugriff(IConfiguration configuration)
    {
        Global.Konfig("WikiUrl", Global.Modus.Update, configuration);
        Global.Konfig("WikiJsonUser", Global.Modus.Update, configuration);
        Global.Konfig("WikiJsonUserKennwort", Global.Modus.Update, configuration);
        
        // Proxy erstellen
        Proxy = XmlRpcProxyGen.Create<IDokuWikiApi>();
        ((XmlRpcClientProtocol)Proxy).Url = configuration["WikiUrl"];

        // Manuelle HTTP-Header setzen
        var credentials = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{configuration["WikiJsonUser"]}:{configuration["WikiJsonUserKennwort"]}"));
        ((XmlRpcClientProtocol)Proxy).Headers.Add("Authorization", "Basic " + credentials);
    }
}