using Microsoft.SharePoint.Client;

namespace DemoProviderWeb.Models
{
    public class TelefonoViewModel
    {
        public int Id { get; set; }
        public string Nombre { get; set; }
        public string Numero { get; set; }

        public static TelefonoViewModel FromListItem(ListItem item)
        {
            var data = new TelefonoViewModel();
            int res;
            int.TryParse(item["ID"].ToString(), out res);
            data.Id = res;
            data.Nombre = item["Title"].ToString();
            data.Numero = item["Numero"].ToString();
            return data;
        }
    }
}