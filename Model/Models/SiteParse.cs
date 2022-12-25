using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;

namespace Model.Models
{
  
    public interface ISiteParse
    {
        void Download(string url, string directory, string filename);
        void Parse();
    }

    public class SiteParse: ISiteParse
    {
        public Uri produkt { get; set; }
        public string PHPSESSID { get; set; }


        public SiteParse(string url, string PHPSESSID)
        {
            produkt = new Uri(url);
            this.PHPSESSID = PHPSESSID;
        }

        public void Download(string url, string directory, string filename)
        {
           
            using (var c = new WebClient())
                c.DownloadFile(url, Path.Combine(directory, filename));

            Console.WriteLine($"Download {directory} - {filename}");
        }

        public void Parse()
        {
            CookieContainer cookieContainer = new CookieContainer();
            cookieContainer.Add(new Cookie("PHPSESSID", PHPSESSID) { Domain = produkt.Host, Expires = new DateTime(2023, 07, 1) });


            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(produkt);
            request.CookieContainer = cookieContainer;

            HttpWebResponse resp = (HttpWebResponse)request.GetResponse();

            var result = String.Empty;
            if (resp.StatusCode == HttpStatusCode.OK)
            {
                var responseStream = resp.GetResponseStream();
                if (responseStream != null)
                {
                    StreamReader streamReader;
                    if (resp.CharacterSet != null)
                        streamReader = new StreamReader(responseStream);
                    else
                        streamReader = new StreamReader(responseStream);
                    result = streamReader.ReadToEnd();
                    streamReader.Close();
                }
                resp.Close();
            }

            var doc = new HtmlDocument();
            doc.LoadHtml(result);

            var cols = doc.DocumentNode.SelectNodes("//form[@class='form2']//tr");

            foreach (var item in cols)
            {

                if (item != null)
                {
                    var link = item.SelectSingleNode($"td[1]/a");
                    var date = item.SelectSingleNode($"td[3]");


                    if (link != null && date != null)
                    {
                        //Console.WriteLine(link.InnerText + " " + " " + link.Attributes["href"].Value);
                        //Console.WriteLine(date.InnerText);

                        Download("http://lib.istu.edu.ua/index.php/" + link.Attributes["href"].Value, "temp", link.InnerText);
                    }
                }
            }
        }
    }
}
