using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using RestSharp;

namespace FiluetExcelAddIn
{
    public class RuPostQueryAPI
    {
        public PostSearch SearchRPO(string BoxID)
        {
            PostSearch ps = new PostSearch();
            string url = string.Format("{0}{1}", Properties.Resources.SearchURL, BoxID);
            var client = new RestClient(url);

            var r = new RestRequest();
            //r.Resource = "1.0/clean/address";
            r.Method = Method.GET;
            r.Timeout = 3000;

            r.AddHeader("Authorization", Properties.Resources.Authorization);
            r.AddHeader("X-User-Authorization", Properties.Resources.XUserAuthorization);
            r.AddHeader("Content-Type", Properties.Resources.ContentType);

            //r.AddJsonBody(SimpleJson.DeserializeObject(output));

            IRestResponse response = client.Execute(r);
            try
            {
                ps = JsonConvert.DeserializeObject<PostSearch>(response.Content);
            }
            catch (Exception ex)
            {
                ThisAddIn.formProgress.SetLog("Error: " + ex.Message);
            }
            return ps;
        }
    }

    public class Dimension
    {

        [JsonProperty("height")]
        public int Height { get; set; }

        [JsonProperty("length")]
        public int Length { get; set; }

        [JsonProperty("width")]
        public int Width { get; set; }
    }

    public class PostSearch
    {

        [JsonProperty("address-type-to")]
        public string AddressTypeTo { get; set; }

        [JsonProperty("area-to")]
        public string AreaTo { get; set; }

        [JsonProperty("barcode")]
        public string Barcode { get; set; }

        [JsonProperty("batch-category")]
        public string BatchCategory { get; set; }

        [JsonProperty("batch-name")]
        public string BatchName { get; set; }

        [JsonProperty("batch-status")]
        public string BatchStatus { get; set; }

        [JsonProperty("building-to")]
        public string BuildingTo { get; set; }

        [JsonProperty("comment")]
        public string Comment { get; set; }

        [JsonProperty("corpus-to")]
        public string CorpusTo { get; set; }

        [JsonProperty("dimension")]
        public Dimension Dimension { get; set; }

        [JsonProperty("ground-rate-with-vat")]
        public int GroundRateWithVat { get; set; }

        [JsonProperty("ground-rate-wo-vat")]
        public int GroundRateWoVat { get; set; }

        [JsonProperty("hotel-to")]
        public string HotelTo { get; set; }

        [JsonProperty("house-to")]
        public string HouseTo { get; set; }

        [JsonProperty("human-operation-name")]
        public string HumanOperationName { get; set; }

        [JsonProperty("id")]
        public int Id { get; set; }

        [JsonProperty("index-to")]
        public int IndexTo { get; set; }

        [JsonProperty("insr-rate-with-vat")]
        public int InsrRateWithVat { get; set; }

        [JsonProperty("insr-rate-wo-vat")]
        public int InsrRateWoVat { get; set; }

        [JsonProperty("insr-value")]
        public int InsrValue { get; set; }

        [JsonProperty("last-oper-attr")]
        public string LastOperAttr { get; set; }

        [JsonProperty("last-oper-date")]
        public DateTime LastOperDate { get; set; }

        [JsonProperty("last-oper-type")]
        public string LastOperType { get; set; }

        [JsonProperty("letter-to")]
        public string LetterTo { get; set; }

        [JsonProperty("location-to")]
        public string LocationTo { get; set; }

        [JsonProperty("mail-category")]
        public string MailCategory { get; set; }

        [JsonProperty("mail-direct")]
        public int MailDirect { get; set; }

        [JsonProperty("mail-type")]
        public string MailType { get; set; }

        [JsonProperty("mass")]
        public int Mass { get; set; }

        [JsonProperty("mass-rate-with-vat")]
        public int MassRateWithVat { get; set; }

        [JsonProperty("mass-rate-wo-vat")]
        public int MassRateWoVat { get; set; }

        [JsonProperty("num-address-type-to")]
        public string NumAddressTypeTo { get; set; }

        [JsonProperty("order-num")]
        public string OrderNum { get; set; }

        [JsonProperty("place-to")]
        public string PlaceTo { get; set; }

        [JsonProperty("postmarks")]
        public IList<string> Postmarks { get; set; }

        [JsonProperty("postoffice-code")]
        public string PostofficeCode { get; set; }

        [JsonProperty("rcp-pays-shipping")]
        public bool RcpPaysShipping { get; set; }

        [JsonProperty("region-to")]
        public string RegionTo { get; set; }

        [JsonProperty("room-to")]
        public string RoomTo { get; set; }

        [JsonProperty("slash-to")]
        public string SlashTo { get; set; }

        [JsonProperty("street-to")]
        public string StreetTo { get; set; }

        [JsonProperty("surname")]
        public string Surname { get; set; }

        [JsonProperty("total-rate-wo-vat")]
        public int TotalRateWoVat { get; set; }

        [JsonProperty("total-vat")]
        public int TotalVat { get; set; }

        [JsonProperty("transport-type")]
        public string TransportType { get; set; }

        [JsonProperty("version")]
        public int Version { get; set; }
    }
}
