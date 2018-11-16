using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;

namespace StockAnalyzerWB
{
    class TDAmeritrade
    {
        static string BaseAddress = "https://api.tdameritrade.com";
        static string ApiKey = null;   //set this to your apikey, or put your api key in the first line of file AuthData.txt"
        // returns the option chain data for a stock symbol into
        //more info at https://developer.tdameritrade.com/option-chains/apis/get/marketdata/chains#
        //strikeCount=1

        private static void ensureAuthorization()
        {
            if (ApiKey == null)
            {
                try
                {
                    string[] lines = System.IO.File.ReadAllLines(System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory.ToString(), @".\AuthData.txt"));
                    if (lines.Length > 0)
                        ApiKey = System.Web.HttpUtility.UrlEncode(lines[0]);
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
        }

        async static Task<HttpResponseMessage> getResponse(string RequestUri)
        {
            using (var client = new HttpClient())
            {

                client.BaseAddress = new Uri(BaseAddress);
                client.DefaultRequestHeaders.Add("User-Agent", "Anything");
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                HttpResponseMessage response = null;
                try
                {
                    response = await client.GetAsync(RequestUri);
                    response.EnsureSuccessStatusCode();
                }
                catch (Exception e)
                {
                    //string sz = $"ERROR: {e.Message} " + e?.InnerException.Message;
                    throw (e);

                }
                return response;

            }
        }

        public async static Task<List<string>> getOptionDates(string symbol)
        {
            if (symbol == null || symbol == "")
                return null;
            HttpResponseMessage response;
            ensureAuthorization();
            string RequestUri = $"/v1/marketdata/chains?apikey={ApiKey}&symbol={symbol}&contractType=PUT&strikeCount=1";
            try
            {
                response = await getResponse(RequestUri);
            }
            catch
            {
                return null;
            }
            string result = await response.Content.ReadAsStringAsync();
            var ochain = response.Content.ReadAsAsync<OptionChain>().Result;
            var optionDates = new List<string>() ;
            foreach (var item in ochain.putExpDateMap)
            {
                optionDates.Add(item.Key.Remove(item.Key.LastIndexOf(':')));

            }
            return optionDates;

        }


        public async static Task<OptionChain> GetOptionChain(string symbol, string contractType)
        {
            if (symbol == null || symbol == "")
                return null;
            HttpResponseMessage response;
            ensureAuthorization();
            
            string RequestUri = $"/v1/marketdata/chains?apikey={ApiKey}&symbol={symbol}&contractType={contractType}&includeQuotes=TRUE&strikeCount=40";
            try
            {
                response = await getResponse(RequestUri);
            }
            catch
            {
                return null;
            }
            string result = await response.Content.ReadAsStringAsync();
            var ochain = response.Content.ReadAsAsync<OptionChain>().Result;
            return ochain;
            
        }
    }



    //classes used for JSON deserialization
    public class StrikeData
    {
        public string putCall { get; set; }
        public string symbol { get; set; }
        public decimal bid { get; set; }
        public decimal ask { get; set; }
        public decimal last { get; set; }
        public override string ToString()
        {
            return $"{symbol}: {last} ";
        }

    }
 
    public class ExpirationDate :  Dictionary<string, StrikeData[]>
    {
    }

    public class UnderLying
    {
        public decimal bid { get; set; }
        public decimal ask { get; set; }
        public decimal last { get; set; }
    }

    public class ExpirationDates : Dictionary<string, ExpirationDate>
    {

    }
    public class OptionChain
    {
        public string symbol { get; set; }
        public string status { get; set; }
        public UnderLying underlying;
        public string isDelayed { get; set; }
        public string underlyingPrice { get; set; }
       // public ExpirationDates callExpDateMap;
      //  public ExpirationDates putExpDateMap;
       public Dictionary<string, ExpirationDate> callExpDateMap;
        public Dictionary<string, ExpirationDate> putExpDateMap;

        public override string ToString()
        {
            return $"{symbol}: {underlyingPrice} ";
        }
    }
}
