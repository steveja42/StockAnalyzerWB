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

        public async static Task<OptionChain> GetOptionChain(string symbol, string contractType)
        {

            using (var client = new HttpClient())
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
                client.BaseAddress = new Uri(BaseAddress);
                client.DefaultRequestHeaders.Add("User-Agent", "Anything");
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                string RequestUri = $"/v1/marketdata/chains?apikey={ApiKey}&symbol={symbol}&contractType={contractType}&includeQuotes=TRUE";
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
                
                string result = await response.Content.ReadAsStringAsync();
                var ochain = response.Content.ReadAsAsync<OptionChain>().Result;
                return ochain;
            }
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
    public class OptionChain
    {
        public string symbol { get; set; }
        public string status { get; set; }
        public UnderLying underlying;
        public string isDelayed { get; set; }
        public string underlyingPrice { get; set; }
        public Dictionary<string, ExpirationDate> callExpDateMap;

        public Dictionary<string, ExpirationDate> putExpDateMap;

        public override string ToString()
        {
            return $"{symbol}: {underlyingPrice} ";
        }
    }
}
