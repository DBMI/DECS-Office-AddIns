using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.Office.Interop.Excel;
using System.Net;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Connects to US Census Bureau's online geocoding service, queries for census info & parses the output.
     */
    internal class Geocode
    {
        private const string URL = @"https://geocoding.geo.census.gov/geocoder/geographies/onelineaddress?address=";
        private const string SUFFIX = @"&benchmark=2020&vintage=2020&format=json";

        // https://stackoverflow.com/a/28546547/18749636
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(
            System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
        );

        /// <summary>
        /// Sends HTTP query containing address & converts response to a @c CensusData object.
        /// </summary>
        /// <param name="address">Address to query</param>
        
        internal CensusData Convert(string address)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            CensusData data = new CensusData();
            string addressEncoded = Uri.EscapeDataString(address.Replace(", ", ","));
            string url = URL + addressEncoded + SUFFIX;

            HttpClient httpClient = new HttpClient();
            httpClient.Timeout = TimeSpan.FromMinutes(5);

            // Call asynchronous network methods in a try/catch block to handle exceptions.
            try
            {
                httpClient.BaseAddress = new Uri(url);

                using (HttpResponseMessage response = httpClient.GetAsync(url).Result)
                {
                    response.EnsureSuccessStatusCode();
                    string responseBody = response.Content.ReadAsStringAsync().Result;

                    // Convert "Census Tracts" to "CensusTracts" (without space) for parsing into CensusData class.
                    responseBody = responseBody.Replace("Census Tracts", "CensusTracts");

                    var serializer = new JavaScriptSerializer();
                    serializer.MaxJsonLength = int.MaxValue;
                    data = serializer.Deserialize<CensusData>(responseBody);
                }
            }
            catch (HttpRequestException e)
            {
                log.Error(e.Message);
            }

            return data;
        }
    }
}
