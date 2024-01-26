
using Newtonsoft.Json;
using System.Text.Json;
using System.Web.UI;

namespace DECS_Excel_Add_Ins
{
    public class CensusData
    {
        public Result result { get; set; }

        public ulong FIPS()
        {
            ulong geoid = 0;

            try
            {
                if (ulong.TryParse(result.addressMatches[0].geographies.CensusTracts[0].GEOID, out ulong temp))
                {
                    geoid = temp;
                }
            }
            catch { }

            return geoid;
        }
    }

    public class Result
    {
        public Input input { get; set; }
        public Addressmatch[] addressMatches { get; set; }
    }

    public class Input
    {
        public Address address { get; set; }
    }

    public class Address
    {
        public string address { get; set; }
    }

    public class Addressmatch
    {
        public Geographies geographies { get; set; }
        public string matchedAddress { get; set; }
    }

    public class Geographies
    {
        [JsonProperty(PropertyName = "Census Tracts")]
        public CensusTract[] CensusTracts { get; set; }
    }

    public class CensusTract
    {
        public string GEOID { get; set; }
    }
}
