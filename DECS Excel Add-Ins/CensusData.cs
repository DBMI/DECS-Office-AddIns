
using Newtonsoft.Json;
using System.Web.UI;

namespace DECS_Excel_Add_Ins
{
    /**
     * @brief Used to deserialize JSON returned from US Census Bureau geocoding service.
     */
    public class CensusData
    {
        public Result result { get; set; }

        /// <summary>
        /// Pulls the census tract number (FIPS code) out of the object.
        /// </summary>
        /// <param name="sheet">ActiveWorksheet.</param>
        /// <returns>ulong</returns>
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

    /**
     * @brief First-level class to deserialize JSON returned from US Census Bureau geocoding service.
     */
    public class Result
    {
        public Input input { get; set; }
        public Addressmatch[] addressMatches { get; set; }
    }

    /**
     * @brief Second-level class to deserialize JSON returned from US Census Bureau geocoding service.
     */
    public class Input
    {
        public Address address { get; set; }
    }

    /**
     * @brief Third-level class to deserialize JSON returned from US Census Bureau geocoding service.
     */
    public class Address
    {
        public string address { get; set; }
    }

    /**
     * @brief Second-level class to deserialize JSON returned from US Census Bureau geocoding service.
     */
    public class Addressmatch
    {
        public Geographies geographies { get; set; }
        public string matchedAddress { get; set; }
    }

    /**
     * @brief Third-level class to deserialize JSON returned from US Census Bureau geocoding service.
     */
    public class Geographies
    {
        [JsonProperty(PropertyName = "Census Tracts")]
        public CensusTract[] CensusTracts { get; set; }
    }

    /**
     * @brief Fourth-level class to deserialize JSON returned from US Census Bureau geocoding service.
     */
    public class CensusTract
    {
        public string GEOID { get; set; }
    }
}
