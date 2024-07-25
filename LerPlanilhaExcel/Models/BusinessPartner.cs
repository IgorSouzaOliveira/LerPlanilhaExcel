using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LerPlanilhaExcel.Models
{
    class BusinessPartner
    {
        public String CardCode { get; set; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public String EmailAddress { get; set; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public String Phone1 { get; set; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public String Phone2 { get; set; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public String U_LGO_Rota { get; set; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public String PayTermsGrpCode { get; set; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public String PriceListNum { get; set; }

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]
        public int SalesPersonCode { get; set; }

    }
}
