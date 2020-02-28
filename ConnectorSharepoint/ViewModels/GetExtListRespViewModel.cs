using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestClientObjectModel.ViewModels
{
    class GetExtListRespViewModel
    {
        [JsonProperty("status")]
        public string Status { get; set; }
        [JsonProperty("list")]
        public ExternalList ExternalList { get; set; }
    }

    class ExternalList
    {
        [JsonProperty("id")]
        public string Id { get; set; }
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("items")]
        public string[] Items { get; set; }

    }
}
