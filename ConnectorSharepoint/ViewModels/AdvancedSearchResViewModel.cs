using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestClientObjectModel.ViewModels
{
    class AdvancedSearchResViewModel
    {
        [JsonProperty("recordsTotal")]
        public int  RecordTotal { get; set; }
        [JsonProperty("data")]
        public List<Data> data { get; set; }
    }

    class Data
    {
        [JsonProperty("_id")]
        public string Id { get; set; }
    }
}

