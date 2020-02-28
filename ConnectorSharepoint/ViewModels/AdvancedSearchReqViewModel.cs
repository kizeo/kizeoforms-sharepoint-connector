using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestClientObjectModel.ViewModels
{
    class AdvancedSearchReqViewModel
    {
        public List<Filters> Filters;
    }

    class Filters
    {
        [JsonProperty("type")]
        public string Type { get; set; }
        [JsonProperty("components")]
        public List<Components> Components;
    }

    class Components
    {
        [JsonProperty("field")]
        public string Field { get; set; }
        [JsonProperty("operator")]
        public string Operator { get; set; }
        [JsonProperty("val")]
        public string Value { get; set; }
    }
}

