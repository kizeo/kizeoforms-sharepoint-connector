using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestClientObjectModel.ViewModels
{
    class MarkDataReqViewModel
    {
        [JsonProperty("data_ids")]
        public List<string> Ids { get; set; } = new List<string>();
    }
}
