using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestClientObjectModel.ViewModels
{
    class FormDatasRespViewModel
    {
        [JsonProperty("status")]
        public string Status { get; set; }
        [JsonProperty("data")]
        public List<FormData> Data { get; set; }
    }

    class FormData
    {
        [JsonProperty("_id")]
        public string Id { get; set; }
        [JsonProperty("_form_id")]
        public string FormID { get; set; }

    }
}
