using Newtonsoft.Json;
using System.Collections.Generic;

namespace TestClientObjectModel.ViewModels
{
    class TransformTextRespViewModel
    {
        [JsonProperty("status")]
        public string Status { get; set; }
        [JsonProperty("data")]
        public List<TextData> TextDatas { get; set; }
    }

    class TextData
    {
        [JsonProperty("data_id")]
        public string Data_id { get; set; }
        [JsonProperty("text")]
        public string[] Text { get; set; }


    }
}
