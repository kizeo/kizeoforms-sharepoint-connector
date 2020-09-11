using System;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using TestClientObjectModel.ViewModels;

namespace TestClientObjectModel
{
    class KizeoFormsApiManager
    {
        public HttpClient HttpClient;
        public string KfApiUrl;
        public static log4net.ILog Log = log4net.LogManager.GetLogger(typeof(KizeoFormsApiManager));

        /// <summary>
        /// Create an instance of KFAPiManager 
        /// </summary>
        /// <param name="baseUrl">url of kizeo Forms</param>
        /// <param name="token">the token to authentificate</param>
        public KizeoFormsApiManager(string baseUrl, string token)
        {
            HttpClient = new HttpClient();
            KfApiUrl = baseUrl;

            Log.Debug($"Configuring Http Client");

            try
            {
                HttpClient.BaseAddress = new Uri(baseUrl);
                HttpClient.DefaultRequestHeaders.Accept.Clear();
                HttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                if (string.IsNullOrEmpty(token))
                    throw new ArgumentException("Kizeo forms authentification token can not be null");

                HttpClient.DefaultRequestHeaders.TryAddWithoutValidation("Authorization", token);

                TestKfApi(baseUrl).Wait();
                Log.Debug($"Configuration succeeded");
            }
            catch (ArgumentNullException ex)
            {
                TOOLS.LogErrorAndExitProgram("url kizeo forms can not be empty" + ex.Message);
            }
            catch (HttpRequestException)
            {
                TOOLS.LogErrorAndExitProgram("Can't connect to Kizeo forms server");
            }
            catch (Exception ex)
            {
                TOOLS.LogErrorAndExitProgram("Error while trying access the server : " + ex.Message);
            }

        }


        /// <summary>
        /// run an http request to test the api
        /// </summary>
        /// <param name="kfUrl">url of kizeoForms</param>
        /// <returns></returns>
        public async Task TestKfApi(string kfUrl)
        {
            var testToken = await HttpClient.GetAsync($"{kfUrl}/rest/v3/testapi/sharepoint");
            if (!testToken.IsSuccessStatusCode)
            {
                throw new ArgumentException("Unauthorized Token");
            }
        }


        /// <summary>
        /// transform an expretion with expretion + data 
        /// </summary>
        /// <param name="formId">the formId</param>
        /// <param name="dataId">the DataId</param>
        /// <param name="columnSelector">expresion to replace into it</param>
        /// <returns></returns>
        public async Task<string[]> TransformText(string formId, string dataId, string columnSelector)
        {
            if (string.IsNullOrEmpty(columnSelector))
            {
                TOOLS.LogErrorAndExitProgram("Path = null " + columnSelector + " 3");
                return null;
            }
            else
            {

                HttpResponseMessage response = await HttpClient.PostAsJsonAsync($"{KfApiUrl}/rest/v3/forms/{formId}/transformText",
                                    new { textToTransform = columnSelector, data_ids = new string[] { dataId } }); 

                TransformTextRespViewModel transformedText = await response.Content.ReadAsAsync<TransformTextRespViewModel>();
                return transformedText.TextDatas.Where(td => td.Data_id == dataId).First().Text;
            }
        }

        public async Task<string[]> TransformTextAddItem(string formId, string dataId, string columnSelector)
        {
            if (string.IsNullOrEmpty(columnSelector))
            {
                TOOLS.LogErrorAndExitProgram("Path = null " + columnSelector + " 3");
                return null;
            }
            else
            {

                HttpResponseMessage response = await HttpClient.PostAsJsonAsync($"{KfApiUrl}/rest/v3/forms/{formId}/transformText",
                                    new { textToTransform = columnSelector, data_ids = new string[] { dataId } });

                TransformTextRespViewModel transformedText = await response.Content.ReadAsAsync<TransformTextRespViewModel>();
               /* string columnValue = transformedText.TextDatas.Where(td => td.Data_id == dataId).First().Text.First();*/
                string[] columnValue = transformedText.TextDatas.Where(td => td.Data_id == dataId).First().Text;

               /* if (columnValue.Contains("##"))
                    TOOLS.LogErrorAndExitProgram($"No column name found in kizeo forms acording to the expression : {columnValue}");
*/
                return columnValue;
            }
        }

        /// <summary>
        /// Create a marker based on the form id and the sharepoint list id
        /// </summary>
        /// <param name="formId">  the id of the kizeo forms </param>
        /// <param name="listId"> the id of sharepoint list or library</param>
        /// <returns> a unique marker</returns>
        public string CreateKFMarker(string formId, Guid listId)
        {

            return formId + listId.ToString().Substring(0, listId.ToString().IndexOf("-"));

        }

    }


}
