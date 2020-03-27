using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using TestClientObjectModel.ViewModels;

namespace TestClientObjectModel
{
    class ContentDisposition
    {
        private static readonly Regex regex = new Regex(
            "^([^;]+);(?:\\s*([^=]+)=((?<q>\"?)[^\"]*\\k<q>);?)*$",
            RegexOptions.Compiled
        );

        private readonly string fileName;
        private readonly StringDictionary parameters;
        private readonly string type;

        public ContentDisposition(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                throw new ArgumentNullException("s");
            }
            Match match = regex.Match(s);
            if (!match.Success)
            {
                throw new FormatException("input is not a valid content-disposition string.");
            }
            var typeGroup = match.Groups[1];
            var nameGroup = match.Groups[2];
            var valueGroup = match.Groups[3];

            int groupCount = match.Groups.Count;
            int paramCount = nameGroup.Captures.Count;

            this.type = typeGroup.Value;
            this.parameters = new StringDictionary();

            for (int i = 0; i < paramCount; i++)
            {
                string name = nameGroup.Captures[i].Value;
                string value = valueGroup.Captures[i].Value;

                if (name.Equals("filename", StringComparison.InvariantCultureIgnoreCase))
                {
                    this.fileName = value;
                }
                else
                {
                    this.parameters.Add(name, value);
                }
            }
        }
        public string FileName
        {
            get
            {
                return this.fileName;
            }
        }
        public StringDictionary Parameters
        {
            get
            {
                return this.parameters;
            }
        }
        public string Type
        {
            get
            {
                return this.type;
            }
        }
    }

    class SharePointManager
    {
        public ClientContext Context { get; set; }
        //Remember only one locky ever 
        public static object locky;

        //Remember only one locky ever 
        public static object lockyFileName;
        public KizeoFormsApiManager KfApiManager;

        private string ExtractDomainNameFromURL(string Url)
        {
            return System.Text.RegularExpressions.Regex.Replace(
                Url,
                @"^([a-zA-Z]+:\/\/)?([^\/]+)\/.*?$",
                "$2"
            );
        }
        private string TrySharePointConnection(string spDomain, string spClientId, string spClientSecret, string spTenantId)
        {
            string access_url = $"https://accounts.accesscontrol.windows.net/{spTenantId}/tokens/OAuth/2";
            const string resource_id = "00000003-0000-0ff1-ce00-000000000000";
            try
            {
                var request = (HttpWebRequest)WebRequest.Create("https://accounts.accesscontrol.windows.net/" + spTenantId + "/tokens/OAuth/2");

               
                var postData = "grant_type=client_credentials";
                postData += $"&client_id={ spClientId}@{spTenantId}";
                postData += "&client_secret=" + spClientSecret;
                postData += $"&resource={resource_id}/{ExtractDomainNameFromURL(spDomain)}@{spTenantId}";

                byte[] data = Encoding.UTF8.GetBytes(postData);

                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = data.Length;

                using (var stream = request.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);

                }
                var response = (HttpWebResponse)request.GetResponse();

                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();


                dynamic json = JsonConvert.DeserializeObject(responseString);
                return json.access_token;
            }
            catch (System.Net.WebException)
            {
                try
                {
                    var cc = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(ExtractDomainNameFromURL(spDomain), spClientId, spClientSecret);
                    client_context_buffer = cc;

                }
                catch (Exception)
                {
                    TOOLS.LogErrorwithoutExitProgram("Impossible de communiquer avec SharePoint");
                    return "undefined";
                }
            }
            catch (Exception)
            {
                TOOLS.LogErrorwithoutExitProgram("Impossible de communiquer avec SharePoint");
                return "undefined";
            }
            return "undefined";
        }

        private ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken)
        {
            Uri targetUri = new Uri(targetUrl);

            ClientContext clientContext = new ClientContext(targetUrl);


            clientContext.AuthenticationMode = ClientAuthenticationMode.Anonymous;
            clientContext.FormDigestHandlingEnabled = false;
            clientContext.ExecutingWebRequest +=
                delegate (object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken;
                };

            return clientContext;
        }

        public static log4net.ILog Log = log4net.LogManager.GetLogger(typeof(SharePointManager));
        private ClientContext client_context_buffer = null;
        /// <summary>
        /// Create instance of Shareoint manager and initialise the context
        /// </summary>
        /// <param name="spUrl"> SharePoint url</param>
        /// <param name="spUser">UserName</param>
        /// <param name="spPwd">Password</param>
        /// 
        public SharePointManager(string spDomain, string spClientId, string spClientSecret, string spTenantId, KizeoFormsApiManager kfApiManager_)
        {
            KfApiManager = kfApiManager_;
            Log.Debug($"Configuring Sharepoint Context");
            locky = new object();
            lockyFileName = new object();
            try
            {
                var token = TrySharePointConnection(spDomain, spClientId, spClientSecret, spTenantId);
                if (!token.Equals("undefined"))
                {
                    if (client_context_buffer != null)
                        Context = client_context_buffer;
                    else
                        Context = GetClientContextWithAccessToken(spDomain, token);
                    var web = Context.Web;
                    lock (locky)
                    {
                        Context.Load(web);
                        Context.ExecuteQuery();
                    }

                    Log.Debug($"Configuration succeeded");
                }
                else
                {
                    throw new Exception();
                }

              
            }
            catch (Exception ex)
            {
                TOOLS.LogErrorAndExitProgram("Error occured while initializing sharepoint config : " + ex.Message);
            }

        }


        /// <summary>
        /// Run an export : download the file and send it to sharePoint
        /// </summary>
        /// <param name="exportType"> Export type used only for logging </param>
        /// <param name="url"> Url of the request that is going to be executed </param>
        /// <param name="isToExport"> Boolean to export or not</param>
        /// <param name="formToSpLibrary"> Form to sp library to get config data </param>
        /// <param name="data"> The data that holds the dataId and the formId</param>
        /// <param name="path"> the kizeoforms expresion ## ## that controls where the file will be sent </param>
        /// <param name="postDataIds">Ids of the data that is going to be exported</param>
        public async void RunExport(string exportType, string url, bool isToExport, FormToSpLibrary formToSpLibrary, FormData data, string path, string[] postDataIds = null)
        {
            if (isToExport)
            {
                if (!string.IsNullOrEmpty(path))
                {
                    HttpResponseMessage response;

                    Log.Debug($"--Export to {exportType}");

                    if (postDataIds == null)
                        response = await KfApiManager.HttpClient.GetAsync($"{KfApiManager.KfApiUrl}/{url}");
                    else
                        response = await KfApiManager.HttpClient.PostAsJsonAsync($"{KfApiManager.KfApiUrl}/{url}", new { data_ids = postDataIds });

                    if (response.IsSuccessStatusCode)
                    {
                        string filePath = await KfApiManager.TransformText(data.FormID, data.Id, path);
                        string fileName = GetFileName(response, formToSpLibrary.SpLibrary, Path.Combine(formToSpLibrary.SpWebSiteUrl, filePath),"");

                        using (var ms = new MemoryStream())
                        {
                            filePath = Path.Combine(formToSpLibrary.LibraryName, filePath);
                            Log.Debug($"-----Downloading file : {fileName} from kizeoForms");
                            await response.Content.CopyToAsync(ms);
                            SendToSpLibrary(ms, formToSpLibrary.SpLibrary, formToSpLibrary.MetaData, data, filePath, fileName);
                        }
                    }
                    else
                    {
                        TOOLS.LogErrorAndExitProgram($"Error loading the export : {exportType} in path : {path}");
                    }

                }
                else
                {
                    TOOLS.LogErrorAndExitProgram($"le champ path de l'export {exportType} est vide");
                }
            }
        }

        /// <summary>
        /// send File to Sharepoint
        /// </summary>
        /// <param name="ms">the memory stream holding the data</param>
        /// <param name="formToSpLibrary"> Form to sp library to get config data </param>
        /// <param name="data"> The data that holds the dataId and the formId</param>
        /// <param name="metadatas">collection<Datamaping> tha holds the metadata</param>
        /// <param name="filePath">the filePath</param>
        /// <param name="fileName_">the fileName</param>
        public void SendToSpLibrary(MemoryStream ms, List formsLibrary, List<DataMapping> metadatas, FormData data, string filePath, string fileName_)
        {
            FileCreationInformation fcInfo = new FileCreationInformation();

            fcInfo.Url = Path.Combine(filePath, fileName_);
            fcInfo.Overwrite = true;
            fcInfo.Content = ms.ToArray();

            Log.Debug("-----Configuring file destination in Sharepoint");
            Log.Debug($"-----File url : {fcInfo.Url}");

            lock (locky)
            {

                try
                {
                    Microsoft.SharePoint.Client.File uploadedFile = formsLibrary.RootFolder.Files.Add(fcInfo);

                    if (metadatas != null)
                    {
                        FillFileMetaDatas(metadatas, uploadedFile, data).Wait();
                    }

                    uploadedFile.ListItemAllFields.Update();
                    uploadedFile.Update();

                    Context.ExecuteQuery();
                    Log.Debug($"File : {fcInfo.Url}  sent to splibrary successfully");

                }
                catch (Exception Ex)
                {

                    TOOLS.LogErrorwithoutExitProgram(Ex.Message + $"\n" + " filepath : error while uploading file :" + fcInfo.Url);
                  /*  TOOLS.LogErrorwithoutExitProgram(Ex.Message + $"\n" + " filepath : error while uploading file :" + fcInfo.Url + "\n Stack Trace : " + Ex.StackTrace);*/
                }


            }
        }


        /// <summary>
        /// fill metadata
        /// </summary>
        /// <param name="metadatas">collection of dataMapping</param>
        /// <param name="uploadedFile">the file that is going to be uploaded</param>
        /// <param name="data">data holds dataId and formId</param>
        /// <returns></returns>
        public async Task FillFileMetaDatas(List<DataMapping> metadatas, Microsoft.SharePoint.Client.File uploadedFile, FormData data)
        {
            Log.Debug("Processing MetaData");
            foreach (var metaData in metadatas)
            {
                string columnValue = await KfApiManager.TransformText(data.FormID, data.Id, metaData.KfColumnSelector);
                TOOLS.ConvertToCorrectTypeAndSet(uploadedFile.ListItemAllFields, metaData, columnValue);
            }
        }


        /// <summary>
        /// get all libraries folders paths
        /// </summary>
        /// <param name="spLibrary">the guid of the SpLibrary</param>
        /// <param name="spWebSiteUrl">SharePoint webSite url</param>
        /// <returns></returns>
        public List<string> GetAllLibraryFolders(List spLibrary, string spWebSiteUrl)
        {
            var folderItems = spLibrary.GetItems(CamlQuery.CreateAllFoldersQuery());

            lock (locky)
            {
                try
                {
                    Context.Load(folderItems, f => f.Include(i => i.Folder));
                    Context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    TOOLS.LogErrorAndExitProgram($"Error while retriving All folders from library : {spLibrary.Id} \n {ex.Message}");
                }

                return folderItems.Select(i => i.Folder.ServerRelativeUrl.Replace(spWebSiteUrl + "/", "")).ToList();
            }

        }
       

        /// <summary>
        /// Add and send item to Sharepointlist including Media 
        /// </summary>
        /// <param name="spList">guid of sharepoint list</param>
        /// <param name="item">the item</param>
        /// <param name="data">data that holds dataId and formId</param>
        /// <param name="dataToMark"></param>
        /// <param name="itemUpdated">In case of update, use this item</param>
        /// <returns></returns>
        public async Task<ListItem> AddItemToList(List spList, List<DataMapping> dataMappings, FormData data, MarkDataReqViewModel dataToMark, ListItem itemUpdated = null)
        {
            Log.Debug($"Processing data : {data.Id}");
            ListItem item;
            if (itemUpdated == null) {
                item = spList.AddItem(new ListItemCreationInformation());
            } else {
                item = itemUpdated;
            }

            foreach (var dataMapping in dataMappings)
            {
                string columnValue = await KfApiManager.TransformText(data.FormID, data.Id, dataMapping.KfColumnSelector);
                TOOLS.ConvertToCorrectTypeAndSet(item, dataMapping, columnValue);
            }

            try
            {
                lock (SharePointManager.locky)
                {
                    item.Update();
                    Context.ExecuteQuery();
                    dataToMark.Ids.Add(data.Id);
                    Log.Debug($"Data {data.Id} sent to Sharepoint successefully");
                }


                var x = $"{KfApiManager.KfApiUrl}/rest/v3/forms/{data.FormID}/data/{data.Id}/all_medias";
                var response = await KfApiManager.HttpClient.GetAsync(x);

                if (response.StatusCode == System.Net.HttpStatusCode.OK)
                {
                    using (var ms = new MemoryStream())
                    {
                        Log.Debug($"processing media for data {data.Id}");
                        await response.Content.CopyToAsync(ms);
                        ms.Seek(0, SeekOrigin.Begin);
                        AttachmentCreationInformation pjInfo = new AttachmentCreationInformation();
                        pjInfo.ContentStream = ms;
                        pjInfo.FileName = response.Content.Headers.ContentDisposition.FileName;

                        Attachment pj = item.AttachmentFiles.Add(pjInfo);

                        lock (locky)
                        {
                            Context.Load(pj);
                            Context.ExecuteQuery();

                        }
                    }
                }


            }
            catch (Exception)
            {
                throw;
            }

            return item;
        }


        /// <summary>
        /// Delete delet existing item used in the case of Unique column in sharepoint list
        /// </summary>
        /// <param name="spList">guid of sp List</param>
        /// <param name="spUniqueColumnName">sharepoint unique column name</param>
        /// <param name="kfUniqueColumnvalue">unique column value</param>
        public ListItem RetrieveExistingItem(List spList, string spUniqueColumnName, string kfUniqueColumnvalue)
        {
            CamlQuery cmQuery = new CamlQuery();
            cmQuery.ViewXml = $"<View><Query><Where><Eq><FieldRef Name='{spUniqueColumnName}'/><Value Type='Text'>{kfUniqueColumnvalue}</Value></Eq></Where></Query><RowLimit>10</RowLimit></View>";

            lock (locky)
            {
                ListItemCollection collListItem = spList.GetItems(cmQuery);
                Context.Load(collListItem);
                Context.ExecuteQuery();

                return collListItem.First();
            }

        }

        /// <summary>
        /// return all sharepoint files 
        /// </summary>
        /// <param name="spLibrary">guid of sharepoint library</param>
        /// <param name="path">the path in the spLibrary where the paths would be retrieved</param>
        /// <returns></returns>
        public List<string> GetSpFiles(List spLibrary, string path)
        {
            try
            {
                var query = CamlQuery.CreateAllItemsQuery();
                query.FolderServerRelativeUrl = path;

                lock (locky)
                {
                    var fileItems = spLibrary.GetItems(query);
                    Context.Load(fileItems, x => x.Include(xx => xx.File.Name));
                    Context.Load(fileItems, x => x.Include(xx => xx.ContentType.Name));
                    Context.ExecuteQuery();
                    return fileItems.Where(f => f.ContentType.Name != "Folder").Select(s => s.File.Name).ToList();
                }


            }
            catch (Exception EX)
            {
                TOOLS.LogErrorAndExitProgram($"Can not get files from {path} to check if the current file already exist.\n {EX.Message}");
                return null;
            }

        }

        /// <summary>
        /// search if this file name already exist if so add (2) in the end 
        /// </summary>
        /// <param name="response">http response that holds data</param>
        /// <param name="spLibrary">spLibrary config</param>
        /// <param name="folderPath">path in the spLibrary</param>
        /// <param name="fileNamePrefix"> file name prefixe</param>
        /// <returns></returns>
        public string GetFileName(HttpResponseMessage response, List spLibrary, string folderPath,string fileNamePrefix)
        {
            string fileNameText;
            if (response.Content.Headers.ContentDisposition == null)
            {
                IEnumerable<string> contentDisposition;
                if (response.Content.Headers.TryGetValues("Content-Disposition", out contentDisposition))
                {
                    fileNameText = contentDisposition.ToArray()[0];
                    var cp = new ContentDisposition(fileNameText);
                    fileNameText = cp.FileName;
                } else
                {
                    fileNameText = "";
                }
            } else
            {
                fileNameText = response.Content.Headers.ContentDisposition.FileName.ToString();
            }
            string fileName = fileNamePrefix +TOOLS.CleanString(fileNameText);
            // var allFiles = GetSpFiles(spLibrary, folderPath);

            int i = 2;
            string fileNameWithoutExt = fileName.Substring(0, fileName.IndexOf("."));
            string extention = fileName.Substring(fileName.IndexOf(".") + 1, fileName.Length - fileName.IndexOf(".") - 1);
            // bool containsFile = true;

            Guid g;
            g = Guid.NewGuid();
            string guidParsed = g.ToString().Substring(0, 13);
            fileName = fileNameWithoutExt + $".{guidParsed}." + extention;

            // do
            // {
            //     if (allFiles.Contains(fileName))
            //     {
            //         fileName = fileNameWithoutExt + $"({i})." + extention;
            //         i++;
            //         containsFile = true;
            //     } else {
            //         containsFile = false;
            //     }

            // } while (containsFile);

            return fileName;
        }

        /// <summary>
        /// fill an expression with data from Sharepoint
        /// </summary>
        /// <param name="dataSchema">string with $$XX$$ to represent data to replace</param>
        /// <param name="item">record in sharepoint </param>
        /// <returns></returns>
        public string TransformSharepointText(string dataSchema, ListItem item)
        {
            int startIndex, endIndex;
            string columnName;

            while (dataSchema.Contains("$$"))
            {
                startIndex = dataSchema.IndexOf("$$");
                endIndex = dataSchema.IndexOf("$$", startIndex + 2);
                columnName = dataSchema.Substring(startIndex + 2, (endIndex - (startIndex + 2)));
                dataSchema = dataSchema.Remove(startIndex, (endIndex - startIndex) + 2);
                try
                {
                    if (item[columnName] != null){
                        dataSchema = dataSchema.Insert(startIndex, item[columnName].ToString());
                    }
                }
                catch (Exception ex)
                {
                    TOOLS.LogErrorAndExitProgram($"Error occured when loading data : {columnName} from sharepoint please check your $$ syntaxe");
                }
            }

            return dataSchema;

        }


        /// <summary>
        /// Load sharepoint lsit
        /// </summary>
        /// <param name="listId">guid of the SharePoint list</param>
        /// <returns></returns>
        public List LoadSpList(Guid listId)
        {
            var spList = Context.Web.Lists.GetById(listId);

            Context.Load(spList);

            try
            {
                lock (locky)
                {
                    Context.ExecuteQuery();
                    return spList;
                }

            }
            catch (Exception ex)
            {
                TOOLS.LogErrorAndExitProgram($"Error while loading Sharepoint list {spList.Id} " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Load SharePoint Library
        /// </summary>
        /// <param name="formToSpLibrary">form to sp library config</param>
        /// <returns></returns>
        public List LoadSpLibrary(FormToSpLibrary formToSpLibrary)
        {
            var spLibrary = Context.Web.Lists.GetById(formToSpLibrary.SpLibraryId);
            Context.Load(spLibrary);
            Context.Load(spLibrary, spl => spl.Title);

            lock (locky)
            {

                try
                {
                    Context.ExecuteQuery();
                    formToSpLibrary.SpWebSiteUrl = ($@"{spLibrary.ParentWebUrl}/{spLibrary.Title}");
                    formToSpLibrary.LibraryName = spLibrary.Title;
                    Log.Debug("-SharePoint library Succesfully Loaded");
                    return spLibrary;
                }
                catch (Exception ex)
                {
                    TOOLS.LogErrorAndExitProgram("Error while loading Sharepoint Library " + ex.Message);
                    return null;
                }
            }

        }

        /// <summary>
        /// Load sp library
        /// </summary>
        /// <param name="periodicExport">periodic export config</param>
        /// <returns></returns>
        public List LoadSpLibrary(PeriodicExport periodicExport)
        {
            var spLibrary = Context.Web.Lists.GetById(periodicExport.SpLibraryId);
            Context.Load(spLibrary);
            Context.Load(spLibrary, spl => spl.Title);

            lock (locky)
            {

                try
                {
                    Context.ExecuteQuery();
                    periodicExport.SpWebSiteUrl = ($@"{spLibrary.ParentWebUrl}/{spLibrary.Title}");
                    periodicExport.LibraryName = spLibrary.Title;
                    Log.Debug("-SharePoint library Succesfully Loaded");
                    return spLibrary;
                }
                catch (Exception ex)
                {
                    TOOLS.LogErrorAndExitProgram("Error while loading Sharepoint Library " + ex.Message);
                    return null;
                }
            }

        }

        /// <summary>
        /// Load sharepoint lite records
        /// </summary>
        /// <param name="spList">guid of the spList</param>
        /// <returns> a listItem collection</returns>
        public ListItemCollection LoadListItems(List spList)
        {
            try
            {
                lock (locky)
                {
                    ListItemCollection items = spList.GetItems(new CamlQuery());
                    Context.Load(items);
                    Context.ExecuteQuery();
                    return items;
                }

            }
            catch (Exception ex)
            {
                TOOLS.LogErrorAndExitProgram($"Error while retrieving items from List {spList.Id} {ex.Message}");

            }

            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="exportType">the export type (only for log)</param>
        /// <param name="url">url of http request</param>
        /// <param name="periodicexport">periodic export config</param>
        /// <param name="path">path</param>
        /// <param name="postDataIds">array of ids to export</param>
        /// <param name="fileNamePrefix">file name prefix</param>
        internal async void RunPeriodicExport(string exportType, string url, PeriodicExport periodicexport, string path, string[] postDataIds, string fileNamePrefix)
        {

            if (!string.IsNullOrEmpty(path))
            {
                HttpResponseMessage response;

                Log.Debug($"--Export to {exportType}");

                response = await KfApiManager.HttpClient.PostAsJsonAsync($"{KfApiManager.KfApiUrl}/{url}", new { data_ids = postDataIds });

                if (response.IsSuccessStatusCode)
                {
                    string filePath = path;
                    string fileName = GetFileName(response, periodicexport.SpLibrary, Path.Combine(periodicexport.SpWebSiteUrl, filePath),fileNamePrefix);

                    using (var ms = new MemoryStream())
                    {
                        filePath = Path.Combine(periodicexport.LibraryName, filePath);
                        Log.Debug($"-----Downloading file : {fileName} from kizeoForms");
                        await response.Content.CopyToAsync(ms);
                        SendPeriodicExportToSpLibrary(ms, periodicexport.SpLibrary, filePath, fileName);
                    }
                }
                else
                {
                    TOOLS.LogErrorAndExitProgram($"Error loading the export : {exportType} in path : {path}");
                }

            }
            else
            {
                TOOLS.LogErrorAndExitProgram($"le champ path de l'export {exportType} est vide");
            }

        }

        /// <summary>
        /// send file to sp library
        /// </summary>
        /// <param name="ms">the memory stream holding the export</param>
        /// <param name="spLibrary">Guid of spLibrary</param>
        /// <param name="filePath">the file path</param>
        /// <param name="fileName_">the fileName</param>
        private void SendPeriodicExportToSpLibrary(MemoryStream ms, List spLibrary, string filePath, string fileName_)
        {
            FileCreationInformation fcInfo = new FileCreationInformation();

            fcInfo.Url = Path.Combine(filePath, fileName_);
            fcInfo.Overwrite = true;
            fcInfo.Content = ms.ToArray();

            Log.Debug("-----Configuring file destination in Sharepoint");
            Log.Debug($"-----File url : {fcInfo.Url}");

            lock (locky)
            {

                try
                {
                    Microsoft.SharePoint.Client.File uploadedFile = spLibrary.RootFolder.Files.Add(fcInfo);

                    uploadedFile.ListItemAllFields.Update();
                    uploadedFile.Update();

                    Context.ExecuteQuery();
                    Log.Debug($"File : {fcInfo.Url}  sent to splibrary successfully");

                }
                catch (Exception Ex)
                {

                    TOOLS.LogErrorwithoutExitProgram(Ex.Message + $"\n" + " filepath : error while uploading file :" + fcInfo.Url + "\n Stack Trace : " + Ex.StackTrace);

                }


            }
        }
    }
}
