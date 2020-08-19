using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using OfficeDevPnP.Core;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using TestClientObjectModel.ViewModels;
using AuthenticationManager = OfficeDevPnP.Core.AuthenticationManager;

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
        public static log4net.ILog Log = log4net.LogManager.GetLogger(typeof(SharePointManager));
        private ClientContext client_context_buffer = null;
        /// <summary>
        /// Create instance of Shareoint manager and initialise the context
        /// </summary>
        /// <param name="spUrl"> SharePoint url</param>
        /// <param name="spUser">UserName</param>
        /// <param name="spPwd">Password</param>
        /// 
        public SharePointManager(string spDomain, string spClientId, string spClientSecret, KizeoFormsApiManager kfApiManager_)
        {
            KfApiManager = kfApiManager_;
            Log.Debug($"Configuring Sharepoint Context");
            locky = new object();
            lockyFileName = new object();
            try
            {
                Context = new AuthenticationManager().GetAppOnlyAuthenticatedContext(spDomain, spClientId, spClientSecret);
                var web = Context.Web;
                lock (locky)
                {
                    Context.Load(web);
                    Context.ExecuteQuery();
                }

                Log.Debug($"Configuration succeeded");
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
                        string filePath = (await KfApiManager.TransformText(data.FormID, data.Id, path)).First();
                        string fileName = GetFileName(response, formToSpLibrary.SpLibrary, Path.Combine(formToSpLibrary.SpWebSiteUrl, filePath), "");

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
                string[] columnValues = await KfApiManager.TransformText(data.FormID, data.Id, metaData.KfColumnSelector);
                TOOLS.ConvertToCorrectTypeAndSet(uploadedFile.ListItemAllFields, metaData, columnValues.First());
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
        public ListItem MustUpdate(string line, List<ListItem> items, DataMapping mapping, List<DataMapping> unique)
        {
            if (unique.Find(_mapping => _mapping.KfColumnSelector.Equals(mapping.KfColumnSelector)) != null)
            {
                for (var i = 0; i < items.Count; i++)
                {
                    foreach (var item in items[i].FieldValues)
                    {
                        if (item.Key.Equals(mapping.SpColumnId))

                            if (item.Value != null && item.Value.Equals(line))
                            {
                                return items[i];
                            }
                    }
                }
            }
            return null;
        }
        private async Task<ListItem> NeedTobeUpdated(List<DataMapping> dataMappings, FormData data, List spList, List<DataMapping> unique)
        {
            foreach (var mapping in dataMappings)
            {
                if (unique.Find(_mapping => _mapping.KfColumnSelector.Equals(mapping.KfColumnSelector)) != null)
                {
                    string columnValue = (await KfApiManager.TransformText(data.FormID, data.Id, mapping.KfColumnSelector)).First();
                    var q = new CamlQuery() { ViewXml = "<View><Query><Where><Eq><FieldRef Name='" + mapping.SpColumnId + "' /><Value Type='Text'>" + columnValue + "</Value></Eq></Where></Query></View>" };
                    var r = spList.GetItems(q);
                    Context.Load(r);
                    Context.ExecuteQuery();
                    if (r.Count > 0)
                        return r.Last();
                }
            }
            return null;
        }

        private void RemoveDuplicated(ref List<ListItem> toAdd, List<DataMapping> dataMappings, List spList, List<DataMapping> unique, ListItemCollection allItems)
        {

            List<int> toRemove = new List<int>();
            List<ListItem> complement = new List<ListItem>();
            int count = 0;
            List<ListItem> queries = new List<ListItem>();
            foreach (ListItem item in toAdd)
            {
                foreach (ListItem spItem in allItems)
                {
                    if (spItem[unique.First().SpColumnId].Equals(item[unique.First().SpColumnId]))
                    {
                        foreach (var key in dataMappings)
                        {
                            spItem[key.SpColumnId] = item[key.SpColumnId];
                        }
                        spItem.Update();
                        complement.Add(spItem);
                        toRemove.Add(count);
                        break;
                    }
                }
                count++;
            }
            int buff = 0;
            foreach (int index in toRemove)
            {
                toAdd.RemoveAt(index - buff++);
            }
            toAdd.AddRange(complement);
        }


        private ListItemCollection getAllListItems(List spList)
        {
            var q = new CamlQuery() { ViewXml = "<View><Query /></View>" };
            var r = spList.GetItems(q);
            Context.Load(r);
            Context.ExecuteQuery();
            return r;
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
        public async Task<ListItem> AddItemToList(List spList, List<DataMapping> dataMappings, FormData data, MarkDataReqViewModel dataToMark, List<DataMapping> uniqueDataMappings)
        {
            bool containsArray = false;
            Log.Debug($"Processing data : {data.Id}");
            List<ListItem> toAdd = new List<ListItem>();
            List<List<string>> lines = new List<List<string>>();
            List<string[]> results = new List<string[]>();
            ListItemCollection allItems = getAllListItems(spList);
            int toCreate = -1;

            ListItem item = spList.AddItem(new ListItemCreationInformation());

            var r = await NeedTobeUpdated(dataMappings, data, spList, uniqueDataMappings);
            if (r != null)
            {
                item = r;
            }
            foreach (var dataMapping in dataMappings)
            {
                string[] columnValues = await KfApiManager.TransformText(data.FormID, data.Id, dataMapping.KfColumnSelector);

                TOOLS.ConvertToCorrectTypeAndSet(item, dataMapping, columnValues.First());
                if (columnValues.Length > 1)
                {
                    if (toCreate.Equals(-1))
                        toCreate = columnValues.Length - 1;

                    containsArray = true;
                    int cvc = 0;
                    foreach (string columnValue in columnValues)
                    {
                        if (columnValue != columnValues.First())
                        {
                            if (toCreate > 0)
                            {
                                var tmp_item = spList.AddItem(new ListItemCreationInformation());
                                foreach (var field in dataMappings)
                                {
                                    if (item.FieldValues.ContainsKey(field.SpColumnId))
                                    {
                                        tmp_item[field.SpColumnId] = item[field.SpColumnId];
                                    }
                                }
                                TOOLS.ConvertToCorrectTypeAndSet(tmp_item, dataMapping, columnValue);
                                toAdd.Add(tmp_item);
                                toCreate--;
                            }
                            else
                            {
                                TOOLS.ConvertToCorrectTypeAndSet(toAdd[cvc++], dataMapping, columnValue);
                            }
                        }
                    }
                }
            }
            toAdd.Insert(0, item);
            if (containsArray)
                RemoveDuplicated(ref toAdd, dataMappings, spList, uniqueDataMappings, allItems);
            try
            {
                lock (locky)
                {
                    foreach (var add in toAdd)
                    {
                        add.Update();
                    }
                    Context.ExecuteQuery();
                    dataToMark.Ids.Add(data.Id);
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

                        Attachment pj = toAdd.Last().AttachmentFiles.Add(pjInfo);

                        lock (locky)
                        {
                            Context.Load(pj);
                            Context.ExecuteQuery();
                        }
                    }
                }

                /*   foreach (var item in items)
                   {
                       item.DeleteObject();
                   }*/
                Context.ExecuteQuery();
                Log.Debug($"Data {data.Id} sent to Sharepoint successefully");
            }
            catch (Exception)
            {
                throw;
            }

            return null;
        }


        /// <summary>
        /// Get all items of a list
        /// </summary>
        /// <param name="spList">guid of sp List</param>
        public List<ListItem> GetItems(List spList)
        {
            CamlQuery cmQuery = new CamlQuery();



            CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
            ListItemCollection items = spList.GetItems(query);
            Context.Load(items);
            Context.ExecuteQuery();


            lock (locky)
            {
                ListItemCollection collListItem = spList.GetItems(cmQuery);
                Context.Load(collListItem);
                Context.ExecuteQuery();

                return collListItem.ToList();
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
        public string GetFileName(HttpResponseMessage response, List spLibrary, string folderPath, string fileNamePrefix)
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
                }
                else
                {
                    fileNameText = "";
                }
            }
            else
            {
                fileNameText = response.Content.Headers.ContentDisposition.FileName.ToString();
            }
            string fileName = fileNamePrefix + TOOLS.CleanString(fileNameText);


            int i = 2;
            string fileNameWithoutExt = fileName.Substring(0, fileName.IndexOf("."));
            string extention = fileName.Substring(fileName.IndexOf(".") + 1, fileName.Length - fileName.IndexOf(".") - 1);


            Guid g;
            g = Guid.NewGuid();
            string guidParsed = g.ToString().Substring(0, 13);
            fileName = fileNameWithoutExt + $".{guidParsed}." + extention;

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
                    if (item[columnName] != null)
                    {
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
                    string fileName = GetFileName(response, periodicexport.SpLibrary, Path.Combine(periodicexport.SpWebSiteUrl, filePath), fileNamePrefix);

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
