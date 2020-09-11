using log4net.Config;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Timers;
using TestClientObjectModel.ViewModels;

namespace TestClientObjectModel
{
    class Program
    {
        public static ClientContext Context { get; set; }
        public static Config Config { get; set; }
        public static HttpClient HttpClient;
        public static log4net.ILog Log = log4net.LogManager.GetLogger(typeof(Program));

        public static SharePointManager SpManager;
        public static KizeoFormsApiManager KfApiManager;

        public static string ConfFile = @"sharepoint_kf_connector_config.json";
        public static int periodicExportHour = 4;

        static void Main(string[] args)
        {
            System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = System.Globalization.CultureInfo.GetCultureInfo("en-US");

            // Please don't choose too short time
            int kfToSpsyncTime = 5;
            int spToKfSyncTime = 30;


            try
            {
                XmlConfigurator.Configure();

                Log.Info("Configuration");

                Log.Debug($"Get and serialize config file from : {ConfFile}.");
                Config = GetConfig(ConfFile);
                Log.Debug($"Configuration succeeded");

                if (Config == null || Config.KizeoConfig == null || Config.SharepointConfig == null)
                {
                    TOOLS.LogErrorAndExitProgram("The config file is empty or some data is missing.");
                }

                KfApiManager = new KizeoFormsApiManager(Config.KizeoConfig.Url, Config.KizeoConfig.Token);
                HttpClient = KfApiManager.HttpClient;

                SpManager = new SharePointManager(Config.SharepointConfig.SPDomain, Config.SharepointConfig.SPClientId, Config.SharepointConfig.SPClientSecret, KfApiManager);
                Context = SpManager.Context;

                FillKfExtListsFromSp().Wait();
                FillSpListsFromKfData().Wait();
                UploadExportsToSpLibraries().Wait();

                initTimers(kfToSpsyncTime, spToKfSyncTime);

                Log.Info($"Synchronisation will be executed every {kfToSpsyncTime} minutes.");
                Console.ReadKey();

            }
            catch (Exception EX)
            {
                Log.Fatal(EX);
            }
            finally
            {
                Context.Dispose();
            }
            while (true)
            {
                Console.ReadKey();
            }

        }





        /// <summary>
        /// Fill kizeo Forms External list with SharePoint data acording to the Config file
        /// </summary>
        /// <returns></returns>
        public static async Task FillKfExtListsFromSp()
        {
            HttpResponseMessage response;

            Log.Info("Updating KF's external lists from Sharepoint's lists");
            foreach (var spListToExtList in Config.SpListsToExtLists)
            {
                Log.Debug($"Loading Sharepoint's list: {spListToExtList.SpListId}");
                var spList = SpManager.LoadSpList(spListToExtList.SpListId);

                ListItemCollection items = SpManager.LoadListItems(spList);

                Log.Debug("Sharepoint's list Succesfully Loaded");

                Log.Debug($"Loading KF's extrenal list: {spListToExtList.ExListId}");

                response = await HttpClient.GetAsync($"{Config.KizeoConfig.Url}/rest/v3/lists/{spListToExtList.ExListId}");

                if (response.IsSuccessStatusCode)
                {
                    GetExtListRespViewModel kfExtList = await response.Content.ReadAsAsync<GetExtListRespViewModel>();
                    if (kfExtList.ExternalList == null)
                        TOOLS.LogErrorAndExitProgram($"Can not find an externalList with for id {spListToExtList.ExListId}, please check if this is a valid id.");
                    Log.Debug("KF's list successfully loaded");

                    List<string> linesToAdd = new List<string>();
                    foreach (var item in items)
                    {
                        Log.Debug($"Processing item : {item["Title"]}");
                        linesToAdd.Add(SpManager.TransformSharepointText(spListToExtList.DataSchema, item));
                    }

                    response = await HttpClient.PutAsJsonAsync($"{Config.KizeoConfig.Url}/rest/v3/lists/{spListToExtList.ExListId}", new { items = linesToAdd });
                    if (!response.IsSuccessStatusCode)
                    {
                        Log.Error("Error when updating external list with data :");
                        foreach (var line in linesToAdd)
                        {
                            TOOLS.LogErrorwithoutExitProgram("\n" + line);
                        }

                        TOOLS.LogErrorAndExitProgram($"Error when updating external list {spListToExtList.ExListId} ");
                    }
                    else
                    {
                        Log.Info($"External list {spListToExtList.ExListId} updated successfully");
                    }
                }
            }
        }


        /// <summary>
        /// Fill Sharepoint list from kizeo forms data acording to the config file
        /// </summary>
        /// <returns></returns>
        public static async Task FillSpListsFromKfData()
        {
            Log.Info("Filling Sharepoint's lists from KF's data");

            FormDatasRespViewModel formData = null;
            HttpResponseMessage response;

            MarkDataReqViewModel dataToMark = new MarkDataReqViewModel();

            foreach (var formToSpList in Config.FormsToSpLists)
            {
                string marker = KfApiManager.CreateKFMarker(formToSpList.FormId, formToSpList.SpListId);
                string formId = formToSpList.FormId;
                Log.Debug($"Processing form : {formId}");

                if (formToSpList.DataMapping != null)
                {
                    response = await HttpClient.GetAsync($"{Config.KizeoConfig.Url}/rest/v3/forms/{formId}/data/unread/{marker}/50?includeupdated");
                    if (response.IsSuccessStatusCode)
                    {
                        formData = await response.Content.ReadAsAsync<FormDatasRespViewModel>();
                        if (formData.Data == null)
                            TOOLS.LogErrorAndExitProgram($"Can not find a form with for id {formId}, please check if this is a valid id.");
                        Log.Debug($"{formData.Data.Count} data retrieved successfully from form.");

                        Log.Debug("Loading Sharepoint's list");
                        var spList = SpManager.LoadSpList(formToSpList.SpListId);
                        ListItemCollection allItems = SpManager.getAllListItems(spList);
                        Log.Debug("Sharepoint's list succesfully loaded");

                        dataToMark = new MarkDataReqViewModel();

                        foreach (var data in formData.Data)
                        {
                            try
                            {
                                var uniqueColumns = formToSpList.DataMapping.Where(dm => dm.SpecialType == "Unique").ToList();
                                await SpManager.AddItemToList(spList, formToSpList.DataMapping, data, dataToMark, uniqueColumns, allItems);
                            }
                            catch (ServerException ex)
                            {
                                TOOLS.LogErrorwithoutExitProgram($"Error while sending item {data.Id} from form {data.FormID} to the Sharepoint's list {spList.Id}  : " + ex.Message);
                            }
                        }
                        if (dataToMark.Ids.Count > 0)
                        {
                            response = await HttpClient.PostAsJsonAsync($"{Config.KizeoConfig.Url}/rest/v3/forms/{formId}/markasreadbyaction/{marker}", dataToMark);
                            Log.Debug($"{dataToMark.Ids.Count} data marked");
                        }
                    }
                }
                else
                    TOOLS.LogErrorAndExitProgram("No datamapping was configured, please add a datamapping");

            }
        }


        /// <summary>
        /// upload repports to sharepoint library acording to config file
        /// </summary>
        /// <returns></returns>
        public static async Task UploadExportsToSpLibraries()
        {
            Log.Info($"Uploading exports to SharePoint's library");

            FormDatasRespViewModel formData = null;
            HttpResponseMessage response;
            MarkDataReqViewModel dataToMark;
            string formId;

            foreach (var formToSpLibrary in Config.FormsToSpLibraries)
            {
                string marker = KfApiManager.CreateKFMarker(formToSpLibrary.FormId, formToSpLibrary.SpLibraryId);
                formId = formToSpLibrary.FormId;
                dataToMark = new MarkDataReqViewModel();

                Log.Debug($"-Processing form : {formId}");
                response = await HttpClient.GetAsync($"{Config.KizeoConfig.Url}/rest/v3/forms/{formId}/data/unread/{marker}/50?includeupdated");

                if (response.IsSuccessStatusCode)
                {
                    Log.Debug($"-Loading Sharepoint's library");
                    formToSpLibrary.SpLibrary = SpManager.LoadSpLibrary(formToSpLibrary);

                    formData = await response.Content.ReadAsAsync<FormDatasRespViewModel>();
                    if (formData.Data == null)
                        TOOLS.LogErrorAndExitProgram($"Can not find a form for id {formId}, check if this is a valid id.");

                    Log.Debug($"{formData.Data.Count} Data retrieved successfully from form");

                    var allSpPaths = SpManager.GetAllLibraryFolders(formToSpLibrary.SpLibrary, formToSpLibrary.SpWebSiteUrl);

                    foreach (var data in formData.Data)
                    {
                        Log.Debug($"-Processing data : {data.Id}");
                        var allExportPaths = await GetAllExportsPath(data, formToSpLibrary);

                        Log.Warn("Creating All Folders hierarchy");
                        foreach (var path in allExportPaths)
                        {
                            if (!string.IsNullOrEmpty(path) && !string.IsNullOrWhiteSpace(path) && !allSpPaths.Contains(path))
                            {
                                Log.Warn($"Creating path : {path}");
                                TOOLS.CreateSpPath(Context, path, formToSpLibrary.SpLibrary);
                                allSpPaths.Add(path);
                            }
                        }

                        Log.Debug($"--Processing data : {data.Id}");

                        SpManager.RunExport("PDF", $"rest/v3/forms/{formId}/data/{data.Id}/pdf", formToSpLibrary.ToStandardPdf, formToSpLibrary, data, formToSpLibrary.StandardPdfPath);
                        SpManager.RunExport("Excel", $"rest/v3/forms/{formId}/data/multiple/excel", formToSpLibrary.ToExcelList, formToSpLibrary, data, formToSpLibrary.ExcelListPath, new string[] { data.Id });
                        SpManager.RunExport("CSV", $"rest/v3/forms/{formId}/data/multiple/csv", formToSpLibrary.ToCsv, formToSpLibrary, data, formToSpLibrary.CsvPath, new string[] { data.Id });
                        SpManager.RunExport("CSV_Custom", $"rest/v3/forms/{formId}/data/multiple/csv_custom", formToSpLibrary.ToCsvCustom, formToSpLibrary, data, formToSpLibrary.CsvCustomPath, new string[] { data.Id });
                        SpManager.RunExport("Excel_Custom", $"rest/v3/forms/{formId}/data/multiple/excel_custom", formToSpLibrary.ToExcelListCustom, formToSpLibrary, data, formToSpLibrary.ExcelListCustomPath, new string[] { data.Id });

                        if (formToSpLibrary.Exports != null && formToSpLibrary.Exports.Count > 0)
                        {
                            foreach (var export in formToSpLibrary.Exports)
                            {
                                Log.Debug($"---Processing export : {export.Id}");

                                SpManager.RunExport("Initial Type", $"rest/v3/forms/{formId}/data/{data.Id}/exports/{export.Id}", export.ToInitialType, formToSpLibrary, data, export.initialTypePath);
                                SpManager.RunExport("Pdf Type", $"rest/v3/forms/{formId}/data/{data.Id}/exports/{export.Id}/pdf", export.ToPdf, formToSpLibrary, data, export.PdfPath);

                            }
                        }

                        dataToMark.Ids.Add(data.Id);
                    }

                    if (dataToMark.Ids.Count > 0)
                    {
                        response = await HttpClient.PostAsJsonAsync($"{Config.KizeoConfig.Url}/rest/v3/forms/{formId}/markasreadbyaction/{marker}", dataToMark);
                        Log.Debug($"-{dataToMark.Ids.Count} data marked");
                    }

                }

            }

        }

        /// <summary>
        /// Upload periodicly reports to sharepoint library
        /// </summary>
        public static void RunPeriodicExports()
        {
            foreach (var periodicExport in Config.PeriodicExports)
            {
                string formId = periodicExport.FormId;
                Log.Debug($"-Processing form : {formId}");

                Log.Debug($"-Loading Sharepoint library");
                periodicExport.SpLibrary = SpManager.LoadSpLibrary(periodicExport);

                var allSpPaths = SpManager.GetAllLibraryFolders(periodicExport.SpLibrary, periodicExport.SpWebSiteUrl);
                var allExportPaths = GetAllPeriodicExportsPath(periodicExport);

                Log.Warn("Creating All Folders hierarchy");
                foreach (var path in allExportPaths)
                {
                    if (!string.IsNullOrEmpty(path) && !string.IsNullOrWhiteSpace(path) && !allSpPaths.Contains(path))
                    {
                        Log.Warn($"Creating path : {path}");
                        TOOLS.CreateSpPath(Context, path, periodicExport.SpLibrary);
                        allSpPaths.Add(path);
                    }
                }

                ExportPeriodicly("CSV", periodicExport, periodicExport.ToCsv, periodicExport.CsvPeriod, periodicExport.CsvPath);
                ExportPeriodicly("CSV Custom", periodicExport, periodicExport.ToCsvCustom, periodicExport.CsvCustomPeriod, periodicExport.CsvCustomPath);
                ExportPeriodicly("Excel List", periodicExport, periodicExport.ToExcelList, periodicExport.ExcelListPeriod, periodicExport.ExcelListPath);
                ExportPeriodicly("ExcelList Custom", periodicExport, periodicExport.ToExcelListCustom, periodicExport.ExcelListCustomPeriod, periodicExport.ExcelListCustomPath);
            }
        }

        public static void ExportPeriodicly(string exportType, PeriodicExport periodicExport, bool toExport, int period, string path)
        {
            if (toExport)
            {
                if ((period == 1 || period == 3) && DateTime.Now.Hour == periodicExportHour)
                {
                    ExportBetweenTwoDates(exportType, DateTime.Now.AddDays(-1), DateTime.Now, periodicExport.FormId, periodicExport, path);

                }
                if ((period == 2 || period == 3) && (DateTime.Now.DayOfWeek == DayOfWeek.Sunday) && DateTime.Now.Hour == periodicExportHour)
                {
                    ExportBetweenTwoDates(exportType, DateTime.Now.AddDays(-7), DateTime.Now, periodicExport.FormId, periodicExport, path);
                }

            }
        }

        public static async void ExportBetweenTwoDates(string exportType, DateTime lower, DateTime upper, string formId, PeriodicExport periodicExport, string path)
        {
            AdvancedSearchReqViewModel req = new AdvancedSearchReqViewModel
            {
                Filters = new List<Filters> { new Filters { Type = "and",
                                Components = new List<Components>{ new Components{Field= "_update_time", Operator=">",Value=$"{lower}"},
                                    new Components { Field = "_update_time", Operator = "<", Value = $"{upper}" } } } }
            };
            var response = await HttpClient.PostAsJsonAsync($"{Config.KizeoConfig.Url}/rest/v3/forms/{formId}/data/advanced", req);
            if (response.IsSuccessStatusCode)
            {
                AdvancedSearchResViewModel resp = await response.Content.ReadAsAsync<AdvancedSearchResViewModel>();
                if (resp.RecordTotal > 0)
                {
                    var filenamePrefix = $"{lower:dd-MM-yyyy} to {upper:dd-MM-yyyy}";
                    SpManager.RunPeriodicExport(exportType, $"rest/v3/forms/{formId}/data/multiple/excel", periodicExport, path, resp.data.Select(x => x.Id).ToArray(), filenamePrefix);
                }
            }
        }

        /// <summary>
        /// Get All paths in the config file for an formToSpLibrary config
        /// </summary>
        /// <param name="data"> the data thas holds dataId and FormId</param>
        /// <param name="formToSpLibrary">The formToSpLibrary config where paths will be retrieved</param>
        /// <returns></returns>
        public async static Task<LockedHashset<string>> GetAllExportsPath(FormData data, FormToSpLibrary formToSpLibrary)
        {
            LockedHashset<string> allConfigPaths = new LockedHashset<string>();

            if (formToSpLibrary.ToStandardPdf) allConfigPaths.Add((await KfApiManager.TransformText(data.FormID, data.Id, formToSpLibrary.StandardPdfPath)).First());
            if (formToSpLibrary.ToExcelList) allConfigPaths.Add((await KfApiManager.TransformText(data.FormID, data.Id, formToSpLibrary.ExcelListPath)).First());
            if (formToSpLibrary.ToCsv) allConfigPaths.Add((await KfApiManager.TransformText(data.FormID, data.Id, formToSpLibrary.CsvPath)).First());
            if (formToSpLibrary.ToCsvCustom) allConfigPaths.Add((await KfApiManager.TransformText(data.FormID, data.Id, formToSpLibrary.CsvCustomPath)).First());
            if (formToSpLibrary.ToExcelListCustom) allConfigPaths.Add((await KfApiManager.TransformText(data.FormID, data.Id, formToSpLibrary.ExcelListCustomPath)).First());

            if (formToSpLibrary.Exports != null && formToSpLibrary.Exports.Count > 0)
            {
                var pdfPaths = formToSpLibrary.Exports.Where(e => e.ToPdf).Select(e => e.PdfPath).ToList();

                foreach (var pdfPath in pdfPaths)
                    allConfigPaths.Add((await KfApiManager.TransformText(data.FormID, data.Id, pdfPath)).First());

                var initialTypepaths = formToSpLibrary.Exports.Where(e => e.ToInitialType).Select(e => e.initialTypePath).ToList();
                foreach (var initialTypePath in initialTypepaths)
                    allConfigPaths.Add((await KfApiManager.TransformText(data.FormID, data.Id, initialTypePath)).First());
            }

            return allConfigPaths;
        }


        /// <summary>
        /// initialise timers and threads that are going to execute the 4 tasks
        /// </summary>
        /// <param name="kfToSpsyncTime">time between each interaction from kizeoForms to Sharepoint</param>
        /// <param name="spToKfSyncTime">Time between each interaction from Sharepoint to kizeo</param>
        private static void initTimers(int kfToSpsyncTime, int spToKfSyncTime)
        {
            System.Timers.Timer kfToSpTimer = new System.Timers.Timer();
            kfToSpTimer.Elapsed += new ElapsedEventHandler(KfToSpEvent);
            kfToSpTimer.Interval = kfToSpsyncTime * 60 * 1000;
            kfToSpTimer.Enabled = true;

            System.Timers.Timer spToKfTimer = new System.Timers.Timer();
            spToKfTimer.Elapsed += new ElapsedEventHandler(SpToKfEvent);
            spToKfTimer.Interval = spToKfSyncTime * 60 * 1000;
            spToKfTimer.Enabled = true;

            System.Timers.Timer periodicExportsTimer = new System.Timers.Timer();
            periodicExportsTimer.Elapsed += new ElapsedEventHandler(PeriodicExportsEvent);
            periodicExportsTimer.Interval = 60 * 60 * 1000;
            periodicExportsTimer.Enabled = true;
        }

        private static async void KfToSpEvent(object source, ElapsedEventArgs e)
        {
            try
            {

                await FillSpListsFromKfData();
                await UploadExportsToSpLibraries();

            }
            catch (Exception ex)
            {

                Log.Fatal(ex);
            }
        }

        private static async void SpToKfEvent(object source, ElapsedEventArgs e)
        {

            try
            {

                await FillKfExtListsFromSp();

            }
            catch (Exception ex)
            {

                Log.Fatal(ex);
            }
        }

        private static void PeriodicExportsEvent(object source, ElapsedEventArgs e)
        {

            try
            {

                if (DateTime.Now.Hour == periodicExportHour)
                {
                    RunPeriodicExports();
                }

            }
            catch (Exception ex)
            {

                Log.Fatal(ex);
            }
        }


        /// <summary>
        /// load and parse the config file
        /// </summary>
        /// <param name="ConfFilePath">path of the config file</param>
        /// <returns>instance of Config that holds the current config</returns>
        public static Config GetConfig(string ConfFilePath)
        {
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Kizeo");
            string filePath = Path.Combine(path, ConfFilePath);

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            if (!System.IO.File.Exists(filePath))
            {
                using (FileStream fs = System.IO.File.Create(filePath))
                {

                }
            }

            try
            {
                String jsonText = new StreamReader(filePath).ReadToEnd();
                Config = JsonConvert.DeserializeObject<Config>(jsonText);
            }
            catch (FileNotFoundException ex)
            {
                TOOLS.LogErrorAndExitProgram("file not found " + ex.Message);
            }
            catch (DirectoryNotFoundException ex)
            {
                TOOLS.LogErrorAndExitProgram("File not found" + ex.Message);
            }
            catch (ArgumentNullException ex)
            {
                TOOLS.LogErrorAndExitProgram("no configFile was given" + ex.Message);
            }
            catch (JsonReaderException ex)
            {
                TOOLS.LogErrorAndExitProgram("Error parsing json config" + ex.Message);
            }
            catch (JsonSerializationException ex)
            {
                TOOLS.LogErrorAndExitProgram("Erreur lors de la transforamtion du fichier config en objet :" + ex.Message);
            }
            catch (Exception ex)
            {
                TOOLS.LogErrorAndExitProgram(ex.Message);
            }

            return Config;
        }



        /// <summary>
        /// Get all paths in the config that's gonna be used to export into for a periodic export
        /// </summary>
        /// <param name="periodicExport">the periodic export</param>
        /// <returns>list of strings representing paths</returns>
        public static List<string> GetAllPeriodicExportsPath(PeriodicExport periodicExport)
        {
            List<string> allConfigPaths = new List<string>();

            allConfigPaths.Add(periodicExport.ExcelListPath);
            allConfigPaths.Add(periodicExport.CsvPath);
            allConfigPaths.Add(periodicExport.CsvCustomPath);
            allConfigPaths.Add(periodicExport.ExcelListCustomPath);

            return allConfigPaths;

        }

    }


}


