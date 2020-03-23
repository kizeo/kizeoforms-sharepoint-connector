using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace TestClientObjectModel
{

    class Config
    {
        [JsonProperty("sharepoint_config")]
        public SharepointConfig SharepointConfig { get; set; }

        [JsonProperty("kizeo_config")]
        public KizeoConfig KizeoConfig { get; set; }

        [JsonProperty("forms_to_bind_with_sp_list")]
        public List<FormToSpList> FormsToSpLists { get; set; }

        [JsonProperty("forms_to_export_to_sp_library")]
        public List<FormToSpLibrary> FormsToSpLibraries { get; set; }

        [JsonProperty("scheduled_exports")]
        public List<PeriodicExport> PeriodicExports { get; set; }
        
        [JsonProperty("sp_lists_to_bind_with_kf_list")]
        public List<SpListToExtList> SpListsToExtLists { get; set; }

    }

    class SharepointConfig
    {
        [JsonProperty("domain")]
        public string SPDomain { get; set; }
        [JsonProperty("tenant")]
        public string SPTenantID { get; set; }
        [JsonProperty("client")]
        public string SPClientId { get; set; }
        [JsonProperty("secret")]
        public string SPClientSecret { get; set; }


    }

    class KizeoConfig
    {
        [JsonProperty("url")]
        public string Url { get; set; }
        [JsonProperty("token")]
        public string Token { get; set; }

       
    }

    class FormToSpList
    {
        [JsonProperty("form_id")]
        public string FormId { get; set; }
        [JsonProperty("sp_list_id")]
        public Guid SpListId { get; set; }
        [JsonProperty("data")]
        public List<DataMapping> DataMapping { get; set; }

    }

    class FormToSpLibrary
    {
        [JsonProperty("form_id")]
        public string FormId { get; set; }
        [JsonProperty("sp_library_id")]
        public Guid SpLibraryId { get; set; }
        [JsonProperty("to_standard_pdf")]
        public bool ToStandardPdf { get; set; }
        [JsonProperty("standard_pdf_path")]
        public string StandardPdfPath { get; set; }
        [JsonProperty("to_excel_list")]
        public bool ToExcelList { get; set; }
        [JsonProperty("excel_list_path")]
        public string ExcelListPath { get; set; }
        [JsonProperty("to_excel_list_custom")]
        public bool ToExcelListCustom { get; set; }
        [JsonProperty("excel_list_custom_path")]
        public string ExcelListCustomPath { get; set; }
        [JsonProperty("to_csv")]
        public bool ToCsv { get; set; }
        [JsonProperty("csv_path")]
        public string CsvPath { get; set; }
        [JsonProperty("to_csv_custom")]
        public bool ToCsvCustom { get; set; }
        [JsonProperty("csv_custom_path")]
        public string CsvCustomPath { get; set; }
        [JsonProperty("exports")]
        public List<Export> Exports { get; set; }
        [JsonProperty("metadata")]
        public List<DataMapping> MetaData { get; set; }

        public string LibraryName { get; set; }
        public string SpWebSiteUrl { get; set; }
        public List SpLibrary { get; set; }

    }

    class PeriodicExport
    {
        [JsonProperty("form_id")]
        public string FormId { get; set; }
        [JsonProperty("sp_library_id")]
        public Guid SpLibraryId { get; set; }
     
        [JsonProperty("to_excel_list")]
        public bool ToExcelList { get; set; }
       [JsonProperty("excel_list_path")]
        public string ExcelListPath { get; set; }
        [JsonProperty("excel_list_period")]
        public int ExcelListPeriod { get; set; }
        [JsonProperty("to_excel_list_custom")]
        public bool ToExcelListCustom { get; set; }
        [JsonProperty("excel_list_custom_path")]
        public string ExcelListCustomPath { get; set; }
        [JsonProperty("excel_list_custom_period")]
        public int ExcelListCustomPeriod { get; set; }
        [JsonProperty("to_csv")]
        public bool ToCsv { get; set; }
        [JsonProperty("csv_path")]
        public string CsvPath { get; set; }
        [JsonProperty("csv_period")]
        public int CsvPeriod { get; set; }
        [JsonProperty("to_csv_custom")]
        public bool ToCsvCustom { get; set; }
        [JsonProperty("csv_custom_path")]
        public string CsvCustomPath { get; set; }
        [JsonProperty("csv_custom_period")]
        public int CsvCustomPeriod { get; set; }

        public string LibraryName { get; set; }
        public string SpWebSiteUrl { get; set; }
        public List SpLibrary { get; set; }

    }

    class Export
    {
        [JsonProperty("id")]
        public string Id { get; set; }
        [JsonProperty("to_initial_type")]
        public bool ToInitialType { get; set; }
        [JsonProperty("initial_type_path")]
        public string initialTypePath { get; set; }
        [JsonProperty("to_pdf")]
        public bool ToPdf { get; set; }
        [JsonProperty("pdf_path")]
        public string PdfPath { get; set; }

    }

    class SpListToExtList
    {
        [JsonProperty("kf_list_id")]
        public string ExListId { get; set; }
        [JsonProperty("sp_list_id")]
        public Guid SpListId { get; set; }
        [JsonProperty("sharepoint_data_schema")]
        public string DataSchema { get; set; }

    }

    public class DataMapping
    {
        [JsonProperty("sp_column_id")]
        public string SpColumnId { get; set; }
        [JsonProperty("kf_column_selector")]
        public string KfColumnSelector { get; set; }
        [JsonProperty("special_type")]
        public string SpecialType { get; set; }
    }


}
