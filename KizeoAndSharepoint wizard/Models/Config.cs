using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Net.Http;
using System.Runtime.CompilerServices;
using System.Web.Script.Serialization;

namespace KizeoAndSharepoint_wizard.Models
{
    public class Config : INotifyPropertyChanged
    {
        private SharepointConfig _SharepointConfig;
        private KizeoConfig _KizeoConfig;

        [JsonProperty("sharepoint_config")] public SharepointConfig SharepointConfig { get { return _SharepointConfig; } set { _SharepointConfig = value; OnPropertyChanged(); } }
        [JsonProperty("kizeo_config")] public KizeoConfig KizeoConfig { get { return _KizeoConfig; } set { _KizeoConfig = value; OnPropertyChanged(); } }

        [JsonProperty("forms_to_bind_with_sp_list")]
        public ObservableCollection<FormToSpList> FormsToSpLists { get; set; }
        [JsonProperty("forms_to_export_to_sp_library")]
        public ObservableCollection<FormToSpLibrary> FormsToSpLibraries { get; set; }
        [JsonProperty("scheduled_exports")]
        public ObservableCollection<PeriodicExport> PeriodicExports { get; set; }
        [JsonProperty("sp_lists_to_bind_with_kf_list")]
        public ObservableCollection<SpListToExtList> SpListsToExtLists { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string caller = "")
        {
            if (this.PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(caller));
            }
        }

    }

    public class SharepointConfig : INotifyPropertyChanged
    {
        private string _SPDomain;
        private string _SPTenantID;
        private string _SPClientId;
        private string _SPClientSecret;

        [JsonProperty("domain")]
        public string SPDomain { get { return _SPDomain; } set { _SPDomain = value; this.OnPropertyChanged(); } }
        [JsonProperty("tenant")]
        public string SPTenantID { get { return _SPTenantID; } set { _SPTenantID = value; this.OnPropertyChanged(); } }
        [JsonProperty("client")]
        public string SPClientId { get { return _SPClientId; } set { _SPClientId = value; this.OnPropertyChanged(); } }
        [JsonProperty("secret")]
        public string SPClientSecret { get { return _SPClientSecret; } set { _SPClientSecret = value; this.OnPropertyChanged(); } }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string caller = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(caller));
            }
        }

        [JsonIgnore]
        public ClientContext Context;
    }

    public class KizeoConfig : INotifyPropertyChanged
    {
        private string _Url;
        private string _Token;

        [JsonProperty("url")] public string Url { get { return _Url; } set { _Url = value; OnPropertyChanged(); } }
        [JsonProperty("token")] public string Token { get { return _Token; } set { _Token = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string caller = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(caller));
            }
        }

        [JsonIgnore]
        public HttpClient HttpClient;
    }

    public class FormToSpList : INotifyPropertyChanged
    {
        private string _FormId;
        private Guid _SpListId;

        [JsonProperty("form_id")]
        public string FormId { get { return _FormId; } set { _FormId = value; OnPropertyChanged(); } }
        [JsonProperty("sp_list_id")]
        public Guid SpListId { get { return _SpListId; } set { _SpListId = value; OnPropertyChanged(); } }
        [JsonProperty("data")]
        public ObservableCollection<DataMapping> DataMapping { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string caller = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(caller));
            }
        }

        public override string ToString()
        {
            return $"Form ID KF : {FormId} =>  SpList : {SpListId} ";
        }

    }

    public class FormToSpLibrary : INotifyPropertyChanged
    {
        private string _FormId;
        private Guid _SpLibraryId;
        private bool _ToStandardPdf;
        private string _StandardPdfPath;
        private bool _ToExcelList;
        private string _ExcelListPath;
        private bool _ToExcelListCustom;
        private string _ExcelListCustomPath;
        private bool _ToCsv;
        private string _CsvPath;
        private bool _ToCsvCustom;
        private string _CsvCustomPath;

        [JsonProperty("form_id")]
        public string FormId { get { return _FormId; } set { _FormId = value; OnPropertyChanged(); } }
        [JsonProperty("sp_library_id")]
        public Guid SpLibraryId { get { return _SpLibraryId; } set { _SpLibraryId = value; OnPropertyChanged(); } }
        [JsonProperty("to_standard_pdf")]
        public bool ToStandardPdf { get { return _ToStandardPdf; } set { _ToStandardPdf = value; OnPropertyChanged(); } }
        [JsonProperty("standard_pdf_path")]
        public string StandardPdfPath { get { return _StandardPdfPath; } set { _StandardPdfPath = value; OnPropertyChanged(); } }
        [JsonProperty("to_excel_list")]
        public bool ToExcelList { get { return _ToExcelList; } set { _ToExcelList = value; OnPropertyChanged(); } }
        [JsonProperty("excel_list_path")]
        public string ExcelListPath { get { return _ExcelListPath; } set { _ExcelListPath = value; OnPropertyChanged(); } }
        [JsonProperty("to_excel_list_custom")]
        public bool ToExcelListCustom { get { return _ToExcelListCustom; } set { _ToExcelListCustom = value; OnPropertyChanged(); } }
        [JsonProperty("excel_list_custom_path")]
        public string ExcelListCustomPath { get { return _ExcelListCustomPath; } set { _ExcelListCustomPath = value; OnPropertyChanged(); } }
        [JsonProperty("to_csv")]
        public bool ToCsv { get { return _ToCsv; } set { _ToCsv = value; OnPropertyChanged(); } }
        [JsonProperty("csv_path")]
        public string CsvPath { get { return _CsvPath; } set { _CsvPath = value; OnPropertyChanged(); } }
        [JsonProperty("to_csv_custom")]
        public bool ToCsvCustom { get { return _ToCsvCustom; } set { _ToCsvCustom = value; OnPropertyChanged(); } }
        [JsonProperty("csv_custom_path")]
        public string CsvCustomPath { get { return _CsvCustomPath; } set { _CsvCustomPath = value; OnPropertyChanged(); } }
        [JsonProperty("exports")]
        public ObservableCollection<Export> Exports { get; set; }
        [JsonProperty("metadata")]
        public ObservableCollection<DataMapping> MetaData { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string caller = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(caller));
            }
        }

        public override string ToString()
        {
            return $"Form ID KF : {FormId} =>  SpLibrary : {SpLibraryId} ";
        }

    }

    public class PeriodicExport : INotifyPropertyChanged
    {
        private string _FormId;
        private Guid _SpLibraryId;

        private bool _ToExcelList { get; set; }
        private string _ExcelListPath { get; set; }
        private int _ExcelListPeriod { get; set; }
        private bool _ToExcelListCustom { get; set; }
        private string _ExcelListCustomPath { get; set; }
        private int _ExcelListCustomPeriod { get; set; }
        private bool _ToCsv { get; set; }
        private string _CsvPath { get; set; }
        private int _CsvPeriod { get; set; }
        private bool _ToCsvCustom { get; set; }
        private string _CsvCustomPath { get; set; }
        private int _CsvCustomPeriod { get; set; }


        [JsonProperty("form_id")]
        public string FormId { get { return _FormId; } set { _FormId = value; OnPropertyChanged(); } }
        [JsonProperty("sp_library_id")]
        public Guid SpLibraryId { get { return _SpLibraryId; } set { _SpLibraryId = value; OnPropertyChanged(); } }
        [JsonProperty("to_excel_list")]
        public bool ToExcelList { get { return _ToExcelList; } set { _ToExcelList = value; OnPropertyChanged(); } }
        [JsonProperty("excel_list_path")]
        public string ExcelListPath { get { return _ExcelListPath; } set { _ExcelListPath = value; OnPropertyChanged(); } }
        [JsonProperty("excel_list_period")]
        public int ExcelListPeriod { get { return _ExcelListPeriod; } set { _ExcelListPeriod = value; OnPropertyChanged(); } }
        [JsonProperty("to_excel_list_custom")]
        public bool ToExcelListCustom { get { return _ToExcelListCustom; } set { _ToExcelListCustom = value; OnPropertyChanged(); } }
        [JsonProperty("excel_list_custom_path")]
        public string ExcelListCustomPath { get { return _ExcelListCustomPath; } set { _ExcelListCustomPath = value; OnPropertyChanged(); } }
        [JsonProperty("excel_list_custom_period")]
        public int ExcelListCustomPeriod { get { return _ExcelListCustomPeriod; } set { _ExcelListCustomPeriod = value; OnPropertyChanged(); } }
        [JsonProperty("to_csv")]
        public bool ToCsv { get { return _ToCsv; } set { _ToCsv = value; OnPropertyChanged(); } }
        [JsonProperty("csv_path")]
        public string CsvPath { get { return _CsvPath; } set { _CsvPath = value; OnPropertyChanged(); } }
        [JsonProperty("csv_period")]
        public int CsvPeriod { get { return _CsvPeriod; } set { _CsvPeriod = value; OnPropertyChanged(); } }
        [JsonProperty("to_csv_custom")]
        public bool ToCsvCustom { get { return _ToCsvCustom; } set { _ToCsvCustom = value; OnPropertyChanged(); } }
        [JsonProperty("csv_custom_path")]
        public string CsvCustomPath { get { return _CsvCustomPath; } set { _CsvCustomPath = value; OnPropertyChanged(); } }
        [JsonProperty("csv_custom_period")]
        public int CsvCustomPeriod { get { return _CsvCustomPeriod; } set { _CsvCustomPeriod = value; OnPropertyChanged(); } }


        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string caller = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(caller));
            }
        }

        public override string ToString()
        {
            return $"Form ID KF : {FormId} =>  SpLibrary : {SpLibraryId} ";
        }

        [JsonProperty("period_choice")]
        public ObservableCollection<PeriodicExportsChoices> PeriodicChoices { get; set; }

    }

    public class Export : INotifyPropertyChanged
    {
        private string _Id { get; set; }
        private bool _ToInitialType { get; set; }
        private string _InitialTypePath { get; set; }
        private bool _ToPdf { get; set; }
        private string _PdfPath { get; set; }

        [JsonProperty("id")]
        public string Id { get { return _Id; } set { _Id = value; OnPropertyChanged(); } }
        [JsonProperty("to_initial_type")]
        public bool ToInitialType { get { return _ToInitialType; } set { _ToInitialType = value; OnPropertyChanged(); } }
        [JsonProperty("initial_type_path")]
        public string InitialTypePath { get { return _InitialTypePath; } set { _InitialTypePath = value; OnPropertyChanged(); } }
        [JsonProperty("to_pdf")]
        public bool ToPdf { get { return _ToPdf; } set { _ToPdf = value; OnPropertyChanged(); } }
        [JsonProperty("pdf_path")]
        public string PdfPath { get { return _PdfPath; } set { _PdfPath = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string caller = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(caller));
            }
        }

        public override string ToString()
        {
            return "ExportId : " + Id + " (" + (ToInitialType ? "Original format " + InitialTypePath + (ToPdf ? " " : "") : "") + (ToPdf ? "PDF " + PdfPath : "") + ")";
        }
    }

    public class SpListToExtList : INotifyPropertyChanged
    {
        private string _ExListId;
        private Guid _SpListId;
        private string _DataSchema;

        [JsonProperty("kf_list_id")]
        public string ExListId { get { return _ExListId; } set { _ExListId = value; OnPropertyChanged(); } }
        [JsonProperty("sp_list_id")]
        public Guid SpListId { get { return _SpListId; } set { _SpListId = value; OnPropertyChanged(); } }
        [JsonProperty("sharepoint_data_schema")]
        public string DataSchema { get { return _DataSchema; } set { _DataSchema = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string caller = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(caller));
            }
        }

        public override string ToString()
        {
            return $"SharePoint List : {SpListId} =>  KF External List {ExListId} ";
        }

    }

    public class DataMapping : INotifyPropertyChanged
    {
        private string _SpColumnId;
        private string _KfColumnSelector;
        private string _SpecialType;

        [JsonProperty("sp_column_id")]
        public string SpColumnId { get { return _SpColumnId; } set { _SpColumnId = value; OnPropertyChanged(); } }
        [JsonProperty("kf_column_selector")]
        public string KfColumnSelector { get { return _KfColumnSelector; } set { _KfColumnSelector = value; OnPropertyChanged(); } }
        [JsonProperty("special_type")]
        public string SpecialType { get { return _SpecialType; } set { _SpecialType = value; OnPropertyChanged(); } }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string caller = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(caller));
            }
        }

        public override string ToString()
        {
            return $"SharePoint Column ID : {SpColumnId} <=> KF String {KfColumnSelector} type : {SpecialType}";
        }
    }

    public class PeriodicExportsChoices
    {
        [JsonProperty("id")]
        public int Id { get; set; }
        [JsonProperty("name")]
        public string Name { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }

}


