﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.34014
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace PointLinker.Properties {
    
    
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "11.0.0.0")]
    internal sealed partial class Settings : global::System.Configuration.ApplicationSettingsBase {
        
        private static Settings defaultInstance = ((Settings)(global::System.Configuration.ApplicationSettingsBase.Synchronized(new Settings())));
        
        public static Settings Default {
            get {
                return defaultInstance;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("The selected Excel file does not meet the requirement formatting for the Link Poi" +
            "nts!")]
        public string notMeetFormat {
            get {
                return ((string)(this["notMeetFormat"]));
            }
            set {
                this["notMeetFormat"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("The selected point data include a Z coordinate for the points. 2D sketches do not" +
            " support 3D points.\r\n\r\n Only the X and Y coordinates will be used. Do you wish t" +
            "o continue?         ")]
        public string hasMoreColumn {
            get {
                return ((string)(this["hasMoreColumn"]));
            }
            set {
                this["hasMoreColumn"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("The currently edited object is not a sketch.\r\n\r\nPlease activate the sketch to imp" +
            "ort points.")]
        public string notSketch {
            get {
                return ((string)(this["notSketch"]));
            }
            set {
                this["notSketch"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Excel does not seem to be available!")]
        public string failToGetExcel {
            get {
                return ((string)(this["failToGetExcel"]));
            }
            set {
                this["failToGetExcel"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("error of opening excel file!")]
        public string badExcelFile {
            get {
                return ((string)(this["badExcelFile"]));
            }
            set {
                this["badExcelFile"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Selected file has been imported previously.\r\n\r\nAre you sure you want to continue " +
            "to import a new item?\r\n          ")]
        public string hasImport {
            get {
                return ((string)(this["hasImport"]));
            }
            set {
                this["hasImport"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("The Excel file does not exist, please resolve")]
        public string fileNotExists {
            get {
                return ((string)(this["fileNotExists"]));
            }
            set {
                this["fileNotExists"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_ADNPlugin_PointLinker_Sketch")]
        public string sketchAttSetName {
            get {
                return ((string)(this["sketchAttSetName"]));
            }
            set {
                this["sketchAttSetName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Point Linker Error")]
        public string errorCaption {
            get {
                return ((string)(this["errorCaption"]));
            }
            set {
                this["errorCaption"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_ADNPlugin_PointLinker_Entity_Points_ID")]
        public string sketchPtAttName {
            get {
                return ((string)(this["sketchPtAttName"]));
            }
            set {
                this["sketchPtAttName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("The points will be updated one by one.\r\n\r\nThe lines or spline will be re-created." +
            "")]
        public string samePtCount {
            get {
                return ((string)(this["samePtCount"]));
            }
            set {
                this["samePtCount"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Some previous points may have been missed, or the previous points count is differ" +
            "ent to the new file.\r\nSome new points may be added.\r\nSome old points may break t" +
            "he link with the excel file.\r\nThe lines or spline will be re-created.")]
        public string difPtCount {
            get {
                return ((string)(this["difPtCount"]));
            }
            set {
                this["difPtCount"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Error when update or resolve points!")]
        public string errorUpdatePt {
            get {
                return ((string)(this["errorUpdatePt"]));
            }
            set {
                this["errorUpdatePt"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_ADNPlugin_PointLinker_Entity_Lines_ID")]
        public string sketchLineAttName {
            get {
                return ((string)(this["sketchLineAttName"]));
            }
            set {
                this["sketchLineAttName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_ADNPlugin_PointLinker_Entity_Spline_ID")]
        public string sketchSLAttName {
            get {
                return ((string)(this["sketchSLAttName"]));
            }
            set {
                this["sketchSLAttName"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_ADNPlugin_PointLinker_Entity*")]
        public string findSketchEntity {
            get {
                return ((string)(this["findSketchEntity"]));
            }
            set {
                this["findSketchEntity"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("No points exist with some items, do you want to delete the attribute information?" +
            "")]
        public string whenClose {
            get {
                return ((string)(this["whenClose"]));
            }
            set {
                this["whenClose"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("No units defined in the excel file, use default unit of document?")]
        public string noUnits {
            get {
                return ((string)(this["noUnits"]));
            }
            set {
                this["noUnits"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Lines need at least 2 points!")]
        public string lessPtForLines {
            get {
                return ((string)(this["lessPtForLines"]));
            }
            set {
                this["lessPtForLines"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Spline needs at least 2 points!")]
        public string lessPtForSpline {
            get {
                return ((string)(this["lessPtForSpline"]));
            }
            set {
                this["lessPtForSpline"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("An error occurred! Please contact the provider of the add-in!")]
        public string generalError {
            get {
                return ((string)(this["generalError"]));
            }
            set {
                this["generalError"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("ADNPlugin-PointLinker")]
        public string messageCaption {
            get {
                return ((string)(this["messageCaption"]));
            }
            set {
                this["messageCaption"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Select Excel File")]
        public string capForOpenDlg {
            get {
                return ((string)(this["capForOpenDlg"]));
            }
            set {
                this["capForOpenDlg"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Excel Files|*.xls;*.xlsx")]
        public string filterForOpenDlg {
            get {
                return ((string)(this["filterForOpenDlg"]));
            }
            set {
                this["filterForOpenDlg"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("The status of this item is up to date!")]
        public string upToDate {
            get {
                return ((string)(this["upToDate"]));
            }
            set {
                this["upToDate"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_ADNPlugin_PointLinker_Entity_Lines*")]
        public string attSearchLines {
            get {
                return ((string)(this["attSearchLines"]));
            }
            set {
                this["attSearchLines"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("_ADNPlugin_PointLinker_Entity_Spline*")]
        public string attSeatchSpline {
            get {
                return ((string)(this["attSeatchSpline"]));
            }
            set {
                this["attSeatchSpline"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Failed to create Line/Spline because two adjoining points were the same! ")]
        public string adjoiningPts {
            get {
                return ((string)(this["adjoiningPts"]));
            }
            set {
                this["adjoiningPts"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Failed to open the Excel file!")]
        public string failToOpenWorkbook {
            get {
                return ((string)(this["failToOpenWorkbook"]));
            }
            set {
                this["failToOpenWorkbook"] = value;
            }
        }
        
        [global::System.Configuration.UserScopedSettingAttribute()]
        [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
        [global::System.Configuration.DefaultSettingValueAttribute("Failed to open the sketch entities")]
        public string failToDelete {
            get {
                return ((string)(this["failToDelete"]));
            }
            set {
                this["failToDelete"] = value;
            }
        }
    }
}