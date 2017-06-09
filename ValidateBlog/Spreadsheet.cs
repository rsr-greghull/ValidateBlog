using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

// Read https://stackoverflow.com/questions/7613898/how-to-read-an-excel-spreadsheet-in-c-sharp-quickly to speed up processing of spreadsheet
// https://stackoverflow.com/questions/14051257/conversion-from-int-array-to-string-array to convert array of objects to strings

namespace ValidateBlog {
    class SheetAccessor : IDisposable {
        Excel.Application App = null;
        Excel.Workbook Workbook = null;
        Excel.Worksheet Sheet = null;
        object[,] SheetData;

        public int Row; // Current row of interest

        public SheetAccessor(string fileName, string sheetName) {
            // Open Excel and fetch the workbook we're interested in
            App = new Excel.Application();
            // Stupid Excel requires an absolute path for this
            string workbookName = Directory.GetCurrentDirectory() + fileName;
            App.DisplayAlerts = false; // No stupid dialog boxes!
            Workbook = App.Workbooks.Open(Filename: workbookName, UpdateLinks: 3, ReadOnly: true);

            // Just get the single sheet we want; the one labeled "stories"
            Sheet = Workbook.Worksheets[sheetName];

            SheetData = Sheet.UsedRange.Value2;
            Sheet = null;
            if (Workbook != null) {
                Workbook.Close();
                Marshal.FinalReleaseComObject(Workbook);
            }
            Workbook = null;
            if (App != null) {
                App.Quit();
                Marshal.FinalReleaseComObject(App);
                App = null;
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // No reason to compute this more than once.
        int ValueOfA = Convert.ToInt32('A');
        // Extract a single value from the current row
        public string GetCell(string col) {
            int colIndex = Convert.ToInt32(col[0]) - ValueOfA;
            if (col.Length == 2) {
                colIndex *= 26;
                colIndex += Convert.ToInt32(col[1]) - ValueOfA + 26; // Starts after 26 1-char labels
            }
            colIndex += 1; // Excel indexes from 1, not 0

            // Excel.Range range = (Excel.Range)Sheet.Cells[Row, col];
            // object val = range.Value2;
            object val = SheetData[Row, colIndex];
            if (val == null) {
                return null;
            }
            if (val.GetType() == typeof(string)) {
                return (string)val;
            }
            if (val.GetType() == typeof(double)) {
                return Convert.ToString(val);
            }
            if (val.GetType() == typeof(int)) {
                return null;
            }
            return null;
        }
        public T GetTypedCell<T>(string col) {
            return (T)((Excel.Range)Sheet.Cells[Row, col]).Value2;
        }

        // Try to dispose of all the COM crap the way you're meant to.

        private bool isDisposed = false;
        public void Dispose() { Dispose(true); GC.SuppressFinalize(this); }
        protected virtual void Dispose(bool disposing) {
            if (!isDisposed) {
                if (disposing) {
                    Sheet = null;
                    if (Workbook != null) {
                        Workbook.Close();
                        Marshal.FinalReleaseComObject(Workbook);
                    }
                    Workbook = null;
                    if (App != null) {
                        App.Quit();
                        Marshal.FinalReleaseComObject(App);
                        App = null;
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            isDisposed = true;
        }
    }


    // Holds the record for a single story
    class SpreadsheetRecord {
        string _Title;
        string[] Authors;
        string[] Editors;
        string[] Translators;
        int Year; // Year of eligibility; for reprints, this is 100-years in the past, so 1912 would mean originally published 2012
        string Magazine; // Name of publication
        string Issue;
        Uri IssueLink;
        string MagIssue;
        Uri StoryLink; // null if story not accessible
        DateTime PublicationDate;
        DateTime ReviewDate;
        int WordCount;
        string Category;
        string SubGenre;
        string Blurb;
        int Rating;
        string Note;
        string Series;
        string Pitch;
        string SffCat;
        string SettingTime;
        string SettingPlace;
        string Tone;
        string[] Keywords;
        string Protagonist;
        string BlogTitle;
        public string[] BlogLabels;
        string PermaLink;
        string Body; // total HTML text of review
        string Review; // inverted portion only
        Uri RSRLink; // Actual link to blog page (dictionary key)
        bool _Reprint;

        public string Title { get { return _Title; } }
        public Uri BloggerLink {get {return RSRLink;} }
        public bool Reprint { get { return _Reprint; } }

        public SpreadsheetRecord(SheetAccessor sheet) {
            string s;
            _Title = sheet.GetCell("A");
            Authors = MakeStringArray(sheet.GetCell("B"), sheet.GetCell("C"));
            Translators = MakeStringArray(sheet.GetCell("D"), sheet.GetCell("E"));
            Editors = MakeStringArray(sheet.GetCell("F"), sheet.GetCell("G"));
            s = sheet.GetCell("H");
            _Reprint = false;
            if (s != null) {
                Year = int.Parse(s);
                if (Year < 2015) {
                    Year += 100;
                    _Reprint = true;
                }
            }
            Magazine = sheet.GetCell("I");
            Issue = sheet.GetCell("J");
            s = sheet.GetCell("K");
            if (s != null) {
                IssueLink = new Uri(s);
            }
            MagIssue = sheet.GetCell("L");
            s = sheet.GetCell("N");
            if (s != null) {
                StoryLink = new Uri(s);
            }
            s = sheet.GetCell("P");
            if (s != null) {
                double x;
                if (Double.TryParse(s, out x)) {
                    PublicationDate = DateTime.FromOADate((double)x);
                }
            }
            s = sheet.GetCell("Q");
            if (s != null) {
                double x;
                if (Double.TryParse(s, out x)) {
                    ReviewDate = DateTime.FromOADate((double)x);
                }
            }
            s = sheet.GetCell("Q");
            if (s != null) {
                WordCount = int.Parse(s);
            }
            Category = sheet.GetCell("S");
            SubGenre = sheet.GetCell("T");
            Blurb = sheet.GetCell("U");
            s = sheet.GetCell("V");
            if (s != null) {
                Rating = int.Parse(s);
            }
            Note = sheet.GetCell("W");
            Series = sheet.GetCell("X");
            Pitch = sheet.GetCell("Y");
            SffCat = sheet.GetCell("Z");
            SettingTime = sheet.GetCell("AA");
            SettingPlace = sheet.GetCell("AB");
            Tone = sheet.GetCell("AC");
            // Too much is packed into the keyword string
            s = sheet.GetCell("AD");
            if (s != null) {
                string[] sa = s.Split('|');
                if (sa[0].Length > 0) {
                    Keywords = sa[0].Split(',');
                }
                if (sa.Length > 1) {
                    Protagonist = sa[1];
                }
            }
            BlogTitle = sheet.GetCell("AE");
            s = sheet.GetCell("AF");
            if (s != null) {
                BlogLabels = s.Split(',');
                for (int i = 0; i < BlogLabels.Length; ++i) {
                    BlogLabels[i] = BlogLabels[i].Trim(' ');
                }
            }
            PermaLink = sheet.GetCell("AG");
            Body = sheet.GetCell("AR");
            Review = sheet.GetCell("AS");
            s = sheet.GetCell("AT");
            if (s != null && s.Length > 0) {
                RSRLink = new Uri(s);
            }

        }

        // Take two strings, possibly containing substrings separated by | characters.
        // Convert to a string array.
        string[] MakeStringArray(string s1, string s2) {
            string s = s1 + "|" + s2;
            s = s.TrimEnd('|'); // Second string might not exist
            string[] sArray = s.Split('|');
            if (s.Length == 0) {
                sArray = null;
            }
            return sArray;
        }
    }
    
}
