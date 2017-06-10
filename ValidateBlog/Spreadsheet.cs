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
    class SheetAccessor {
        object[,] SheetData;
        public int LastRow => SheetData.GetUpperBound(0); // 1 would get number of columns, but we don't need that.
        public int Row; // Current row of interest

        public SheetAccessor(string fileName, string sheetName) {
            // Open Excel and fetch the workbook we're interested in
            Excel.Application app = new Excel.Application();
            // Stupid Excel requires an absolute path for this
            string workbookName = Directory.GetCurrentDirectory() + "\\" + fileName;
            app.DisplayAlerts = false; // No stupid dialog boxes!
            Excel.Workbook workbook = app.Workbooks.Open(Filename: workbookName, UpdateLinks: 3, ReadOnly: true);

            // Just get the single sheet we want; the one labeled "stories"
            Excel.Worksheet sheet = workbook.Worksheets[sheetName];

            // Copy the whole sheet at once into an array of objects. This is about 1000 times faster than iterating.
            SheetData = sheet.UsedRange.Value2;

            // Now carefully shut down all the COM crap we just created.
            sheet = null;
            if (workbook != null) {
                workbook.Close(SaveChanges: false);
                Marshal.FinalReleaseComObject(workbook);
            }
            workbook = null;
            if (app != null) {
                app.Quit();
                Marshal.FinalReleaseComObject(app);
                app = null;
            }

            // Not sure why we have to do this twice, but once definitely doesn't work.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        // No reason to compute this more than once. It's always 65.
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
                string s = (string)val;
                return s.Length > 0 ? s : null;
            }
            if (val.GetType() == typeof(double)) {
                return Convert.ToString(val);
            }
            if (val.GetType() == typeof(int)) {
                return null;
            }
            return null;
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
        string _BlogTitle;
        public string[] BlogLabels;
        string PermaLink;
        string Body; // total HTML text of review
        string Review; // inverted portion only
        Uri RSRLink; // Actual link to blog page (dictionary key)
        bool _Reprint;

        public string Title { get { return _Title; } }
        public Uri BloggerLink {get {return RSRLink;} }
        public bool Reprint { get { return _Reprint; } }
        public string BlogTitle => _BlogTitle; 

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
            if (s != null && s != "0") {
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
            _BlogTitle = sheet.GetCell("AE");
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
