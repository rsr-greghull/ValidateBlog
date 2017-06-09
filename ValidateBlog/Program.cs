using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.ServiceModel.Syndication;



namespace ValidateBlog {
    class Program {
        static bool NonReviewFlag = false; // Display pages that aren't reviews
        static bool VerboseFlag = false; // Print more extensive results

        static void Main(string[] args) {
            // Name of file with records of EPH data per category is the only argument
            // This file is formatted by a Perl script, so we can depend on it to be well-formed

            // Walk through any switches
            int argument = 0;
            bool error = false;
            while (!error && args.Length > argument && args[argument][0] == '-') {
                switch (args[argument]) {
                case "-nonreview":
                    NonReviewFlag = true;
                    break;
                case "-verbose":
                    VerboseFlag = true;
                    break;
                default:
                    error = true;
                    System.Environment.Exit(-1);
                    break;
                }
                ++argument;
            }

            if (error || args.Length - argument != 1) {
                Console.Error.WriteLine("usage: {0} [-verbose -nonreview] blog.xml\n", System.AppDomain.CurrentDomain.FriendlyName);
                System.Environment.Exit(-1);
            }

            // Get the name of the file that contains the blogger output
            String blogFileName = args[argument++];

            // Read it and parse it
            Console.WriteLine("Reading Blogger backup feed");
            Syndicator syn = new Syndicator(blogFileName);
            int errors = syn.Errors;
            var synDict = syn.Validate(NonReviewFlag, VerboseFlag);

            // Make a new dictionary to hold what we get from the spreadsheet
            Dictionary<Uri, SpreadsheetRecord> SheetDict = new Dictionary<Uri, SpreadsheetRecord>(synDict.Count);

            // Now open the spreadsheet
            Console.WriteLine("Opening Excel Spreadsheet");
            using (SheetAccessor sheet = new SheetAccessor("\\RSR - Stories Issues Magazines.xlsx", "stories")) { 

                // Find out how many rows there are
                sheet.Row = 3;
                while (true) {
                    string title = sheet.GetCell("A");
                    if (title == null || title.Length == 0) {
                        break;
                    }
                    ++sheet.Row;
                }
                int rowCount = sheet.Row; // First row in spreadsheet with a blank title

                Console.WriteLine("Reading Excel Spreadsheet");
                // Now loop through all the rows, processing them as needed
                int notReviewed = 0;
                int reviewed = 0;
                for (sheet.Row = 3; sheet.Row < rowCount; ++sheet.Row) {
                    SpreadsheetRecord rec = new SpreadsheetRecord(sheet);
                    if (rec.BloggerLink != null && !rec.Reprint) {
                        if (synDict.ContainsKey(rec.BloggerLink)) {
                            if (SheetDict.ContainsKey(rec.BloggerLink)) {
                                Console.WriteLine("Unexpected Duplicate Url: {0} is used for {1} and {2}\n", rec.BloggerLink, rec.Title, SheetDict[rec.BloggerLink].Title);
                                ++errors;
                            } else {
                                SheetDict.Add(rec.BloggerLink, rec);
                                ++reviewed;
                            }
                        } else {
                            Console.WriteLine("Spreadsheet contains URL {0} not in blog: {1}", rec.BloggerLink, rec.Title);
                            ++errors;
                        }
                    } else {
                        ++notReviewed;
                    }
                }
                Console.WriteLine("Spreadsheet contains {0} records, {1} reviewed and {2} not reviewed", rowCount - 2, reviewed, notReviewed);
            }

            Console.WriteLine("Cross-reference against Blog");
            // Now check for spreadsheet items not in the blog
            foreach (Uri blogUrl in synDict.Keys) {
                if (SheetDict.ContainsKey(blogUrl)) {
                    var sheetItem = SheetDict[blogUrl];
                    var blogItem = synDict[blogUrl];
                    IEnumerable<string> differenceQuery = blogItem.Labels.Except(sheetItem.BlogLabels);
                    bool isDifference = false;
                    foreach (string s in differenceQuery) {
                        Console.WriteLine("Blog contains label {0} not in spreadsheet: {1} {2}", s, blogUrl, synDict[blogUrl].Title);
                        isDifference = true;
                        ++errors;
                    }
                    // We expect blog to be a subset of spreadsheet, normally, so we only print this when we learn it's not.
                    if (isDifference) {
                        IEnumerable<string> differenceQuery2 = sheetItem.BlogLabels.Except(blogItem.Labels);
                        foreach (string s in differenceQuery2) {
                            Console.WriteLine("Spreadsheet contains label {0} not in blog: {1} {2}", s, blogUrl, synDict[blogUrl].Title);
                        }
                    }
                } else {
                    Console.WriteLine("Blog contains URL {0} not in spreadsheet: {1}", blogUrl, synDict[blogUrl].Title);
                    ++errors;
                }
            }
            Console.WriteLine("Total errors: {0}", errors);

        }
    }
}
    


