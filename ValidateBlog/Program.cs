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
            Console.OutputEncoding = Encoding.UTF8;

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

            if (error || args.Length - argument != 2) {
                Console.Error.WriteLine("usage: {0} [-verbose -nonreview] blog.xml spreadsheet.xls\n", System.AppDomain.CurrentDomain.FriendlyName);
                System.Environment.Exit(-1);
            }

            // Get the name of the file that contains the blogger output
            String blogFileName = args[argument++];
            String spreadsheetFileName = args[argument++];

            // Read it and parse it
            Console.WriteLine("Reading Blogger backup feed");
            Syndicator syn = new Syndicator(blogFileName);
            int synErrors = syn.Errors;
            var synDict = syn.Validate(NonReviewFlag, VerboseFlag);

            // Make a new dictionary to hold what we get from the spreadsheet
            Dictionary<Uri, SpreadsheetRecord> SheetDict = new Dictionary<Uri, SpreadsheetRecord>(synDict.Count);

            // Now open the spreadsheet
            Console.WriteLine("\r\nOpening Excel Spreadsheet");
            SheetAccessor sheet = new SheetAccessor(spreadsheetFileName, "stories");
            // Find out how many rows there are
            int rowCount = sheet.LastRow;

            Console.WriteLine("Compare Spreadsheet URLs vs. Blog");
            // Now loop through all the rows, processing them as needed
            int notReviewed = 0;
            int reviewed = 0;
            int urlErrors = 0;
            for (sheet.Row = 3; sheet.Row <= rowCount; ++sheet.Row) {
                SpreadsheetRecord rec = new SpreadsheetRecord(sheet);
                if (rec.BloggerLink != null && !rec.Reprint) {
                    if (synDict.ContainsKey(rec.BloggerLink)) {
                        if (SheetDict.ContainsKey(rec.BloggerLink)) {
                            Console.WriteLine("Unexpected Duplicate Url: {0} is used for {1} and {2}\n", rec.BloggerLink, rec.Title, SheetDict[rec.BloggerLink].Title);
                            ++urlErrors;
                        } else {
                            SheetDict.Add(rec.BloggerLink, rec);
                            ++reviewed;
                        }
                    } else {
                        Console.WriteLine("Spreadsheet contains URL {0} not in blog: {1}", rec.BloggerLink, rec.Title);
                        ++urlErrors;
                    }
                } else {
                    ++notReviewed;
                }
            }
            Console.WriteLine("Spreadsheet contains {0} records, {1} reviewed and {2} not reviewed", rowCount - 2, reviewed, notReviewed);

            Console.WriteLine("Compare Blog URLs vs. Spreadsheet");
            // Now check for spreadsheet items not in the blog
            foreach (Uri blogUrl in synDict.Keys) {
                if (!SheetDict.ContainsKey(blogUrl)) {
                    Console.WriteLine("Blog contains URL {0} not in spreadsheet: {1}!", blogUrl, synDict[blogUrl].Title);
                    ++urlErrors;
                }
            }

            Console.WriteLine("\r\nErrors in labels");
            // Now check for spreadsheet items not in the blog
            int labelErrors = 0;
            foreach (Uri blogUrl in synDict.Keys) {
                if (!SheetDict.ContainsKey(blogUrl)) {
                    continue;
                }
                var sheetItem = SheetDict[blogUrl];
                var blogItem = synDict[blogUrl];
                IEnumerable<string> differenceQuery = blogItem.Labels.Except(sheetItem.BlogLabels);
                bool isDifference = false;
                foreach (string s in differenceQuery) {
                    Console.WriteLine("Blog\t\t{0}", s);
                    isDifference = true;
                    ++labelErrors;
                }
                // We expect blog to be a subset of spreadsheet, normally, so we only print this when we learn it's not.
                if (isDifference) {
                    IEnumerable<string> differenceQuery2 = sheetItem.BlogLabels.Except(blogItem.Labels);
                    foreach (string s in differenceQuery2) {
                        Console.WriteLine("Spreadsheet\t{0}", s);
                    }
                    Console.WriteLine("{0}", synDict[blogUrl].Title);
                    Console.WriteLine("{0}", blogUrl);
                    Console.WriteLine();
                }
            }
            
            Console.WriteLine("\r\nErrors in Titles");
            // Now check for spreadsheet items whose titles don't match
            int titleErrors = 0;
            foreach (Uri blogUrl in synDict.Keys) {

                // These are reported on elsewhere
                if (!SheetDict.ContainsKey(blogUrl)) {
                    continue;
                }
                string sheetTitle = SheetDict[blogUrl].BlogTitle;
                string blogTitle = synDict[blogUrl].Title;
                if (sheetTitle != blogTitle) {
                    Console.WriteLine("Blog Title:\t\t{0}", blogTitle);
                    Console.WriteLine("SpreadSheet Title:\t{0}", sheetTitle);
                    Console.WriteLine("{0}", blogUrl);

                    ++titleErrors;
                }

            }
            Console.WriteLine("Syndication errors: {0}", synErrors);
            Console.WriteLine("URL errors: {0}", urlErrors);
            Console.WriteLine("Label errors: {0}", labelErrors);
            Console.WriteLine("Title errors: {0}", titleErrors);
        }
    }
}
    


