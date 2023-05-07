using OfficeOpenXml;
using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;

namespace YouTubeXLSParser
{
    class Program
    {
        static void Main(string[] args)
        {
            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Specify the paths and other settings
            const string inputPath = "input.xlsx";
            const string outputPath = "output.txt";
            const int minimumCommentChars = 206;

            // Load the Excel workbook
            using var package = new ExcelPackage(new FileInfo(inputPath));
            // Get the second worksheet in the workbook
            var worksheet = package.Workbook.Worksheets[1];

            // Create a StringBuilder to store the concatenated values
            var sb = new System.Text.StringBuilder();
            var topLevelCommentCounter = 0;
            var replyNumber = 1;

            // Loop through each row in the worksheet
            for (var row = 1; row <= worksheet.Dimension.Rows; row++)
            {
                //reset replace text
                var replace = "";
                var column3 = worksheet.Cells[row, 3].Value?.ToString();
                var column8 = worksheet.Cells[row, 8].Value?.ToString();

                // guard clause to skip over short comments (except for first)
                if (
                    topLevelCommentCounter != 1 && (
                    (string.IsNullOrWhiteSpace(column3) && column8.Length < minimumCommentChars)
                    ||
                    (string.IsNullOrWhiteSpace(column8) && column3.Length < minimumCommentChars)
                    )
                )
                {
                    continue;
                }
                // top level comments
                if (!string.IsNullOrEmpty(column3))
                {
                    replace = ScrubHtmlAndEmojis(column3).Insert(0, "Comment " + topLevelCommentCounter + ": ");

                    //increment top level comment count
                    topLevelCommentCounter++;

                    //if it's a top level comment than reply number should be reset
                    replyNumber = 1;
                }
                //replies to comment:
                else if (!string.IsNullOrEmpty(column8))
                {
                    replace = ScrubHtmlAndEmojis(column8).Insert(0, "     Reply " + replyNumber + ": ");

                    //increment reply number count
                    replyNumber++;
                }

                // Append top level or comment reply to StringBuilder and add a new line
                if (topLevelCommentCounter == 2 || replace.Length > minimumCommentChars)
                    sb.AppendLine(replace + Environment.NewLine);

            }

            // Write the contents of the StringBuilder to the output file
            File.WriteAllText(outputPath, sb.ToString());

            // Display a message to indicate that the operation is complete
            Console.WriteLine("XLS concatenation complete.");
        }

        public static string ScrubHtmlAndEmojis(string rawComment)
        {
            // this does most of the work of cleaning the comments:
            // removing non-ascii, removing html tags, converting html-encoded items to ascii, ret
            var cleanedComment = Regex.Replace(rawComment, @"<[^>]*(>|$)|&nbsp;|&zwnj;|&raquo;|&laquo;", string.Empty).Trim();
            cleanedComment = HttpUtility.HtmlDecode(Regex.Replace(cleanedComment, @"[^\u0000-\u007F]+", string.Empty).Replace("<br>", Environment.NewLine));
            return cleanedComment;
        }
    }


}


