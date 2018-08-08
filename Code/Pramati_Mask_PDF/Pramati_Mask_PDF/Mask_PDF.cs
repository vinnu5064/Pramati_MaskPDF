using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Activities;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Drawing;

namespace Pramati_Mask_PDF
{
    public class Mask_PDF : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> InputFilePath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> OutputDirPath { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> MaskingKeyWord { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<string> Occurrence { get; set; }

        [Category("Input")]
        [RequiredArgument]
        public InArgument<bool> MaskOnlyKeyword { get; set; }

        [Category("Output")]
        public OutArgument<string> Result { get; set; }

        public static string maskingStartCoordinates = string.Empty;
        public static bool maskingSuccessfull = false;
        protected override void Execute(CodeActivityContext context)
        {
            Result.Set(context, "Not Masked");
            string inputFilePath = InputFilePath.Get(context);
            string outputDirPath = OutputDirPath.Get(context);
            string maskingKeyWord = MaskingKeyWord.Get(context);
            string occurrenceValue = Occurrence.Get(context);
            bool maskOnlyKeyword = MaskOnlyKeyword.Get(context);
            if (!string.IsNullOrEmpty(inputFilePath) && !string.IsNullOrEmpty(outputDirPath) && !string.IsNullOrEmpty(maskingKeyWord)
                && maskingKeyWord.IndexOf(" ") < 0 && !string.IsNullOrEmpty(occurrenceValue) && File.Exists(inputFilePath) &&
                Directory.Exists(outputDirPath))
            {
                string redactedPDFNameEscaped = Regex.Replace(Path.GetFileName(inputFilePath), "%", "%%");
                string redactedPDFPath = $@"{outputDirPath}\{redactedPDFNameEscaped}";
                string jpgFilesDir = ConvertPDFToJPGs(inputFilePath);
                if (maskOnlyKeyword)
                {
                    if (!string.IsNullOrEmpty(occurrenceValue) && occurrenceValue.Equals("ALL", StringComparison.OrdinalIgnoreCase))
                    {
                        int occurrence = 0;
                        var redactedJPGFiles = RedactOnlyKeyword(jpgFilesDir, maskingKeyWord, occurrence, maskOnlyKeyword);
                        if (maskingSuccessfull)
                        {
                            ConvertJPGsToPDF(redactedJPGFiles, redactedPDFPath);
                            Result.Set(context, "Success");
                        }
                    }
                    else if (!string.IsNullOrEmpty(occurrenceValue))
                    {
                        if (IsDigitsOnly(occurrenceValue))
                        {
                            int occurrence = Convert.ToInt32(occurrenceValue);
                            var redactedJPGFiles = RedactOnlyKeyword(jpgFilesDir, maskingKeyWord, occurrence, maskOnlyKeyword);
                            if (maskingSuccessfull)
                            {
                                ConvertJPGsToPDF(redactedJPGFiles, redactedPDFPath);
                                Result.Set(context, "Success");
                            }
                        }
                        else
                        {
                            throw new Exception("Occurrence should contain only digits.");
                        }
                    }
                }
                else if (!string.IsNullOrEmpty(occurrenceValue))
                {
                    if (IsDigitsOnly(occurrenceValue))
                    {
                        int occurrence = Convert.ToInt32(occurrenceValue);
                        var redactedJPGFiles = RedactTextTillEnd(jpgFilesDir, maskingKeyWord, occurrence, maskOnlyKeyword);
                        if (maskingSuccessfull)
                        {
                            ConvertJPGsToPDF(redactedJPGFiles, redactedPDFPath);
                            Result.Set(context, "Success");
                        }
                    }
                    else
                    {
                        throw new Exception("Occurrence should contain only digits.");
                    }
                }
                Directory.Delete(jpgFilesDir, recursive: true);
            }
            else
            {
                if (!string.IsNullOrEmpty(inputFilePath) && File.Exists(inputFilePath))
                {
                    throw new Exception("Input File Path is not correct.");
                }
                else if (!string.IsNullOrEmpty(outputDirPath) && Directory.Exists(outputDirPath))
                {
                    throw new Exception("Directory Path is not correct.");
                }
                else if (!string.IsNullOrEmpty(maskingKeyWord) && maskingKeyWord.IndexOf(" ") < 0)
                {
                    throw new Exception("Masking Keyword is not correct.");
                }
                else if (!string.IsNullOrEmpty(occurrenceValue))
                {
                    throw new Exception("Occurrence Value should not be null.");
                }
            }
        }

        private static List<string> RedactOnlyKeyword(string jpgFilesDir, string maskingKeyWord, int occurrence, bool maskOnlyKeyword)
        {
            int maskingTextCount = 0;
            var jpgFiles = Directory.EnumerateFiles(jpgFilesDir)
                .OrderBy(file => Utils.ParsePageNumber(file))
                .ToList();

            foreach (string jpgFile in jpgFiles)
            {
                string hocrFile = ConvertJPGToHocr(jpgFile);
                string hocrContent = File.ReadAllText(hocrFile);
                if (hocrContent.ToLower().Contains(maskingKeyWord.ToLower()))
                {
                    HtmlDocument matchesHtmlDoc = new HtmlDocument();
                    matchesHtmlDoc.LoadHtml(hocrContent);
                    string spanClass = "ocr_line";
                    HtmlNodeCollection innerSpans = matchesHtmlDoc.DocumentNode.SelectNodes("//span[@class = '" + spanClass + "']/span");
                    for (int i = 0; i < innerSpans.Count; i++)
                    {
                        if (innerSpans[i].InnerHtml.ToLower().Contains(maskingKeyWord.ToLower()))
                        {
                            maskingTextCount++;
                            if (occurrence > 0 && occurrence == maskingTextCount)
                            {
                                string divOuterHTMLKeyword = innerSpans[i].OuterHtml;
                                string startcorKeyword = divOuterHTMLKeyword.Substring(divOuterHTMLKeyword.IndexOf("bbox") + 5);
                                maskingStartCoordinates = startcorKeyword.Substring(0, startcorKeyword.IndexOf("; x_wconf"));
                                break;
                            }
                            else if (occurrence == 0)
                            {
                                string divOuterHTMLKeyword = innerSpans[i].OuterHtml;
                                string startcorKeyword = divOuterHTMLKeyword.Substring(divOuterHTMLKeyword.IndexOf("bbox") + 5);
                                maskingStartCoordinates = startcorKeyword.Substring(0, startcorKeyword.IndexOf("; x_wconf"));
                                if (!string.IsNullOrEmpty(maskingStartCoordinates))
                                {
                                    string[] arrFinalStartCoordinates = maskingStartCoordinates.Split(' ').ToArray();
                                    Coordinates coordinates = new Coordinates(Convert.ToInt32(arrFinalStartCoordinates[0]),
                                        Convert.ToInt32(arrFinalStartCoordinates[1]),
                                        Convert.ToInt32(arrFinalStartCoordinates[2]), Convert.ToInt32(arrFinalStartCoordinates[3]));
                                    maskingStartCoordinates = string.Empty;
                                    ChangePixelColor(jpgFile, coordinates);
                                }
                            }
                        }
                    }
                    if (!string.IsNullOrEmpty(maskingStartCoordinates))
                    {
                        string[] arrFinalStartCoordinates = maskingStartCoordinates.Split(' ').ToArray();
                        Coordinates coordinates = new Coordinates(Convert.ToInt32(arrFinalStartCoordinates[0]),
                            Convert.ToInt32(arrFinalStartCoordinates[1]),
                            Convert.ToInt32(arrFinalStartCoordinates[2]), Convert.ToInt32(arrFinalStartCoordinates[3]));
                        maskingStartCoordinates = string.Empty;
                        ChangePixelColor(jpgFile, coordinates);
                        break;
                    }
                }
            }
            return jpgFiles;
        }

        private static List<string> RedactTextTillEnd(string jpgFilesDir, string maskingKeyWord, int occurrence, bool maskOnlyKeyword)
        {
            if (occurrence == 0)
            {
                throw new Exception("Occurence should be greater than 0.");
            }
            else
            {
                int leftXValue = 0;
                int leftXValueOfMaskingWord = 0;
                bool occured = false;
                int maskingTextCount = 0;
                var jpgFiles = Directory.EnumerateFiles(jpgFilesDir)
                    .OrderBy(file => Utils.ParsePageNumber(file))
                    .ToList();

                foreach (string jpgFile in jpgFiles)
                {
                    string hocrFile = ConvertJPGToHocr(jpgFile);
                    string hocrContent = File.ReadAllText(hocrFile);
                    if (hocrContent.ToLower().Contains(maskingKeyWord.ToLower()))
                    {
                        HtmlDocument matchesHtmlDoc = new HtmlDocument();
                        matchesHtmlDoc.LoadHtml(hocrContent);
                        string spanClass = "ocr_line";
                        HtmlNodeCollection innerSpans = matchesHtmlDoc.DocumentNode.SelectNodes("//span[@class = '" + spanClass + "']/span");
                        for (int i = 0; i < innerSpans.Count; i++)
                        {
                            if (innerSpans[i].InnerHtml.ToLower().Contains(maskingKeyWord.ToLower()))
                            {
                                maskingTextCount++;
                                if (occurrence == maskingTextCount)
                                {
                                    occured = true;
                                    string divOuterHTMLKeyword = innerSpans[i].OuterHtml;
                                    string startcorKeyword = divOuterHTMLKeyword.Substring(divOuterHTMLKeyword.IndexOf("bbox") + 5);
                                    string startCoordinatesKeyword = startcorKeyword.Substring(0, startcorKeyword.IndexOf("; x_wconf"));
                                    string[] arrStartCoordinatesKeyword = startCoordinatesKeyword.Split(' ').ToArray();
                                    arrStartCoordinatesKeyword[1] = Convert.ToString(Convert.ToInt32(arrStartCoordinatesKeyword[1]));
                                    arrStartCoordinatesKeyword[3] = Convert.ToString(Convert.ToInt32(arrStartCoordinatesKeyword[3]));

                                    string divOuterHTML = innerSpans[i + 1].OuterHtml;
                                    string startcor = divOuterHTML.Substring(divOuterHTML.IndexOf("bbox") + 5);
                                    string startCoordinates = startcor.Substring(0, startcor.IndexOf("; x_wconf"));
                                    string[] arrStartCoordinates = startCoordinates.Split(' ').ToArray();
                                    leftXValueOfMaskingWord = Convert.ToInt32(arrStartCoordinatesKeyword[0]);
                                    leftXValue = Convert.ToInt32(arrStartCoordinates[0]);

                                    if (leftXValueOfMaskingWord < leftXValue)
                                    {
                                        maskingStartCoordinates = String.Concat(arrStartCoordinates[0], " ",
                                        arrStartCoordinatesKeyword[1], " ", arrStartCoordinates[2], " ", arrStartCoordinatesKeyword[3]);
                                    }
                                    else
                                    {
                                        maskingStartCoordinates = String.Concat(Convert.ToString(leftXValueOfMaskingWord), " ",
                                        arrStartCoordinatesKeyword[1], " ", arrStartCoordinates[2], " ", arrStartCoordinatesKeyword[3]);
                                    }
                                }
                            }
                            else
                            {
                                if (occured)
                                {
                                    if (leftXValueOfMaskingWord < leftXValue)
                                    {
                                        string divOuterHTML = innerSpans[i + 1].OuterHtml;
                                        string startcor = divOuterHTML.Substring(divOuterHTML.IndexOf("bbox") + 5);
                                        string startCoordinates = startcor.Substring(0, startcor.IndexOf("; x_wconf"));
                                        string[] arrStartCoordinates = startCoordinates.Split(' ').ToArray();

                                        string[] endYCoordinates = maskingStartCoordinates.Split(' ').ToArray();

                                        if (Convert.ToInt32(endYCoordinates[2]) < Convert.ToInt32(arrStartCoordinates[2]))
                                        {
                                            int rightYCoordinates = Convert.ToInt32(endYCoordinates[2]) < Convert.ToInt32(arrStartCoordinates[2]) ?
                                                Convert.ToInt32(arrStartCoordinates[2]) : Convert.ToInt32(endYCoordinates[2]);

                                            maskingStartCoordinates = String.Concat(endYCoordinates[0], " ",
                                            endYCoordinates[1], " ", rightYCoordinates, " ", endYCoordinates[3]);
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                        if (!string.IsNullOrEmpty(maskingStartCoordinates))
                        {
                            string[] arrFinalStartCoordinates = maskingStartCoordinates.Split(' ').ToArray();
                            Coordinates coordinates = new Coordinates(Convert.ToInt32(arrFinalStartCoordinates[0]),
                                Convert.ToInt32(arrFinalStartCoordinates[1]),
                                Convert.ToInt32(arrFinalStartCoordinates[2]), Convert.ToInt32(arrFinalStartCoordinates[3]));
                            maskingStartCoordinates = string.Empty;
                            ChangePixelColor(jpgFile, coordinates);
                            break;
                        }
                    }
                }
                return jpgFiles;
            }
        }

        private static string ConvertPDFToJPGs(string pdfFilePath)
        {
            string pdfFileDirectory = Path.GetDirectoryName(pdfFilePath);
            string pdfFileName = Path.GetFileName(pdfFilePath);

            Guid guid = Guid.NewGuid();
            string jpegFilesDirName = guid.ToString().Substring(0, 3) + "_" + Path.GetFileNameWithoutExtension(pdfFilePath);
            string jpegFilesDirPath = Path.Combine(pdfFileDirectory, jpegFilesDirName);
            Directory.CreateDirectory(jpegFilesDirPath);

            string magickCommand = @"C:\Program Files\ImageMagick-7.0.7-Q16\magick.exe";
            string magickCommandArgs = $@"convert -limit memory unlimited -density 300 -quality 100 ""{pdfFilePath}"" ""{jpegFilesDirPath}\{jpegFilesDirName}.jpg""";
            Utils.RunCommand(magickCommand, magickCommandArgs, ConfigurationSettings.AppSettings["image_magick_install_dir"]);
            return jpegFilesDirPath;
        }

        private static void ConvertJPGsToPDF(List<String> jpgFilePaths, string redactedPDFPath)
        {
            string spaceSepJPGFiles = jpgFilePaths.Aggregate("", (acc, jpgFilePath) => $@"{acc}""{jpgFilePath}"" ").Trim();

            string magickCommand = @"C:\Program Files\ImageMagick-7.0.7-Q16\magick.exe";
            string magickCommandArgs = $@"convert -limit memory unlimited -interlace Plane -sampling-factor 4:2:0 -quality 70 -density 150x150 -units PixelsPerInch -resize 1241x1754 -repage 1241x1754 {spaceSepJPGFiles} ""{redactedPDFPath}""";
            Utils.RunCommand(magickCommand, magickCommandArgs, ConfigurationSettings.AppSettings["image_magick_install_dir"]);
        }

        private static void ChangePixelColor(string jpgFile, Coordinates coordinates)
        {
            try
            {
                //Bitmap sourceBitmap = (Bitmap)Image.FromFile(jpgFile);
                //Bitmap dupBitmap = new Bitmap(sourceBitmap.Width, sourceBitmap.Height, sourceBitmap.PixelFormat);

                Bitmap sourceBitmap = (Bitmap)Image.FromFile(jpgFile);
                Bitmap dupBitmap = new Bitmap(sourceBitmap.Width, sourceBitmap.Height);

                using (var gr = Graphics.FromImage(dupBitmap))
                    gr.DrawImage(sourceBitmap, new Rectangle(0, 0, sourceBitmap.Width, sourceBitmap.Height));

                sourceBitmap.Dispose();
                File.Delete(jpgFile);

                for (int i = Convert.ToInt32(coordinates.LeftX); i <= Convert.ToInt32(coordinates.RightX); i++)
                    for (int j = Convert.ToInt32(coordinates.LeftY); j <= Convert.ToInt32(coordinates.RightY); j++)
                        dupBitmap.SetPixel(i, j, Color.White);

                dupBitmap.Save(jpgFile);
                maskingSuccessfull = true;
            }
            catch (Exception exception)
            {
                Console.WriteLine("Exception occured while chaning pixel color");
                Console.WriteLine(exception.StackTrace);
                throw exception;
            }
        }

        private static string ConvertJPGToHocr(string jpegFilePath)
        {
            string hocrFilePath = jpegFilePath.Replace(Path.GetExtension(jpegFilePath), "");

            string tesseractCommand = @"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe";
            string tesseractArgs = $@"""{jpegFilePath}"" ""{hocrFilePath}"" -l eng -psm 1 hocr";
            Utils.RunCommand(tesseractCommand, tesseractArgs, ConfigurationSettings.AppSettings["tesseract_install_dir"]);
            return $"{hocrFilePath}.hocr";
        }

        private static bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                    return false;
            }
            return true;
        }
    }
}

class Coordinates
{
    public int LeftX { get; }
    public int LeftY { get; }
    public int RightX { get; }
    public int RightY { get; }

    public Coordinates(int leftX, int leftY, int rightX, int rightY)
    {
        this.LeftX = leftX;
        this.LeftY = leftY;
        this.RightX = rightX;
        this.RightY = rightY;
    }

    public override string ToString()
    {
        return $"{LeftX} {LeftY} {RightX} {RightY}";
    }
}

class CooridnatesYComparer : IComparer<Coordinates>
{
    public int Compare(Coordinates coord1, Coordinates coord2)
    {
        if (Math.Abs(coord1.LeftY - coord2.LeftY) <= 30 || coord1.LeftY == coord2.LeftY)
            return 0;
        else if (coord1.LeftY < coord2.LeftY)
            return -1;
        else
            return 1;
    }
}

class CoordinatesXComparer : IComparer<Coordinates>
{
    public int Compare(Coordinates coord1, Coordinates coord2)
    {
        if (coord1.LeftX < coord2.LeftX)
            return -1;
        else if (coord1.LeftX == coord2.LeftX)
            return 0;
        else
            return 1;
    }
}

class Utils
{
    public static string prem_text = "";

    public static void RunCommand(string Command, string Args, string CommandDir)
    {
        string output = "";
        string error = "";
        try
        {

            ProcessStartInfo startInfo = new ProcessStartInfo(Command, Args);
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
            startInfo.CreateNoWindow = true;
            startInfo.UseShellExecute = false;
            //startInfo.RedirectStandardOutput = true;
            //startInfo.RedirectStandardError = true;

            Process process = new Process();
            process.StartInfo = startInfo;
            process.Start();
            //output = process.StandardOutput.ReadToEnd();
            //error = process.StandardError.ReadToEnd();
            process.WaitForExit();

            GC.Collect();
        }
        catch (Exception exception)
        {
            Console.WriteLine("Standard Output " + output);
            Console.WriteLine("Standard Error " + error);
            Console.WriteLine($"Exception occured while executing {Command} {Args}");
            Console.WriteLine(exception.StackTrace);
            throw exception;
        }
    }

    public static int ParsePageNumber(string FileName)
    {
        string trimmedFileName = Path.GetFileNameWithoutExtension(FileName);
        if (int.TryParse(trimmedFileName.Substring(trimmedFileName.LastIndexOf("-") + 1), out int n))
        {
            return int.Parse(trimmedFileName.Substring(trimmedFileName.LastIndexOf("-") + 1));
        }
        else
        {
            return 1;
        }
    }

    public static Coordinates ParseCoordinates(string Title)
    {
        var coordinates = Title.Replace("bbox", string.Empty).Trim().Split(';').First().Split(' ')
            .Select(coord => int.Parse(coord.Trim())).ToArray();

        return new Coordinates(coordinates[0], coordinates[1], coordinates[2], coordinates[3]);
    }

    private static int WordIndexFromOffset(string Text, int Offset)
    { // change this to a functional style using zip with indexes and count with predicate
        int wordIndex = 0;
        for (int i = 0; i <= Offset; i++)
            if (Text[i] == '~') wordIndex++;
        return wordIndex;
    }

}