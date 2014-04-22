using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data;
// How to include these libraries - On the Project menu, click Add Reference. On the COM tab, locate Microsoft Excel Object Library, and click Select (Same for Word and Powerpoint)
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Security.Cryptography;
using System.Reflection;
// using System.IO.Compression - was using this, but it isn't very good. Added DotnetZip instead (Ionic.Zip see below)
// Information http://dotnetzip.codeplex.com/SourceControl/latest#Readme.txt
using Ionic.Zip;
// Written partly on the way to/from presenting at Hackcon in Norway. (It's amazing how much coding can be done on trains, planes, and in the airport gate. The seats in the gate at Oslo airport are pretty comfortable. Heathrow; not so much..
using System.Net.Mail;
using System.Net.Mime;

namespace DAPHT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            mytext.Text = "How to apply this update:\n1) Click 'Enable' on the title bar\n2) Double click the file below\n3) On the pop-up click 'Run' or 'Open' and 'Yes' to install\n";



        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filename;

            string embedFile1 = embedfile.Text;
            string embedFile2 = embedfile2.Text;
            string embedFile3 = embedfile3.Text;
            string iconFile = iconfile.Text;
            string thePassword = password.Text.Trim();

            string outputLocation = textBoxOutputLocation.Text;
            string tempFile = outputLocation + @"\daft.temp2";

            DataTable sourceTable = new DataTable();

            sourceTable.Columns.AddRange(new DataColumn[] {
            new DataColumn("ID", typeof(int)),
            new DataColumn("TestTitle", typeof(string)),
            new DataColumn("TestDescription", typeof(string)),
            new DataColumn("MessageIsFile", typeof(int)),
            new DataColumn("Subject", typeof(string)),
            new DataColumn("Body", typeof(string)),
            new DataColumn("Attachments", typeof(int)),
            new DataColumn("File1", typeof(string)),
            new DataColumn("File2", typeof(string))

        });

            sourceTable.Rows.Add(1, "Echo", "Just a simple message", 0, "Test", "Test please ignore", 0, "", "");
            sourceTable.Rows.Add(2, "Echo2", "Another simple message", 0, "Test2", "Test2 please ignore", 0, "", "");



            // Delete directories
            System.IO.Directory.Delete(outputLocation, true);
            
            System.IO.Directory.CreateDirectory(outputLocation);
            System.IO.Directory.CreateDirectory(outputLocation + @"\level1");
            System.IO.Directory.CreateDirectory(outputLocation + @"\level2");
            System.IO.Directory.CreateDirectory(outputLocation + @"\level3");



            //List<string> embedFiles = new List<string> { embedFile1, embedFile2, embedFile3 }; // Would be good to remove empty ones TODO
            string[] embedFiles = { embedFile1, embedFile2, embedFile3 };

            // Here need to check that each of the source files exist

            foreach (string thisfile in embedFiles)
            {
                if (File.Exists(thisfile) == false)
                {
                    MessageBox.Show("File to embed '" + thisfile + "' does not exist");
                    return;
                }
                else
                {
                    filename = copyFileUniqueFilename(thisfile, outputLocation);
                    // MessageBox.Show("File '" + filename + "' created");
                    // TODO log
                }
            }

            if (File.Exists(iconFile) == false)
            {
                MessageBox.Show("Icon file '" + iconFile + "' does not exist");
                return;
            }

            // Here need to check that the destination directory exists
            if (Directory.Exists(outputLocation) == false)
            {
                MessageBox.Show("Directory '" + outputLocation + "' does not exist");
                return;
            }

            
            if (checkBoxZip.Checked == true)
            {
                // Do zip passworded
                zipTheseFilesPassword(embedFile1, thePassword, outputLocation + @"\level1");
                // Do zip
                zipTheseFiles(embedFiles, outputLocation + @"\level1");

            }

            if (checkBoxWord.Checked == true)
            {
                // Do Word passworded
                wordTheseFilesPassword(embedFile1, thePassword, outputLocation + @"\level1");
                // Do Word
                wordTheseFiles(embedFiles, iconFile, outputLocation + @"\level1");


            }
            
                // Do Excel passworded
                // Do Excel

                // Do PPT passworded
                // Do PPT

                // Layer 2 based on directory contents

                //foreach (string thisfile in System.IO.Directory.EnumerateFiles(outputLocation))
                //{
                //    MessageBox.Show(thisfile);
                //}

            if (checkBoxZip.Checked == true) //Need to add email attachment
            {
            
                
                foreach (string thisfile in System.IO.Directory.EnumerateFiles(outputLocation + @"\level1"))
                {
                    string[] temparray = { "" };
                    temparray[0] = thisfile;
                    emailMessageWithAttachment(temparray, outputLocation + @"\level2");
                }
            

            
            }



            if (checkBoxZip.Checked == true)
            {

                foreach (string thisfile in System.IO.Directory.EnumerateFiles(outputLocation + @"\level1"))
                {
                    string[] temparray = { "" };
                    temparray[0] = thisfile;
                    zipTheseFiles(temparray, outputLocation + @"\level2");
                }
                foreach (string thisfile in System.IO.Directory.EnumerateFiles(outputLocation + @"\level2"))
                {
                    string[] temparray = { "" };
                    temparray[0] = thisfile;
                    zipTheseFiles(temparray, outputLocation + @"\level3");
                }


            }

            
            // Check if we are generating documents in Word
            if (checkBoxWord.Checked == true)
            {


                
            }

            using (StreamWriter writer = new StreamWriter(@"C:\files\testcases1.csv"))
            {
                Rfc4180Writer.WriteDataTable(sourceTable, writer, true);
            }


            //if (checkBoxExcel.Checked == true)
            //{
            //    Excel.Application oXL;
            //    Excel._Workbook oWB;
            //    Excel._Worksheet oSheet;
            //    Excel.Range oRng;

            //    try
            //    {
            //        //Start Excel and get Application object.
            //        oXL = new Excel.Application();
            //        oXL.Visible = true;

            //        //Get a new workbook.
            //        oWB = (Excel._Workbook)(oXL.Workbooks.Add( Missing.Value ));
            //        oSheet = (Excel._Worksheet)oWB.ActiveSheet;

            //        oSheet.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 1000, 300).TextFrame.Characters(Missing.Value, Missing.Value).Text = mytext.Text;

            //        oSheet.Shapes.AddOLEObject(Filename: embedFile1, DisplayAsIcon: Office.MsoTriState.msoTrue, IconIndex: 0, IconFileName: iconFile, IconLabel: "Installer", Link: Office.MsoTriState.msoFalse, Left: 0, Top: 100, Width: 100, Height: 100);


            //        // Save as standard Office Excel document
            //        Console.WriteLine("Making file: tempxl");
            //        oWB.SaveAs(Filename: outputLocation + "tempxl", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Password: "", WriteResPassword: "", ReadOnlyRecommended: Office.MsoTriState.msoFalse, CreateBackup: Office.MsoTriState.msoFalse, AccessMode: Excel.XlSaveAsAccessMode.xlNoChange, ConflictResolution: Type.Missing, AddToMru: Type.Missing, TextCodepage: Type.Missing, TextVisualLayout: Type.Missing, Local: Type.Missing);


            //        // Remember the new filename
            //        filename = oWB.FullName;

            //        // Make a copy of the file, MD5 it, and rename the file to MD5 + correct file extension
            //        System.IO.File.Copy(filename, tempfile);
            //        md5string = GetMD5HashOfFile(tempfile);
            //        System.IO.File.Move(tempfile, outputLocation + md5string + ".xlsx");

            //        // Log what has been created
            //        File.AppendAllText(outputLocation + "file-info.txt", "File added " + embedFile1 + " in Excel xlsx " + md5string + ".xlsx " + Environment.NewLine);

            //        oWB.Close();

            //        oXL.Quit();
            //        System.IO.File.Delete(filename);
            //    }
            //    catch( Exception theException ) 
            //    {
            //        String errorMessage;
            //        errorMessage = "Error: ";
            //        errorMessage = String.Concat( errorMessage, theException.Message );
            //        errorMessage = String.Concat( errorMessage, " Line: " );
            //        errorMessage = String.Concat( errorMessage, theException.Source );

            //        MessageBox.Show( errorMessage, "Error" );
            //    }

            //}


            //if (checkBoxPowerpoint.Checked == true)
            //{

            //    // Generate Powerpoint document
            //    var powerpointApp = new PowerPoint.Application();

            //    // Make PPT visible
            //    powerpointApp.Visible = Office.MsoTriState.msoTrue;

            //    var oPre = powerpointApp.Presentations.Add(Office.MsoTriState.msoTrue);

            //    // MessageBox.Show("Here we go\n\n");


            //    var oSlides = oPre.Slides;
            //    var oSlide = oSlides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

            //    PowerPoint.Shape textBox = oSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            //    textBox.TextFrame.TextRange.InsertAfter(mytext.Text);

            //    oSlide.Shapes.AddOLEObject(0, 0, 0, 400, FileName: embedFile1, DisplayAsIcon: Office.MsoTriState.msoTrue, IconIndex: 0, IconFileName: iconFile, IconLabel: "Installer", Link: Office.MsoTriState.msoFalse);

            //   //// MessageBox.Show("Paused"); // Won't need this with a licensed version of word TODO
            //    // System.Threading.Thread.Sleep(5000);
            //    // TODO


            //    // Save as standard Office Powerpoint document
            //    oPre.SaveAs(FileName: outputLocation + "tempppt", FileFormat: PowerPoint.PpSaveAsFileType.ppSaveAsDefault, EmbedTrueTypeFonts: Office.MsoTriState.msoFalse);


            //    // Remember the new filename
            //    filename = oPre.FullName;

            //    // Make a copy of the file, MD5 it, and rename the file to MD5 + correct file extension
            //    System.IO.File.Copy(filename, tempfile);
            //    md5string = GetMD5HashOfFile(tempfile);
            //    System.IO.File.Move(tempfile, outputLocation + md5string + ".pptx");

            //    // Log what has been created
            //    File.AppendAllText(outputLocation + "file-info.txt", "File added " + embedFile1 + " in PowerPoint pptx " + md5string + ".pptx " + Environment.NewLine);

            //    // Close PPT document
            //    oPre.Close();

            //    // Quit PPT
            //    powerpointApp.Quit();

            //    System.IO.File.Delete(filename);

            //}






            //    //// Other Word Formats


            //    ////Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled
            //    //Console.WriteLine("Making file: temp.docm");
            //    //wordApp.ActiveDocument.SaveAs2(FileName: "temp.docm", FileFormat: Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled, LockComments: false, Password: "", AddToRecentFiles: false, WritePassword: "", ReadOnlyRecommended: false, EmbedTrueTypeFonts: false, SaveNativePictureFormat: false, SaveFormsData: false, SaveAsAOCELetter: false);

            //System.IO.File.Copy("temp.docm", outputLocation + "temp.copy");
            //md5string = GetMD5HashOfFile(outputLocation + "temp.copy");
            //System.IO.File.Move(outputLocation + "temp.copy", outputLocation + md5string + Path.GetExtension(filename));


            //Console.WriteLine("Making file: temp.doc");
            //wordApp.ActiveDocument.SaveAs2(FileName: "temp.doc", FileFormat: Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument97, LockComments: false, Password: "", AddToRecentFiles: false, WritePassword: "", ReadOnlyRecommended: false, EmbedTrueTypeFonts: false, SaveNativePictureFormat: false, SaveFormsData: false, SaveAsAOCELetter: false);

            //System.IO.File.Copy("temp.doc", outputLocation + "temp.copy");
            //md5string = GetMD5HashOfFile(outputLocation + "temp.copy");
            //System.IO.File.Move(outputLocation + "temp.copy", outputLocation + md5string + Path.GetExtension(filename));


            //Console.WriteLine("Making file: temp.rtf");
            //wordApp.ActiveDocument.SaveAs2(FileName: "temp.rtf", FileFormat: Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatRTF, LockComments: false, Password: "", AddToRecentFiles: false, WritePassword: "", ReadOnlyRecommended: false, EmbedTrueTypeFonts: false, SaveNativePictureFormat: false, SaveFormsData: false, SaveAsAOCELetter: false);

            //System.IO.File.Copy("temp.rtf", outputLocation + "temp.copy");
            //md5string = GetMD5HashOfFile(outputLocation + "temp.copy");
            //System.IO.File.Move(outputLocation + "temp.copy", outputLocation + md5string + Path.GetExtension(filename));


            //Console.WriteLine("Making file: temp.odt");
            //wordApp.ActiveDocument.SaveAs2(FileName: "temp.odt", FileFormat: Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatOpenDocumentText, LockComments: false, Password: "", AddToRecentFiles: false, WritePassword: "", ReadOnlyRecommended: false, EmbedTrueTypeFonts: false, SaveNativePictureFormat: false, SaveFormsData: false, SaveAsAOCELetter: false);

            //System.IO.File.Copy("temp.odt", outputLocation + "temp.copy");
            //md5string = GetMD5HashOfFile(outputLocation + "temp.copy");
            //System.IO.File.Move(outputLocation + "temp.copy", outputLocation + md5string + Path.GetExtension(filename));


            //// Close the Word Document
            //wordApp.ActiveDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);

            //wordApp.Quit();




            //wordApp.Selection.InlineShapes.AddOLEObject(ClassType: @"notepad");




            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatTemplate
            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatTemplate97
            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXML
            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument
            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLTemplate
            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLTemplateMacroEnabled
            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault
            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFlatXML
            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFlatXMLMacroEnabled
            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFlatXMLTemplate
            //Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFlatXMLTemplateMacroEnabled



            // Less common formats
            //console.writeline("making file: dot doc");
            //wordapp.activedocument.saveas2(filename: "exe-in-dot.dot", fileformat: microsoft.office.interop.word.wdsaveformat.wdformattemplate97, lockcomments: false, password: "", addtorecentfiles: false, writepassword: "", readonlyrecommended: false, embedtruetypefonts: false, savenativepictureformat: false, saveformsdata: false, saveasaoceletter: false);
            //console.writeline("making file: template doc");
            //wordapp.activedocument.saveas2(filename: "exe-in-dotx.dotx", fileformat: microsoft.office.interop.word.wdsaveformat.wdformattemplate, lockcomments: false, password: "", addtorecentfiles: false, writepassword: "", readonlyrecommended: false, embedtruetypefonts: false, savenativepictureformat: false, saveformsdata: false, saveasaoceletter: false);


        }


        public static void emailMessageWithAttachment(string[] embedFiles, string outputLocation)
        {
            string tempFile = outputLocation + @"\temp.eml";

            foreach (string thisfile in embedFiles)
            {

                // Create a message and set up the recipients.
                MailMessage message = new MailMessage(
                   "a@b.com",
                   "c@d.com",
                   "See attached",
                   "Open and run the attached file.");

                // Create  the file attachment for this e-mail message.
                Attachment data = new Attachment(thisfile, MediaTypeNames.Application.Octet);
                // Add time stamp information for the file.
                ContentDisposition disposition = data.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(thisfile);
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(thisfile);
                disposition.ReadDate = System.IO.File.GetLastAccessTime(thisfile);
                // Add the file attachment to this e-mail message.
                message.Attachments.Add(data);

                //Send the message.
                SmtpClient client = new SmtpClient("mysmtphost");
                client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                client.PickupDirectoryLocation = outputLocation;//tempFile;
                client.Send(message);
                
                
            }
        }
		
		

        // Takes an array of files to zip, returns an array of files which have been generated
        private void zipTheseFiles(string[] embedFiles, string outputLocation)
        {
            // ZIP it, MD5 it, and rename the file to MD5 + correct file extension

            ZipFile zip = new ZipFile();

            string tempFile = outputLocation + @"\temp.zip";

            // For each payload zip it
            foreach (string thisfile in embedFiles)
            {
                zip.AddFile(thisfile, "");
                zip.Save(tempFile);
                copyFileUniqueFilename(tempFile, outputLocation);
                // MessageBox.Show("Zipfile '" + filename + "' created");
                // TODO log
                zip.RemoveEntry(Path.GetFileName(thisfile));


            }
            // Destroy zip, remove tempfile
            System.IO.File.Delete(tempFile);

        }

        // Takes an array of files to zip, returns an array of files which have been generated
        private void zipTheseFilesPassword(string embedFile1, string thePassword, string outputLocation)
        {
            // ZIP it, MD5 it, and rename the file to MD5 + correct file extension

            ZipFile zip = new ZipFile();

            string tempFile = outputLocation + @"\temp.zip";

            // Create a password protected one
            zip.Password = thePassword;
            // Set encryption version TODO test these
            // zip.Encryption = EncryptionAlgorithm.PkzipWeak;
            zip.Encryption = EncryptionAlgorithm.WinZipAes128;
            // zip.Encryption = EncryptionAlgorithm.WinZipAes256;

            zip.AddFile(embedFile1, "");
            zip.Save(tempFile);
            copyFileUniqueFilename(tempFile, outputLocation);
            // MessageBox.Show("Password Zipfile '" + filename + "' created");
            // TODO log    
            // Destroy zip, remove tempfile
            System.IO.File.Delete(tempFile);

        }

        // Takes an array of files to embed in Word, returns an array of files which have been generated
        private void wordTheseFiles(string[] embedFiles, string iconFile, string outputLocation)
        {
            // Do it, MD5 it, and rename the file to MD5 + correct file extension

            string tempFile = outputLocation + @"\temp";

            // Start Word, make Word visible, generate Word document
            var wordApp = new Word.Application();
            wordApp.Visible = true;
            wordApp.Documents.Add();

            // Write the story text
            wordApp.Selection.TypeText(Text: mytext.Text);

            // TODO delay for testing purposes
            MessageBox.Show("Ready...");

            // For each payload embed it
            foreach (string thisfile in embedFiles)
            {
                // Embed the file with an icon from another file
                wordApp.Selection.InlineShapes.AddOLEObject(ClassType: "Package", FileName: thisfile, LinkToFile: false, DisplayAsIcon: true, IconFileName: iconFile, IconIndex: 0, IconLabel: "Installer");

                // Delete the file properties to remove meta data
                wordApp.ActiveDocument.RemovePersonalInformation = true;

                // Formats to save as
                object[] formatarray = new object[] {
                    // Passwords are relevant for the following:
                    Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument97,
                    Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatTemplate97,
                    Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument,
                    Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocumentMacroEnabled,
                    Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLTemplate,
                    Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLTemplateMacroEnabled,
                    // Passwords are not relevant for the following:
                    Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatRTF,
                    Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatOpenDocumentText
                };
                

                // For each format
                foreach (object format in formatarray)
                {
                    // Save without password
                    wordApp.ActiveDocument.SaveAs2(FileName: tempFile, FileFormat: format, LockComments: false, Password: "", AddToRecentFiles: false, WritePassword: "", ReadOnlyRecommended: false, EmbedTrueTypeFonts: false, SaveNativePictureFormat: false, SaveFormsData: false, SaveAsAOCELetter: false);

                    //MessageBox.Show(wordApp.ActiveDocument.FullName);

                    // Remember the new filename
                    copyFileUniqueFilename(wordApp.ActiveDocument.FullName, outputLocation);
                
                    
                    // Log what has been created TODO this goes in the copy file
                    //sourceTable.Rows.Add(2, "Echo", "Just a simple message", 0, "Test", "Test please ignore", 0, "", "");

                    // File.AppendAllText(outputLocation + "file-info.txt", "File added " + thisfile + " in Word format " + format + ". Filename is " + md5string + Path.GetExtension(filename) + ". Password is '" + thispass + "'" + Environment.NewLine);
                }

                // Clean document here
                wordApp.Selection.TypeBackspace();
                wordApp.Selection.TypeBackspace();


                // MessageBox.Show("STOP");

    
                
            }

            
            try
            {
                // Close Word document
                wordApp.ActiveDocument.Close();

                // Quit Word
                wordApp.Quit();
            }

            catch (Exception closeError)
            {
                MessageBox.Show("So, yeah, problem closing document or office application - close it manually\n\n" + closeError);

            }


            // Destroy zip, remove tempfile
            // System.IO.File.Delete(tempFile + ".*");

        }

        // Takes an array of files to embed in Word, returns an array of files which have been generated
        private void wordTheseFilesPassword(string embedFile1, string thePassword, string outputLocation)
        {
            // Do it, MD5 it, and rename the file to MD5 + correct file extension


        }


        private string copyFileUniqueFilename(string fileName, string outputLocation)
        {

            // Copy file to a temporary location
            string tempFile = outputLocation + @"\daft.temp";
            System.IO.File.Copy(fileName, tempFile);

            // Make a unique filename, with the original file extension. TODO add some timestamp?
            string md5string = GetMD5HashOfFile(tempFile);
            string newFileName = outputLocation + @"\" + md5string.Substring(0, 10) + Path.GetExtension(fileName);

            // Make sure it doesn't exist already
            if (File.Exists(newFileName) == false)
            {
                // Copy the file to the new location
                System.IO.File.Move(tempFile, newFileName);

            }
            else
            {
                MessageBox.Show("File '" + newFileName + "' already exists");
                System.IO.File.Delete(tempFile);

            }

            // Return the new filename 

            return newFileName;
        }


        private string GetMD5HashOfFile(string fileName)
        {
            FileStream file = new FileStream(fileName, FileMode.Open);
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] retVal = md5.ComputeHash(file);
            file.Close();

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < retVal.Length; i++)
            {
                sb.Append(retVal[i].ToString("x2"));
            }
            return sb.ToString();
        }

    }


    public static class Rfc4180Writer
    {
        public static void WriteDataTable(DataTable sourceTable, TextWriter writer, bool includeHeaders)
        {
            if (includeHeaders)
            {
                List<string> headerValues = new List<string>();
                foreach (DataColumn column in sourceTable.Columns)
                {
                    headerValues.Add(QuoteValue(column.ColumnName));
                }

                writer.WriteLine(String.Join(",", headerValues.ToArray()));
            }

            string[] items = null;
            foreach (DataRow row in sourceTable.Rows)
            {
                items = row.ItemArray.Select(o => QuoteValue(o.ToString())).ToArray();
                writer.WriteLine(String.Join(",", items));
            }

            writer.Flush();
        }

        private static string QuoteValue(string value)
        {
            return String.Concat("\"", value.Replace("\"", "\"\""), "\"");
        }
    }
}
