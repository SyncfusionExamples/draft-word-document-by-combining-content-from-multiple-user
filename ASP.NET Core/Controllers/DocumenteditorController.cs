using DocumentEditorApp.Models;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.IO;
using EJ2DocumentEditor = Syncfusion.EJ2.DocumentEditor;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using DocIOWordDocument = Syncfusion.DocIO.DLS.WordDocument;

namespace DocumentEditorApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DocumenteditorController : ControllerBase
    {
        private IHostingEnvironment hostEnvironment;
        public DocumenteditorController(IHostingEnvironment environment)
        {
            this.hostEnvironment = environment;
        }
        //Import file from client side.
        [Route("Import")]
        public string Import(IFormCollection data)
        {
            if (data.Files.Count == 0)
                return null;
            Stream stream = new MemoryStream();
            IFormFile file = data.Files[0];
            int index = file.FileName.LastIndexOf('.');
            string type = index > -1 && index < file.FileName.Length - 1 ?
                file.FileName.Substring(index) : ".docx";
            file.CopyTo(stream);
            stream.Position = 0;
            DocIOWordDocument document = new DocIOWordDocument(stream, Syncfusion.DocIO.FormatType.Automatic);

            foreach (WSection section in document.Sections)
            {
                //Accesses the Body of section where all the contents in document are apart.
                WTextBody sectionBody = section.Body;
                IterateTextBody(sectionBody);
            }
            EJ2DocumentEditor.WordDocument worddocument = EJ2DocumentEditor.WordDocument.Load(document);
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(worddocument);
            worddocument.Dispose();
            return json;
        }
        public class CustomClipboarParameter
        {
            public string content { get; set; }
            public string type { get; set; }
        }


        [Route("SystemClipboard")]
        public string SystemClipboard([FromBody]CustomClipboarParameter param)
        {
            if (param.content != null && param.content != "")
            {
                try
                {
                    Syncfusion.EJ2.DocumentEditor.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.LoadString(param.content, GetFormatType(param.type.ToLower()));
                    string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                    document.Dispose();
                    return json;
                }
                catch (Exception)
                {
                    return "";
                }
            }
            return "";
        }

        public class CustomRestrictParameter
        {
            public string passwordBase64 { get; set; }
            public string saltBase64 { get; set; }
            public int spinCount { get; set; }
        }

        [Route("RestrictEditing")]
        public string[] RestrictEditing([FromBody]CustomRestrictParameter param)
        {
            if (param.passwordBase64 == "" && param.passwordBase64 == null)
                return null;
            return Syncfusion.EJ2.DocumentEditor.WordDocument.ComputeHash(param.passwordBase64, param.saltBase64, param.spinCount);
        }

        //Import documents from web server.
        [Route("ImportFile")]
        public string ImportFile([FromBody]CustomParams param)
        {
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + param.fileName;
            try
            {
                Stream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.ReadWrite);
                Syncfusion.EJ2.DocumentEditor.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(stream, GetFormatType(path));
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                document.Dispose();
                stream.Dispose();
                return json;
            }
            catch
            {
                return "Failure";
            }
        }
        [Route("ImportUser1File")]
        public string ImportUser1File([FromBody]CustomParams param)
        {
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + param.fileName;
            try
            {
                Stream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.ReadWrite);
                WordDocument document = new WordDocument(stream, FormatType.Docx);
                //Creates the bookmark navigator instance to access the bookmark
                BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                //Moves the virtual cursor to the location before the end of the bookmark "Northwind"
                bookmarkNavigator.MoveToBookmark("User1");
                //Gets the bookmark content as WordDocumentPart
                WordDocumentPart wordDocumentPart = bookmarkNavigator.GetContent();
                //Saves the WordDocumentPart as separate Word document
                WordDocument newDocument = wordDocumentPart.GetAsWordDocument();
                //Close the WordDocumentPart instance
                wordDocumentPart.Close();
                Syncfusion.EJ2.DocumentEditor.WordDocument worddocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(newDocument);
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(worddocument);
                document.Dispose();
                stream.Dispose();
                return json;
            }
            catch
            {
                return "Failure";
            }
        }
        [Route("ImportUser2File")]
        public string ImportUser2File([FromBody]CustomParams param)
        {
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + param.fileName;
            try
            {
                Stream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.ReadWrite);
                WordDocument document = new WordDocument(stream, FormatType.Docx);
                //Creates the bookmark navigator instance to access the bookmark
                BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(document);
                //Moves the virtual cursor to the location before the end of the bookmark "Northwind"
                bookmarkNavigator.MoveToBookmark("User2");
                //Gets the bookmark content as WordDocumentPart
                WordDocumentPart wordDocumentPart = bookmarkNavigator.GetContent();
                //Saves the WordDocumentPart as separate Word document
                WordDocument newDocument = wordDocumentPart.GetAsWordDocument();
                //Close the WordDocumentPart instance
                wordDocumentPart.Close();
                Syncfusion.EJ2.DocumentEditor.WordDocument worddocument = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(newDocument);
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(worddocument);
                document.Dispose();
                stream.Dispose();
                return json;
            }
            catch
            {
                return "Failure";
            }
        }
        public class SaveParam
        {
            public string content { get; set; }
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [Route("ExportPdf")]
        public FileStreamResult ExportPdf([FromBody] SaveParameter data)
        {
            // Converts the sfdt to stream
            Stream document = EJ2DocumentEditor.WordDocument.Save(data.Content, EJ2DocumentEditor.FormatType.Docx);
            Syncfusion.DocIO.DLS.WordDocument doc = new Syncfusion.DocIO.DLS.WordDocument(document, Syncfusion.DocIO.FormatType.Docx);
            //Instantiation of DocIORenderer for Word to PDF conversion 
            DocIORenderer render = new DocIORenderer();
            //Converts Word document into PDF document 
            PdfDocument pdfDocument = render.ConvertToPDF(doc);
            Stream stream = new MemoryStream();
            
            //Saves the PDF file
            pdfDocument.Save(stream);
            stream.Position = 0;
            pdfDocument.Close();         
            document.Close();
            return new FileStreamResult(stream, "application/pdf")
            {
                FileDownloadName = data.FileName
            };
        }
        public class SaveParameter
        {
            public string Content { get; set; }
            public string FileName { get; set; }
            public string UserName { get; set; }
        }

        [Route("ExportSfdt")]
        public void ExportSfdt([FromBody] SaveParameter data)
        {
            string name = data.FileName;
            //string format = GetFormatType(name);
            if (string.IsNullOrEmpty(name))
            {
                name = "Document1.doc";
            }
            Syncfusion.DocIO.DLS.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.Save(data.Content);

            //Gets the bookmark content as WordDocumentPart
            WordDocumentPart wordDocumentPart = new WordDocumentPart(document);
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + data.FileName;

                Stream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.ReadWrite);
                //Loads the Word document with bookmark NorthwindDB
                WordDocument replacedocument = new WordDocument(stream, FormatType.Docx);
            //Creates the bookmark navigator instance to access the bookmark
            BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(replacedocument);
            //Moves the virtual cursor to the location before the end of the bookmark "NorthwindDB"
            bookmarkNavigator.MoveToBookmark(data.UserName);
            //Replaces the bookmark content with word body part
            bookmarkNavigator.ReplaceContent(wordDocumentPart);
            stream.Dispose();
            FileStream merged = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            replacedocument.Save(merged, FormatType.Docx);
            merged.Close();
        }

        //Save document in web server.
        [Route("Save")]
        public string Save([FromBody]CustomParameter param)
        {
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + param.fileName;
            Byte[] byteArray = Convert.FromBase64String(param.documentData);
            Stream stream = new MemoryStream(byteArray);
            EJ2DocumentEditor.FormatType type = GetFormatType(path);
            try
            {
                FileStream fileStream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);

                if (type != EJ2DocumentEditor.FormatType.Docx)
                {
                    Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(stream, Syncfusion.DocIO.FormatType.Docx);
                    document.Save(fileStream, GetDocIOFomatType(type));
                    document.Close();
                }
                else
                {
                    stream.Position = 0;
                    stream.CopyTo(fileStream);
                }
                stream.Dispose();
                fileStream.Dispose();
                return "Sucess";
            }
            catch
            {
                Console.WriteLine("err");
                return "Failure";
            }
        }

        internal static EJ2DocumentEditor.FormatType GetFormatType(string fileName)
        {
            int index = fileName.LastIndexOf('.');
            string format = index > -1 && index < fileName.Length - 1 ? fileName.Substring(index + 1) : "";

            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 Document editor does not support this file format.");
            switch (format.ToLower())
            {
                case "dotx":
                case "docx":
                case "docm":
                case "dotm":
                    return EJ2DocumentEditor.FormatType.Docx;
                case "dot":
                case "doc":
                    return EJ2DocumentEditor.FormatType.Doc;
                case "rtf":
                    return EJ2DocumentEditor.FormatType.Rtf;
                case "txt":
                    return EJ2DocumentEditor.FormatType.Txt;
                case "xml":
                    return EJ2DocumentEditor.FormatType.WordML;
                default:
                    throw new NotSupportedException("EJ2 Document editor does not support this file format.");
            }
        }

        internal static Syncfusion.DocIO.FormatType GetDocIOFomatType(EJ2DocumentEditor.FormatType type)
        {
            switch (type)
            {
                case EJ2DocumentEditor.FormatType.Docx:
                    return FormatType.Docx;
                case EJ2DocumentEditor.FormatType.Doc:
                    return FormatType.Doc;
                case EJ2DocumentEditor.FormatType.Rtf:
                    return FormatType.Rtf;
                case EJ2DocumentEditor.FormatType.Txt:
                    return FormatType.Txt;
                case EJ2DocumentEditor.FormatType.WordML:
                    return FormatType.WordML;
                default:
                    throw new NotSupportedException("DocIO does not support this file format.");
            }
        }
        #region Helper methods
        /// <summary>
        /// Iterates textbody child elements.
        /// </summary>
        private static void IterateTextBody(WTextBody textBody)
        {
            //Iterates through each of the child items of WTextBody.
            for (int i = 0; i < textBody.ChildEntities.Count; i++)
            {
                //IEntity is the basic unit in DocIO DOM. 
                //Accesses the body items as IEntity.
                IEntity bodyItemEntity = textBody.ChildEntities[i];
                //Decides the element type by using EntityType.
                switch (bodyItemEntity.EntityType)
                {
                    case EntityType.Paragraph:
                        WParagraph paragraph = bodyItemEntity as WParagraph;
                        //Processes the paragraph contents.
                        //Iterates through the paragraph's DOM.
                        IterateParagraph(paragraph.Items);
                        break;
                    case EntityType.Table:
                        //Table is a collection of rows and cells.
                        //Iterates through table's DOM.
                        IterateTable(bodyItemEntity as WTable);
                        break;
                }
            }
        }

        /// <summary>
        /// Iterates table child elements.
        /// </summary>
        private static void IterateTable(WTable table)
        {
            //Iterates the row collection in a table.
            foreach (WTableRow row in table.Rows)
            {
                //Iterates the cell collection in a table row.
                foreach (WTableCell cell in row.Cells)
                {
                    //Table cell is derived from (also a) TextBody.
                    //Reusing the code meant for iterating TextBody.
                    IterateTextBody(cell);
                }
            }
        }

        /// <summary>
        /// Iterates paragraph child elements.
        /// </summary>
        private static void IterateParagraph(ParagraphItemCollection paraItems)
        {
            for (int i = 0; i < paraItems.Count; i++)
            {
                Entity entity = paraItems[i];
                switch (entity.EntityType)
                {
                    //Gets the merge field.
                    case EntityType.MergeField:
                        WMergeField field = entity as WMergeField;
                        WCharacterFormat charFormatOfResult = GetCharacterFormatOfResult(field);
                        if (charFormatOfResult != null)
                            ApplyFormatForFieldCode(field, charFormatOfResult);
                        break;
                }
            }
        }
        /// <summary>
        /// Gets character format from Field Result.
        /// </summary>
        private static WCharacterFormat GetCharacterFormatOfResult(WMergeField field)
        {
            Entity entity = field;
            bool isSeparatorFound = false;
            //Iterates to sibling items until Field End 
            while (entity.NextSibling != null)
            {
                if (entity is WTextRange && isSeparatorFound)
                    //Sets character format for text ranges
                    return (entity as WTextRange).CharacterFormat;
                else if ((entity is WFieldMark) && (entity as WFieldMark).Type == FieldMarkType.FieldSeparator)
                    isSeparatorFound = true;
                else if ((entity is WFieldMark) && (entity as WFieldMark).Type == FieldMarkType.FieldEnd)
                    break;
                //Gets next sibling item
                entity = entity.NextSibling as Entity;

            }
            return null;
        }
        /// <summary>
        /// Applies character format to Field Code.
        /// </summary>
        private static void ApplyFormatForFieldCode(WMergeField field, WCharacterFormat charFormat)
        {
            Entity entity = field.NextSibling as Entity;
            //Iterates to sibling items until Field End 
            while (entity.NextSibling != null)
            {
                if (entity is WTextRange)
                    CompareAndApply(entity as WTextRange, charFormat);
                else if ((entity is WFieldMark) && (entity as WFieldMark).Type == FieldMarkType.FieldSeparator)
                    break;
                else if ((entity is WFieldMark) && (entity as WFieldMark).Type == FieldMarkType.FieldEnd)
                    break;
                //Gets next sibling item
                entity = entity.NextSibling as Entity;
            }
        }
        /// <summary>
        /// Compares and apply character format.
        /// </summary>
        private static void CompareAndApply(WTextRange textrange, WCharacterFormat charFormat)
        {
            if (textrange.CharacterFormat.AllCaps != charFormat.AllCaps)
                textrange.CharacterFormat.AllCaps = charFormat.AllCaps;
            if (textrange.CharacterFormat.Bidi != charFormat.Bidi)
                textrange.CharacterFormat.Bidi = charFormat.Bidi;
            if (textrange.CharacterFormat.Bold != charFormat.Bold)
                textrange.CharacterFormat.Bold = charFormat.Bold;
            if (textrange.CharacterFormat.BoldBidi != charFormat.BoldBidi)
                textrange.CharacterFormat.BoldBidi = charFormat.BoldBidi;
            if (textrange.CharacterFormat.CharacterSpacing != charFormat.CharacterSpacing)
                textrange.CharacterFormat.CharacterSpacing = charFormat.CharacterSpacing;
            if (textrange.CharacterFormat.ComplexScript != charFormat.ComplexScript)
                textrange.CharacterFormat.ComplexScript = charFormat.ComplexScript;
            if (textrange.CharacterFormat.DoubleStrike != charFormat.DoubleStrike)
                textrange.CharacterFormat.DoubleStrike = charFormat.DoubleStrike;
            if (textrange.CharacterFormat.Emboss != charFormat.Emboss)
                textrange.CharacterFormat.Emboss = charFormat.Emboss;
            if (textrange.CharacterFormat.Engrave != charFormat.Engrave)
                textrange.CharacterFormat.Engrave = charFormat.Engrave;
            if (textrange.CharacterFormat.FontName != charFormat.FontName)
                textrange.CharacterFormat.FontName = charFormat.FontName;
            if (textrange.CharacterFormat.FontSize != charFormat.FontSize)
                textrange.CharacterFormat.FontSize = charFormat.FontSize;
            if (textrange.CharacterFormat.FontSizeBidi != charFormat.FontSizeBidi)
                textrange.CharacterFormat.FontSizeBidi = charFormat.FontSizeBidi;
            if (textrange.CharacterFormat.Hidden != charFormat.Hidden)
                textrange.CharacterFormat.Hidden = charFormat.Hidden;
            if (textrange.CharacterFormat.HighlightColor != charFormat.HighlightColor)
                textrange.CharacterFormat.HighlightColor = charFormat.HighlightColor;
            if (textrange.CharacterFormat.Italic != charFormat.Italic)
                textrange.CharacterFormat.Italic = charFormat.Italic;
            if (textrange.CharacterFormat.ItalicBidi != charFormat.ItalicBidi)
                textrange.CharacterFormat.ItalicBidi = charFormat.ItalicBidi;
            if (textrange.CharacterFormat.Ligatures != charFormat.Ligatures)
                textrange.CharacterFormat.Ligatures = charFormat.Ligatures;
            if (textrange.CharacterFormat.LocaleIdASCII != charFormat.LocaleIdASCII)
                textrange.CharacterFormat.LocaleIdASCII = charFormat.LocaleIdASCII;
            if (textrange.CharacterFormat.LocaleIdBidi != charFormat.LocaleIdBidi)
                textrange.CharacterFormat.LocaleIdBidi = charFormat.LocaleIdBidi;
            if (textrange.CharacterFormat.LocaleIdFarEast != charFormat.LocaleIdFarEast)
                textrange.CharacterFormat.LocaleIdFarEast = charFormat.LocaleIdFarEast;
            if (textrange.CharacterFormat.NumberForm != charFormat.NumberForm)
                textrange.CharacterFormat.NumberForm = charFormat.NumberForm;
            if (textrange.CharacterFormat.NumberSpacing != charFormat.NumberSpacing)
                textrange.CharacterFormat.NumberSpacing = charFormat.NumberSpacing;
            if (textrange.CharacterFormat.OutLine != charFormat.OutLine)
                textrange.CharacterFormat.OutLine = charFormat.OutLine;
            if (textrange.CharacterFormat.Position != charFormat.Position)
                textrange.CharacterFormat.Position = charFormat.Position;
            if (textrange.CharacterFormat.Shadow != charFormat.Shadow)
                textrange.CharacterFormat.Shadow = charFormat.Shadow;
            if (textrange.CharacterFormat.SmallCaps != charFormat.SmallCaps)
                textrange.CharacterFormat.SmallCaps = charFormat.SmallCaps;
            if (textrange.CharacterFormat.Strikeout != charFormat.Strikeout)
                textrange.CharacterFormat.Strikeout = charFormat.Strikeout;
            if (textrange.CharacterFormat.StylisticSet != charFormat.StylisticSet)
                textrange.CharacterFormat.StylisticSet = charFormat.StylisticSet;
            if (textrange.CharacterFormat.SubSuperScript != charFormat.SubSuperScript)
                textrange.CharacterFormat.SubSuperScript = charFormat.SubSuperScript;
            if (textrange.CharacterFormat.TextBackgroundColor != charFormat.TextBackgroundColor)
                textrange.CharacterFormat.TextBackgroundColor = charFormat.TextBackgroundColor;
            if (textrange.CharacterFormat.TextColor != charFormat.TextColor)
                textrange.CharacterFormat.TextColor = charFormat.TextColor;
            if (textrange.CharacterFormat.UnderlineStyle != charFormat.UnderlineStyle)
                textrange.CharacterFormat.UnderlineStyle = charFormat.UnderlineStyle;
            if (textrange.CharacterFormat.UseContextualAlternates != charFormat.UseContextualAlternates)
                textrange.CharacterFormat.UseContextualAlternates = charFormat.UseContextualAlternates;
        }
        #endregion
    }
}
