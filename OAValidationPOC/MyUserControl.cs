using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OecdAuthoring;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.IO;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;
using OpenXmlPowerTools;

namespace OAValidationPOC
{
    public partial class MyUserControl : UserControl
    {
        public MyUserControl()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
#if false
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            int numberStoryRanges = doc.StoryRanges.Count;
            MessageBox.Show(string.Format("numberStoryRanges: {0}", numberStoryRanges));

            int i = 0;
            foreach (Word.Range range in doc.StoryRanges)
            {
                i++;
                MessageBox.Show(string.Format("range #{0}", i));

                int numberParagraphs = range.Paragraphs.Count;
                MessageBox.Show(string.Format("number of paragraphs in range: {0}", numberParagraphs));

                for (int p = 1; p <= numberParagraphs; ++p)
                {
                    range.Paragraphs[p].Range.Select();
                    MessageBox.Show(string.Format("{0}", p));
                }
            }
#endif
            var doc = Globals.ThisAddIn.Application.ActiveDocument;

            Object start = doc.Content.Start;
            Object end = doc.Content.End;

            //select entire document 
            var range = doc.Range(ref start, ref end);
            var packageStream = (MemoryStream)range.GetPackageStreamFromRange();
            var byteArray = packageStream.ToArray();

            WmlDocument wmlDoc = new WmlDocument("file.docx", byteArray);

            var result = ValidationWebService.Validate(wmlDoc);

            if (! result.ValidationErrors.Any())
            {
                if (this.errorButtonList != null)
                {
                    foreach (var eb in this.errorButtonList)
                        this.Controls.Remove(eb);
                    this.errorButtonList = new List<Button>();
                    this.cachedErrors = new List<WmlToXmlValidationError>();
                }
                MessageBox.Show("No Validation Errors");
            }
            else
            {
#if false
                var first = result.ValidationErrors.First();
                int paraToSelect;
                if (int.TryParse(first.BlockLevelContentIdentifier, out paraToSelect))
                {
                    try
                    {
                        doc.Paragraphs[paraToSelect].Range.Select();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Caught exception: " + ex.ToString());
                    }
                }
                MessageBox.Show(first.ErrorMessage);
#endif

                this.SuspendLayout();

                if (this.errorButtonList != null)
                {
                    foreach (var eb in this.errorButtonList)
                        this.Controls.Remove(eb);
                    this.errorButtonList = new List<Button>();
                }

                this.errorButtonList = new List<Button>();
                this.cachedErrors = new List<WmlToXmlValidationError>();

                var errorsToShow = result.ValidationErrors.Take(7).ToList();

                ///////////////////////////////////////////////////////////////////////////////
                int errorButtonLeft = 10;
                int errorButtonsTop = 60;
                int errorButtonWidth = 170;
                int errorButtonHeight = 50;
                int distanceBetweenButtons = 15;
                ///////////////////////////////////////////////////////////////////////////////

                int i = 0;
                foreach (var err in errorsToShow)
                {
                    var button = new System.Windows.Forms.Button();
                    button.Location = new System.Drawing.Point(errorButtonLeft, errorButtonsTop + (i * (errorButtonHeight + distanceBetweenButtons)));
                    button.Name = "error" + i.ToString();
                    button.Size = new System.Drawing.Size(errorButtonWidth, errorButtonHeight);
                    button.TabIndex = i + 1;
                    button.Text = err.ErrorMessage;
                    button.UseVisualStyleBackColor = true;
                    button.Click += new System.EventHandler(error_Button_Click);
                    this.errorButtonList.Add(button);
                    this.Controls.Add(button);
                    cachedErrors.Add(err);
                    i++;
                }

                this.ResumeLayout(false);
            }
        }

        private void error_Button_Click(object sender, EventArgs e)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            Button b = (Button)sender;
            var errorNumberString = b.Name.Substring(5);
            int errorNumber;
            if (int.TryParse(errorNumberString, out errorNumber))
            {
                var err = this.cachedErrors.Skip(errorNumber).FirstOrDefault();
                if (err != null)
                {
                    try
                    {
                        int paraToSelect;
                        if (int.TryParse(err.BlockLevelContentIdentifier, out paraToSelect))
                        {
                            doc.Paragraphs[paraToSelect].Range.Select();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Caught exception: " + ex.ToString());
                    }
                }
            }
        }
    }

    public static class OpcHelper
    {
        /// <summary>
        /// Returns the part contents in xml
        /// </summary>
        /// <param name="part">System.IO.Packaging.Packagepart</param>
        /// <returns></returns>
        static XElement GetContentsAsXml(PackagePart part)
        {
            XNamespace pkg =
               "http://schemas.microsoft.com/office/2006/xmlPackage";
            if (part.ContentType.EndsWith("xml"))
            {
                using (Stream partstream = part.GetStream())
                using (StreamReader streamReader = new StreamReader(partstream))
                {
                    string streamString = streamReader.ReadToEnd();
                    XElement newXElement =
                        new XElement(pkg + "part", new XAttribute(pkg + "name", part.Uri),
                            new XAttribute(pkg + "contentType", part.ContentType),
                            new XElement(pkg + "xmlData", XElement.Parse(streamString)));
                    return newXElement;
                }
            }
            else
            {
                using (Stream str = part.GetStream())
                using (BinaryReader binaryReader = new BinaryReader(str))
                {
                    int len = (int)binaryReader.BaseStream.Length;
                    byte[] byteArray = binaryReader.ReadBytes(len);
                    // the following expression creates the base64String, then chunks
                    // it to lines of 76 characters long
                    string base64String = (System.Convert.ToBase64String(byteArray))
                        .Select
                        (
                            (c, i) => new
                            {
                                Character = c,
                                Chunk = i / 76
                            }
                        )
                        .GroupBy(c => c.Chunk)
                        .Aggregate(
                            new StringBuilder(),
                            (s, i) =>
                                s.Append(
                                    i.Aggregate(
                                        new StringBuilder(),
                                        (seed, it) => seed.Append(it.Character),
                                        sb => sb.ToString()
                                    )
                                )
                                .Append(Environment.NewLine),
                            s => s.ToString()
                        );

                    return new XElement(pkg + "part",
                        new XAttribute(pkg + "name", part.Uri),
                        new XAttribute(pkg + "contentType", part.ContentType),
                        new XAttribute(pkg + "compression", "store"),
                        new XElement(pkg + "binaryData", base64String)
                    );
                }
            }
        }
        /// <summary>
        /// Returns an XDocument
        /// </summary>
        /// <param name="package">System.IO.Packaging.Package</param>
        /// <returns></returns>
        public static XDocument OpcToFlatOpc(Package package)
        {
            XNamespace
                pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
            XDeclaration
                declaration = new XDeclaration("1.0", "UTF-8", "yes");
            XDocument doc = new XDocument(
                declaration,
                new XProcessingInstruction("mso-application", "progid=\"Word.Document\""),
                new XElement(pkg + "package",
                    new XAttribute(XNamespace.Xmlns + "pkg", pkg.ToString()),
                    package.GetParts().Select(part => GetContentsAsXml(part))
                )
            );
            return doc;
        }

        /// <summary>
        /// Returns a System.IO.Packaging.Package stream for the given range.
        /// </summary>
        /// <param name="range">Range in word document</param>
        /// <returns></returns>
        public static Stream GetPackageStreamFromRange(this Word.Range range)
        {
            XDocument doc = XDocument.Parse(range.WordOpenXML);
            XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
            XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
            Package InmemoryPackage = null;
            MemoryStream memStream = new MemoryStream();
            using (InmemoryPackage = Package.Open(memStream, FileMode.Create))
            {
                // add all parts (but not relationships)
                foreach (var xmlPart in doc.Root
                    .Elements()
                    .Where(p =>
                        (string)p.Attribute(pkg + "contentType") !=
                        "application/vnd.openxmlformats-package.relationships+xml"))
                {
                    string name = (string)xmlPart.Attribute(pkg + "name");
                    string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                    if (contentType.EndsWith("xml"))
                    {
                        Uri u = new Uri(name, UriKind.Relative);
                        PackagePart part = InmemoryPackage.CreatePart(u, contentType,
                            CompressionOption.SuperFast);
                        using (Stream str = part.GetStream(FileMode.Create))
                        using (XmlWriter xmlWriter = XmlWriter.Create(str))
                            xmlPart.Element(pkg + "xmlData")
                                .Elements()
                                .First()
                                .WriteTo(xmlWriter);
                    }
                    else
                    {
                        Uri u = new Uri(name, UriKind.Relative);
                        PackagePart part = InmemoryPackage.CreatePart(u, contentType,
                            CompressionOption.SuperFast);
                        using (Stream str = part.GetStream(FileMode.Create))
                        using (BinaryWriter binaryWriter = new BinaryWriter(str))
                        {
                            string base64StringInChunks =
                           (string)xmlPart.Element(pkg + "binaryData");
                            char[] base64CharArray = base64StringInChunks
                                .Where(c => c != '\r' && c != '\n').ToArray();
                            byte[] byteArray =
                                System.Convert.FromBase64CharArray(base64CharArray,
                                0, base64CharArray.Length);
                            binaryWriter.Write(byteArray);
                        }
                    }
                }
                foreach (var xmlPart in doc.Root.Elements())
                {
                    string name = (string)xmlPart.Attribute(pkg + "name");
                    string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                    if (contentType ==
                        "application/vnd.openxmlformats-package.relationships+xml")
                    {
                        // add the package level relationships
                        if (name == "/_rels/.rels")
                        {
                            foreach (XElement xmlRel in
                                xmlPart.Descendants(rel + "Relationship"))
                            {
                                string id = (string)xmlRel.Attribute("Id");
                                string type = (string)xmlRel.Attribute("Type");
                                string target = (string)xmlRel.Attribute("Target");
                                string targetMode =
                                    (string)xmlRel.Attribute("TargetMode");
                                if (targetMode == "External")
                                    InmemoryPackage.CreateRelationship(
                                        new Uri(target, UriKind.Absolute),
                                        TargetMode.External, type, id);
                                else
                                    InmemoryPackage.CreateRelationship(
                                        new Uri(target, UriKind.Relative),
                                        TargetMode.Internal, type, id);
                            }
                        }
                        else
                        // add part level relationships
                        {
                            string directory = name.Substring(0, name.IndexOf("/_rels"));
                            string relsFilename = name.Substring(name.LastIndexOf('/'));
                            string filename =
                                relsFilename.Substring(0, relsFilename.IndexOf(".rels"));
                            PackagePart fromPart = InmemoryPackage.GetPart(
                                new Uri(directory + filename, UriKind.Relative));
                            foreach (XElement xmlRel in
                                xmlPart.Descendants(rel + "Relationship"))
                            {
                                string id = (string)xmlRel.Attribute("Id");
                                string type = (string)xmlRel.Attribute("Type");
                                string target = (string)xmlRel.Attribute("Target");
                                string targetMode =
                                    (string)xmlRel.Attribute("TargetMode");
                                if (targetMode == "External")
                                    fromPart.CreateRelationship(
                                        new Uri(target, UriKind.Absolute),
                                        TargetMode.External, type, id);
                                else
                                    fromPart.CreateRelationship(
                                        new Uri(target, UriKind.Relative),
                                        TargetMode.Internal, type, id);
                            }
                        }
                    }
                }
                InmemoryPackage.Flush();
            }
            return memStream;
        }
    }
}
