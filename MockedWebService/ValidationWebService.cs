using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;

namespace OecdAuthoring
{
    public class ValidationResults
    {
        public List<WmlToXmlValidationError> ValidationErrors;
    }

    public static class ValidationWebService
    {
        public static ValidationResults Validate(WmlDocument wmlSourceDocument)
        {
            var tempPath = Path.GetTempPath();

            WmlToXmlOecdSettings oecdSettings = new WmlToXmlOecdSettings();
            oecdSettings.WriteImageFiles = true;
            oecdSettings.DoNotAutomateWord = true;
            oecdSettings.ContentTypeRegexExtension = null;
            oecdSettings.Sources = new[] { wmlSourceDocument };

            var now = DateTime.Now;
            string subDirName = string.Format("Temp-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}-{6:000}", now.Year - 2000, now.Month, now.Day, now.Hour, now.Minute, now.Second, now.Millisecond);
            oecdSettings.TempDi = new DirectoryInfo(Path.Combine(tempPath, subDirName));
            oecdSettings.TempDi.Create();

            oecdSettings.ImageBase = oecdSettings.TempDi;

            // determine default language for first document.
#if false
<w:styles xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" mc:Ignorable="w14 w15 w16se">
	<w:docDefaults>
		<w:rPrDefault>
			<w:rPr>
				<w:rFonts w:ascii="Georgia" w:eastAsiaTheme="minorHAnsi" w:hAnsi="Georgia" w:cs="Times New Roman"/>
				<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
			</w:rPr>
		</w:rPrDefault>
		<w:pPrDefault/>
	</w:docDefaults>
#endif
            string defaultLanguage = null;
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlSourceDocument.DocumentByteArray, 0, wmlSourceDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    if (wDoc.MainDocumentPart.StyleDefinitionsPart != null)
                    {
                        var sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                        defaultLanguage = (string)sXDoc.Root.Elements(W.docDefaults).Elements(W.rPrDefault).Elements(W.rPr).Elements(W.lang).Attributes(W.val).FirstOrDefault();
                    }
                }
            }
            oecdSettings.DefaultLang = defaultLanguage;

            var results = oecdSettings.Sources.Select(src => WmlToXmlOecd.ApplyContentTypes(src, oecdSettings));

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Produce -ContentType-New.xml
            var contentTypeXml = results.Select(res => WmlToXmlOecd.ProduceContentTypeXml(res, oecdSettings)).First();

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // Validate per validation rules
            var errorList = WmlToXmlOecd.ValidateContentTypeXml(results.First(), contentTypeXml, oecdSettings);

            var r = new ValidationResults();
            r.ValidationErrors = errorList;
            return r;
        }
    }
}
