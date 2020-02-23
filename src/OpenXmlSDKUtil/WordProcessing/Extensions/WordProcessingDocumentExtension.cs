using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using JetBrains.Annotations;
// ReSharper disable All

namespace OpenXmlSDKUtil.WordProcessing
{
    public static class WordProcessingDocumentExtension
    {
        public static bool HasStyleId([NotNull]this WordprocessingDocument doc, string styleId)
        {
            var styles = doc.MainDocumentPart.StyleDefinitionsPart.Styles;
            if (!styles.Elements<Style>().Any())
                return false;
            var style = styles.Elements<Style>().FirstOrDefault(x => x.StyleId == styleId);
            return style != null;
        }

        public static string GetStyleIdByStyleName([NotNull] this WordprocessingDocument doc, string styleName)
        {
            var styles = doc.MainDocumentPart.StyleDefinitionsPart.Styles;
            var styleId = styles.Descendants<StyleName>().Where(x => x.Val.Value.Equals(styleName))
                .Select(x => ((Style) x.Parent).StyleId).FirstOrDefault();
            return styleId;
        }
        public static string AddImage([NotNull] this WordprocessingDocument doc, string imageFile)
        {
            var mainPart = doc.MainDocumentPart;
            var imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(imageFile, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }
            return mainPart.GetIdOfPart(imagePart);
        }
    }


}