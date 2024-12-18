using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLCore
{
    public abstract class BaseOpenXMLProcessor : IOpenXMLProcessor
    {
        protected string FilePath;
        protected string ImgPath;

        public BaseOpenXMLProcessor(string? filePath = null, string? imgPath = null)
        {
            FilePath = filePath ?? string.Empty;
            ImgPath = imgPath ?? string.Empty;
        }

        protected WordprocessingDocument OpenDocument()
        {
            return WordprocessingDocument.Open(FilePath, true);
        }

        public abstract void AddParagraphText(string text);
        public abstract void SetParagraphProperty();
        public abstract void SetParagraphFont();
        public abstract void AddFixedTable(string[,] tableData);
        public abstract void AddImg(string imgPath);
        public abstract void AddTitleSytle();
        public abstract void SetParagraphStyle_Title_Level(string styleId);
    }
}
