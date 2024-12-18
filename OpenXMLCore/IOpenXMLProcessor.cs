using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLCore
{
    public interface IOpenXMLProcessor
    {
        void AddParagraphText(string text);

        void SetParagraphProperty();

        void SetParagraphFont();

        void AddFixedTable(string[,] tableData);

        void AddImg(string imgPath);

        void AddTitleSytle();

        void SetParagraphStyle_Title_Level(string styleId);
    }
}
