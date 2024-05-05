using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace PP_Project.Models
{
    public class FormatChecks
    {
        public static void StylizeNormalParagraph(Word.Paragraph paragraph)
        {
            paragraph.Range.Select();
            paragraph.Range.Font.Size = 14;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Bold = 0;
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            paragraph.LineSpacing = 18f;
            paragraph.FirstLineIndent = 35.433f;
            paragraph.SpaceAfter = 0;
            paragraph.SpaceBefore = 0;
            paragraph.Range.Select();
        }

        public static void StylizeImageParagraph(Word.Paragraph paragraph)
        {
            paragraph.Range.Select();
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.LineSpacing = 18f;
            paragraph.FirstLineIndent = 0.0f;
            paragraph.SpaceAfter = 0;
            paragraph.SpaceBefore = 0;
        }

        public static void StylizeImageNameParagraph(Word.Paragraph paragraph)
        {
            paragraph.Range.Select();
            paragraph.Range.Font.Size = 14;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Bold = 0;
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.LineSpacing = 18f;
            paragraph.FirstLineIndent = 0.0f;
            paragraph.SpaceAfter = 6;
            paragraph.SpaceBefore = 0;
        }

        public static void StylizeListParagraph(Word.Paragraph paragraph)
        {
            paragraph.Range.Select();
            paragraph.Range.Font.Size = 14;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Bold = 0;
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
            paragraph.LineSpacing = 18f;
            paragraph.FirstLineIndent = 35.433f;
            paragraph.SpaceAfter = 6;
            paragraph.SpaceBefore = 0;
            paragraph.LeftIndent = 0.0f;
            paragraph.RightIndent = 0.0f;
        }

        public static void StylizeTableNameParagraph(Word.Paragraph paragraph)
        {
            paragraph.Range.Select();
            paragraph.Range.Font.Size = 14;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Range.Font.Bold = 0;
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraph.FirstLineIndent = 0;
            paragraph.SpaceAfter = 6;
            paragraph.SpaceBefore = 6;
        }

        public static void StylizeTableCellParagraph(Word.Paragraph paragraph)
        {
            paragraph.Range.Select();
            paragraph.Range.Font.Size = 12;
            paragraph.Range.Font.Name = "Times New Roman";
            paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            paragraph.LineSpacing = 12f;
            paragraph.FirstLineIndent = 0;
            paragraph.SpaceAfter = 0;
            paragraph.SpaceBefore = 0;
        }
        //TODO:
        //Check Headers

        //Check Correctness of a Table of Contents: Remake it if not all headers are there, apply correct style.
        //Table of contents will be saved in memory until the end of a document, where by the end of it,
        //it will have a complete list of headers, which it will then check.
        //If some names for whatever reason do not coincide, notify the user.
        //Screw this. I hate Interop

    }
}
