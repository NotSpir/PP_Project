using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Page = Microsoft.Office.Interop.Word.Page;
using Word = Microsoft.Office.Interop.Word;

namespace PP_Project.Models
{
    public class GetWordData
    {
        public static void InitWord(string filePath)
        {
            try
            {
                Word.Application app = new Word.Application();
                Word.Document doc = null;

                object miss = Type.Missing;
                object oTrue = true;
                object oFalse = false;

                app.Visible = true;
                object file = filePath;
                doc = app.Documents.Open(ref file,
                    ref miss, ref miss, ref miss, ref miss,
                    ref miss, ref miss, ref miss, ref miss,
                    ref miss, ref miss, ref miss, ref miss,
                    ref miss, ref miss, ref miss);

                var paragraphs = doc.Paragraphs;
                doc.TrackRevisions = true;
                doc.Revisions.AcceptAll();

                ListTemplate customListTemplate = doc.ListTemplates.Add(false);

                // Задаем желаемый формат списка
                for (int i = 1; i <= customListTemplate.ListLevels.Count; i++)
                {
                    ListLevel listLevel = customListTemplate.ListLevels[i];
                    listLevel.NumberFormat = "%" + i.ToString();
                }

                ListTemplate sourcesListTemplate = doc.ListTemplates.Add(false);
                for (int i = 1; i <= sourcesListTemplate.ListLevels.Count; i++)
                {
                    ListLevel listLevel = customListTemplate.ListLevels[i];
                    listLevel.NumberFormat = "%" + i.ToString() + ".";
                }
                //var headerStyle = CreateWordHeader(doc);

                //Add skipping of selected pages

                //Finally, the part where Format

                //Стили в оглавлении
                foreach (TableOfContents toc in doc.TablesOfContents)
                {
                    toc.Range.Font.Size = 14;
                }
                bool prevIsPicture = false;
                bool isTableNameFormated = false;
                bool formatingSources = false;

                foreach (Paragraph para in paragraphs)
                {
                    if (formatingSources)
                    {
                        para.Range.Select();
                        FormatChecks.StylizeListParagraph(para);
                        para.Range.ListFormat.ApplyListTemplateWithLevel(sourcesListTemplate, true, WdListApplyTo.wdListApplyToSelection, WdDefaultListBehavior.wdWord10ListBehavior);
                        continue;
                    }

                    bool isTabled = false;
                    //Check for tables
                    if (IsInTable(para))
                    {
                        if (para.Previous() != null)
                        {
                            if (!isTableNameFormated)
                            {
                                para.Previous().Range.Select();
                                FormatChecks.StylizeTableNameParagraph(para.Previous());
                                isTableNameFormated = true;
                            }
                        }
                        else isTableNameFormated = true;

                        para.Range.Select();
                        FormatChecks.StylizeTableCellParagraph(para);
                        continue;
                    }
                    isTableNameFormated = false;

                    //Garbage. Responsible for lists. I hate this one.
                    if (para.Range.ListFormat.ListType != WdListType.wdListNoNumbering)
                    {
                        int level = para.Range.ListFormat.ListLevelNumber;

                        if (level <= customListTemplate.ListLevels.Count)
                        {
                            para.Range.Select();
                            para.Range.ListFormat.ApplyListTemplateWithLevel(customListTemplate, true, WdListApplyTo.wdListApplyToSelection, WdDefaultListBehavior.wdWord10ListBehavior);
                            FormatChecks.StylizeListParagraph(para);
                        }
                    }

                    if (prevIsPicture)
                    {
                        FormatChecks.StylizeImageNameParagraph(para);
                        prevIsPicture = false;
                    }
                    else
                    {
                        if (para.Range.InlineShapes.Count > 0)
                        {
                            foreach (Word.InlineShape ils in para.Range.InlineShapes)
                            {
                                // validate the object
                                if (ils != null)
                                {
                                    // validate this is a picture
                                    if (ils.Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapePicture)
                                    {
                                        FormatChecks.StylizeImageParagraph(para);
                                        prevIsPicture = true;
                                    }
                                }
                            }
                        }
                        else
                        {
                            FormatChecks.StylizeNormalParagraph(para);
                            if (para.Range.Text.ToLower().Contains("Список использованных источников".ToLower()))
                            {
                                formatingSources = true;
                            }
                        }

                    }


                }

                TableOfContentsController.CheckHeaders(doc);
                doc.Save();
                app.Quit();
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так!", "Ошибка");
            }
        }

        static bool IsInTable(Paragraph para)
        {
            if (para.Range.Tables.Count > 0)
            {
                return true;
            }

            return false;
        }
        public static Style CreateWordHeader(Word.Document doc)
        {
            Style headerStyle = doc.Styles.Add("ЗАГОЛОВОК_E", Type.Missing);
            headerStyle.Font.Bold = 1;
            headerStyle.Font.Size = 16;
            headerStyle.Font.Name = "Arial";
            headerStyle.ParagraphFormat.SpaceAfter = 12;
            return headerStyle;
        }
    }
}
