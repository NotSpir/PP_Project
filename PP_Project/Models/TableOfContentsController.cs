using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PP_Project.Models
{
    public static class TableOfContentsController
    {

        public static void CheckHeaders(Document doc)
        {
            //DOES NOT WORK. AT ALL.
            try
            {
                bool headerExists;
                string[] headersToCheck = { "СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ",
                                         "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "ПРИЛОЖЕНИЕ" };

                foreach (string header in headersToCheck)
                {
                    headerExists = false;
                    foreach (Section section in doc.Sections)
                    {
                        foreach (HeaderFooter headerFooter in section.Headers)
                        {
                            if (headerFooter.Range.Text.ToUpper().Contains(header.ToUpper()))
                            {
                                Console.WriteLine($"Found '{header}' header!");
                                headerExists = true;
                                break; 
                            }
                        }

                        if (headerExists) break;

                        foreach (HeaderFooter footer in section.Footers)
                        {
                            if (footer.Range.Text.ToUpper().Contains(header.ToUpper()))
                            {
                                Console.WriteLine($"Found '{header}' header!");
                                headerExists = true;
                                break; 
                            }
                        }

                        if (headerExists) break; 
                    }

                    if (!headerExists)
                    {
                        MessageBox.Show($"Заголовок '{header}' не найден.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
