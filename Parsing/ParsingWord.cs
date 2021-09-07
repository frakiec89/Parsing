using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Parsing
{
   public class ParsingWord
    {
        public   string Metod(string path)
        {
            string content = string.Empty;
            object FileName = path;
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Word.Document doc = null;
            Word.Range range = null;
            try
            {
                doc = app.Documents.Open(ref FileName, ref MissingObj, ref rOnly, ref MissingObj,
                ref MissingObj, ref MissingObj, ref MissingObj, ref MissingObj,
                ref MissingObj, ref MissingObj, ref MissingObj, ref MissingObj,
                ref MissingObj, ref MissingObj, ref MissingObj, ref MissingObj);
                object StartPosition = 0;
                object EndPositiojn = doc.Characters.Count;
                range = doc.Range(ref StartPosition, ref EndPositiojn);

                // Получение основного текста со страниц (без учёта сносок и колонтитулов)
                string MainText = (range == null || range.Text == null) ? null : range.Text;
                if (MainText != null)
                {
                    content += MainText;
                }

                // Получение текста из нижних и верхних колонтитулов
                foreach (Word.Section section in doc.Sections)
                {
                    // Нижние колонтитулы
                    foreach (Word.HeaderFooter footer in section.Footers)
                    {
                        string FooterText = (footer.Range == null || footer.Range.Text == null) ? null : footer.Range.Text;
                        if (FooterText != null)
                        {
                            content += FooterText;
                        }
                    }

                    // Верхние колонтитулы
                    foreach (Word.HeaderFooter header in section.Headers)
                    {
                        string HeaderText = (header.Range == null || header.Range.Text == null) ? null : header.Range.Text;
                        if (HeaderText != null)
                        {
                            content += HeaderText;
                        }
                    }
                }
                // Получение текста сносок
                if (doc.Footnotes.Count != 0)
                {
                    foreach (Word.Footnote footnote in doc.Footnotes)
                    {
                        string FooteNoteText = (footnote.Range == null || footnote.Range.Text == null) ? null : footnote.Range.Text;
                        if (FooteNoteText != null)
                        {
                            content += FooteNoteText;
                        }
                    }
                }

                return content;
            }
            catch (Exception ex)
            {
                throw new Exception("Error");
            }
            finally
            {
                /* Очистка неуправляемых ресурсов */
                if (doc != null)
                {
                    doc.Close(ref SaveChanges);
                }
                if (range != null)
                {
                    Marshal.ReleaseComObject(range);
                    range = null;
                }
                if (app != null)
                {
                    app.Quit();
                    Marshal.ReleaseComObject(app);
                    app = null;
                }
            }
        }


        public string Reg (string content)
        {
            string rez = string.Empty; // пустая  строка 

            var contentMassiv = content.Split('\r'); // разбиваю входную  строку на  массив строк 

            string pattern = @"[A-ЯЁ][а-яё]+\s[A-ЯЁ][а-яё]+"; // регулярное выражение
          
            Regex rgx = new Regex(pattern); 
            Stopwatch sw;

            List<string> piople = new List<string>(); // сюда сложу людей кого найду

            foreach (string input in contentMassiv) // поиск
            {
                sw = Stopwatch.StartNew();
                Match match = rgx.Match(input);
                sw.Stop();
                if (match.Success)
                {
                    piople.Add(string.Format($"{match.Value} ")); // складываем 
                }
            }
            List<string> noDupes = piople.Distinct().OrderBy(x=>x).ToList(); // поиск дублирования  и сортировка 
            noDupes.ForEach(x => rez += $"{x} \n"); // лист  в  строку 
            return rez; // вернем  строку 
        }
    }
}
