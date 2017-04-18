using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KeyWords;
using Facet.Combinatorics;
using System.Text.RegularExpressions;
using System.IO;
using Novacode;
using Microsoft.Office.Interop.Word;
//using EPocalipse.IFilter;
namespace KeyWords
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
     public String enter = "\n";
        public List<String> KeyPhrases = new List<String>();
        String path = "";
        public int KeyWordsNumb;
        public String source_text;
        public bool keywordsIsExtracted;
        

       
        private void btnOpen_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                if (ofd.ShowDialog(this) == DialogResult.OK)
                {
                    tbFile.Text = ofd.FileName;
                    rtbText.Clear();
                }
            }
            //FolderBrowserDialog fbd = new FolderBrowserDialog() { Description = "Select your path." }
        }

        private void btnExtract_Click(object sender, EventArgs e)
        {

            try
            {
                
                FindKeyPhrases();
                keywordsIsExtracted = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                String keyPhrases = "";
                KeyPhrases.Clear();
                //находим ключевые фразы:
                FindKeyPhrases();

                // keyPhrases = enter + "Key Phrases: " + enter;
                //выбераем только уникальные ключевые фразы:
                var keyPhrases2 =
                    KeyPhrases.Distinct();

                foreach (var w in keyPhrases2)
                {
                    keyPhrases += w + "; ";
                }

                path = tbFile.Text;
                //дописываем в конец документа ключевые фразы:

                AppendToWordDocx(path, keyPhrases);
                keywordsIsExtracted = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = "10";
            radioButton1.Checked = true;
        }
        private void HighlightText(object fileName, List<String> textToFind)
        {
            path = tbFile.Text;
            //object fileName = path;
            //textToFind = "test";
            object readOnly = false;
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = new Microsoft.Office.Interop.Word.Document();
            object missing = Type.Missing;
            try
            {
                doc = word.Documents.Open(ref fileName, ref missing, ref readOnly,
                                          ref missing, ref missing, ref missing,
                                          ref missing, ref missing, ref missing,
                                          ref missing, ref missing, ref missing,
                                          ref missing, ref missing, ref missing,
                                          ref missing);
                doc.Activate();


                object matchPhrase = false;
                object matchCase = false;
                object matchPrefix = false;
                object matchSuffix = false;
                object matchWholeWord = false;
                object matchWildcards = false;
                object matchSoundsLike = false;
                object matchAllWordForms = false;
                object matchByte = false;
                object ignoreSpace = false;
                object ignorePunct = false;

                object highlightedColor = Microsoft.Office.Interop.Word.WdColor.wdColorGreen;
                object textColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightOrange;

                Microsoft.Office.Interop.Word.Range range = doc.Range();
                foreach (object c in textToFind)
                {
                    bool highlighted = range.Find.HitHighlight(c,
                                                               highlightedColor,
                                                               textColor,
                                                               matchCase,
                                                               matchWholeWord,
                                                               matchPrefix,
                                                               matchSuffix,
                                                               matchPhrase,
                                                               matchWildcards,
                                                               matchSoundsLike,
                                                               matchAllWordForms,
                                                               matchByte,
                                                               false,
                                                               false,
                                                               false,
                                                               false,
                                                               false,
                                                               ignoreSpace,
                                                               ignorePunct,
                                                               false);
                }
                System.Diagnostics.Process.Start(fileName.ToString());

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
                Console.ReadKey(true);
            }

            
            
                }
       
       
        public void AppendToWordDocx(String path, String textToAdd)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(path);
            object missing = System.Reflection.Missing.Value;
            doc.Content.Font.Bold = 0;
            doc.Content.Text +="Key Words:"+enter+ textToAdd;
            app.Visible = true;    //Optional
            doc.Save();

            //this.Close();
        }
        public String ExtractTextFromMSWord(object path)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            object miss = System.Reflection.Missing.Value;
            //object path = @"C:\DOC\myDocument.docx";
            object readOnly = true;
            Microsoft.Office.Interop.Word.Document docs = word.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            string totaltext = "";
            for (int i = 0; i < docs.Paragraphs.Count; i++)
            {
                totaltext += " \r\n " + docs.Paragraphs[i + 1].Range.Text.ToString();
            }
            //Console.WriteLine(totaltext);
            docs.Close();
            word.Quit();
            return totaltext;
        }
        public String ExtractTxtFile(String path)
        {
            String result = "";
            // Open the file to read from.
            if (radioButton1.Checked)
            {
                string[] readText = File.ReadAllLines(path);
                foreach (string s in readText)
                {
                    result += s;
                }
            }
            if (radioButton2.Checked)
            {
                StreamReader file = new StreamReader(path, Encoding.GetEncoding(1251));
                result = file.ReadToEnd();
            }
            return result;
        }
        //подсчитывает колличество вхождений строки s0 в строку s
        public int CountWords(string s, string s0)
        {
            int count = (s.Length - s.Replace(s0, "").Length) / s0.Length;
            return count;
        }
        //удаляет исключительные слова(разделенные '|') из исходной строки:
        public String StringReplacer(String initialString, String sToRemove)
        {
            //строку с исключительными словами преобразуем в список:
            List<String> SToRemove = new List<String>(sToRemove.Split('|'));
            //Console.WriteLine("The initial string: '{0}'", initialString);
            foreach (var w in SToRemove)
            {
                //считаем количество вхождений слова(без пробелов) в исходную строку:
                int count = CountWords(initialString, w.Replace(" ", ""));

                //удаляем данное слово столько раз сколько оно встречается в исходн. строке:
                for (int i = 0; i < count; i++)
                {
                    initialString = initialString.Replace(w, " ");
                }
            }
            return initialString;
        }
        public void FindKeyPhrases()
        {
            try
            {

            rtbText.Text = "";//очищаем textBox
                              // Extract text from an input file.
            path = tbFile.Text;
            KeyWordsNumb = Int32.Parse(textBox1.Text);//кол-во ключевых слов
                                                      //1 - st Method:
                                                      // DocxToText dtt = new DocxToText(path);

            // //  String MyText= dtt.ExtractText();
            // //var words = "two one three one three one";
            // // String source_text = "2 1 3 1 3 1";
            //source_text = dtt.ExtractText();
            //2-nd Method:
            //TextReader reader = new EPocalipse.IFilter.FilterReader(path);
            //using (reader)
            //{
            //    source_text = reader.ReadToEnd();
            //    //label1.Text = "Text loaded from " + openFileDialog1.FileName;
            //}
            //3-d Method:
            source_text = ExtractTextFromMSWord(path);

            //очищаем текст от лишних пробелов:
            string cleanedString = System.Text.RegularExpressions.Regex.Replace(source_text, @"\s+", " ");
            rtbText.Text += "Source text:" + enter + enter;
            rtbText.Text += cleanedString + enter;
            rtbText.Text += "---------------------------------------------------------------------------------" + enter;

            //находим 10 самых часто встречаемых слов в тексте:
            //----------------


            // string patternToReplace = @"( of| and| as| the| a| with| on| in| at| to| for| under| after| it| is| are| their| her| she| he| they| when| where| by| for|Of |And |As |The |A |With |On |In |At |To |For |Under |After |It |Is |Are |Their |Her |She |He |They |When |Where|By |For )";

            //очищаем текст от специальных символов
            String cleanedStringWUC = Regex.Replace(cleanedString, @"[^0-9a-zA-ZА-Яа-яёЁъЪэЭыЫа-щА-ЩЬьЮюЯяЇїІіЄєҐґ -]+", " ");
            //очищаем текст от лишних пробелов:
            cleanedStringWUC = System.Text.RegularExpressions.Regex.Replace(cleanedStringWUC, @"\s+", " ");

            string patternToReplace = ExtractTxtFile("Words to be deleted.txt");//считываем исключающие слова с файла
                //очищаем текст от лишних пробелов:
            patternToReplace = System.Text.RegularExpressions.Regex.Replace(patternToReplace, @"\s+", " ");

            //patternToReplace = @"(" + patternToReplace + ")";

            //удаляем все исключающие слова из текста:

            //variant1:
            //cleanedStringWUC = Regex.Replace(cleanedStringWUC, patternToReplace, " ");
            //variant2:
            cleanedStringWUC = StringReplacer(cleanedStringWUC, patternToReplace);
            rtbText.Text += "Cleaned string:" + enter+enter;
            rtbText.Text += cleanedStringWUC+enter;
            rtbText.Text += enter+"Common words:";
            //String cleanedStringWUC = Regex.Replace(cleanedString, "/[^ a - zA - Z] / g", "");
            var orderedWords = cleanedStringWUC
                  .Split(' ')
                  .GroupBy(x => x)
                  .Select(x => new {
                      KeyField = x.Key,
                      Count = x.Count()
                  })
                  .OrderByDescending(x => x.Count)
                  .Take(KeyWordsNumb);




            String[] CommonWords1 = new String[KeyWordsNumb];

            int i = 0;
            foreach (var item in orderedWords)
            {

                CommonWords1[i] = item.KeyField;
                i++;

            }
            int CommonWordsLength = 0;
            for (int w = 0; w < CommonWords1.Length; w++) if (CommonWords1[w] != null) CommonWordsLength++;

            String[] CommonWords2 = new String[CommonWordsLength];
            int w2 = 0;
            for (int w = 0; w < CommonWords1.Length; w++)
            {
                if (CommonWords1[w] != null)
                {
                    CommonWords2[w2] = CommonWords1[w];

                    w2++;
                }
            }
            //--------------
            List<String> list = new List<String>(CommonWords2);
           
            List<String> ToRemove1 = new List<String>();
            

            foreach (String w in list) if (w == "") ToRemove1.Add(w);
            foreach (String w in ToRemove1) list.Remove(w);

            String[] CommonWords3 = list.ToArray();
            rtbText.Text += enter;
            foreach (String w in CommonWords3) rtbText.Text += w + "; ";
            //перебираем все размеще́ния по по 2:
            Variations<String> variationsTo2 = new Variations<String>(CommonWords3, 2, GenerateOption.WithRepetition);
            Variations<String> variationsTo3 = new Variations<String>(CommonWords3, 3, GenerateOption.WithRepetition);

            //записываем словосочетания по 2 слова в массив:
            String[] combinationsWith_2 = new String[variationsTo2.Count];//массив равный длине соответств. коллекции


            //rtbText.Text += enter;
            //rtbText.Text += "Combination to 2:";
            int comb = 0;
            foreach (IList<String> v in variationsTo2)
            {
                combinationsWith_2[comb] = v[0] + " " + v[1];
                //rtbText.Text += combinationsWith_2[comb] + enter;
                //Console.WriteLine(String.Format("{{{0} {1}}}", v[0], v[1]));
                comb++;
            }

            //записываем словосочетания по 3 слова в массив:
            String[] combinationsWith_3 = new String[variationsTo3.Count];//массив равный длине соответств. коллекции

            //rtbText.Text += enter;
            //rtbText.Text += "Combination to 3:";
            comb = 0;
            foreach (IList<String> v in variationsTo3)
            {
                combinationsWith_3[comb] = v[0] + " " + v[1] + " " + v[2];
                //rtbText.Text += combinationsWith_3[comb] + enter;
                comb++;
            }
            //находим список из существующих в тексте словосочетаний по 2 слова:
            //rtbText.Text += enter;
            //rtbText.Text += "existing combination to 2:";
            List<String> variationsTo2exist = new List<String>();
            for (comb = 0; comb < combinationsWith_2.Length; comb++)
            {
                var regex = new Regex(combinationsWith_2[comb]);
                if (regex.IsMatch(cleanedString))
                {
                    variationsTo2exist.Add(combinationsWith_2[comb]);
                    //rtbText.Text += combinationsWith_2[comb] + enter;
                }

            }




            //записываем в соответствующий массив все словосочет. по 2 которые есть в тексте:
            //String[] combinationsWith_2exist = new String[variationsTo2exist.Count];//массив равный длине соответств. коллекции
            //comb = 0;
            //foreach (String v in variationsTo2exist)
            //{
            //    combinationsWith_2exist[comb] = v;
            //    comb++;
            //}

            //---------------------------------------------
            //находим список из существующих в тексте словосочетаний по 3 слова:
            //rtbText.Text += enter;
            //rtbText.Text += "existing combination to 3:";
            List<String> variationsTo3exist = new List<String>();
            for (comb = 0; comb < combinationsWith_3.Length; comb++)
            {
                var regex = new Regex(combinationsWith_3[comb]);
                if (regex.IsMatch(cleanedString))
                {
                    variationsTo3exist.Add(combinationsWith_3[comb]);
                    //rtbText.Text += combinationsWith_3[comb] + enter;
                }
            }
            //записываем в соответствующий массив все словосочет. по 3 которые есть в тексте:
            //String[] combinationsWith_3exist = new String[variationsTo3exist.Count];//массив равный длине соответств. коллекции
            //comb = 0;
            //foreach (String v in variationsTo3exist)
            //{
            //    combinationsWith_3exist[comb] = v;
            //    comb++;
            //}

            //записываем часто встречаемые слова в массив:
            List<string> commonWords = new List<string>(CommonWords3);
            //проверяем есть ли часто встречаемые слова в списке по 2 слова, если есть удаляем данные слова из списка часто встреч. слов:
            List<String> ToRemove = new List<String>();
            foreach (String comb2 in variationsTo2exist)
            {

                foreach (String w in commonWords)
                {
                    var regex = new Regex(w);
                    if (regex.IsMatch(comb2))
                    {
                        //commonWords.Remove(w);

                        ToRemove.Add(w);

                    }
                }

            }

            //удаляем ненужные слова из списка:
            IEnumerable<String> ToRemoveDict = ToRemove.Distinct();
            foreach (String w in ToRemoveDict)
            { commonWords.Remove(w); }

            rtbText.Text += enter+enter;
            rtbText.Text += "All common single words:" + enter;
            foreach (String w in commonWords)
            {
                rtbText.Text += w + enter;
                KeyPhrases.Add(w);
            }
            //проверяем есть ли словосочетания по 2 в списке по 3 слова, если есть удаляем данные слов-я из списка:
            List<String> ToRemove2 = new List<String>();
            foreach (String comb3 in variationsTo3exist)
            {
                foreach (String comb2 in variationsTo2exist)
                {
                    var regex = new Regex(comb2);
                    if (regex.IsMatch(comb3))
                    {
                        //variationsTo2exist.Remove(comb2);
                        ToRemove2.Add(comb2);
                    }
                }

            }
            //удаляем ненужные словосочетания из списка:
            IEnumerable<String> ToRemoveDict2 = ToRemove2.Distinct();
            foreach (String com2 in ToRemoveDict2)
            { variationsTo2exist.Remove(com2); }
            //выводим ключевые фразы:
            rtbText.Text += enter;
            rtbText.Text += "All common combinations:" + enter;
            //foreach (String w in commonWords) rtbText.Text += w+ enter;
            foreach (String comb2 in variationsTo2exist)
            {
                rtbText.Text += comb2 + enter;
                KeyPhrases.Add(comb2);
            }
            foreach (String comb3 in variationsTo3exist)
            {
                rtbText.Text += comb3 + enter;
                KeyPhrases.Add(comb3);
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (keywordsIsExtracted)
            {
                try
                {
                    String path = tbFile.Text;

                    HighlightText(path, KeyPhrases);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message,
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }else
            {
                MessageBox.Show("Please, first click \"Extract Text\" button or \"Save keywords to the file\" button");
            }

        }
    }
}
