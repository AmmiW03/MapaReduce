using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MapaReduce.Modelos
{
    internal class clsMetodos
    {

        /*public static List<mdlCategorias> getCategories(String filePath)
        {
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = new XSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheetAt(0);
            DataFormatter dataf = new DataFormatter();

            if (sheet == null)
                return new List<mdlCategorias>();

            List<mdlCategorias> categorias = new List<mdlCategorias>();

            for(int i = 0; i < sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                mdlCategorias categorie = new mdlCategorias();
                categorie.setMin(int.Parse(dataf.FormatCellValue(row.GetCell(0))));
                categorie.setName(dataf.FormatCellValue(row.GetCell(1)));
                categorias.Add(categorie);
            }
            fs.Close();
            return categorias;
        }*/

        /*public static void readExcel(List<mdlCategorias> Categories, String filePath,String sheetName, int headRow)
        {
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = new XSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheet(sheetName);
            DataFormatter dataf = new DataFormatter();
            bool flag;

            if (sheet == null)
                return ;
            IRow headers = sheet.GetRow(headRow);

            for (int i = headRow+1; i < sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                flag = true;
                foreach (mdlCategorias category in Categories)
                {
                    if (category.getName() == dataf.FormatCellValue(row.GetCell(0)))
                    {
                        flag = false;
                    }
                }
                if (flag)
                {
                    for (int j = 0; j < row.Cells.Count; j++)
                    {

                    }
                }
            }
        }*/

        public static void firstExcel()
        {
            //FileStream fs = new FileStream(@"D:\Recursos\AProyectos\Datos\Extra\Excel resultante.xlsx", FileMode.Open, FileAccess.ReadWrite);
            using (FileStream fs = new FileStream(@"D:\Recursos\AProyectos\Datos\Extra\HolaMundo.xlsx", FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new XSSFWorkbook(fs);
                ISheet sheet = workbook.GetSheetAt(0);
                DataFormatter dataf = new DataFormatter();
                List<string> commonW = new List<string> { "a", "ante", "bajo", "con", "contra", "de", "desde", "en", "entre", "hacia", "hasta", "para", "por", "según", "sin", "sobre", "tras", "y", "del", "la", };
                if (sheet == null)
                {
                    return;
                }
                for (int i = 0; i < sheet.LastRowNum; i++)
                {
                    String res = String.Empty;
                    List<String> result = new List<String>();
                    List<String> firstCell = extractWords(dataf.FormatCellValue(sheet.GetRow(i).GetCell(1)), commonW);
                    List<String> secondCell = extractWords(dataf.FormatCellValue(sheet.GetRow(i).GetCell(8)), commonW);
                    foreach (String word in firstCell)
                    {
                        if (result.Find(D => D == word) == null)
                        {
                            result.Add(word);
                        }
                    }
                    foreach (String word in secondCell)
                    {
                        if (result.Find(D => D == word) == null)
                        {
                            result.Add(word);
                        }
                    }
                    foreach (String word in result)
                    {
                        if (commonW.Find(D => D == word) == null)
                        {
                            res += word + " ";
                        }
                    }
                    sheet.GetRow(i).CreateCell(9).SetCellValue(res);
                }

                workbook.Write(fs,false);
            }
        }

        public static List<String> extractWords(String cell, List<String> common)
        {
            if(cell == null || cell == "")
            {
                return new List<String>();
            }
            List<String> words = new List<String>();
            string[] auxWords = cell.Split(',');
            String auxString = ""; 

            foreach (string word in auxWords)
            {
                if(words.Find(D => D == word) == null && common.Find(D => D == word) == null)
                {
                    auxString += word;
                }             
            }
            string[] auxWords_ = auxString.Split(' ');
            foreach (string word in auxWords_)
            {
                if (words.Find(D => D == word) == null && common.Find(D => D == word) == null)
                {
                    words.Add(word);
                }
            }
            return words;
        }
    }
}
