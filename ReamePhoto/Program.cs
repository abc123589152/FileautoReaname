using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
namespace FileautoReaname
{
    internal class Program
    {
        private static void SortAsFileCreationTime(ref FileInfo[] arrFi)
        {
            Array.Sort(arrFi, delegate (FileInfo x, FileInfo y) { return x.CreationTime.CompareTo(y.CreationTime); });//Array.Sort進行排序陣列
        }
        private static void Main(string[] args)
        {
            
            string orignatefilePath,CopyfilePath,ExcelfilePath;
            string[] array1 = new string[300];
            string TodayTime = DateTime.Now.ToString("yyyy-MM-dd");
            int ExcelSheetNum;
            int SUM_realexportNUM = 0;
            Console.Write("輸入Excel位置(檔案名稱:報稅名稱.xlsx):");
            ExcelfilePath = Console.ReadLine();
            Console.Write("輸入sheet號碼(1.name1 2.name2 3.name3):");
            ExcelSheetNum = int.Parse(Console.ReadLine());
            Console.Write("輸入原始檔案位置:");
            orignatefilePath = Console.ReadLine();
            Console.Write("輸入copy到的檔案位置:");
            CopyfilePath = Console.ReadLine();
            IWorkbook workbook = null;
            FileStream fs = new FileStream(ExcelfilePath, FileMode.Open, FileAccess.Read);
            if (ExcelfilePath.IndexOf(".xlsx") > 0)
                workbook = new XSSFWorkbook(fs);
            else if (ExcelfilePath.IndexOf(".xls") > 0)
                workbook = new HSSFWorkbook(fs);
            ISheet sheet = workbook.GetSheetAt(ExcelSheetNum - 1);//取得sheet number，輸入sheet數字時以十進位所以要-1
            for (int i = 0; i<=sheet.LastRowNum; i++)
            {
                try
                {
                    IRow row = sheet.GetRow(i);//取得行數              
                    array1[i] = CopyfilePath + @"\" + row.GetCell(1).ToString() + "_" + TodayTime + ".jpeg";//row.GetCell(1):取得第一行第二格
                    Console.WriteLine(array1[i]);//顯示所有複製的名稱及詳細位址
                    DirectoryInfo di = new DirectoryInfo(orignatefilePath);
                    FileInfo[] arrFi = di.GetFiles("*.*");//取得有相似字的相關檔案
                    SortAsFileCreationTime(ref arrFi);//丟到function去做創建時間排序後丟回
                    File.Copy(orignatefilePath + @"\" + arrFi[i].Name, array1[i]);//File.Copy("old FilePath and Name","New FilePath and Name")                    
                    SUM_realexportNUM = i;
                }
                catch (Exception ex) {
                    break;
                }
            }
            Console.WriteLine("需要匯出的總共有" + (sheet.LastRowNum+1)+ "筆");
            Console.WriteLine("實際匯出有"+(SUM_realexportNUM+1)+"筆");
            Console.WriteLine("輸入Enter退出");
            Console.ReadLine();
        }
    }
}
