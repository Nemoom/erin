using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;

namespace Erin_s
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        List<CS_staff> CS_staffs = new List<CS_staff>();

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                CS_staffs = new List<CS_staff>();
                DirectoryInfo folder = new DirectoryInfo(@"Q:\CNGrp095\FEC&I\Working Hours Management");
                //获取文件夹下所有的文件
                FileInfo[] fileList = folder.GetFiles();
                foreach (FileInfo file in fileList)
                {
                    //判断文件的扩展名是否为 .gif
                    if (file.Extension == ".csv")
                    {                        
                        try
                        {
                            CS_staff aCS_staff = new CS_staff();
                            aCS_staff.teamName = file.Name.Split('-')[0];
                            aCS_staff.name = (file.Name.Split('-')[1]).Split('.')[0];
                            using (StreamReader sReader = new StreamReader(file.FullName))
                            {
                                while (sReader.Peek() >= 0)
                                {
                                    string mStr = sReader.ReadLine();
                                    if (mStr.StartsWith("20"))
                                    {
                                        string[] Array_mStr = mStr.Split(',');
                                        string strDate = Array_mStr[0].Split(' ')[0];
                                        DateTime dt;
                                        DateTimeFormatInfo dtFormat = new System.Globalization.DateTimeFormatInfo();
                                        //dtFormat.ShortDatePattern = "yyyyMMdd";
                                        try
                                        {
                                            dt = DateTime.ParseExact(strDate, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture);
                                        }
                                        catch (Exception)
                                        {
                                            dtFormat.ShortDatePattern = "yyyy/MM/dd";
                                            dt = Convert.ToDateTime(strDate, dtFormat);
                                        }                                       
                                        
                                        string mstr = Array_mStr[0].Split(' ')[1];
                                        if (dt.Year==DateTime.Now.Year)
                                        {
                                            switch (mstr)
                                            {
                                                case "上午":
                                                    aCS_staff.workHours[dt.Month - 1, dt.Day - 1] = aCS_staff.workHours[dt.Month - 1, dt.Day - 1] + 4;
                                                    aCS_staff.str_workHours[dt.Month - 1, dt.Day - 1] = aCS_staff.str_workHours[dt.Month - 1, dt.Day - 1] + "上午4";
                                                    break;
                                                case "下午":
                                                    aCS_staff.workHours[dt.Month - 1, dt.Day - 1] = aCS_staff.workHours[dt.Month - 1, dt.Day - 1] + 4;
                                                    aCS_staff.str_workHours[dt.Month - 1, dt.Day - 1] = aCS_staff.str_workHours[dt.Month - 1, dt.Day - 1] + "下午4";
                                                    break;
                                                case "上午&下午":
                                                    aCS_staff.workHours[dt.Month - 1, dt.Day - 1] = 8;
                                                    aCS_staff.str_workHours[dt.Month - 1, dt.Day - 1] = "8";
                                                    break;
                                                default:
                                                    break;
                                            }
                                        }
                                    }
                                }
                            }
                            CS_staffs.Add(aCS_staff);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            string csvFilePath = Path.Combine(DateTime.Now.ToString("yyyyMMdd") + ".csv");
            if (File.Exists(csvFilePath))
            { File.Delete(csvFilePath); }
            string line = string.Empty;
            using (StreamWriter csvFile = new StreamWriter(csvFilePath, true, Encoding.UTF8))
            {
                line = "Name,Sub_department,Jan,Feb,Mar,Apr,May,June,July,Aug,Sept,Oct,Nov,Dec,Total";
                csvFile.WriteLine(line);
                for (int i = 0; i < CS_staffs.Count; i++)
                {
                    line = CS_staffs[i].name + "," + CS_staffs[i].teamName + "," + getSum(i, 1) + ","
                        + getSum(i, 2) + ","
                        + getSum(i, 3) + ","
                        + getSum(i, 4) + ","
                        + getSum(i, 5) + ","
                        + getSum(i, 6) + ","
                        + getSum(i, 7) + ","
                        + getSum(i, 8) + ","
                        + getSum(i, 9) + ","
                        + getSum(i, 10) + ","
                        + getSum(i, 11) + ","
                        + getSum(i, 12) + ","
                        + getSum(i, 0) ;
                    csvFile.WriteLine(line);
                }
            }
            MessageBox.Show("完成，夸我！");
        }

        public int getSum(int staffIndex,int Month)
        {
            int mSum=0;
            List<int> list = new List<int>();
            if (Month==0)
            {
                for (int i = 0; i < 12; i++)
                {
                    for (int j = 0; j < 31; j++)
                    {
                        list.Add(CS_staffs[staffIndex].workHours[i,j]);
                    }
                }
            }
            else
            {
                for (int j = 0; j < 31; j++)
                {
                    list.Add(CS_staffs[staffIndex].workHours[Month-1, j]);
                }
            }
            mSum = list.Sum();
            return mSum;
        }

        public class CS_staff
        {
            public int[,] workHours = new int[12, 31];
            public string[,] str_workHours = new string[12, 31];
            public string name;
            public string teamName;
        }
    }
}
