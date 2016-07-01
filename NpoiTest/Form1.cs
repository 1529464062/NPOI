using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI;
using NPOI.HSSF;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOIHelp;

namespace NpoiTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            /*
            DateTime dt1 = DateTime.Now;
            NPOIHelp.NPOIHelp nh = new NPOIHelp.NPOIHelp();
            nh.open_xls("dd1dsfs1f.xls");
            for (int i = 0; i < 5534; i++)
            {
                nh.set_row(i);
                for ( int j = 0; j < 256; j++)
                {
                    nh.set_data(j, i +"-"+j+"sdfsgsdsdgadgefsdfafedfasfsfsdfsdffswefsefsefsefsfefss");
                    if (i == j)
                    {
                        nh.set_alarm(j);
                    }                    
                }
                
            }
                nh.close();
                DateTime dt2 = DateTime.Now;
                MessageBox.Show("输出用时"+(dt2 - dt1).TotalMilliseconds.ToString()+"毫秒");
             * */
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.ShowDialog();
            openFileName = ofd.FileName;
            open_xls(openFileName);
        }
        #region xls操作
        string openFileName;
        HSSFWorkbook hssfworkbook;
        public string open_xls(string filePath)
        {
            FileStream fs = new FileStream(openFileName, FileMode.Open);
            hssfworkbook = new HSSFWorkbook(fs);

            return "";
        }
        #endregion
    }
}
