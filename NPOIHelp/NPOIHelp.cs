using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.SS;
using Microsoft.VisualBasic;
using NPOI.SS.UserModel;
namespace NPOIHelp
{
    /// <summary>
    /// NPOI操作类库
    /// </summary>
    public class NPOIHelp
    {
        #region 公共变量
        /// <summary>
        /// 文件是否打开标志
        /// </summary>
        private bool _hasOpen;
        /// <summary>
        /// 当前工作行数
        /// </summary>
        private int working_row = 0;
        /// <summary>
        /// 最大的行数
        /// </summary>
        private int max_row_num = 1000;
        /// <summary>
        /// 定义了一个消息
        /// </summary>
        private string message = "";
        /// <summary>
        /// 操作的工作簿
        /// </summary>
        private HSSFWorkbook xlBook;
        /// <summary>
        /// 操作的工作表
        /// </summary>
        private NPOI.SS.UserModel.ISheet xlSheet;
        /// <summary>
        /// 打开，保存的文件名
        /// </summary>
        private string strFilename;
        #endregion
        #region 公共方法
        /// <summary>
        /// 向指定的行和列的单元格写值
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="value">值</param>
        /// <returns></returns>
        private bool setData(int rowIndex, int columnIndex, string value)
        {
            //rowIndex--;
            //columnIndex--;
            NPOI.SS.UserModel.IRow row = null;
            if (xlSheet.GetRow(rowIndex) != null)
            {
                row = xlSheet.GetRow(rowIndex);
            }
            else
            {
                row = xlSheet.CreateRow(rowIndex);
            }
            /*
            if (xlSheet.LastRowNum == 0)
            {
                row = xlSheet.CreateRow(0);
            }
            else
            {
                if (xlSheet.LastRowNum >= rowIndex)
                {
                    if (xlSheet.GetRow(rowIndex) != null)
                    {
                        row = xlSheet.GetRow(rowIndex);
                    }
                    else
                    {
                        row = xlSheet.CreateRow(rowIndex);
                    }
                }
                else
                {
                    row = xlSheet.CreateRow(rowIndex);
                }
            }
             * */
            NPOI.SS.UserModel.ICell cell = null;
            if (row.GetCell(columnIndex) != null)
            {
                cell = row.GetCell(columnIndex);
            }
            else
            {
                cell = row.CreateCell(columnIndex);
            }
            /*
            if (row.LastCellNum >= columnIndex)
            {
                if (row.GetCell(columnIndex) != null)
                {
                    cell = row.GetCell(columnIndex);
                }
                else
                {
                    cell = row.CreateCell(columnIndex);
                }

            }
            else
            {
                cell = row.CreateCell(columnIndex);
            }
             * */
            cell.SetCellValue(value);
            return true;
        }
        /// <summary>
        /// 获取指定行列的数据
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns>该单元格的内容</returns>
        private string getData(int rowIndex, int columnIndex)
        {
            //rowIndex--;
            //columnIndex--;
            if (xlSheet.LastRowNum >= rowIndex)
            {
                NPOI.SS.UserModel.IRow row = null;
                if (xlSheet.LastRowNum == 0)
                {
                    row = xlSheet.CreateRow(rowIndex);
                }
                else
                {
                    row = xlSheet.GetRow(rowIndex);
                }
                if (row.LastCellNum >= columnIndex)
                {
                    if (row.GetCell(columnIndex) != null)
                    {
                        return row.GetCell(columnIndex).ToString();
                    }
                    else
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }
        /// <summary>
        /// 设置指定行列的单元格的背景色
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="color">颜色</param>
        private void setAlarm(int rowIndex, int columnIndex)
        {
            // rowIndex--;
            //columnIndex--;
            NPOI.SS.UserModel.IRow row = null;
            NPOI.SS.UserModel.ICell cell = null;
            //if (xlSheet.LastRowNum >= rowIndex)
            //{
            if (xlSheet.GetRow(rowIndex) != null)
            {
                row = xlSheet.GetRow(rowIndex);
            }
            else
            {
                row = xlSheet.CreateRow(rowIndex);
            }
            //if (xlSheet.LastRowNum == 0)
            //{
            //    row = xlSheet.CreateRow(rowIndex);
            //}
            //else
            //{
            //    row = xlSheet.GetRow(rowIndex);
            //}
            //}
            //else
            //{
            //    row = xlSheet.CreateRow(rowIndex);
            //}
            if (row.GetCell(columnIndex) != null)
            {
                cell = row.GetCell(columnIndex);
            }
            else
            {
                cell = row.CreateCell(columnIndex);
            }
            ICellStyle style = xlBook.CreateCellStyle();
            style.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Yellow.Index;
            style.FillPattern = FillPattern.SolidForeground;
            cell.CellStyle = style;
        }
        /// <summary>
        /// 删除指定工作表名的所有行
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        private void sheetClear(string sheetName)
        {
            xlSheet = xlBook.GetSheet(sheetName);
            for (int i = 0; i < xlSheet.LastRowNum; i++)
            {
                NPOI.SS.UserModel.IRow row=xlSheet.GetRow(i);
                xlSheet.RemoveRow(row);
            }
        }
        #endregion
        public int getValidedRowCount()
        {
            return -1;
        }
        public string getURL()
        {
            return "";
        }
        public bool hasOpen()
        {
            return _hasOpen;
        }
        /// <summary>
        /// 判断文件是否支持读写操作
        /// </summary>
        /// <param name="fileName">文件名全路径</param>
        /// <returns>被占用返回false，反正返回true</returns>
        private bool isFileCanWrite(string fileName)
        {
            System.IO.FileStream fs = null;
            try
            {
                fs = new System.IO.FileStream(fileName, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.Write);
                if (!fs.CanWrite)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
            finally
            {
                fs.Close();
                fs.Dispose();
            }
        }
        /// <summary>
        /// 打开xls文件，我修改成了，如果没有这个文件，就创建这个文件
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="sheetName">工作表名</param>
        /// <param name="clear">清空标识？</param>
        /// <returns></returns>
        public string open_xls(string fileName = "buyma-need-sell.xls", string sheetName = "data", bool clear = true)
        {
            try
            {
                //文件名 全路径
                strFilename = System.IO.Directory.GetCurrentDirectory() + "\\" + fileName;
                if (System.IO.File.Exists(strFilename) == false)
                {
                    //这里有修改项，如果没有这个xls文件，我就进行创建操作
                    //System.IO.File.Create(strFilename);
                    xlBook = new HSSFWorkbook();
                    xlBook.CreateSheet(sheetName);
                }
                else
                {
                    if (isFileCanWrite(strFilename) == false)
                    {
                        return "当該ファイル（" + fileName + ")は既に他のアプリより開いている状態であるため、まず関連アプリを閉じてください";
                    }
                    System.IO.FileStream fs = new System.IO.FileStream(strFilename, System.IO.FileMode.Open);
                    xlBook = new HSSFWorkbook(fs);
                    xlBook.SetSheetName(0, sheetName);
                }
                //在工作簿中以修改第一个工作表为指定工作表名

                xlSheet = xlBook.GetSheet(sheetName);
                working_row = 1;
                _hasOpen = true;
                return "";
            }
            catch (Exception ex)
            {
                return "sheet [data] is not exit." + ex.ToString();
                throw;
            }
        }
        /// <summary>
        /// 获取当前工作行数
        /// </summary>
        /// <returns></returns>
        public int get_row()
        {
            return this.working_row;
        }
        /// <summary>
        /// 指定列数，获取当前工作行中指定列，单元格的值
        /// </summary>
        /// <param name="read_lie">指定的列索引</param>
        /// <returns></returns>        
        public string get_data(int read_lie)
        {
            string result = null;
            result = getData(this.working_row, read_lie);
            if (result.Length > 0)
            {
                return result;
            }
            else
            {
                return "";
            }
        }
        /// <summary>
        /// 使用指定列和当前工作行，向单元格中写值
        /// </summary>
        /// <param name="set_lie">指定的列索引</param>
        /// <param name="str">要设定的值</param>
        /// <returns></returns>
        public bool set_data(int set_lie, string str)
        {
            setData(this.working_row, set_lie, str);
            return true;
        }
        /// <summary>
        /// 设定当前工作行
        /// </summary>
        /// <param name="row_value">当前工作行的索引</param>
        public void set_row(int row_value)
        {
            working_row = row_value;
        }
        /// <summary>
        /// 设置最大的行数
        /// </summary>
        /// <param name="value">最大行数的值</param>
        public void set_max_row_num(int value)
        {
            max_row_num = value;
        }
        public void set_alarm(int set_lie)
        {
            setAlarm(this.working_row, set_lie);
            /*
            string loc_txt = "A1";
            if (set_lie <= 26)
            {
             //测试下更改是否可行
             //测试vs下提交是否可行
                loc_txt = Strings.Chr(64 + set_lie) + Conversion.Str(working_row);

            }
            else if (set_lie <= 52)
            {
                loc_txt = "A" + Strings.Chr(64 + set_lie - 26) + Conversion.Str(working_row);
            }
            else if (set_lie <= 78)
            {
                loc_txt = "B" + Strings.Chr(64 + set_lie - 52) + Conversion.Str(working_row);
            }
            else if (set_lie <= 104)
            {
                loc_txt = "C" + Strings.Chr(64 + set_lie - 78) + Conversion.Str(working_row);
            }
            loc_txt = Strings.Replace(loc_txt, " ", "", 1, -1, CompareMethod.Binary);
           NPOI.SS.UserModel.ICell cell=xlSheet.
            //Excel.Range xlRange = xlSheet.Range(loc_txt);
            //dynamic xlInterior = xlRange.Interior;
            //xlInterior.ColorIndex = 6;
            //黄色
             * */
        }
        /// <summary>
        /// 下移工作行
        /// </summary>
        /// <returns>当前工作行与最大工作行的大小关系</returns>
        public bool move_2_next()
        {
            if (working_row <= max_row_num)
            {
                working_row = working_row + 1;
                return true;
            }
            else
            {
                return false;
            }
        }
        public void close()
        {
            if (hasOpen())
            {
                System.IO.FileStream fs = new System.IO.FileStream(strFilename, System.IO.FileMode.OpenOrCreate);
                xlBook.Write(fs);
                xlBook.Close();
                xlSheet = null;
                xlBook = null;
                _hasOpen = false;
                fs.Close();
                fs.Dispose();
            }

        }
    }
}
