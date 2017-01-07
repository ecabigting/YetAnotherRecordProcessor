using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace YetAnotherRecordProcessor
{
    public partial class Form1 : Form
    {
        Excel.Application xlAppSource;
        Excel.Workbook xlWorkBookSource;
        Excel.Worksheet xlWorkSheetSource;
        Excel.Range xlRangeSource;

        Excel.Application xlAppTarget;
        Excel.Workbook xlWorkBookTarget;
        Excel.Worksheet xlWorkSheetTarget;
        Excel.Range xlRangeTarget;
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            xlAppSource = new Excel.Application();
            xlWorkBookSource = xlAppSource.Workbooks.Open("D:\\exf\\source-file.xlsx", 0, true, 5, "", "", true,
                                                            Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                                                            false, false, 0, true, 1, 0);
            xlWorkSheetSource = (Excel.Worksheet)xlWorkBookSource.Worksheets.get_Item(1);

            xlAppTarget = new Excel.Application();
            xlWorkBookTarget = xlAppTarget.Workbooks.Open("D:\\exf\\90-91-target.xlsx", 0, false, 5, "", "", true,
                                                            Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t",
                                                            false, false, 0, true, 1, 0);
            xlWorkSheetTarget = (Excel.Worksheet)xlWorkBookTarget.Worksheets.get_Item(1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            xlAppSource.Visible = true;
            xlAppTarget.Visible = true;
            string studName = "";
            object misValue = System.Reflection.Missing.Value;
            //xlRangeSource = xlWorkSheetSource.UsedRange;
            xlRangeTarget = xlWorkSheetTarget.UsedRange;
            xlRangeSource = xlWorkSheetSource.UsedRange;

            for (int i = 2; i <= xlRangeTarget.Rows.Count; i++)
            {
                string currPupNum = xlRangeTarget.Cells[i, 1].Text;
                studName = xlRangeTarget.Cells[i, 3].Text;
                string[] studNameSplt = studName.ToString().Split(' ');
                int studNameLen = studNameSplt.Length;

                if(currPupNum.Trim() == "")
                {
                    string pupNumber = findtheFuckingPupilNumber(studNameSplt[0],
                                                             studNameSplt[1].Replace(" ",""),
                                                             studNameSplt[studNameLen - 1].Replace(" ", ""),
                                                             studNameSplt[studNameLen - 2].Replace(" ", "") + " " + studNameSplt[studNameLen - 1].Replace(" ", "")
                                                            );
                    if(pupNumber != null || pupNumber != " " || pupNumber != "")
                    {
                        xlRangeTarget.Cells[i, 1].Value = pupNumber;
                    }
                }
            }
            
            MessageBox.Show("fucking done! goodluck!");

            
        }

        private string findtheFuckingPupilNumber(string fname,string mname,string lname, string lname2nd)
        {
            //Excel.Range TempRange = tryToFindThisBitch(fname);
            for (int yy = 2; yy <= xlRangeSource.Rows.Count; yy++)
            {
                string currLName = xlRangeSource.Cells[yy, 6].Text;//col last name
                string currMName = xlRangeSource.Cells[yy, 5].Text;//col middlename
                if (currLName.Trim().Replace(" ", "") == lname)
                {
                    string currFName = xlRangeSource.Cells[yy, 4].Text;
                    if (currFName.Trim().Replace(" ", "") == fname)
                    {
                        string pupNum = xlRangeSource.Cells[yy, 3].Text;
                        return pupNum;
                    }
                }
                else
                {
                    if (currMName.Trim().Replace(" ", "") == mname)
                    {
                        string pupNum = xlRangeSource.Cells[yy, 3].Text;
                        return pupNum;
                    }
                }
            }
            return "";
        }

        public Excel.Range tryToFindThisBitch(string fname)
        {
            return null;
            //try
            //{
                //return xlRangeSource.Find(fname.Trim(), Type.Missing,
                //                         Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                //                         Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                //                         Type.Missing, Type.Missing).EntireRow;
                
            //}catch
            //{
            //    return null;
            //}
        }


        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            releaseObject(xlAppTarget);
            releaseObject(xlWorkBookTarget);
            releaseObject(xlWorkSheetTarget);

            releaseObject(xlAppSource);
            releaseObject(xlWorkBookSource);
            releaseObject(xlWorkSheetSource);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            releaseObject(xlAppTarget);
            releaseObject(xlWorkBookTarget);
            releaseObject(xlWorkSheetTarget);

            releaseObject(xlAppSource);
            releaseObject(xlWorkBookSource);
            releaseObject(xlWorkSheetSource);
        }
    }
}
