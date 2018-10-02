using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using System.IO;
using JHSchool.Data;
using JHSchool;
using System.Windows.Forms;
using System.Diagnostics;

namespace StuinCoursePlugins
{
    class CourseStudentList
    {
        private Workbook _book;
        private Workbook _template;
        private List<CourseRecord> _CourseList;

        public CourseStudentList()
        {
            this._book = new Workbook(new MemoryStream(Properties.Resources.課程修課學生清單));
            this._template = new Workbook(new MemoryStream(Properties.Resources.課程修課學生清單));

            this._CourseList = Course.Instance.SelectedList;
            _CourseList.Sort();
        }

        public void Execute()
        {
            SmartSchool.Customization.PlugIn.Global.SetStatusBarMessage("處理中，請稍候...", 0);

            this._book.Worksheets[0].Cells.Clear();

            int row = 0;
            int page = 1;

            // 將所選取的班級，資料取出
            foreach (CourseRecord cr in this._CourseList)
            {
                #region 標題
                {
                    this._book.Worksheets[0].Cells.CopyRow(this._template.Worksheets[0].Cells, 0, row);
                    this._book.Worksheets[0].Cells[row++, 0].PutValue(cr.SchoolYear + "學年度第" + cr.Semester + "學期 課程學生修課清單");
                }
                #endregion

                #region 第二行
                {
                    this._book.Worksheets[0].Cells.CopyRow(this._template.Worksheets[0].Cells, 1, row);
                    // 課程名稱
                    this._book.Worksheets[0].Cells[row, 1].PutValue(cr.Name);
                    // 節次/權數
                    this._book.Worksheets[0].Cells[row, 6].PutValue(cr.Period);
                    #region 授課教師
                    {
                        if (cr.GetFirstTeacher() != null)
                        {
                            this._book.Worksheets[0].Cells[row, 8].PutValue(cr.GetFirstTeacher().Name);
                        }
                        if (cr.GetSecondTeacher() != null)
                        {
                            this._book.Worksheets[0].Cells[row, 8].PutValue(this._book.Worksheets[0].Cells[row, 6].StringValue + "," + cr.GetSecondTeacher().Name);
                        }
                        if (cr.GetThirdTeacher() != null)
                        {
                            this._book.Worksheets[0].Cells[row, 8].PutValue(this._book.Worksheets[0].Cells[row, 6].StringValue + "," + cr.GetThirdTeacher().Name);
                        }

                        //欄位重新定位
                        this._book.Worksheets[0].AutoFitColumn(8, row, row);
                    }
                    #endregion
                    row++;
                }
                #endregion

                IEnumerable<JHSchool.Data.JHSCAttendRecord> scr = JHSchool.Data.JHSCAttend.SelectByCourseIDs(new string[] { cr.ID });
                IEnumerable<string> studentids = from screc in scr select screc.RefStudentID;
                // 取得一般生
                List<JHSchool.Data.JHStudentRecord> students = JHStudent.SelectByIDs(studentids).Where(x => x.Status == K12.Data.StudentRecord.StudentStatus.一般).ToList();
                students.Sort(ParseStudent);

                #region 第三行
                {
                    this._book.Worksheets[0].Cells.CopyRow(this._template.Worksheets[0].Cells, 2, row);

                    // 所屬科目
                    this._book.Worksheets[0].Cells[row, 1].PutValue(cr.Subject);

                    #region 修課人數
                    {
                        if (cr.Class == null)
                        {
                            this._book.Worksheets[0].Cells[row, 8].PutValue(null);
                        }
                        else
                        {
                            // 修課學生人數
                            this._book.Worksheets[0].Cells[row, 8].PutValue(students.Count);
                        }
                    }
                    #endregion

                    row++;
                }
                #endregion

                this._book.Worksheets[0].Cells.CopyRow(this._template.Worksheets[0].Cells, 3, row++);

                #region 填入課程的修課學生資料
                {
                    int n = 0;
                    foreach (JHStudentRecord student in students)
                    {
                        n++;
                        if (students.Count == n)
                        {
                            this._book.Worksheets[0].Cells.CopyRow(this._template.Worksheets[0].Cells, 6, row);
                        }
                        else
                        {
                            this._book.Worksheets[0].Cells.CopyRow(this._template.Worksheets[0].Cells, 5, row);
                        }

                        #region 學生班級
                        {
                            if (student.Class != null)
                            {
                                this._book.Worksheets[0].Cells[row, 0].PutValue(student.Class.Name);
                            }
                            else
                            {
                                this._book.Worksheets[0].Cells[row, 0].PutValue("");
                            }
                        }
                        #endregion
                        this._book.Worksheets[0].Cells[row, 1].PutValue(student.SeatNo); // 座號
                        this._book.Worksheets[0].Cells[row, 2].PutValue(student.StudentNumber); // 學號
                        this._book.Worksheets[0].Cells[row, 3].PutValue(student.EnglishName); // 英文姓名
                        this._book.Worksheets[0].Cells[row, 5].PutValue(student.Name); // 姓名
                        this._book.Worksheets[0].Cells[row, 7].PutValue(student.Gender); // 性別
                        row++;
                    }
                }
                #endregion

                this._book.Worksheets[0].Cells.CopyRow(this._template.Worksheets[0].Cells, 7, row);
                this._book.Worksheets[0].Cells[row, 7].PutValue("第" + page + "頁/共" + _CourseList.Count + "頁");
                row++;
                page++;

                this._book.Worksheets[0].HorizontalPageBreaks.Add(row, 0);
            }

            Save();
        }

        private void Save()
        {
            try
            {
                SmartSchool.Customization.PlugIn.Global.SetStatusBarMessage("請選擇儲存位置", 100);
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                SaveFileDialog1.Filter = "Excel (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
                SaveFileDialog1.FileName = "課程修課清單";

                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    this._book.Save(SaveFileDialog1.FileName);
                    Process.Start(SaveFileDialog1.FileName);
                }
                else
                {
                    MessageBox.Show("檔案未儲存");

                }
            }
            catch
            {
                MessageBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
            }

            SmartSchool.Customization.PlugIn.Global.SetStatusBarMessage("已完成");
        }

        //排序功能
        private int ParseStudent(JHStudentRecord x, JHStudentRecord y)
        {
            //取得班級名稱
            string Xstring = x.Class != null ? x.Class.Name : "";
            string Ystring = y.Class != null ? y.Class.Name : "";

            //取得座號
            string Xint = x.SeatNo.HasValue ? x.SeatNo.ToString() : "";
            string Yint = y.SeatNo.HasValue ? y.SeatNo.ToString() : "";
            //班級名稱加:號加座號(靠右對齊補0)
            Xstring += ":" + Xint.PadLeft(2, '0');
            Ystring += ":" + Yint.PadLeft(2, '0');

            return Xstring.CompareTo(Ystring);

        }
    }
}
