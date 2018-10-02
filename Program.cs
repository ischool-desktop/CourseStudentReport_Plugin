using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Presentation;
using FISCA.Permission;
using JHSchool;
using FISCA.Presentation.Controls;
using Framework;

namespace StuinCoursePlugins
{
    public class Program
    {
        [MainMethod(" StuinCoursePlugins")]
        public static void Nain()
        {
            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["學生修課清單"].Visible = false;
            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["學生修課清單"].Text = "_學生修課清單";

            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["學生修課清單"].Enable = User.Acl["JHSchool.Course.Report0000"].Executable;
            MotherForm.RibbonBarItems["課程", "資料統計"]["報表"]["學生修課清單"].Click += delegate
            {
                if (Course.Instance.SelectedList.Count >= 1)
                {
                    (new CourseStudentList()).Execute();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇課程");
                }
            };
        }
    }
}
