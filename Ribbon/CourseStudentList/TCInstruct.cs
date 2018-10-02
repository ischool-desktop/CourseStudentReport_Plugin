using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Framework;
using FISCA.DSAUtil;
using FISCA.Presentation;

namespace StuinCoursePlugins
{
    public class TCInstruct : CacheManager<TCInstructRecord>
    {
        private static TCInstruct _Instance = null;
        public static TCInstruct Instance
        {
            get
            {
                if (_Instance == null) _Instance = new TCInstruct();
                return _Instance;
            }
        }

        /// <summary>
        /// 從(課程)查詢授課教師。
        /// </summary>
        private Dictionary<string, List<string>> _course_teachers = new Dictionary<string, List<string>>();
        /// <summary>
        /// 從(教師)查詢教授課程。
        /// </summary>
        private Dictionary<string, List<string>> _teacher_courses = new Dictionary<string, List<string>>();

        private ListPaneField teacherField;

        private TCInstruct()
        {
            ItemLoaded += delegate
            {
                #region 重建課程及學生修課反查表
                _course_teachers.Clear();
                _teacher_courses.Clear();
                foreach (var item in Items)
                {
                    if (!_course_teachers.ContainsKey(item.RefCourseID))
                        _course_teachers.Add(item.RefCourseID, new List<string>());

                    if (!_teacher_courses.ContainsKey(item.RefTeacherID))
                        _teacher_courses.Add(item.RefTeacherID, new List<string>());

                    _teacher_courses[item.RefTeacherID].Add(item.ID);
                    _course_teachers[item.RefCourseID].Add(item.ID);
                }
                #endregion

                teacherField.Reload();
            };

            ItemUpdated += delegate (object sender, ItemUpdatedEventArgs e)
            {
                #region 更新(課成)及(教師)修課反查表
                List<string> keys = new List<string>(e.PrimaryKeys);
                keys.Sort(); //排序後的資料才可以進行 BinarySearch。

                #region 掃描每一個(課程)的(所有授課教師)是否有在(e.PrimaryKeys)中出現。
                foreach (var cid in _course_teachers.Keys)
                {
                    List<string> removeItems = new List<string>();
                    foreach (var eachInstruct in _course_teachers[cid])
                    {
                        if (keys.BinarySearch(eachInstruct) >= 0)
                            removeItems.Add(eachInstruct);
                    }

                    foreach (var eachInstruct in removeItems)
                        _course_teachers[cid].Remove(eachInstruct);
                }
                #endregion

                #region 掃描每一個(教師)的(所有課程)是否有在(e.PrimaryKeys)中出現。
                foreach (var cid in _teacher_courses.Keys)
                {
                    List<string> removeItems = new List<string>();
                    foreach (var eachInstruct in _teacher_courses[cid])
                    {
                        if (keys.BinarySearch(eachInstruct) >= 0)
                            removeItems.Add(eachInstruct);
                    }

                    foreach (var item in removeItems)
                        _teacher_courses[cid].Remove(item);
                }
                #endregion

                #region 將新資料加入到原有的集合中。
                foreach (var key in e.PrimaryKeys)
                {
                    var item = Items[key];
                    if (item != null)
                    {
                        if (!_course_teachers.ContainsKey(item.RefCourseID))
                            _course_teachers.Add(item.RefCourseID, new List<string>());

                        if (!_teacher_courses.ContainsKey(item.RefTeacherID))
                            _teacher_courses.Add(item.RefTeacherID, new List<string>());

                        _course_teachers[item.RefCourseID].Add(item.ID);
                        _teacher_courses[item.RefTeacherID].Add(item.ID);
                    }
                }
                #endregion

                teacherField.Reload();
                #endregion
            };
        }

        protected override Dictionary<string, TCInstructRecord> GetAllData()
        {
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("All");

            DSRequest dsreq = new DSRequest(helper);
            Dictionary<string, TCInstructRecord> result = new Dictionary<string, TCInstructRecord>();
            string srvname = "SmartSchool.Course.GetTCInstruct";
            foreach (var item in FISCA.Authentication.DSAServices.CallService(srvname, dsreq).GetContent().GetElements("TCInstruct"))
            {
                helper = new DSXmlHelper(item);
                var teacherid = helper.GetText("RefTeacherID");
                var courseid = helper.GetText("RefCourseID");
                var id = item.GetAttribute("ID");
                var sequence = helper.GetText("Sequence");

                TCInstructRecord record = new TCInstructRecord(teacherid, courseid, id, sequence);
                result.Add(record.ID, record);
            }
            return result;
        }

        protected override Dictionary<string, TCInstructRecord> GetData(IEnumerable<string> primaryKeys)
        {
            // 指示是否需要呼叫 Service。
            bool execute_require = false;

            //建立 Request。
            DSXmlHelper helper = new DSXmlHelper("Request");
            helper.AddElement("All");
            helper.AddElement("Condition");
            foreach (string id in primaryKeys)
            {
                helper.AddElement("Condition", "ID", id);
                execute_require = true;
            }

            //儲存最後結果的集合。
            Dictionary<string, TCInstructRecord> result = new Dictionary<string, TCInstructRecord>();

            if (execute_require)
            {
                string srvname = "SmartSchool.Course.GetTCInstruct";
                DSRequest dsreq = new DSRequest(helper);

                foreach (var item in FISCA.Authentication.DSAServices.CallService(srvname, dsreq).GetContent().GetElements("TCInstruct"))
                {
                    helper = new DSXmlHelper(item);
                    var teacherid = helper.GetText("RefTeacherID");
                    var courseid = helper.GetText("RefCourseID");
                    var id = item.GetAttribute("ID");
                    var sequence = helper.GetText("Sequence");

                    TCInstructRecord record = new TCInstructRecord(teacherid, courseid, id, sequence);
                    result.Add(record.ID, record);
                }
            }

            return result;
        }

        /// <summary>
        /// 取得課程的所有授課教師。
        /// </summary>
        /// <param name="courseID">課程編號。</param>
        /// <returns>授課教師清單。</returns>
        public List<TCInstructRecord> GetCourseTeachers(string courseID)
        {
            List<TCInstructRecord> result = new List<TCInstructRecord>();
            if (_course_teachers.ContainsKey(courseID))
            {
                foreach (var eachInstructID in _course_teachers[courseID])
                {
                    var objInstruct = Items[eachInstructID];
                    if (objInstruct.Course != null && objInstruct.Teacher != null)
                        result.Add(Items[eachInstructID]);
                }
            }
            return result;
        }
    }
}
