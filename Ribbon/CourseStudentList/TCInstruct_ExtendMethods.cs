using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool;

namespace StuinCoursePlugins
{
    public static class TCInstruct_ExtendMethods
    {
        /// <summary>
        /// 取得課程的第一位授課教師。
        /// </summary>
        public static TeacherRecord GetFirstTeacher(this CourseRecord course)
        {
            if (course != null)
            {
                TCInstructRecord tc = course.GetFirstInstruct();
                if (tc != null) return tc.Teacher;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的第一位授課教師。
        /// </summary>
        internal static TCInstructRecord GetFirstInstruct(this CourseRecord course)
        {
            if (course != null)
            {
                foreach (TCInstructRecord each in course.GetInstructs())
                    if (each.Sequence == "1") return each;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的所有上課教師關聯資料。
        /// </summary>
        public static List<TCInstructRecord> GetInstructs(this CourseRecord course)
        {
            if (course != null)
                return TCInstruct.Instance.GetCourseTeachers(course.ID);
            else
                return null;
        }

        /// <summary>
        /// 取得課程的第二位授課教師。
        /// </summary>
        public static TeacherRecord GetSecondTeacher(this CourseRecord course)
        {
            if (course != null)
            {
                TCInstructRecord tc = course.GetSecondInstruct();
                if (tc != null) return tc.Teacher;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的第二位授課教師。
        /// </summary>
        internal static TCInstructRecord GetSecondInstruct(this CourseRecord course)
        {
            if (course != null)
            {
                foreach (TCInstructRecord each in course.GetInstructs())
                    if (each.Sequence == "2") return each;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的第三位授課教師。
        /// </summary>
        public static TeacherRecord GetThirdTeacher(this CourseRecord course)
        {
            if (course != null)
            {
                TCInstructRecord tc = course.GetThirdInstruct();
                if (tc != null) return tc.Teacher;
            }
            return null;
        }

        /// <summary>
        /// 取得課程的第三位授課教師。
        /// </summary>
        internal static TCInstructRecord GetThirdInstruct(this CourseRecord course)
        {
            if (course != null)
            {
                foreach (TCInstructRecord each in course.GetInstructs())
                    if (each.Sequence == "3") return each;
            }
            return null;
        }
    }
}
