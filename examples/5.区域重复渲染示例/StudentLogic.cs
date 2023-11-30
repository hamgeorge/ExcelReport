using System.Collections.Generic;

namespace _5.区域重复渲染示例
{
    public static class StudentLogic
    {
        public static List<StudentInfo> GetList()
        {
            List<StudentInfo> list = new List<StudentInfo>();

            list.Add(new StudentInfo() { Class = "一班", Name = "XXX01", Gender = true, RecordNo = "YYY0001", Phone = "158******01", Email = "xxx01@live.cn" });
            list.Add(new StudentInfo() { Class = "二班", Name = "XXX02", Gender = false, RecordNo = "YYY0002", Phone = "158******02", Email = "xxx02@live.cn" });
            list.Add(new StudentInfo() { Class = "一班", Name = "XXX03", Gender = true, RecordNo = "YYY0003", Phone = "158******03", Email = "xxx03@live.cn" });
            list.Add(new StudentInfo() { Class = "一班", Name = "XXX04", Gender = true, RecordNo = "YYY0004", Phone = "158******04", Email = "xxx04@live.cn" });

            return list;
        }

        public static List<Teacher> GetTeacherList()
        {
            List<Teacher> list = new List<Teacher>();

            list.Add(new Teacher() { Class = "一班", Name = "Wang", Gender = true});
            list.Add(new Teacher() { Class = "二班", Name = "Lee", Gender = false });
            list.Add(new Teacher() { Class = "一班", Name = "Han", Gender = true });
            //list.Add(new Teacher() { Class = "一班", Name = "Zhao", Gender = true});

            return list;
        }
    }
}