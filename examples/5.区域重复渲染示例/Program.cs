using ExcelReport;
using ExcelReport.Driver.NPOI;
using ExcelReport.Renderers;
using System;

namespace _5.区域重复渲染示例
{
    class Program
    {
        static void Main(string[] args)
        {
            Configurator.Put(".xls", new WorkbookLoader());
            var re1 = new RegionRepeaterRenderer<StudentInfo>("rptStudentInfo", StudentLogic.GetList(),
                            new ParameterRenderer<StudentInfo>("Name", t => t.Name),
                            new ParameterRenderer<StudentInfo>("Gender", t => t.Gender ? "男" : "女"),
                            new ParameterRenderer<StudentInfo>("Class", t => t.Class),
                            new ParameterRenderer<StudentInfo>("RecordNo", t => t.RecordNo),
                            new ParameterRenderer<StudentInfo>("Phone", t => t.Phone),
                            new ParameterRenderer<StudentInfo>("Email", t => t.Email)
                            );
            var re2 = new RegionRepeaterRenderer<Teacher>("rptTeacherInfo", StudentLogic.GetTeacherList(),
                            new ParameterRenderer<Teacher>("tName", t => t.Name),
                            new ParameterRenderer<Teacher>("tGender", t => t.Gender ? "男" : "女"),
                            new ParameterRenderer<Teacher>("tClass", t => t.Class)
                            );
            var sheet = new SheetRenderer("区域重复渲染示例", re1, re2);
            try
            {
                ExportHelper.ExportToLocal(@"Template\Template.xls", "out.xls", sheet);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Console.WriteLine("finished!");
        }
    }
}
