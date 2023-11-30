using System.Collections.Generic;

namespace ExcelReport.Driver
{
    public interface ISheet : IEnumerable<IRow>, IAdapter
    {
        string SheetName { get; }

        IRow this[int rowIndex]
        {
            get;
        }

        int CopyRows(int start, int end);
        int ReplaceRegion(int startRow,int endRow,int startColum,int endColumn);

        int RemoveRows(int start, int end);
        int ClearRegion(int startRow, int endRow, int startColum, int endColumn);
    }
}