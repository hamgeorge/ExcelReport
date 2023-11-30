using ExcelReport.Driver.NPOI.Extends;
using NPOI.Extend;
using System.Collections;
using System.Collections.Generic;
using NpoiRow = NPOI.SS.UserModel.IRow;
using NpoiSheet = NPOI.SS.UserModel.ISheet;
using NpoiCell = NPOI.SS.UserModel.ICell;
using NpoiUseModel = NPOI.SS.UserModel;
using NPOI.Extend.Formula;
using System.Text.RegularExpressions;

namespace ExcelReport.Driver.NPOI
{
    public class Sheet : ISheet
    {
        public NpoiSheet NpoiSheet { get; private set; }

        public Sheet(NpoiSheet npoiSheet)
        {
            NpoiSheet = npoiSheet;
        }

        public IRow this[int rowIndex] => NpoiSheet.GetRow(rowIndex).GetAdapter();

        public string SheetName => NpoiSheet.SheetName;

        public int CopyRows(int start, int end)
        {
            return NpoiSheet.CopyRows(start, end);
        }

        public int RemoveRows(int start, int end)
        {
            return NpoiSheet.RemoveRows(start, end);
        }

        public IEnumerator<IRow> GetEnumerator()
        {
            foreach (NpoiRow npoiRow in NpoiSheet)
            {
                yield return npoiRow.GetAdapter();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public object GetOriginal()
        {
            return NpoiSheet;
        }

        public int ReplaceRegion(int startRow, int endRow, int startColum, int endColumn)
        {
            return ReplaceRegionInternal(startRow, endRow, startColum, endColumn);
        }

        private int ReplaceRegionInternal(int startRowIndex, int endRowIndex, int startColumIndex, int endColumnIndex)
        {
            int span = endRowIndex - startRowIndex + 1;

            int newStartRowIndex = startRowIndex + span;
            //插入空行
            //sheet.InsertRows(newStartRowIndex, span);
            //复制行
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                NpoiRow sourceRow =NpoiSheet.GetRow(i);
                NpoiRow targetRow = NpoiSheet.GetRow(i + span);

                targetRow.Height = sourceRow.Height;
                targetRow.ZeroHeight = sourceRow.ZeroHeight;

                #region 替换单元格
                foreach (NpoiCell sourceCell in sourceRow.Cells)
                {
                    if (sourceCell.ColumnIndex< startColumIndex || sourceCell.ColumnIndex > endColumnIndex)
                    {
                        continue;
                    }
                    NpoiCell targetCell = targetRow.GetCell(sourceCell.ColumnIndex);
                    if (null == targetCell)
                    {
                        targetCell = targetRow.CreateCell(sourceCell.ColumnIndex);
                    }
                    if (null != sourceCell.CellStyle) targetCell.CellStyle = sourceCell.CellStyle;
                    if (null != sourceCell.CellComment) targetCell.CellComment = sourceCell.CellComment;
                    if (null != sourceCell.Hyperlink) targetCell.Hyperlink = sourceCell.Hyperlink;
                    NpoiUseModel.IConditionalFormattingRule[] cfrs = sourceCell.GetConditionalFormattingRules(); //复制条件样式
                    if (null != cfrs && cfrs.Length > 0)
                    {
                        targetCell.AddConditionalFormattingRules(cfrs);
                    }
                    targetCell.SetCellType(sourceCell.CellType);

                    #region 复制值

                    switch (sourceCell.CellType)
                    {
                        case NpoiUseModel.CellType.Numeric:
                            targetCell.SetCellValue(sourceCell.NumericCellValue);
                            break;
                        case NpoiUseModel.CellType.String:
                            targetCell.SetCellValue(sourceCell.RichStringCellValue);
                            break;
                        case NpoiUseModel.CellType.Formula:
                            var formula = CopyFormula(sourceCell.CellFormula, span);
                            targetCell.SetCellFormula(formula);
                            break;
                        case NpoiUseModel.CellType.Blank:
                            targetCell.SetCellValue(sourceCell.StringCellValue);
                            break;
                        case NpoiUseModel.CellType.Boolean:
                            targetCell.SetCellValue(sourceCell.BooleanCellValue);
                            break;
                        case NpoiUseModel.CellType.Error:
                            targetCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                            break;
                    }

                    #endregion
                }

                #endregion
            }
            //获取模板行内的合并区域
            List<MergedRegionInfo> regionInfoList = NpoiSheet.GetMergedRegionInfos(startRowIndex, endRowIndex, null,
                null);
            //复制合并区域
            foreach (MergedRegionInfo regionInfo in regionInfoList)
            {
                regionInfo.FirstRow += span;
                regionInfo.LastRow += span;
                NpoiSheet.AddMergedRegion(regionInfo);
            }
            //获取模板行内的图片
            List<PictureInfo> picInfoList = NpoiSheet.GetAllPictureInfos(startRowIndex, endRowIndex, null, null);
            //复制图片
            foreach (PictureInfo picInfo in picInfoList)
            {
                picInfo.MaxRow += span;
                picInfo.MinRow += span;
                NpoiSheet.AddPicture(picInfo);
            }

            return span;
        }

        private static string CopyFormula(string formula, int span)
        {
            FormulaContext context = new FormulaContext(formula);
            context.Parse();
            return context.Formula.ToString(part =>
            {
                if (part.Type.Equals(PartType.Formula))
                {
                    Regex regex = new Regex(@"([A-Z]+)(\d+)");
                    return regex.Replace(part.ToString(), (m) => $"{m.Groups[1].Value}{int.Parse(m.Groups[2].Value) + span}");
                }
                else
                {
                    return part.ToString();
                }
            });
        }

        public int ClearRegion(int startRowIndex, int endRowIndex, int startColumIndex, int endColumnIndex)
        {
            int span = endRowIndex - startRowIndex + 1;
            for (int i = startRowIndex; i <= endRowIndex; i++)
            {
                NpoiRow sourceRow = NpoiSheet.GetRow(i);
                #region 替换单元格
                foreach (NpoiCell sourceCell in sourceRow.Cells)
                {
                    if (sourceCell.ColumnIndex < startColumIndex || sourceCell.ColumnIndex > endColumnIndex)
                    {
                        continue;
                    }                   
                    NpoiCell targetCell = sourceRow.GetCell(sourceCell.ColumnIndex);
                    sourceRow.RemoveCell(targetCell);
                }

                #endregion
            }

            return span;
        }
    }
}