using System;
using System.Collections.Generic;
using System.Text;
using ExcelReport.Contexts;
using ExcelReport.Exceptions;
using ExcelReport.Extends;
using ExcelReport.Meta;
using System.Linq;

namespace ExcelReport.Renderers
{

    public class RegionRepeaterRenderer<TItem> : Named, IElementRenderer
    {
        protected IEnumerable<TItem> DataSource { set; get; }

        protected IList<IEmbeddedRenderer<TItem>> RendererList { set; get; }

        public RegionRepeaterRenderer(string name, IEnumerable<TItem> dataSource, params IEmbeddedRenderer<TItem>[] renderers)
        {
            Name = name;
            DataSource = dataSource;
            RendererList = new List<IEmbeddedRenderer<TItem>>(renderers);
        }

        public int SortNum(SheetContext sheetContext)
        {
            Repeater repeater = sheetContext.WorksheetContainer.Repeaters[Name];
            if (repeater.Start.IsNull())
            {
                throw new TemplateException($"RepeaterRenderer[{repeater.Name}] start non-existent.");
            }
            if (repeater.End.IsNull())
            {
                throw new TemplateException($"RepeaterRenderer[{repeater.Name}] end non-existent.");
            }
            return repeater.Start.RowIndex;
        }

        public virtual void Render(SheetContext sheetContext)
        {
            Repeater repeater = sheetContext.WorksheetContainer.Repeaters[Name];
            if (RendererList.IsNullOrEmpty())
            {
                throw new RenderException($"RepeaterRenderer[{repeater.Name}] is empty");
            }
            sheetContext.ResetAccumulation();
            if (!DataSource.IsNull())
            {
                foreach (var item in DataSource)
                {
                    sheetContext.CopyRepeaterTemplateByReplace(repeater, () =>
                    {
                        foreach (var renderer in RendererList.OrderBy(renderer => renderer.SortNum(sheetContext)))
                        {
                            renderer.Render(sheetContext, item);
                        }
                    });
                }
            }

            sheetContext.RemoveRegionRepeaterTemplate(repeater);
        }

        public void Append(IEmbeddedRenderer<TItem> renderer)
        {
            RendererList.Add(renderer);
        }
    }

    public class RegionRepeaterRenderer<TSource, TItem> : Named, IEmbeddedRenderer<TSource>
    {
        protected Func<TSource, IEnumerable<TItem>> DgSetDataSource { set; get; }

        protected IList<IEmbeddedRenderer<TItem>> RendererList { set; get; }

        public RegionRepeaterRenderer(string name, Func<TSource, IEnumerable<TItem>> dgSetDataSource, params IEmbeddedRenderer<TItem>[] renderers)
        {
            Name = name;
            DgSetDataSource = dgSetDataSource;
            RendererList = new List<IEmbeddedRenderer<TItem>>(renderers);
        }

        public int SortNum(SheetContext sheetContext)
        {
            Repeater repeater = sheetContext.WorksheetContainer.Repeaters[Name];
            if (repeater.Start.IsNull())
            {
                throw new TemplateException($"RepeaterRenderer[{repeater.Name}] start non-existent.");
            }
            if (repeater.End.IsNull())
            {
                throw new TemplateException($"RepeaterRenderer[{repeater.Name}] end non-existent.");
            }
            return repeater.Start.RowIndex;
        }

        public virtual void Render(SheetContext sheetContext, TSource dataSource)
        {
            Repeater repeater = sheetContext.WorksheetContainer.Repeaters[Name];
            if (RendererList.IsNullOrEmpty())
            {
                throw new RenderException($"RepeaterRenderer[{repeater.Name}] is empty");
            }
            sheetContext.ResetAccumulation();
            var items = DgSetDataSource(dataSource);
            if (!items.IsNull())
            {
                foreach (var item in items)
                {
                    sheetContext.CopyRepeaterTemplateByReplace(repeater, () =>
                    {
                        foreach (var renderer in RendererList.OrderBy(renderer => renderer.SortNum(sheetContext)))
                        {
                            renderer.Render(sheetContext, item);
                        }
                    });
                }
            }

            sheetContext.RemoveRegionRepeaterTemplate(repeater);
        }

        public void Append(IEmbeddedRenderer<TItem> renderer)
        {
            RendererList.Add(renderer);
        }
    }
}

