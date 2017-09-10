using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;


namespace VisioAutomation.Shapes
{
    public class ControlCells : CellGroupMultiRow
    {
        public CellValueLiteral CanGlue { get; set; }
        public CellValueLiteral Tip { get; set; }
        public CellValueLiteral X { get; set; }
        public CellValueLiteral Y { get; set; }
        public CellValueLiteral YBehavior { get; set; }
        public CellValueLiteral XBehavior { get; set; }
        public CellValueLiteral XDynamics { get; set; }
        public CellValueLiteral YDynamics { get; set; }

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.ControlCanGlue, this.CanGlue);
                yield return SrcValuePair.Create(SrcConstants.ControlTip, this.Tip);
                yield return SrcValuePair.Create(SrcConstants.ControlX, this.X);
                yield return SrcValuePair.Create(SrcConstants.ControlY, this.Y);
                yield return SrcValuePair.Create(SrcConstants.ControlYBehavior, this.YBehavior);
                yield return SrcValuePair.Create(SrcConstants.ControlXBehavior, this.XBehavior);
                yield return SrcValuePair.Create(SrcConstants.ControlXDynamics, this.XDynamics);
                yield return SrcValuePair.Create(SrcConstants.ControlYDynamics, this.YDynamics);
            }
        }

        public static List<List<ControlCells>> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }

        public static List<ControlCells> GetCells(IVisio.Shape shape, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, cvt);
        }
        
        private static readonly System.Lazy<ControlCellsReader> lazy_query = new System.Lazy<ControlCellsReader>();

        class ControlCellsReader : ReaderMultiRow<ControlCells>
        {
            public SectionQueryColumn CanGlue { get; set; }
            public SectionQueryColumn Tip { get; set; }
            public SectionQueryColumn X { get; set; }
            public SectionQueryColumn Y { get; set; }
            public SectionQueryColumn YBehavior { get; set; }
            public SectionQueryColumn XBehavior { get; set; }
            public SectionQueryColumn XDynamics { get; set; }
            public SectionQueryColumn YDynamics { get; set; }

            public ControlCellsReader()
            {
                var sec = this.query.SectionQueries.Add(IVisio.VisSectionIndices.visSectionControls);

                this.CanGlue = sec.Columns.Add(SrcConstants.ControlCanGlue, nameof(SrcConstants.ControlCanGlue));
                this.Tip = sec.Columns.Add(SrcConstants.ControlTip, nameof(SrcConstants.ControlTip));
                this.X = sec.Columns.Add(SrcConstants.ControlX, nameof(SrcConstants.ControlX));
                this.Y = sec.Columns.Add(SrcConstants.ControlY, nameof(SrcConstants.ControlY));
                this.YBehavior = sec.Columns.Add(SrcConstants.ControlYBehavior, nameof(SrcConstants.ControlYBehavior));
                this.XBehavior = sec.Columns.Add(SrcConstants.ControlXBehavior, nameof(SrcConstants.ControlXBehavior));
                this.XDynamics = sec.Columns.Add(SrcConstants.ControlXDynamics, nameof(SrcConstants.ControlXDynamics));
                this.YDynamics = sec.Columns.Add(SrcConstants.ControlYDynamics, nameof(SrcConstants.ControlYDynamics));

            }

            public override ControlCells CellDataToCellGroup(Utilities.ArraySegment<string> row)
            {
                var cells = new ControlCells();
                cells.CanGlue = row[this.CanGlue];
                cells.Tip = row[this.Tip];
                cells.X = row[this.X];
                cells.Y = row[this.Y];
                cells.YBehavior = row[this.YBehavior];
                cells.XBehavior = row[this.XBehavior];
                cells.XDynamics = row[this.XDynamics];
                cells.YDynamics = row[this.YDynamics];
                return cells;
            }
        }

    }
}