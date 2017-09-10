using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using IVisio= Microsoft.Office.Interop.Visio;
using VisioAutomation.ShapeSheet;
using VisioAutomation.ShapeSheet.Query;

namespace VisioAutomation.Shapes
{
    public class ShapeLayoutCells : CellGroupSingleRow
    {
        public CellValueLiteral ConnectorFixedCode { get; set; }
        public CellValueLiteral LineJumpCode { get; set; }
        public CellValueLiteral LineJumpDirX { get; set; }
        public CellValueLiteral LineJumpDirY { get; set; }
        public CellValueLiteral LineJumpStyle { get; set; }
        public CellValueLiteral LineRouteExt { get; set; }
        public CellValueLiteral ShapeFixedCode { get; set; }
        public CellValueLiteral ShapePermeablePlace { get; set; }
        public CellValueLiteral ShapePermeableX { get; set; }
        public CellValueLiteral ShapePermeableY { get; set; }
        public CellValueLiteral ShapePlaceFlip { get; set; }
        public CellValueLiteral ShapePlaceStyle { get; set; }
        public CellValueLiteral ShapePlowCode { get; set; }
        public CellValueLiteral ShapeRouteStyle { get; set; }
        public CellValueLiteral ShapeSplit { get; set; }
        public CellValueLiteral ShapeSplittable { get; set; }
        public CellValueLiteral ShapeDisplayLevel { get; set; } // new in visio 2010
        public CellValueLiteral Relationships { get; set; } // new in visio 2010

        public override IEnumerable<SrcValuePair> SrcValuePairs
        {
            get
            {
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutConnectorFixedCode, this.ConnectorFixedCode);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutLineJumpCode, this.LineJumpCode);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutLineJumpDirX, this.LineJumpDirX);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutLineJumpDirY, this.LineJumpDirY);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutLineJumpStyle, this.LineJumpStyle);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutLineRouteExt, this.LineRouteExt);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapeFixedCode, this.ShapeFixedCode);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePermeablePlace, this.ShapePermeablePlace);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePermeableX, this.ShapePermeableX);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePermeableY, this.ShapePermeableY);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePlaceFlip, this.ShapePlaceFlip);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePlaceStyle, this.ShapePlaceStyle);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapePlowCode, this.ShapePlowCode);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapeRouteStyle, this.ShapeRouteStyle);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapeSplit, this.ShapeSplit);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapeSplittable, this.ShapeSplittable);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutShapeDisplayLevel, this.ShapeDisplayLevel);
                yield return SrcValuePair.Create(SrcConstants.ShapeLayoutRelationships, this.Relationships);
            }
        }
        
        public static List<ShapeLayoutCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(page, shapeids, cvt);
        }

        public static ShapeLayoutCells GetCells(IVisio.Shape shape, CellValueType cvt)
        {
            var query = lazy_query.Value;
            return query.GetValues(shape, cvt);
        }

        private static readonly System.Lazy<ShapeLayoutCellsReader> lazy_query = new System.Lazy<ShapeLayoutCellsReader>();

        class ShapeLayoutCellsReader : ReaderSingleRow<ShapeLayoutCells>
        {
            public CellColumn ConnectorFixedCode { get; set; }
            public CellColumn LineJumpCode { get; set; }
            public CellColumn LineJumpDirX { get; set; }
            public CellColumn LineJumpDirY { get; set; }
            public CellColumn LineJumpStyle { get; set; }
            public CellColumn LineRouteExt { get; set; }
            public CellColumn ShapeFixedCode { get; set; }
            public CellColumn ShapePermeablePlace { get; set; }
            public CellColumn ShapePermeableX { get; set; }
            public CellColumn ShapePermeableY { get; set; }
            public CellColumn ShapePlaceFlip { get; set; }
            public CellColumn ShapePlaceStyle { get; set; }
            public CellColumn ShapePlowCode { get; set; }
            public CellColumn ShapeRouteStyle { get; set; }
            public CellColumn ShapeSplit { get; set; }
            public CellColumn ShapeSplittable { get; set; }
            public CellColumn ShapeDisplayLevel { get; set; }
            public CellColumn Relationships { get; set; }

            public ShapeLayoutCellsReader()
            {
                this.ConnectorFixedCode = this.query.Columns.Add(SrcConstants.ShapeLayoutConnectorFixedCode, nameof(SrcConstants.ShapeLayoutConnectorFixedCode));
                this.LineJumpCode = this.query.Columns.Add(SrcConstants.ShapeLayoutLineJumpCode, nameof(SrcConstants.ShapeLayoutLineJumpCode));
                this.LineJumpDirX = this.query.Columns.Add(SrcConstants.ShapeLayoutLineJumpDirX, nameof(SrcConstants.ShapeLayoutLineJumpDirX));
                this.LineJumpDirY = this.query.Columns.Add(SrcConstants.ShapeLayoutLineJumpDirY, nameof(SrcConstants.ShapeLayoutLineJumpDirY));
                this.LineJumpStyle = this.query.Columns.Add(SrcConstants.ShapeLayoutLineJumpStyle, nameof(SrcConstants.ShapeLayoutLineJumpStyle));
                this.LineRouteExt = this.query.Columns.Add(SrcConstants.ShapeLayoutLineRouteExt, nameof(SrcConstants.ShapeLayoutLineRouteExt));
                this.ShapeFixedCode = this.query.Columns.Add(SrcConstants.ShapeLayoutShapeFixedCode, nameof(SrcConstants.ShapeLayoutShapeFixedCode));
                this.ShapePermeablePlace = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePermeablePlace, nameof(SrcConstants.ShapeLayoutShapePermeablePlace));
                this.ShapePermeableX = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePermeableX, nameof(SrcConstants.ShapeLayoutShapePermeableX));
                this.ShapePermeableY = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePermeableY, nameof(SrcConstants.ShapeLayoutShapePermeableY));
                this.ShapePlaceFlip = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePlaceFlip, nameof(SrcConstants.ShapeLayoutShapePlaceFlip));
                this.ShapePlaceStyle = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePlaceStyle, nameof(SrcConstants.ShapeLayoutShapePlaceStyle));
                this.ShapePlowCode = this.query.Columns.Add(SrcConstants.ShapeLayoutShapePlowCode, nameof(SrcConstants.ShapeLayoutShapePlowCode));
                this.ShapeRouteStyle = this.query.Columns.Add(SrcConstants.ShapeLayoutShapeRouteStyle, nameof(SrcConstants.ShapeLayoutShapeRouteStyle));
                this.ShapeSplit = this.query.Columns.Add(SrcConstants.ShapeLayoutShapeSplit, nameof(SrcConstants.ShapeLayoutShapeSplit));
                this.ShapeSplittable = this.query.Columns.Add(SrcConstants.ShapeLayoutShapeSplittable, nameof(SrcConstants.ShapeLayoutShapeSplittable));
                this.ShapeDisplayLevel = this.query.Columns.Add(SrcConstants.ShapeLayoutShapeDisplayLevel, nameof(SrcConstants.ShapeLayoutShapeDisplayLevel));
                this.Relationships = this.query.Columns.Add(SrcConstants.ShapeLayoutRelationships, nameof(SrcConstants.ShapeLayoutRelationships));
            }

            public override ShapeLayoutCells CellDataToCellGroup(Utilities.ArraySegment<string> row)
            {
                var cells = new ShapeLayoutCells();
                cells.ConnectorFixedCode = row[this.ConnectorFixedCode];
                cells.LineJumpCode = row[this.LineJumpCode];
                cells.LineJumpDirX = row[this.LineJumpDirX];
                cells.LineJumpDirY = row[this.LineJumpDirY];
                cells.LineJumpStyle = row[this.LineJumpStyle];
                cells.LineRouteExt = row[this.LineRouteExt];
                cells.ShapeFixedCode = row[this.ShapeFixedCode];
                cells.ShapePermeablePlace = row[this.ShapePermeablePlace];
                cells.ShapePermeableX = row[this.ShapePermeableX];
                cells.ShapePermeableY = row[this.ShapePermeableY];
                cells.ShapePlaceFlip = row[this.ShapePlaceFlip];
                cells.ShapePlaceStyle = row[this.ShapePlaceStyle];
                cells.ShapePlowCode = row[this.ShapePlowCode];
                cells.ShapeRouteStyle = row[this.ShapeRouteStyle];
                cells.ShapeSplit = row[this.ShapeSplit];
                cells.ShapeSplittable = row[this.ShapeSplittable];
                cells.ShapeDisplayLevel = row[this.ShapeDisplayLevel];
                cells.Relationships = row[this.Relationships];
                return cells;
            }
        }
    }
}