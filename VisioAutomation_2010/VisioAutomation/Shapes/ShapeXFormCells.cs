﻿using System.Collections.Generic;
using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet;

using IVisio = Microsoft.Office.Interop.Visio;


namespace VisioAutomation.Shapes
{
    public class ShapeXFormCells : CellGroup
    {
        public CellValueLiteral PinX { get; set; }
        public CellValueLiteral PinY { get; set; }
        public CellValueLiteral LocPinX { get; set; }
        public CellValueLiteral LocPinY { get; set; }
        public CellValueLiteral Width { get; set; }
        public CellValueLiteral Height { get; set; }
        public CellValueLiteral Angle { get; set; }

        public override IEnumerable<CellMetadataItem> CellMetadata
        {
            get
            {


                yield return CellMetadataItem.Create(nameof(this.PinX), SrcConstants.XFormPinX, this.PinX);
                yield return CellMetadataItem.Create(nameof(this.PinY), SrcConstants.XFormPinY, this.PinY);
                yield return CellMetadataItem.Create(nameof(this.LocPinX), SrcConstants.XFormLocPinX, this.LocPinX);
                yield return CellMetadataItem.Create(nameof(this.LocPinY), SrcConstants.XFormLocPinY, this.LocPinY);
                yield return CellMetadataItem.Create(nameof(this.Width), SrcConstants.XFormWidth, this.Width);
                yield return CellMetadataItem.Create(nameof(this.Height), SrcConstants.XFormHeight, this.Height);
                yield return CellMetadataItem.Create(nameof(this.Angle), SrcConstants.XFormAngle, this.Angle);
            }
        }


        public static List<ShapeXFormCells> GetCells(IVisio.Page page, IList<int> shapeids, CellValueType type)
        {
            var reader = ShapeXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(page, shapeids, type);
        }

        public static ShapeXFormCells GetCells(IVisio.Shape shape, CellValueType type)
        {
            var reader = ShapeXFormCells_lazy_builder.Value;
            return reader.GetCellsSingleRow(shape, type);
        }

        private static readonly System.Lazy<ShapeXFormCellsBuilder> ShapeXFormCells_lazy_builder = new System.Lazy<ShapeXFormCellsBuilder>();

        class ShapeXFormCellsBuilder : CellGroupBuilder<ShapeXFormCells>
        {
            public ShapeXFormCellsBuilder() : base(VisioAutomation.ShapeSheet.CellGroups.CellGroupBuilderType.SingleRow)
            {
            }

            public override ShapeXFormCells ToCellGroup(ShapeSheet.Internal.ArraySegment<string> row, VisioAutomation.ShapeSheet.Query.ColumnList cols)
            {
                var cells = new ShapeXFormCells();

                string getcellvalue(string name)
                {
                    return row[cols[name].Ordinal];
                }

                cells.PinX = getcellvalue(nameof(ShapeXFormCells.PinX));
                cells.PinY = getcellvalue(nameof(ShapeXFormCells.PinY));
                cells.LocPinX = getcellvalue(nameof(ShapeXFormCells.LocPinX));
                cells.LocPinY = getcellvalue(nameof(ShapeXFormCells.LocPinY));
                cells.Width = getcellvalue(nameof(ShapeXFormCells.Width));
                cells.Height = getcellvalue(nameof(ShapeXFormCells.Height));
                cells.Angle = getcellvalue(nameof(ShapeXFormCells.Angle));

                return cells;
            }
        }

    }
}