using VisioAutomation.ShapeSheet.CellGroups;
using VisioAutomation.ShapeSheet.Query;
using SRCCON = VisioAutomation.ShapeSheet.SRCConstants;

namespace VisioAutomation.Pages
{
    class PageCellsReader : SingleRowReader<VisioAutomation.Pages.PageCells>
    {
        public CellColumn PageLeftMargin { get; set; }
        public CellColumn CenterX { get; set; }
        public CellColumn CenterY { get; set; }
        public CellColumn OnPage { get; set; }
        public CellColumn PageBottomMargin { get; set; }
        public CellColumn PageRightMargin { get; set; }
        public CellColumn PagesX { get; set; }
        public CellColumn PagesY { get; set; }
        public CellColumn PageTopMargin { get; set; }
        public CellColumn PaperKind { get; set; }
        public CellColumn PrintGrid { get; set; }
        public CellColumn PrintPageOrientation { get; set; }
        public CellColumn ScaleX { get; set; }
        public CellColumn ScaleY { get; set; }
        public CellColumn PaperSource { get; set; }
        public CellColumn DrawingScale { get; set; }
        public CellColumn DrawingScaleType { get; set; }
        public CellColumn DrawingSizeType { get; set; }
        public CellColumn InhibitSnap { get; set; }
        public CellColumn PageHeight { get; set; }
        public CellColumn PageScale { get; set; }
        public CellColumn PageWidth { get; set; }
        public CellColumn ShdwObliqueAngle { get; set; }
        public CellColumn ShdwOffsetX { get; set; }
        public CellColumn ShdwOffsetY { get; set; }
        public CellColumn ShdwScaleFactor { get; set; }
        public CellColumn ShdwType { get; set; }
        public CellColumn UIVisibility { get; set; }
        public CellColumn XGridDensity { get; set; }
        public CellColumn XGridOrigin { get; set; }
        public CellColumn XGridSpacing { get; set; }
        public CellColumn XRulerDensity { get; set; }
        public CellColumn XRulerOrigin { get; set; }
        public CellColumn YGridDensity { get; set; }
        public CellColumn YGridOrigin { get; set; }
        public CellColumn YGridSpacing { get; set; }
        public CellColumn YRulerDensity { get; set; }
        public CellColumn YRulerOrigin { get; set; }
        public CellColumn AvenueSizeX { get; set; }
        public CellColumn AvenueSizeY { get; set; }
        public CellColumn BlockSizeX { get; set; }
        public CellColumn BlockSizeY { get; set; }
        public CellColumn CtrlAsInput { get; set; }
        public CellColumn DynamicsOff { get; set; }
        public CellColumn EnableGrid { get; set; }
        public CellColumn LineAdjustFrom { get; set; }
        public CellColumn LineAdjustTo { get; set; }
        public CellColumn LineJumpCode { get; set; }
        public CellColumn LineJumpFactorX { get; set; }
        public CellColumn LineJumpFactorY { get; set; }
        public CellColumn LineJumpStyle { get; set; }
        public CellColumn LineRouteExt { get; set; }
        public CellColumn LineToLineX { get; set; }
        public CellColumn LineToLineY { get; set; }
        public CellColumn LineToNodeX { get; set; }
        public CellColumn LineToNodeY { get; set; }
        public CellColumn PageLineJumpDirX { get; set; }
        public CellColumn PageLineJumpDirY { get; set; }
        public CellColumn PageShapeSplit { get; set; }
        public CellColumn PlaceDepth { get; set; }
        public CellColumn PlaceFlip { get; set; }
        public CellColumn PlaceStyle { get; set; }
        public CellColumn PlowCode { get; set; }
        public CellColumn ResizePage { get; set; }
        public CellColumn RouteStyle { get; set; }
        public CellColumn AvoidPageBreaks { get; set; }
        public CellColumn DrawingResizeType { get; set; }

        public PageCellsReader()
        {
            this.PageLeftMargin = this.query.AddCell(SRCCON.PageLeftMargin, nameof(SRCCON.PageLeftMargin));
            this.CenterX = this.query.AddCell(SRCCON.CenterX, nameof(SRCCON.CenterX));
            this.CenterY = this.query.AddCell(SRCCON.CenterY, nameof(SRCCON.CenterY));
            this.OnPage = this.query.AddCell(SRCCON.OnPage, nameof(SRCCON.OnPage));
            this.PageBottomMargin = this.query.AddCell(SRCCON.PageBottomMargin, nameof(SRCCON.PageBottomMargin));
            this.PageRightMargin = this.query.AddCell(SRCCON.PageRightMargin, nameof(SRCCON.PageRightMargin));
            this.PagesX = this.query.AddCell(SRCCON.PagesX, nameof(SRCCON.PagesX));
            this.PagesY = this.query.AddCell(SRCCON.PagesY, nameof(SRCCON.PagesY));
            this.PageTopMargin = this.query.AddCell(SRCCON.PageTopMargin, nameof(SRCCON.PageTopMargin));
            this.PaperKind = this.query.AddCell(SRCCON.PaperKind, nameof(SRCCON.PaperKind));
            this.PrintGrid = this.query.AddCell(SRCCON.PrintGrid, nameof(SRCCON.PrintGrid));
            this.PrintPageOrientation = this.query.AddCell(SRCCON.PrintPageOrientation, nameof(SRCCON.PrintPageOrientation));
            this.ScaleX = this.query.AddCell(SRCCON.ScaleX, nameof(SRCCON.ScaleX));
            this.ScaleY = this.query.AddCell(SRCCON.ScaleY, nameof(SRCCON.ScaleY));
            this.PaperSource = this.query.AddCell(SRCCON.PaperSource, nameof(SRCCON.PaperSource));
            this.DrawingScale = this.query.AddCell(SRCCON.DrawingScale, nameof(SRCCON.DrawingScale));
            this.DrawingScaleType = this.query.AddCell(SRCCON.DrawingScaleType, nameof(SRCCON.DrawingScaleType));
            this.DrawingSizeType = this.query.AddCell(SRCCON.DrawingSizeType, nameof(SRCCON.DrawingSizeType));
            this.InhibitSnap = this.query.AddCell(SRCCON.InhibitSnap, nameof(SRCCON.InhibitSnap));
            this.PageHeight = this.query.AddCell(SRCCON.PageHeight, nameof(SRCCON.PageHeight));
            this.PageScale = this.query.AddCell(SRCCON.PageScale, nameof(SRCCON.PageScale));
            this.PageWidth = this.query.AddCell(SRCCON.PageWidth, nameof(SRCCON.PageWidth));
            this.ShdwObliqueAngle = this.query.AddCell(SRCCON.ShdwObliqueAngle, nameof(SRCCON.ShdwObliqueAngle));
            this.ShdwOffsetX = this.query.AddCell(SRCCON.ShdwOffsetX, nameof(SRCCON.ShdwOffsetX));
            this.ShdwOffsetY = this.query.AddCell(SRCCON.ShdwOffsetY, nameof(SRCCON.ShdwOffsetY));
            this.ShdwScaleFactor = this.query.AddCell(SRCCON.ShdwScaleFactor, nameof(SRCCON.ShdwScaleFactor));
            this.ShdwType = this.query.AddCell(SRCCON.ShdwType, nameof(SRCCON.ShdwType));
            this.UIVisibility = this.query.AddCell(SRCCON.UIVisibility, nameof(SRCCON.UIVisibility));
            this.XGridDensity = this.query.AddCell(SRCCON.XGridDensity, nameof(SRCCON.XGridDensity));
            this.XGridOrigin = this.query.AddCell(SRCCON.XGridOrigin, nameof(SRCCON.XGridOrigin));
            this.XGridSpacing = this.query.AddCell(SRCCON.XGridSpacing, nameof(SRCCON.XGridSpacing));
            this.XRulerDensity = this.query.AddCell(SRCCON.XRulerDensity, nameof(SRCCON.XRulerDensity));
            this.XRulerOrigin = this.query.AddCell(SRCCON.XRulerOrigin, nameof(SRCCON.XRulerOrigin));
            this.YGridDensity = this.query.AddCell(SRCCON.YGridDensity, nameof(SRCCON.YGridDensity));
            this.YGridOrigin = this.query.AddCell(SRCCON.YGridOrigin, nameof(SRCCON.YGridOrigin));
            this.YGridSpacing = this.query.AddCell(SRCCON.YGridSpacing, nameof(SRCCON.YGridSpacing));
            this.YRulerDensity = this.query.AddCell(SRCCON.YRulerDensity, nameof(SRCCON.YRulerDensity));
            this.YRulerOrigin = this.query.AddCell(SRCCON.YRulerOrigin, nameof(SRCCON.YRulerOrigin));
            this.AvenueSizeX = this.query.AddCell(SRCCON.AvenueSizeX, nameof(SRCCON.AvenueSizeX));
            this.AvenueSizeY = this.query.AddCell(SRCCON.AvenueSizeY, nameof(SRCCON.AvenueSizeY));
            this.BlockSizeX = this.query.AddCell(SRCCON.BlockSizeX, nameof(SRCCON.BlockSizeX));
            this.BlockSizeY = this.query.AddCell(SRCCON.BlockSizeY, nameof(SRCCON.BlockSizeY));
            this.CtrlAsInput = this.query.AddCell(SRCCON.CtrlAsInput, nameof(SRCCON.CtrlAsInput));
            this.DynamicsOff = this.query.AddCell(SRCCON.DynamicsOff, nameof(SRCCON.DynamicsOff));
            this.EnableGrid = this.query.AddCell(SRCCON.EnableGrid, nameof(SRCCON.EnableGrid));
            this.LineAdjustFrom = this.query.AddCell(SRCCON.LineAdjustFrom, nameof(SRCCON.LineAdjustFrom));
            this.LineAdjustTo = this.query.AddCell(SRCCON.LineAdjustTo, nameof(SRCCON.LineAdjustTo));
            this.LineJumpCode = this.query.AddCell(SRCCON.LineJumpCode, nameof(SRCCON.LineJumpCode));
            this.LineJumpFactorX = this.query.AddCell(SRCCON.LineJumpFactorX, nameof(SRCCON.LineJumpFactorX));
            this.LineJumpFactorY = this.query.AddCell(SRCCON.LineJumpFactorY, nameof(SRCCON.LineJumpFactorY));
            this.LineJumpStyle = this.query.AddCell(SRCCON.LineJumpStyle, nameof(SRCCON.LineJumpStyle));
            this.LineRouteExt = this.query.AddCell(SRCCON.LineRouteExt, nameof(SRCCON.LineRouteExt));
            this.LineToLineX = this.query.AddCell(SRCCON.LineToLineX, nameof(SRCCON.LineToLineX));
            this.LineToLineY = this.query.AddCell(SRCCON.LineToLineY, nameof(SRCCON.LineToLineY));
            this.LineToNodeX = this.query.AddCell(SRCCON.LineToNodeX, nameof(SRCCON.LineToNodeX));
            this.LineToNodeY = this.query.AddCell(SRCCON.LineToNodeY, nameof(SRCCON.LineToNodeY));
            this.PageLineJumpDirX = this.query.AddCell(SRCCON.PageLineJumpDirX, nameof(SRCCON.PageLineJumpDirX));
            this.PageLineJumpDirY = this.query.AddCell(SRCCON.PageLineJumpDirY, nameof(SRCCON.PageLineJumpDirY));
            this.PageShapeSplit = this.query.AddCell(SRCCON.PageShapeSplit, nameof(SRCCON.PageShapeSplit));
            this.PlaceDepth = this.query.AddCell(SRCCON.PlaceDepth, nameof(SRCCON.PlaceDepth));
            this.PlaceFlip = this.query.AddCell(SRCCON.PlaceFlip, nameof(SRCCON.PlaceFlip));
            this.PlaceStyle = this.query.AddCell(SRCCON.PlaceStyle, nameof(SRCCON.PlaceStyle));
            this.PlowCode = this.query.AddCell(SRCCON.PlowCode, nameof(SRCCON.PlowCode));
            this.ResizePage = this.query.AddCell(SRCCON.ResizePage, nameof(SRCCON.ResizePage));
            this.RouteStyle = this.query.AddCell(SRCCON.RouteStyle, nameof(SRCCON.RouteStyle));
            this.AvoidPageBreaks = this.query.AddCell(SRCCON.AvoidPageBreaks, nameof(SRCCON.AvoidPageBreaks));
            this.DrawingResizeType = this.query.AddCell(SRCCON.DrawingResizeType, nameof(SRCCON.DrawingResizeType));
        }


        public override Pages.PageCells CellDataToCellGroup(VisioAutomation.Utilities.ArraySegment<ShapeSheet.CellData> row)
        {
            var cells = new Pages.PageCells();
            cells.PageLeftMargin = row[this.PageLeftMargin];
            cells.CenterX = row[this.CenterX];
            cells.CenterY = row[this.CenterY];
            cells.OnPage = row[this.OnPage];
            cells.PageBottomMargin = row[this.PageBottomMargin];
            cells.PageRightMargin = row[this.PageRightMargin];
            cells.PagesX = row[this.PagesX];
            cells.PagesY = row[this.PagesY];
            cells.PageTopMargin = row[this.PageTopMargin];
            cells.PaperKind = row[this.PaperKind];
            cells.PrintGrid = row[this.PrintGrid];
            cells.PrintPageOrientation = row[this.PrintPageOrientation];
            cells.ScaleX = row[this.ScaleX];
            cells.ScaleY = row[this.ScaleY];
            cells.PaperSource = row[this.PaperSource];
            cells.DrawingScale = row[this.DrawingScale];
            cells.DrawingScaleType = row[this.DrawingScaleType];
            cells.DrawingSizeType = row[this.DrawingSizeType];
            cells.InhibitSnap = row[this.InhibitSnap];
            cells.PageHeight = row[this.PageHeight];
            cells.PageScale = row[this.PageScale];
            cells.PageWidth = row[this.PageWidth];
            cells.ShdwObliqueAngle = row[this.ShdwObliqueAngle];
            cells.ShdwOffsetX = row[this.ShdwOffsetX];
            cells.ShdwOffsetY = row[this.ShdwOffsetY];
            cells.ShdwScaleFactor = row[this.ShdwScaleFactor];
            cells.ShdwType = row[this.ShdwType];
            cells.UIVisibility = row[this.UIVisibility];
            cells.XGridDensity = row[this.XGridDensity];
            cells.XGridOrigin = row[this.XGridOrigin];
            cells.XGridSpacing = row[this.XGridSpacing];
            cells.XRulerDensity = row[this.XRulerDensity];
            cells.XRulerOrigin = row[this.XRulerOrigin];
            cells.YGridDensity = row[this.YGridDensity];
            cells.YGridOrigin = row[this.YGridOrigin];
            cells.YGridSpacing = row[this.YGridSpacing];
            cells.YRulerDensity = row[this.YRulerDensity];
            cells.YRulerOrigin = row[this.YRulerOrigin];
            cells.AvenueSizeX = row[this.AvenueSizeX];
            cells.AvenueSizeY = row[this.AvenueSizeY];
            cells.BlockSizeX = row[this.BlockSizeX];
            cells.BlockSizeY = row[this.BlockSizeY];
            cells.CtrlAsInput = row[this.CtrlAsInput];
            cells.DynamicsOff = row[this.DynamicsOff];
            cells.EnableGrid = row[this.EnableGrid];
            cells.LineAdjustFrom = row[this.LineAdjustFrom];
            cells.LineAdjustTo = row[this.LineAdjustTo];
            cells.LineJumpCode = row[this.LineJumpCode];
            cells.LineJumpFactorX = row[this.LineJumpFactorX];
            cells.LineJumpFactorY = row[this.LineJumpFactorY];
            cells.LineJumpStyle = row[this.LineJumpStyle];
            cells.LineRouteExt = row[this.LineRouteExt];
            cells.LineToLineX = row[this.LineToLineX];
            cells.LineToLineY = row[this.LineToLineY];
            cells.LineToNodeX = row[this.LineToNodeX];
            cells.LineToNodeY = row[this.LineToNodeY];
            cells.PageLineJumpDirX = row[this.PageLineJumpDirX];
            cells.PageLineJumpDirY = row[this.PageLineJumpDirY];
            cells.PageShapeSplit = row[this.PageShapeSplit];
            cells.PlaceDepth = row[this.PlaceDepth];
            cells.PlaceFlip = row[this.PlaceFlip];
            cells.PlaceStyle = row[this.PlaceStyle];
            cells.PlowCode = row[this.PlowCode];
            cells.ResizePage = row[this.ResizePage];
            cells.RouteStyle = row[this.RouteStyle];
            cells.AvoidPageBreaks = row[this.AvoidPageBreaks];
            cells.DrawingResizeType = row[this.DrawingResizeType];
            return cells;
        }
    }
}