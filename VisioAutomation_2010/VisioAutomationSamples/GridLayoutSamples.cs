﻿using VA = VisioAutomation;
using IVisio = Microsoft.Office.Interop.Visio;
using VisioAutomation.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace VisioAutomationSamples
{
    public static class GridLayoutSamples
    {
        public static void ColorGrid()
        {
            // Draws a grid rectangles and then formats the shapes
            // with different colors

            // Demonstrates:
            // How use the GridLayout object to quickly drop a grid
            // How to use ShapeFormatCells to apply formatting to shapes
            // How UpdateBase can be used to modfiy multiple shapes at once

            int[] colors = {
                    0x0A3B76, 0x4395D1, 0x99D9EA, 0x0D686B, 0x00A99D, 0x7ACCC8, 0x82CA9C,
                    0x74A402,
                    0xC4DF9B, 0xD9D56F, 0xFFF468, 0xFFF799, 0xFFC20E, 0xEB6119, 0xFBAF5D,
                    0xE57300, 0xC14000, 0xB82832, 0xD85171, 0xFEDFEC, 0x563F7F, 0xA186BE,
                    0xD9CFE5
                };

            const int num_cols = 5;
            const int num_rows = 5;

            var page = SampleEnvironment.Application.ActiveDocument.Pages.Add();

            var page_size = new VA.Drawing.Size(10, 10);
            SampleEnvironment.SetPageSize(page,page_size);

            var stencil = SampleEnvironment.Application.Documents.OpenStencil("basic_u.vss");
            var master = stencil.Masters["Rectangle"];

            var layout = new VA.Models.Grid.GridLayout(num_cols, num_rows, new VA.Drawing.Size(1, 1), master);
            layout.Origin = new VA.Drawing.Point(0, 0);
            layout.CellSpacing = new VA.Drawing.Size(0, 0);
            layout.RowDirection = VA.Models.Grid.RowDirection.BottomToTop;

            layout.PerformLayout();
            layout.Render(page);

            var fmtcells = new VA.Format.ShapeFormatCells();
            int i = 0;
            var update = new VA.ShapeSheet.Update();
            foreach (var node in layout.Nodes)
            {
                var shapeid = node.ShapeID;
                int color_index = i%colors.Length;
                var color = colors[color_index];
                fmtcells.FillForegnd = new VA.Drawing.ColorRGB(color).ToFormula();
                fmtcells.LinePattern = 0;
                fmtcells.LineWeight = 0;
                update.SetFormulas(shapeid, fmtcells);
                i++;
            }

            update.Execute(page);

            var bordersize = new VA.Drawing.Size(1,1);
            page.ResizeToFitContents(bordersize);
        }
    }
}