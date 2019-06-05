using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioAutomation.Models.Layouts.DirectedGraph
{
    class VisioRenderer
    {


        public VisioRenderer()
        {
        }

        public void Render(IVisio.Page page, DirectedGraphLayout directed_graph_layout, VisioLayoutOptions options)
        {
            // This is Visio-based render - it does NOT use MSAGL
            if (page == null)
            {
                throw new System.ArgumentNullException(nameof(page));
            }

            if (options== null)
            {
                throw new System.ArgumentNullException(nameof(options));
            }

            var page_node = new Dom.Page();
            double x = 0;
            double y = 1;
            foreach (var shape in directed_graph_layout.Shapes)
            {
                var shape_nodes = page_node.Shapes.Drop(shape.MasterName, shape.StencilName, x, y);
                shape.DomNode = shape_nodes;
                shape.DomNode.Text = new VisioAutomation.Models.Text.Element(shape.Label);
                x += 1.0;
            }

            foreach (var connector in directed_graph_layout.Connectors)
            {
                var connector_node = page_node.Shapes.Connect(options.EdgeMasterName, options.EdgeStencilName, connector.From.DomNode, connector.To.DomNode);
                connector.DomNode = connector_node;
                connector.DomNode.Text = new VisioAutomation.Models.Text.Element(connector.Label);
            }

            page_node.ResizeToFit = true;
            page_node.ResizeToFitMargin = new VisioAutomation.Geometry.Size(0.5, 0.5);
            if (options.Layout != null)
            {
                page_node.Layout = options.Layout;
            }
            page_node.Render(page);
        }
    }
}