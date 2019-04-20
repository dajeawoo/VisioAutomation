using VASS=VisioAutomation.ShapeSheet;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPowerShell.Commands.VisioControl
{
    [SMA.Cmdlet(SMA.VerbsCommon.Get, Nouns.VisioControl)]
    public class GetVisioControl : VisioCmdlet
    {
        // CONTEXT:SHAPE
        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;
        
        protected override void ProcessRecord()
        {
            var targetshapes = new VisioScripting.TargetShapes(this.Shapes);
            var type = VASS.CellValueType.Formula;
            var dic_shape_to_listofcontrolscells = this.Client.Control.GetControls(targetshapes, type);
            this.WriteObject(dic_shape_to_listofcontrolscells);
        }
    }
}