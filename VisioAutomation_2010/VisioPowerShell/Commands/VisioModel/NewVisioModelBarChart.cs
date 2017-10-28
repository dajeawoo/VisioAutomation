using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, VisioPowerShell.Commands.Nouns.VisioModelBarChart)]
    public class NewVisioModelBarChart : VisioCmdlet
    {
        [SMA.Parameter(Position = 0, Mandatory = true)]
        public double X0 { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public double Y0 { get; set; }

        [SMA.Parameter(Position = 2, Mandatory = true)]
        public double X1 { get; set; }

        [SMA.Parameter(Position = 3, Mandatory = true)]
        public double Y1 { get; set; }

        [SMA.Parameter(Mandatory = true)]
        public double[] Values;

        [SMA.Parameter(Mandatory = false)]
        public string[] Labels;

        protected override void ProcessRecord()
        {
            var rect = this.GetRectangle();
            var chart = new VisioAutomation.Models.Charting.BarChart(rect);
            chart.DataPoints = new VisioAutomation.Models.Charting.DataPointList(this.Values, this.Labels);
            this.WriteObject(chart);
        }

        protected VisioAutomation.Geometry.Rectangle GetRectangle()
        {
            return new VisioAutomation.Geometry.Rectangle(this.X0, this.Y0, this.X1, this.Y1);
        }
    }
}