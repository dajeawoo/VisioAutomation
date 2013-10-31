using System.Collections.Generic;
using VisioAutomation.Shapes.CustomProperties;
using VA = VisioAutomation;
using SMA = System.Management.Automation;
using IVisio = Microsoft.Office.Interop.Visio;

namespace VisioPS.Commands
{
    [SMA.Cmdlet(SMA.VerbsCommon.Set, "VisioCustomProperty")]
    public class Set_VisioCustomProperty : VisioPS.VisioPSCmdlet
    {
        private int _LangID = -1;
        private int _sortKey = -1;
        private int _type = 0; // 0 = string
        private int _verify = -1;

        [SMA.Parameter(Position = 0, Mandatory = true)]
        public string Name { get; set; }

        [SMA.Parameter(Position = 1, Mandatory = true)]
        public string Value { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string Label { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string Prompt { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public int LangId
        {
            get { return _LangID; }
            set { _LangID = value; }
        }

        [SMA.Parameter(Mandatory = false)]
        public int SortKey
        {
            get { return _sortKey; }
            set { _sortKey = value; }
        }

        [SMA.Parameter(Mandatory = false)]
        public int Type
        {
            get { return _type; }
            set { _type = value; }
        }

        [SMA.Parameter(Mandatory = false)]
        public int Verify
        {
            get { return _verify; }
            set { _verify = value; }
        }

        [SMA.Parameter(Mandatory = false)]
        public IVisio.Shape[] Shapes;

        protected override void ProcessRecord()
        {
            var cp = new CustomPropertyCells();
            cp.Value = this.Value;
            if (this.Label != null)
            {
                cp.Label = this.Label;
            }

            if (this._LangID >= 0)
            {
                cp.LangId = this._LangID;
            }

            if (this.Prompt != null)
            {
                cp.Prompt = this.Prompt;
            }

            if (this._sortKey >= 0)
            {
                cp.SortKey = this._sortKey;
            }

            cp.Type = (int) this._type;

            if (this._verify >= 0)
            {
                cp.Verify = this._verify;
            }

            var scriptingsession = this.ScriptingSession;
            scriptingsession.CustomProp.Set(this.Shapes, this.Name , cp);
        }
    }
}