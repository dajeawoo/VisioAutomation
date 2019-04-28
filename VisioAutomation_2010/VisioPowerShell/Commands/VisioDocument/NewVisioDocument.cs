using SMA = System.Management.Automation;

namespace VisioPowerShell.Commands.VisioDocument
{
    [SMA.Cmdlet(SMA.VerbsCommon.New, Nouns.VisioDocument)]
    public class NewVisioDocument : VisioCmdlet
    {
        [SMA.Parameter(Mandatory = false)]
        public string Template { get; set; }

        [SMA.Parameter(Mandatory = false)]
        public string[] Stencil { get; set; }


        protected override void ProcessRecord()
        {
            if (!this.Client.Application.HasAttachedApplication)
            {
                this.Client.Application.NewAttachedApplication();
            }
            else
            {
                if (!this.Client.Application.ValidateAttachedApplication())
                {
                    this.Client.Application.NewAttachedApplication();
                }
            }

            var doc = this.Client.Document.NewDocumentFromTemplate(this.Template);

            if (this.Stencil != null)
            {
                foreach (string stencil in this.Stencil)
                {
                    var stencildoc = this.Client.Document.OpenStencilDocument(stencil);
                }

            }

            this.WriteObject(doc);
        }
    }
}