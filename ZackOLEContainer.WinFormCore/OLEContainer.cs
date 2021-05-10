using System;
using System.Windows.Forms;
using ZackOLEContainerCore;

namespace ZackOLEContainer.WinFormCore
{
    public class OLEContainer:Control
    {
        private PreviewHandlerHost previewHandler;

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            this.previewHandler = new PreviewHandlerHost(this.Handle);
        }
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if(this.previewHandler!=null)
            {
                this.previewHandler.Resize(this.ClientRectangle);
            }            
        }

        public void OpenFile(String filename)
        {
            this.previewHandler.Open(filename);
            this.previewHandler.Resize(this.ClientRectangle);
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            this.previewHandler.Dispose();
        }
    }
}
