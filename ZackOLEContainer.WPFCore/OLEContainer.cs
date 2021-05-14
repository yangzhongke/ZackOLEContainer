using PInvoke;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Interop;
using System.Windows.Media;
using ZackOLEContainerCore;
using static PInvoke.User32;

namespace ZackOLEContainer.WPFCore
{
    public class OLEContainer : HwndHost
    {
        private PreviewHandlerHost previewHandler;
        protected override HandleRef BuildWindowCore(HandleRef hwndParent)
        {
            var hwndHost = IntPtr.Zero;

            int HOST_ID = 0x00000002;
            hwndHost = CreateWindowEx(0, "Static", "",
                                      WindowStyles.WS_CHILD | WindowStyles.WS_VISIBLE,
                                      0, 0,
                                      (int)this.ActualWidth, (int)this.ActualHeight,
                                      hwndParent.Handle,
                                      (IntPtr)HOST_ID,
                                      IntPtr.Zero,
                                      IntPtr.Zero);
            this.previewHandler = new PreviewHandlerHost(hwndHost);
            return new HandleRef(this, hwndHost);
        }

        protected override void OnWindowPositionChanged(Rect rcBoundingBox)
        {
            base.OnWindowPositionChanged(rcBoundingBox);
            if (DesignerProperties.GetIsInDesignMode(this))
            {
                return;
            }
            var rect = BoundsRelativeTo(this, Window.GetWindow(this));
            this.previewHandler.Resize(rect);
            UpdateWindow(this.Handle);
        }
        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);
            if (DesignerProperties.GetIsInDesignMode(this))
            {
                return;
            }
            var rect = BoundsRelativeTo(this, Window.GetWindow(this));
            this.previewHandler.Resize(rect);
            //UpdateWindow(this.Handle);
        }


        private static Rectangle BoundsRelativeTo(FrameworkElement element,
                                         Visual relativeTo)
        {
            var rect =
              element.TransformToVisual(relativeTo)
                     .TransformBounds(LayoutInformation.GetLayoutSlot(element));
            return new Rectangle((int)rect.X, (int)rect.Y, (int)rect.Width, (int)rect.Height);
        }

        public void OpenFile(String filename)
        {
            this.previewHandler.Open(filename);
            this.previewHandler.Resize(BoundsRelativeTo(this,Window.GetWindow(this)));
        }

        protected override void DestroyWindowCore(HandleRef hwnd)
        {
            DestroyWindow(hwnd.Handle);
            this.previewHandler.Dispose();
        }
    }
}
