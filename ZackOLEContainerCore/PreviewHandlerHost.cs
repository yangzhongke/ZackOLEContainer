using Microsoft.Win32;
using PInvoke;
using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;

namespace ZackOLEContainerCore
{
    /// <summary>
    /// adapted from https://www.codeproject.com/Tips/487566/OLE-container-surrogate-for-NET
    /// </summary>
    public class PreviewHandlerHost:IDisposable
    {
        private object _currentPreviewHandler;
        private Guid _currentPreviewHandlerGuid  = Guid.Empty;
        private Stream _currentPreviewHandlerStream;

        private IntPtr handle;
        public PreviewHandlerHost(IntPtr handle)
        {
            this.handle = handle;
        }

        public dynamic ComObject
        {
            get
            {
                return this._currentPreviewHandler;
            }
        }

        private Guid GetPreviewHandlerGUID(string filename)
        {
            // open the registry key corresponding to the file extension
            RegistryKey regKeyExt = Registry.ClassesRoot.OpenSubKey(Path.GetExtension(filename));
            if(regKeyExt==null)
            {
                return Guid.Empty;
            }
            // open the key that indicates the GUID of the preview handler type
            RegistryKey regKey = regKeyExt.OpenSubKey("shellex\\{8895b1c6-b41f-4c1c-a562-0d564250836f}");
            if (regKey != null) return new Guid(Convert.ToString(regKey.GetValue(null)));

            // sometimes preview handlers are declared on key for the class
            string className = Convert.ToString(regKeyExt.GetValue(null));
            if(className==null)
            {
                return Guid.Empty;
            }
            regKey = Registry.ClassesRoot.OpenSubKey(className + "\\shellex\\{8895b1c6-b41f-4c1c-a562-0d564250836f}");
            if(regKey==null)
            {
                return Guid.Empty;
            }
            else
            {
                return new Guid(Convert.ToString(regKey.GetValue(null)));
            }

        }

        /// <summary>
        /// Resizes the hosted preview handler when this PreviewHandlerHost is resized.
        /// </summary>
        /// <param name="e"></param>
        public void Resize(Rectangle rectangle)
        {
            if(_currentPreviewHandler!=null)
            {
                IPreviewHandler previewHandler = (IPreviewHandler)_currentPreviewHandler;
                previewHandler.SetRect(ref rectangle);
            }            
        }

        public bool Open(string filename)
        {
            UnloadPreviewHandler();

            // try to get GUID for the preview handler
            Guid guid = GetPreviewHandlerGUID(filename);
            if(guid==Guid.Empty)
            {
                throw new ArgumentException("NoPreviewAvailable");
            }

            if (guid != _currentPreviewHandlerGuid)
            {
                _currentPreviewHandlerGuid = guid;

                // need to instantiate a different COM type (file format has changed)
                if (_currentPreviewHandler != null) Marshal.FinalReleaseComObject(_currentPreviewHandler);

                // use reflection to instantiate the preview handler type
                Type comType = Type.GetTypeFromCLSID(_currentPreviewHandlerGuid);
                _currentPreviewHandler = Activator.CreateInstance(comType);
            }

            if (_currentPreviewHandler is IInitializeWithFile)
            {
                // some handlers accept a filename
                ((IInitializeWithFile)_currentPreviewHandler).Initialize(filename, 0);
            }
            else if (_currentPreviewHandler is IInitializeWithStream)
            {
                // other handlers want an IStream (in this case, a file stream)
                _currentPreviewHandlerStream = File.Open(filename, FileMode.Open);
                StreamWrapper stream = new StreamWrapper(_currentPreviewHandlerStream);
                ((IInitializeWithStream)_currentPreviewHandler).Initialize(stream, 0);
            }

            if (_currentPreviewHandler is IPreviewHandler)
            {
                // bind the preview handler to the control's bounds and preview the content
                Rectangle r = GetWindowRect(this.handle);
                ((IPreviewHandler)_currentPreviewHandler).SetWindow(this.handle, ref r);
                ((IPreviewHandler)_currentPreviewHandler).DoPreview();
                return true;
            }
            return false;
        }

        private Rectangle GetWindowRect(IntPtr handle)
        {
            if(User32.GetWindowRect(handle, out RECT r))
            {
                return new Rectangle(r.left,r.top,r.right-r.left,r.bottom-r.top);
            }
            else
            {
                throw new ArgumentException("Cannot GetWindowRect for "+handle);
            }
        }


        public void UnloadPreviewHandler()
        {
            if (_currentPreviewHandler is IPreviewHandler)
            {
                try
                {
                    // explicitly unload the content
                    ((IPreviewHandler)_currentPreviewHandler).Unload();
                }
                catch (Exception)
                {
                }
            }

            if (_currentPreviewHandlerStream != null)
            {
                _currentPreviewHandlerStream.Close();
                _currentPreviewHandlerStream = null;
            }

            _currentPreviewHandlerGuid = new Guid();
        }

        public void Dispose()
        {
            UnloadPreviewHandler();
            if (_currentPreviewHandler != null)
            {
                Marshal.FinalReleaseComObject(_currentPreviewHandler);
                _currentPreviewHandler = null;
                GC.Collect();
            }
        }
    }


   
}
