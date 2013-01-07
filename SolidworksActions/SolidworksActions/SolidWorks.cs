using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SldWorks;
using SwConst;
using Centipede;

namespace SolidworksActions
{
    public class SolidWorksWrapper : IDisposable
    {
        private SolidWorksWrapper()
        {
            App = new SldWorks.SldWorks();
        }

        private static SldWorks.SldWorks App = null;

        void IDisposable.Dispose()
        {
            try
            {
                App.CloseAllDocuments(true);
                App = null;
            }
            catch
            { }
        }
        
        public ModelDoc2 OpenFile(String filename)
        {
            int err = 0, warn = 0;
            
            ModelDoc2 doc = App.OpenDoc6(filename,
                (int)GetDocumentTypeFromFilename(filename, true),
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "",
                ref err,
                ref warn);

            
            if (err != 0)
            {
                Object messages = new Object();
                Object msgIds = new Object(), msgTypes = new Object();
                int count = App.GetErrorMessages(out messages, out msgIds, out msgTypes);

                String exceptionMessage = String.Format("{0} error(s) occurred opening {1}:\n{2}", count, filename, String.Join("\n", messages));

                throw new SolidWorksException(exceptionMessage);
            }
            return doc;
        } 

        public enum FileType
        {
            Assembly = swDocumentTypes_e.swDocASSEMBLY,
            Drawing = swDocumentTypes_e.swDocDRAWING,
            Part = swDocumentTypes_e.swDocPART,

            Unknown = -1
        }

        public IEnumerable<String> GetErrors(out int count)
        {
            Object messages, msgIds, msgTypes;
            count = App.GetErrorMessages(out messages, out msgIds, out msgTypes);

            return messages as String[];
        }

        public IEnumerable<String> GetErrors()
        {
            int count;
            return GetErrors(out count);
        }

        public Exception GetErrorsAsException()
        {
            return new SolidWorksException();
        }

        public FileType GetDocumentTypeFromFilename(String filename, Boolean throwOnUnknown=false)
        {
            switch (System.IO.Path.GetExtension(filename).ToUpper())
            {
                case ".SLDPRT":
                    return FileType.Part;
                case ".SLDASM":
                    return FileType.Assembly;
                case ".SLDDRW":
                    return FileType.Part;
                default:
                    //return unknown, or throw??
                    if (throwOnUnknown)
                    {
                        throw new SolidWorksException(String.Format(@"Unknown filetype ""{0}""", filename));
                    }
                    else
                    {
                        return FileType.Unknown;
                    }
                    
            }
        }

        #region Singleton handling code
        private static Object _lockObject = new Object();
        private static SolidWorksWrapper _instance;
        public static SolidWorksWrapper Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lockObject)
                    {
                        if (_instance == null)
                        {
                            _instance = new SolidWorksWrapper();
                        }
                    }
                }

                return _instance;
            }
        }
        #endregion


        internal ModelDoc2 ActivateDoc(string DocName)
        {
            int errors = 0;
            ModelDoc2 doc = App.ActivateDoc2(DocName, true, ref errors);

            if (errors > 0)
            {
                throw new SolidWorksException();
            }

            return doc;
        }
    }

    [Serializable]
    public class SolidWorksException : Exception
    {
        public SolidWorksException()
            : base(String.Format("Error from Solidworks:\n{0}",
                String.Join("\n", SolidWorksWrapper.Instance.GetErrors())))
        { }
        public SolidWorksException(string message) : base(message) { }
        public SolidWorksException(string message, Exception inner) : base(message, inner) { }
        protected SolidWorksException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }
    }
}
