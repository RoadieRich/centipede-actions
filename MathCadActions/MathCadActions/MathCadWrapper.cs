using System;

namespace MathCADActions
{
    public class MathCadWrapper : IDisposable
    {

        private MathCadWrapper()
        {
            this._application = new Mathcad.Application();
        }

        static readonly object _LockObject = new object();
        private volatile static MathCadWrapper _instance;
        private Mathcad.Application _application;

        public static MathCadWrapper Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_LockObject)
                    {
                        if (_instance == null)
                        {
                            _instance = new MathCadWrapper();
                        }

                    }
                }
                return _instance;
            }
        }

        public Mathcad.Application Application
        {
            get { return this._application; }
        }


        void IDisposable.Dispose()
        {
            if(this._application != null)
            {
                this._application.CloseAll();
                this._application = null;
            }
        }
    }
}