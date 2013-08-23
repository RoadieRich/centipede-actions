using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using Mathcad;

namespace MathCADActions
{
    public class MathCadWrapper : IDisposable
    {

        private MathCadWrapper()
        {
            this._application = new Application();
        }

        static readonly object _LockObject = new object();
        private volatile static MathCadWrapper _instance;
        private Application _application;

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

    public class WorkSheetWrapper
    {
        private readonly Worksheet _worksheet;

        public WorkSheetWrapper(Worksheet worksheet)
        {
            _worksheet = worksheet;
            this._lazyRegions = new Lazy<IEnumerable<IMathcadRegion2>>(() => this._worksheet.Regions.Cast<IMathcadRegion2>());
            this._lazyVarNameMap = new Lazy<IDictionary<Variable, IMathcadRegion2>>(this.GenerateVarNameMap);
        }

        private readonly Lazy<IEnumerable<IMathcadRegion2>> _lazyRegions;
        private readonly Lazy<IDictionary<Variable, IMathcadRegion2>> _lazyVarNameMap;

        public struct Variable : IEquatable<Variable>
        {

            public Variable(string name, string subscript)
                : this(name)
            {
                this.Subscript = subscript;
            }
            
            public readonly string Name;
            public readonly string Subscript;

            public Variable(string name)
            {
                Name = name;
                this.Subscript = null;
            }

            #region IEquatable<Variable>

            /// <summary>
            /// Indicates whether the current object is equal to another object of the same type.
            /// </summary>
            /// <returns>
            /// true if the current object is equal to the <paramref name="other"/> parameter; otherwise, false.
            /// </returns>
            /// <param name="other">An object to compare with this object.</param>
            bool IEquatable<Variable>.Equals(Variable other)
            {
                return string.Equals(this.Name, other.Name) && string.Equals(this.Subscript, other.Subscript);
            }

            /// <summary>
            /// Indicates whether this instance and a specified object are equal.
            /// </summary>
            /// <returns>
            /// true if <paramref name="obj"/> and this instance are the same type and represent the same value; otherwise, false.
            /// </returns>
            /// <param name="obj">Another object to compare to. </param><filterpriority>2</filterpriority>
            public override bool Equals(object obj)
            {
                if (ReferenceEquals(null, obj))
                {
                    return false;
                }
                return obj is Variable && Equals((Variable)obj);
            }

            /// <summary>
            /// Returns the hash code for this instance.
            /// </summary>
            /// <returns>
            /// A 32-bit signed integer that is the hash code for this instance.
            /// </returns>
            /// <filterpriority>2</filterpriority>
            public override int GetHashCode()
            {
                unchecked
                {
                    return ((this.Name != null ? this.Name.GetHashCode() : 0) * 397) ^
                           (this.Subscript != null ? this.Subscript.GetHashCode() : 0);
                }
            }

            #endregion
        }

        public IEnumerable<IMathcadRegion2> Regions
        {
            get { return this._lazyRegions.Value; }
        }

        public IDictionary<Variable, IMathcadRegion2> VarNameMap
        {
            get { return this._lazyVarNameMap.Value; }
        }

        public IMathcadRegion2 GetRegionByTag(string tag)
        {
            return this.Regions.FirstOrDefault(region => region.Tag.Equals(tag));
        }

        public IMathcadRegion2 GetRegionByVar(String varName, string subscript = null)
        {
            Variable v = new Variable(varName, subscript);
            return VarNameMap[v];
        }

        private Dictionary<Variable, IMathcadRegion2> GenerateVarNameMap()
        {
            return this.Regions.ToDictionary(this.GetVarNameInRegion);
        }

        private Variable GetVarNameInRegion(IMathcadRegion2 region)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(region.MathInterface.XML);

            XmlElement element;
            try
            {
                element = xmlDoc.GetElementsByTagName("ml:id").OfType<XmlElement>().First();
            }
            catch (InvalidOperationException)
            {
                return new Variable();
            }

            string varName = element.Value;
            string subscript = element.GetAttribute(@"subscript");

            return new Variable(varName, !String.IsNullOrEmpty(subscript) ? subscript : null);
        }

        public void SaveAs(string fname, MCFileFormat format = MCFileFormat.mcMcadCurrentVersion)
        {
            _worksheet.SaveAs(fname, format);
        }

        public void Save()
        {
            _worksheet.Save();
        }

        public void Close(MCSaveOption options = MCSaveOption.mcPromptToSaveChanges)
        {
            _worksheet.Close(options);
        }

        public object GetValue(string varName)
        {
            return _worksheet.GetValue(varName);
        }
    }
}