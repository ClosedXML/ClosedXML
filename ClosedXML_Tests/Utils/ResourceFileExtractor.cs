using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace ClosedXML_Tests
{
    /// <summary>
    /// Summary description for ResourceFileExtractor.
    /// </summary>
    public sealed class ResourceFileExtractor
    {
        #region Static
        #region Private fields
        private static readonly Dictionary<string, ResourceFileExtractor> ms_defaultExtractors =
                new Dictionary<string, ResourceFileExtractor>();
        #endregion
        #region Public properties
        /// <summary>Instance of resource extractor for executing assembly </summary>
        public static ResourceFileExtractor Instance
        {
            get
            {
                ResourceFileExtractor _return;
                Assembly _assembly = Assembly.GetCallingAssembly();
                string _key = _assembly.GetName().FullName;
                if (!ms_defaultExtractors.TryGetValue(_key, out _return))
                {
                    lock (ms_defaultExtractors)
                    {
                        if (!ms_defaultExtractors.TryGetValue(_key, out _return))
                        {
                            _return = new ResourceFileExtractor(_assembly, true, null);
                            ms_defaultExtractors.Add(_key, _return);
                        }
                    }
                }
                return _return;
            }
        }
        #endregion
        #region Public methods
        #endregion
        #endregion
        #region Private fields
        private readonly Assembly m_assembly;
        private readonly ResourceFileExtractor m_baseExtractor;
        private readonly string m_assemblyName;

        private bool m_isStatic;
        private string m_resourceFilePath;
        #endregion
        #region Constructors
        /// <summary>
        /// Create instance
        /// </summary>
        /// <param name="resourceFilePath"><c>ResourceFilePath</c> in assembly. Example: .Properties.Scripts.</param>
        /// <param name="baseExtractor"></param>
        public ResourceFileExtractor(string resourceFilePath, ResourceFileExtractor baseExtractor)
                : this(Assembly.GetCallingAssembly(), baseExtractor)
        {
            m_resourceFilePath = resourceFilePath;
        }
        /// <summary>
        /// Create instance
        /// </summary>
        /// <param name="baseExtractor"></param>
        public ResourceFileExtractor(ResourceFileExtractor baseExtractor)
                : this(Assembly.GetCallingAssembly(), baseExtractor)
        {
        }
        /// <summary>
        /// Create instance
        /// </summary>
        /// <param name="resourcePath"><c>ResourceFilePath</c> in assembly. Example: .Properties.Scripts.</param>
        public ResourceFileExtractor(string resourcePath)
                : this(Assembly.GetCallingAssembly(), resourcePath)
        {
        }
        /// <summary>
        /// Instance constructor
        /// </summary>
        /// <param name="assembly"></param>
        /// <param name="resourcePath"></param>
        public ResourceFileExtractor(Assembly assembly, string resourcePath)
                : this(assembly ?? Assembly.GetCallingAssembly())
        {
            m_resourceFilePath = resourcePath;
        }
        /// <summary>
        /// Instance constructor
        /// </summary>
        public ResourceFileExtractor()
                : this(Assembly.GetCallingAssembly())
        {
        }
        /// <summary>
        /// Instance constructor
        /// </summary>
        /// <param name="assembly"></param>
        public ResourceFileExtractor(Assembly assembly)
                : this(assembly ?? Assembly.GetCallingAssembly(), (ResourceFileExtractor) null)
        {
        }
        /// <summary>
        /// Instance constructor
        /// </summary>
        /// <param name="assembly"></param>
        /// <param name="baseExtractor"></param>
        public ResourceFileExtractor(Assembly assembly, ResourceFileExtractor baseExtractor)
                : this(assembly ?? Assembly.GetCallingAssembly(), false, baseExtractor)
        {
        }
        /// <summary>
        /// Instance constructor
        /// </summary>
        /// <param name="assembly"></param>
        /// <param name="isStatic"></param>
        /// <param name="baseExtractor"></param>
        /// <exception cref="ArgumentNullException">Argument is null.</exception>
        private ResourceFileExtractor(Assembly assembly, bool isStatic, ResourceFileExtractor baseExtractor)
        {
            #region Check
            if (ReferenceEquals(assembly, null))
            {
                throw new ArgumentNullException("assembly");
            }
            #endregion
            m_assembly = assembly;
            m_baseExtractor = baseExtractor;
            m_assemblyName = Assembly.GetName().Name;
            IsStatic = isStatic;
            m_resourceFilePath = ".Resources.";
        }
        #endregion
        #region Public properties
        /// <summary> Work assembly </summary>
        public Assembly Assembly
        {
            [DebuggerStepThrough]
            get { return m_assembly; }
        }
        /// <summary> Work assembly name </summary>
        public string AssemblyName
        {
            [DebuggerStepThrough]
            get { return m_assemblyName; }
        }
        /// <summary>
        /// Path to read resource files. Example: .Resources.Upgrades.
        /// </summary>
        public string ResourceFilePath
        {
            [DebuggerStepThrough]
            get { return m_resourceFilePath; }
            [DebuggerStepThrough]
            set { m_resourceFilePath = value; }
        }
        public bool IsStatic
        {
            [DebuggerStepThrough]
            get { return m_isStatic; }
            [DebuggerStepThrough]
            set { m_isStatic = value; }
        }
        public IEnumerable<string> GetFileNames()
        {
            string _path = AssemblyName + m_resourceFilePath;
            foreach (string _resourceName in Assembly.GetManifestResourceNames())
            {
                if (_resourceName.StartsWith(_path))
                {
                    yield return _resourceName.Replace(_path, string.Empty);
                }
            }
        }
        #endregion
        #region Public methods
        public string ReadFileFromRes(string fileName)
        {
            Stream _stream = ReadFileFromResToStream(fileName);
            string _result;
            StreamReader sr = new StreamReader(_stream);
            try
            {
                _result = sr.ReadToEnd();
            }
            finally
            {
                sr.Close();
            }
            return _result;
        }

        public string ReadFileFromResFormat(string fileName, params object[] formatArgs)
        {
            return string.Format(ReadFileFromRes(fileName), formatArgs);
        }

        /// <summary>
        /// Read file in current assembly by specific path
        /// </summary>
        /// <param name="specificPath">Specific path</param>
        /// <param name="fileName">Read file name</param>
        /// <returns></returns>
        public string ReadSpecificFileFromRes(string specificPath, string fileName)
        {
            ResourceFileExtractor _ext = new ResourceFileExtractor(Assembly, specificPath);
            return _ext.ReadFileFromRes(fileName);
        }
        /// <summary>
        /// Read file in current assembly by specific file name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="ApplicationException"><c>ApplicationException</c>.</exception>
        public Stream ReadFileFromResToStream(string fileName)
        {
            string _nameResFile = AssemblyName + m_resourceFilePath + fileName;
            Stream _stream = Assembly.GetManifestResourceStream(_nameResFile);
            #region Not found
            if (ReferenceEquals(_stream, null))
            {
                #region Get from base extractor
                if (!ReferenceEquals(m_baseExtractor, null))
                {
                    return m_baseExtractor.ReadFileFromResToStream(fileName);
                }
                #endregion
                throw new ApplicationException("Can't find resource file " + _nameResFile);
            }
            #endregion
            return _stream;
        }
        #endregion
    }
}