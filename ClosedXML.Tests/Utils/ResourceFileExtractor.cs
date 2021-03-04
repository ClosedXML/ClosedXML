using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ClosedXML.Tests
{
    /// <summary>
    /// Summary description for ResourceFileExtractor.
    /// </summary>
    public sealed class ResourceFileExtractor
    {
        #region Static

        #region Private fields

        private static readonly IDictionary<string, ResourceFileExtractor> extractors = new ConcurrentDictionary<string, ResourceFileExtractor>();

        #endregion Private fields

        #region Public properties

        /// <summary>Instance of resource extractor for executing assembly </summary>
        public static ResourceFileExtractor Instance
        {
            get
            {
                Assembly _assembly = Assembly.GetCallingAssembly();
                string _key = _assembly.GetName().FullName;
                if (!extractors.TryGetValue(_key, out ResourceFileExtractor extractor)
                    && !extractors.TryGetValue(_key, out extractor))
                {
                    extractor = new ResourceFileExtractor(_assembly, true, null);
                    extractors.Add(_key, extractor);
                }

                return extractor;
            }
        }

        #endregion Public properties

        #endregion Static

        #region Private fields

        private readonly Assembly m_assembly;
        private readonly ResourceFileExtractor m_baseExtractor;

        private bool m_isStatic;
        //private string ResourceFilePath { get; }

        #endregion Private fields

        #region Constructors

        /// <summary>
        /// Create instance
        /// </summary>
        /// <param name="resourceFilePath"><c>ResourceFilePath</c> in assembly. Example: .Properties.Scripts.</param>
        /// <param name="baseExtractor"></param>
        public ResourceFileExtractor(string resourceFilePath, ResourceFileExtractor baseExtractor)
                : this(Assembly.GetCallingAssembly(), baseExtractor)
        {
            ResourceFilePath = resourceFilePath;
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
            ResourceFilePath = resourcePath;
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
                : this(assembly ?? Assembly.GetCallingAssembly(), (ResourceFileExtractor)null)
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

            if (assembly is null)
            {
                throw new ArgumentNullException("assembly");
            }

            #endregion Check

            Assembly = assembly;
            m_baseExtractor = baseExtractor;
            AssemblyName = Assembly.GetName().Name;
            IsStatic = isStatic;
            ResourceFilePath = ".Resources.";
        }

        #endregion Constructors

        #region Public properties

        /// <summary> Work assembly </summary>
        public Assembly Assembly { get; }

        /// <summary> Work assembly name </summary>
        public string AssemblyName { get; }

        /// <summary>
        /// Path to read resource files. Example: .Resources.Upgrades.
        /// </summary>
        public string ResourceFilePath { get; }

        public bool IsStatic { get; set; }

        public IEnumerable<string> GetFileNames(Func<String, Boolean> predicate = null)
        {
            predicate = predicate ?? (s => true);

            string _path = AssemblyName + ResourceFilePath;
            foreach (string _resourceName in Assembly.GetManifestResourceNames())
            {
                if (_resourceName.StartsWith(_path) && predicate(_resourceName))
                {
                    yield return _resourceName.Replace(_path, string.Empty);
                }
            }
        }

        #endregion Public properties

        #region Public methods

        public string ReadFileFromResource(string fileName)
        {
            Stream _stream = ReadFileFromResourceToStream(fileName);
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

        public string ReadFileFromResourceFormat(string fileName, params object[] formatArgs)
        {
            return string.Format(ReadFileFromResource(fileName), formatArgs);
        }

        /// <summary>
        /// Read file in current assembly by specific path
        /// </summary>
        /// <param name="specificPath">Specific path</param>
        /// <param name="fileName">Read file name</param>
        /// <returns></returns>
        public string ReadSpecificFileFromResource(string specificPath, string fileName)
        {
            ResourceFileExtractor _ext = new ResourceFileExtractor(Assembly, specificPath);
            return _ext.ReadFileFromResource(fileName);
        }

        /// <summary>
        /// Read file in current assembly by specific file name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        /// <exception cref="ApplicationException"><c>ApplicationException</c>.</exception>
        public Stream ReadFileFromResourceToStream(string fileName)
        {
            string _nameResFile = AssemblyName + ResourceFilePath + fileName;
            Stream _stream = Assembly.GetManifestResourceStream(_nameResFile);

            #region Not found

            if (_stream is null)
            {
                #region Get from base extractor

                if (!(m_baseExtractor is null))
                {
                    return m_baseExtractor.ReadFileFromResourceToStream(fileName);
                }

                #endregion Get from base extractor

                throw new ArgumentException("Can't find resource file " + _nameResFile, nameof(fileName));
            }

            #endregion Not found

            return _stream;
        }

        #endregion Public methods
    }
}
