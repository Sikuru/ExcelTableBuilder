﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Sikuru.ExcelTableBuilder.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "4.0.0.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    internal class Resources {
        
        private static global::System.Resources.ResourceManager resourceMan;
        
        private static global::System.Globalization.CultureInfo resourceCulture;
        
        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode")]
        internal Resources() {
        }
        
        /// <summary>
        ///   Returns the cached ResourceManager instance used by this class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Resources.ResourceManager ResourceManager {
            get {
                if (object.ReferenceEquals(resourceMan, null)) {
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("Sikuru.ExcelTableBuilder.Properties.Resources", typeof(Resources).Assembly);
                    resourceMan = temp;
                }
                return resourceMan;
            }
        }
        
        /// <summary>
        ///   Overrides the current thread's CurrentUICulture property for all
        ///   resource lookups using this strongly typed resource class.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        internal static global::System.Globalization.CultureInfo Culture {
            get {
                return resourceCulture;
            }
            set {
                resourceCulture = value;
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to public interface TableBinConvertable
        ///	{
        ///	}
        ///
        ///	public static class TableBinConverter
        ///	{
        ///		public const short BinClassDataConverterVersion = 8;
        ///		private static readonly byte[] ZeroLengthBytes = BitConverter.GetBytes((int)0);
        ///		private static readonly Type TypeInt16 = typeof(System.Int16);
        ///		private static readonly Type TypeInt32 = typeof(System.Int32);
        ///		private static readonly Type TypeInt64 = typeof(System.Int64);
        ///		private static readonly Type TypeSingle = typeof(System.Single);
        ///		private stati [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string TableBinConverterSource {
            get {
                return ResourceManager.GetString("TableBinConverterSource", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to public class TableProxy&lt;TKey, TValue&gt; where TValue : TableBase, new()
        ///	{
        ///		private Dictionary&lt;TKey, int&gt; _table_map_dic;
        ///		private MemoryStream _ms;
        ///
        ///		private Dictionary&lt;TKey, TValue&gt; _table_parsed = new Dictionary&lt;TKey, TValue&gt;();
        ///		private int _count;
        ///
        ///		public TValue this[TKey id] { get { return Get(id); } }
        ///
        ///		public int Count { get { return _count; } }
        ///		public Dictionary&lt;TKey, int&gt;.KeyCollection Keys { get { return _table_map_dic.Keys; } }
        ///		public List&lt;TValue&gt; Values
        ///		{
        ///			get
        ///			{
        /// [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string TableProxy {
            get {
                return ResourceManager.GetString("TableProxy", resourceCulture);
            }
        }
    }
}
