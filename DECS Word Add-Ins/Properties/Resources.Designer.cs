﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace DecsWordAddIns.Properties {
    using System;
    
    
    /// <summary>
    ///   A strongly-typed resource class, for looking up localized strings, etc.
    /// </summary>
    // This class was auto-generated by the StronglyTypedResourceBuilder
    // class via a tool like ResGen or Visual Studio.
    // To add or remove a member, edit your .ResX file then rerun ResGen
    // with the /str option, or rebuild your VS project.
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0")]
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
                    global::System.Resources.ResourceManager temp = new global::System.Resources.ResourceManager("DecsWordAddIns.Properties.Resources", typeof(Resources).Assembly);
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
        ///   Looks up a localized resource of type System.Drawing.Bitmap.
        /// </summary>
        internal static System.Drawing.Bitmap clipboard {
            get {
                object obj = ResourceManager.GetObject("clipboard", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   Looks up a localized resource of type System.Drawing.Bitmap.
        /// </summary>
        internal static System.Drawing.Bitmap crane {
            get {
                object obj = ResourceManager.GetObject("crane", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   Looks up a localized resource of type System.Drawing.Bitmap.
        /// </summary>
        internal static System.Drawing.Bitmap icd_10_zoom {
            get {
                object obj = ResourceManager.GetObject("icd_10_zoom", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   Looks up a localized resource of type System.Drawing.Bitmap.
        /// </summary>
        internal static System.Drawing.Bitmap ICD_case_statements {
            get {
                object obj = ResourceManager.GetObject("ICD_case_statements", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   Looks up a localized resource of type System.Drawing.Bitmap.
        /// </summary>
        internal static System.Drawing.Bitmap ICD_list_statement {
            get {
                object obj = ResourceManager.GetObject("ICD_list_statement", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &lt;!DOCTYPE html&gt;
        ///&lt;html&gt;
        ///&lt;body&gt;
        ///    &lt;p&gt;&lt;br&gt;&lt;/p&gt;
        ///    &lt;p&gt;Hello {{ cookiecutter.__requestor_salutation }},&lt;/p&gt;
        ///    &lt;p&gt;I am pleased to report that your DECS request is completed and the data are now available via OneDrive:&lt;/p&gt;
        ///    &lt;ul&gt;&lt;li&gt;&lt;strong&gt;Request #:&lt;/strong&gt; DECS-{{ cookiecutter.task_number }}&lt;/li&gt;
        ///    &lt;li&gt;&lt;strong&gt;OneDrive Link:&lt;/strong&gt; link&lt;/li&gt;&lt;/ul&gt;
        ///    &lt;p&gt;All data from this request are governed by UCSD policies and IRB rules and regulations.&lt;/p&gt;
        ///    &lt;h2 id=&quot;DECSRequestReady(SlicerDicer)-DataR [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string one_drive_email_body {
            get {
                return ResourceManager.GetString("one_drive_email_body", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized resource of type System.Drawing.Bitmap.
        /// </summary>
        internal static System.Drawing.Bitmap school_of_medicine {
            get {
                object obj = ResourceManager.GetObject("school_of_medicine", resourceCulture);
                return ((System.Drawing.Bitmap)(obj));
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to befox: Brian E. Fox
        ///g1zhang: Ge Zhang
        ///kjdelaney: Kevin J. Delaney
        ///mjmarshall: Michael J. Marshall
        ///padesai: Paresh Desai
        ///pshipman: Perry Shipman
        ///tmelander: Troy M. Melander.
        /// </summary>
        internal static string usernames {
            get {
                return ResourceManager.GetString("usernames", resourceCulture);
            }
        }
        
        /// <summary>
        ///   Looks up a localized string similar to &lt;!DOCTYPE html&gt;
        ///&lt;html&gt;
        ///&lt;body&gt;
        ///    &lt;p&gt;Hello {{ cookiecutter.__requestor_salutation }},&lt;/p&gt;
        ///    &lt;p&gt;I am pleased to report that your DECS request is complete and your results are available on Virtual Research Desktop (VRD):&lt;/p&gt;
        ///    &lt;ul&gt;
        ///    &lt;li&gt;&lt;strong&gt;Request #:&lt;/strong&gt; DECS-{{ cookiecutter.task_number }}&lt;/li&gt;
        ///    &lt;li&gt;&lt;strong&gt;VRD Shared folder:&lt;/strong&gt; SecureDrop/{{ cookiecutter.__directory_name }}&amp;mdash;
        ///    Please transfer the results file(s) from the shared folder to your VRD personal folder.&lt;/li [rest of string was truncated]&quot;;.
        /// </summary>
        internal static string vrd_email_body {
            get {
                return ResourceManager.GetString("vrd_email_body", resourceCulture);
            }
        }
    }
}
