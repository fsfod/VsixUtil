using System;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.ExtensionManager;
using ImpromptuInterface;
using ImpromptuInterface.Optimization;
using ImpromptuInterface.Build;
using System.Collections.Generic;

namespace VsixUtil
{
    public sealed class ExtensionManagerFactory
    {
        private static Assembly LoadImplementationAssembly(VsVersion version)
        {
            var format = "Microsoft.VisualStudio.ExtensionManager.Implementation, Version={0}.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
            var strongName = string.Format(format, VsVersionUtil.GetVersionNumber(version));
            return Assembly.Load(strongName);
        }

        private static Assembly LoadSettingsAssembly(VsVersion version)
        {
            var format = "Microsoft.VisualStudio.Settings{0}, Version={1}.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
            string suffix = "";
            switch (version)
            {
                case VsVersion.Vs2010:
                    suffix = "";
                    break;
                case VsVersion.Vs2012:
                    suffix = ".11.0";
                    break;
                case VsVersion.Vs2013:
                    suffix = ".12.0";
                    break;
                case VsVersion.Vs2015:
                    suffix = ".14.0";
                    break;
                case VsVersion.Vs2017:
                    suffix = ".15.0";
                    break;
                case VsVersion.Vs2019:
                    suffix = ".15.0";
                    break;
                case VsVersion.Vs2022:
                    suffix = ".15.0";
                    break;
                default:
                    throw new Exception("Bad Version");
            }

            var strongName = string.Format(format, suffix, VsVersionUtil.GetVersionNumber(version));
            return Assembly.Load(strongName);
        }

        private static Type GetExtensionManagerServiceType(VsVersion version)
        {
            var assembly = LoadImplementationAssembly(version);
            return assembly.GetType("Microsoft.VisualStudio.ExtensionManager.ExtensionManagerService");
        }

        public static IVsExtensionManager CreateExtensionManager(InstalledVersion installedVersion, string rootSuffix, ExtensionManagerServiceMode loadMode = ExtensionManagerServiceMode.Default)
        {
            var settingsAssembly = LoadSettingsAssembly(installedVersion.VsVersion);

            var externalSettingsManagerType = settingsAssembly.GetType("Microsoft.VisualStudio.Settings.ExternalSettingsManager");
            var settingsManager = externalSettingsManagerType
                .GetMethods()
                .Where(x => x.Name == "CreateForApplication")
                .Where(x =>
                {
                    var parameters = x.GetParameters();
                    return
                        parameters.Length == 2 &&
                        parameters[0].ParameterType == typeof(string) &&
                        parameters[1].ParameterType == typeof(string);
                })
                .FirstOrDefault()
                .Invoke(null, new[] { installedVersion.ApplicationPath, rootSuffix });
            
            var extensionManagerServiceType = GetExtensionManagerServiceType(installedVersion.VsVersion);

            object extensionManager = null;

            bool safeMode = false;
            bool skipSdkDirectories = false;
            bool scanAlways = false;
            bool verboseLog = true;

            if (false)
            {
                extensionManager = Activator.CreateInstance(extensionManagerServiceType, settingsManager);
            }
            else
            {
                var assembly = extensionManagerServiceType.Assembly;
                var modeType = assembly.GetType("Microsoft.VisualStudio.ExtensionManager.ExtensionManagerServiceMode");

                int mode = 0;

                if (safeMode && !TrySetEnumFlag(modeType, ExtensionManagerServiceMode.SafeMode, ref mode))
                {
                    throw new InvalidOperationException("Safe mode option no longer exists");
                }

                if (skipSdkDirectories && !TrySetEnumFlag(modeType, ExtensionManagerServiceMode.DoNotScanSdkDirectories, ref mode))
                {
                    throw new InvalidOperationException("SkipSdkDirectories mode option no longer exists");
                }

                if (scanAlways && !TrySetEnumFlag(modeType, ExtensionManagerServiceMode.ScanAlways, ref mode))
                {
                    throw new InvalidOperationException("ScanAlways mode option no longer exists");
                }



                var constructors = extensionManagerServiceType.GetConstructors().ToList();

                var loggerType = constructors.SelectMany(c => c.GetParameters()).
                                 Where(p => p.ParameterType.Name == "ILogger").
                                 Select(p => p.ParameterType).
                                 FirstOrDefault();

                var logger = new Logger(verboseLog);
                var vslogger =  ActLike(logger, loggerType);
                var modeValue = Enum.ToObject(modeType, mode);

                extensionManager = Activator.CreateInstance(extensionManagerServiceType, settingsManager, modeValue, vslogger);
            }          

            var wrap = extensionManager.ActLike<IVsExtensionManager>();

            return (IVsExtensionManager)wrap;
        }

        static int? GetEnumValue<T>(Type type, T value) where T : Enum {
            string name = Enum.GetName(typeof(T), value);
            var a =  Enum.Parse(type, name);
            return a != null ? (int?)Convert.ToInt32(a) : null;
        }

        static bool TrySetEnumFlag<T>(Type type, T value, ref int flags) where T : Enum {
            var enumValue = GetEnumValue(type, value);
            if (enumValue.HasValue)
            {
                flags |= enumValue.Value;
            }
            return enumValue.HasValue;
        }

        private static object ActLike(object originalDynamic, Type targetType, params Type[] otherInterfaces)
        {
            Type tContext;
            bool tDummy;
            originalDynamic = originalDynamic.GetTargetContext(out tContext, out tDummy);
            tContext = tContext.FixContext();

            var tProxy = BuildProxy.BuildType(tContext, targetType, otherInterfaces);

            return InitializeProxy(tProxy, originalDynamic, new[] { targetType }.Concat(otherInterfaces));
        }

        internal static object InitializeProxy(Type proxytype, object original, IEnumerable<Type> interfaces = null, IDictionary<string, Type> propertySpec = null)
        {
            var tProxy = (IActLikeProxyInitialize)Activator.CreateInstance(proxytype);
            tProxy.Initialize(original, interfaces, propertySpec);
            return tProxy;
        }
    }

    [Flags]
    public enum ExtensionManagerServiceMode
    {
        Default = 0,
        SafeMode = 1,
        DoNotLoadUserExtensions = 2,
        ScanAlways = 4,
        DoNotScanSdkDirectories = 8
    }

    public interface VSILogger
    {
        void LogInformation(string message);

        void LogInformation(string message, string path);

        void LogWarning(string message);

        void LogWarning(string message, string path);

        void LogError(string message);

        void LogError(string errorMessage, string path);
    }


    public class Logger : VSILogger
    {
        public int Level { get; set; } = 1;

        public Logger(bool verboseLog)
        {
            Level = verboseLog ? 2 : 1;
        }

        public void LogError(string msg)
        {
            Console.WriteLine($"Error: {msg}");
        }

        public void LogError(string msg, string path)
        {
            Console.WriteLine($"Error: {msg}, path = {path}");
        }

        public void LogInformation(string msg)
        {
            if (Level < 2)
            {
                return;
            }
            Console.WriteLine($"Info: {msg}");
        }

        public void LogInformation(string msg, string path)
        {
            if (Level < 2)
            {
                return;
            }
            Console.WriteLine($"Info: {msg}, path = {path}");
        }

        public void LogWarning(string msg)
        {
            if (Level < 1)
            {
                return;
            }
            Console.WriteLine($"Warning: {msg}");
        }

        public void LogWarning(string msg, string path)
        {
            if (Level < 1)
            {
                return;
            }
            Console.WriteLine($"Warning: {msg}, path = {path}");
        }
    }
}
