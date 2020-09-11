using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

using Microsoft.VisualStudio.Setup.Configuration;
using Microsoft.Win32;

namespace BrokenEvent.VisualStudioOpener
{
  /// <summary>
  /// The Visual Studio Detector entry point.
  /// </summary>
  public static class VisualStudioDetector
  {
    private static List<IVisualStudioInfo> visualStudioCache;

    private static string CheckOldVisualStudio(RegistryKey key)
    {
      using (RegistryKey subKey = key.OpenSubKey(@"Setup\VS"))
      {
        if (subKey == null)
          return null;

        return subKey.GetValue("EnvironmentPath") as string;
      }
    }

    private static IEnumerable<IVisualStudioInfo> DetectOldVisualStudios()
    {
      using (RegistryKey registryHive = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32))
      using (RegistryKey key = registryHive.OpenSubKey(@"SOFTWARE\Microsoft\VisualStudio", false))
      {
        foreach (string subKeyName in key.GetSubKeyNames())
        {
          Version version;
          if (!Version.TryParse(subKeyName, out version))
            continue;

          using (RegistryKey subKey = key.OpenSubKey(subKeyName, false))
          {
            if (subKey == null)
              continue;
            string path = CheckOldVisualStudio(subKey);
            if (path == null || !File.Exists(path))
              continue;

            yield return new VisualStudioInfo(version, path);
          }
        }
      }
    }

    private static IEnumerable<IVisualStudioInfo> DetectNewVisualStudios()
    {
      SetupConfiguration configuration;
      try
      {
        configuration = new SetupConfiguration();
      }
      catch (COMException ex)
      {
        // class not registered, no VS2017+ installations
        if ((uint)ex.HResult == 0x80040154)
          yield break;

        throw;
      }
      IEnumSetupInstances e = configuration.EnumAllInstances();

      int fetched;
      ISetupInstance[] instances = new ISetupInstance[1];
      do
      {
        e.Next(1, instances, out fetched);
        if (fetched <= 0)
          continue;

        ISetupInstance2 instance2 = (ISetupInstance2)instances[0];
        string filename = Path.Combine(instance2.GetInstallationPath(), @"Common7\IDE\devenv.exe");
        if (File.Exists(filename))
          yield return new VisualStudio2017Info(Version.Parse(instance2.GetInstallationVersion()),  instance2.GetDisplayName(), filename);
      }
      while (fetched > 0);
    }

    /// <summary>
    /// Detects all existing Visual Studio installations.
    /// No caching provided, every call will invoke new detection procedure.
    /// </summary>
    /// <returns>Visual Studio informations enumeration</returns>
    public static IEnumerable<IVisualStudioInfo> DetectVisualStudios()
    {
      yield return new NotepadInfo();

      IVisualStudioInfo vsCodeInfo = VsCodeInfo.Probe();
      if (vsCodeInfo != null)
        yield return vsCodeInfo;

      foreach (IVisualStudioInfo info in DetectOldVisualStudios())
        yield return info;

      foreach (IVisualStudioInfo info in DetectNewVisualStudios())
        yield return info;
    }

    /// <summary>
    /// Gets all existing Visual Studio installations. Caches results.
    /// </summary>
    /// <param name="rebuildCache">True to rebuild Visual Studios cache</param>
    /// <returns>Visual Studio informations enumeration</returns>
    public static IEnumerable<IVisualStudioInfo> GetVisualStudios(bool rebuildCache = false)
    {
      if (visualStudioCache == null || rebuildCache)
      {
        Dispose();
        visualStudioCache = new List<IVisualStudioInfo>(DetectVisualStudios());
      }

      return visualStudioCache;
    }

    /// <summary>
    /// Searches for existing Visual Studio installation by its description.
    /// </summary>
    /// <param name="description">The Visual Studio description</param>
    /// <returns>Visual Studio information or null if not found for this description.</returns>
    /// <remarks>If this call is first, the <see cref="GetVisualStudios"/> will be invoked to obtain list.</remarks>
    public static IVisualStudioInfo GetVisualStudioInfo(string description)
    {
      if (visualStudioCache == null)
        visualStudioCache = new List<IVisualStudioInfo>(DetectVisualStudios());

      foreach (IVisualStudioInfo info in visualStudioCache)
        if (info.Description == description)
          return info;

      return null;
    }

    /// <summary>
    /// Gets the Visual Studio with highest version available.
    /// </summary>
    /// <returns>Visual Studio info. If no Visual Studio installed, the Noteped info will be returned.</returns>
    /// <remarks>If this call is first, the <see cref="GetVisualStudios"/> will be invoked to obtain list.</remarks>
    public static IVisualStudioInfo GetHighestVisualStudio()
    {
      if (visualStudioCache == null)
        visualStudioCache = new List<IVisualStudioInfo>(DetectVisualStudios());

      IVisualStudioInfo result = null;

      foreach (IVisualStudioInfo info in visualStudioCache)
        if (result == null || info.Version > result.Version)
          result = info;

      return result;
    }

    /// <summary>
    /// Disponses all created COM objects. Unnecessary call.
    /// </summary>
    public static void Dispose()
    {
      if (visualStudioCache != null)
        foreach (IVisualStudioInfo info in visualStudioCache)
          info.Dispose();
    }
  }
}
