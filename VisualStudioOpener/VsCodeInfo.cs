using System;

using Microsoft.Win32;

namespace BrokenEvent.VisualStudioOpener
{
  internal class VsCodeInfo: IVisualStudioInfo
  {
    private static VsCodeInfo Probe(RegistryKey root, string uninstallKey)
    {
      // the IS key may change in future versions...
      RegistryKey key = root.OpenSubKey($@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{uninstallKey}");

      if (key == null)
        return null;
      try
      {

        object o = key.GetValue("DisplayIcon");
        if (!(o is string s))
          return null;

        return new VsCodeInfo(s);
      }
      catch
      {
        return null;
      }
      finally
      {
        key.Dispose();
      }
    }

    public static VsCodeInfo Probe()
    {
      VsCodeInfo info;

      info = Probe(Registry.ClassesRoot, "{EA457B21-F73E-494C-ACAB-524FDE069978}");
      if (info != null)
        return info;

      info = Probe(Registry.ClassesRoot, "{EA457B21-F73E-494C-ACAB-524FDE069978}_is1");
      if (info != null)
        return info;

      info = Probe(Registry.LocalMachine, "{EA457B21-F73E-494C-ACAB-524FDE069978}");
      if (info != null)
        return info;

      info = Probe(Registry.LocalMachine, "{EA457B21-F73E-494C-ACAB-524FDE069978}_is1");
      if (info != null)
        return info;

      return null;
    }

    private VsCodeInfo(string path)
    {
      Path = path;
    }

    public void OpenFile(string filename, int lineNumber = -1)
    {
      if (lineNumber != -1)
      {
        System.Diagnostics.Process.Start(Path, $"-g {filename}:{lineNumber}");
        return;
      }

      System.Diagnostics.Process.Start(Path, filename);
    }

    public string Description
    {
      get { return "Visual Studio Code"; }
    }

    public string Path { get; }

    public Version Version
    {
      get { return new Version(0, 0); }
    }

    public void Dispose() { }

    public override string ToString()
    {
      return "VsCode";
    }
  }
}
