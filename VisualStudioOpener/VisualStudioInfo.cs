using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

using BrokenEvent.VisualStudioOpener.DarkMagic;

using EnvDTE;

namespace BrokenEvent.VisualStudioOpener
{
  internal class VisualStudioInfo: IVisualStudioInfo
  {
    private Version version;
    private string description;
    private string path;
    private DTE dte;

    protected VisualStudioInfo(Version version, string description, string path)
    {
      this.version = version;
      this.path = path;
      this.description = description;
    }

    public VisualStudioInfo(Version version, string path):
      this(version, FileVersionInfo.GetVersionInfo(path).FileDescription, path)
    {
    }

    protected virtual DTE GetDTE()
    {
      string vsVersion = string.Format("VisualStudio.DTE.{0}.0", version.Major);
      Type vsType = Type.GetTypeFromProgID(vsVersion);
      DTE dte;

      try
      {
        dte = (DTE)Marshal.GetActiveObject(vsVersion);
      }
      catch (COMException e)
      {
        if (e.Message.Contains("0x800401E3"))
          return (DTE)Activator.CreateInstance(vsType);
        throw;
      }

      return dte;
    }

    public void OpenFile(string filename, int lineNumber)
    {
      if (dte == null)
        dte = GetDTE();
      if (dte == null)
        throw new InvalidOperationException("Unable to obtain DTE for " + description);

      // Register the IOleMessageFilter to handle any threading errors.
      MessageFilter.Register();

      // Display the Visual Studio IDE.
      dte.MainWindow.Activate();
      dte.UserControl = true;

      Window window = dte.ItemOperations.OpenFile(filename, Constants.vsViewKindTextView);
      window.SetFocus();
      if (lineNumber != -1)
        ((TextSelection)dte.ActiveDocument.Selection).GotoLine(lineNumber, true);

      // Turn off the filter
      MessageFilter.Revoke();
    }

    public string Description
    {
      get { return description; }
    }

    public string Path
    {
      get { return path; }
    }

    public override string ToString()
    {
      return description;
    }

    public void Dispose()
    {
      if (dte != null)
        Marshal.ReleaseComObject(dte);
      dte = null;
    }

    public Version Version
    {
      get { return version; }
    }
  }
}
