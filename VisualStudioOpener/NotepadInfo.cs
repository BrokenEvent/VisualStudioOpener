using System;

namespace BrokenEvent.VisualStudioOpener
{
  internal class NotepadInfo: IVisualStudioInfo
  {
    public void Dispose() {}

    public void OpenFile(string filename, int lineNumber)
    {
      System.Diagnostics.Process.Start("notepad.exe", filename);
    }

    public string Description
    {
      get { return "Notepad"; }
    }

    public string Path
    {
      get { return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Notepad.exe"); }
    }

    public Version Version
    {
      get { return Version.Parse("0.0"); }
    }
  }
}
