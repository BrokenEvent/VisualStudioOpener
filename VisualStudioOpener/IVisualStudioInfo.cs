using System;

namespace BrokenEvent.VisualStudioOpener
{
  /// <summary>
  /// Visual Studio Info
  /// </summary>
  public interface IVisualStudioInfo: IDisposable
  {
    /// <summary>
    /// Open file in this visual studio
    /// </summary>
    /// <param name="filename">Path to file to open</param>
    /// <param name="lineNumber">Line number</param>
    void OpenFile(string filename, int lineNumber = -1);

    /// <summary>
    /// Visual Studio description
    /// </summary>
    string Description { get; }

    /// <summary>
    /// Path to the exectuable
    /// </summary>
    string Path { get; }

    /// <summary>
    /// Visual Studio version
    /// </summary>
    Version Version { get; }
  }
}
