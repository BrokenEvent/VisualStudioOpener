[![Build Status](https://img.shields.io/appveyor/ci/BrokenEvent/visualstudioopener/master.svg?style=flat-square)](https://ci.appveyor.com/project/BrokenEvent/visualstudioopener)
[![GitHub License](https://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square)](https://raw.githubusercontent.com/BrokenEvent/VisualStudioOpener/master/LICENSE)
[![Github Releases](https://img.shields.io/github/downloads/BrokenEvent/VisualStudioOpener/total.svg?style=flat-square)](https://github.com/BrokenEvent/VisualStudioOpener/releases)

# VisualStudioOpener
The library is used to provide "Open in Visual Studio" functionality. Old style DTE COM engine is used.

## Features

* Visual Studio installations detection.
* OpenFile(string filename, int lineNumber) for any Visual Studio.
* All versions from VS 2003 to 2017 are supported.
* Parallel VS2017 installations are supported.
* VS instance reusing.
* All COM-related black magic (errors, instance waiting, etc.) is handled.
* Notepad as fallback, if no VS found.

## Usage

The entry point is **BrokenEvent.VisualStudioOpener.VisualStudioDetector** static class. You can use it to detect all existing Visual Studio installations:

      foreach (IVisualStudioInfo info in VisualStudioDetector.GetVisualStudios())
      {
        // code
      }

You can also get the latest VS or VS by its name (called *Description*). Each **IVisualStudioInfo** have

    void OpenFile(string filename, int lineNumber = -1);

method to open files.

## Future Plans
Unity changelog tells, that they abandoned DTE and used something another to open files in VS. DTE is slow and unstable COM thing and if someone has a clue how to do it in another way, you are welcome.

## Credits
(C) 2017, Broken Event. [brokenevent.com](http://brokenevent.com)
