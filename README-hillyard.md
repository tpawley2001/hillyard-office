# Hillyard Office

Hillyard Office is LibreOffice 26.8 changed by Hillyard Tech to look and work
more like Microsoft Office: Word, Excel and PowerPoint ribbons, a File
Backstage, Office keyboard shortcuts and templates, LAMBDA and newer Excel
functions, Batch Replace, an Office compatibility check, metric-compatible
fonts and many Word/Excel import fixes.

This repository is the source code for the Hillyard Office builds.

## History

- The first commit is LibreOffice's `libreoffice-26.8.0.3` tag, unchanged
  (upstream: https://git.libreoffice.org/core).
- Every later commit is a Hillyard Office change. A build's number is the
  number of commits after the first one, and each commit's `Source-Commit:`
  line is the build id the installer reports when it checks for updates.

## Licences

LibreOffice's files, including the ones we changed, keep their upstream
licences (mostly MPL-2.0; see `COPYING*`). Files Hillyard Tech wrote outright
are MIT. `LICENSE-hillyard.txt` lists which is which.

## Building

Build it the way you build LibreOffice:
https://wiki.documentfoundation.org/Development/BuildingOnLinux and
https://wiki.documentfoundation.org/Development/BuildingOnWindows

The update feed address in `instsetoo_native/util/openoffice.lst.in` is a
placeholder here. Replace it with your own feed, or switch updates off.
