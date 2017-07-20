PerfectXL-Rubberduck
====================

PerfectXL VbaCodeAnalyzer uses code of Rubberduck. [Rubberduck](http://rubberduckvba.com/) is an open-source project that analyzes and inspects [VBA](https://en.wikipedia.org/wiki/Visual_Basic_for_Applications) code. It is licensed under the GPLv3 license and the source code is available at [https://github.com/rubberduck-vba/Rubberduck](https://github.com/rubberduck-vba/Rubberduck).

Similarly, PerfectXL VbaCodeAnalyzer is open source, with its source code publicly available at [https://github.com/PerfectXL/PerfectXL-Rubberduck](https://github.com/PerfectXL/PerfectXL-Rubberduck).

Our modifications to the original software are available in a forked repository at [https://github.com/PerfectXL/Rubberduck](https://github.com/PerfectXL/Rubberduck). The main change is that we eliminated EasyHook because we don't require any COM calls.

The analyzer can be installed as a Windows service. It provides an HTTP API where you can submit VBA code and retrieve inspection results.

Packaging the software
----------------------

If you want to install this program yourself, you can find a batch script to build and package the software at [https://github.com/PerfectXL/PerfectXL-Rubberduck/blob/develop/deploy/build-and-package. Cmd] (https://github.com/PerfectXL/PerfectXL-Rubberduck/blob/develop/deploy/build-and-package.cmd). This builds an setup program that you can run on a Windows device to install the software.

License
-------

Copyright 2017 Infotron B.V.

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see [http://www.gnu.org/licenses/](http://www.gnu.org/licenses/).

------------------------------------------------------------------------
