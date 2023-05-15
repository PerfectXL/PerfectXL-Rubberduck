PerfectXL-Rubberduck
====================

This is a port of Rubberduck to .NET Standard. [Rubberduck](http://rubberduckvba.com/) is an open-source project that parses and inspects VBA code. It is licensed under the GPLv3 license and the source code is available at [https://github.com/rubberduck-vba/Rubberduck](https://github.com/rubberduck-vba/Rubberduck).

Similarly, PerfectXL-Rubberduck is open source, with its source code publicly available at https://github.com/PerfectXL/PerfectXL-Rubberduck.

The codebase of PerfectXL-Rubberduck is built from the ground up, starting with an empty solution. Files from Rubberduck have been imported project by project, each with a new project file targeting `netstandard2.0`. Our changes are: add the correct dependencies for .NET Standard, disable certain parts that use specific COM or Windows APIs, exclude unnecessary parts such as front-end and UI elements, and make small code changes to pass all unit tests. Our changes result in a successful build of every project in .NET Standard. The unit tests have been transferred as far as possible and are working. The namespaces and project names of Rubberduck are unchanged. See the commit history for all details of this port.

The end result is a class library `PerfectXL.Rubberduck` targeting `netstandard2.0`. All Rubberduck assemblies are merged into one file `PerfectXL.Rubberduck.dll` using ILRepack. A package is generated that is ready to be published at NuGet.

License
-------

Copyright 2023 Infotron B.V.

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program. If not, see [http://www.gnu.org/licenses/](http://www.gnu.org/licenses/).
