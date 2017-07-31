What is Virmire?
================

Virmire allows users to configure global hotkeys prefixed `ctrl`+`alt`. These can be used to to open files or folders.  `ctrl`+`alt`+`F` could open **Firefox** and `ctrl`+`alt`+`c` could open the **command prompt**.

![Virmire GUI](https://raw.githubusercontent.com/fledo/virmire/4bed844827a435df1c4ffe100bd674530936172d/vimrire.png)

Usage
-----

 - Right click buttons to choouse a target file or right click to choose a target folder.
 - Save the configuration and start the background listener. The configuration interface can now be closed.
 - Right/Left click to clear previous configuraiton from a single button.
 - Click save to store your settings
 - Start Listener and test hotkeys

How?
----

Virmire is written in PowerShell (and thus only works on Windows) and relies on the module [PsEventingPlus](http://pseventing.codeplex.com/releases/view/66587) to register global hotkeys. Settings are by default saved in `%appdata%\virmire` and are used to create a separate background process which listens for the configured hotkeys.
 
Dependencies
------------

 - PowerShell 3+
 - [PsEventingPlus module](http://pseventing.codeplex.com/releases/view/66587) which is released under the following [license](http://pseventing.codeplex.com/license).

Known issues
------------

 - Global hotkeys will not work if they were registered in an elevated instance of PowerShell and you have UAC enabled. **Do not use an elevated instance of PowerShell**.

License
-------

```
The MIT License (MIT)

Copyright (c) 2014-2017 Teddy Wong, Fred Uggla

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
