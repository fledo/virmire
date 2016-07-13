## What is virmire?
**Virmire** is a PowerShell script which allows users to configure global `ctrl`+`alt`+`"key"` hotkeys to open folder or folders. `Ctrl`+`alt`+`F` could open **Firefox** and `ctrl`+`alt`+`c` could open the **command prompt**. This can be achieved ~~from commandline or~~ via a GUI. The GUI also gives an overview to which actions are bound to which keys.

## How?
Virmire uses [PsEventingPlus](http://pseventing.codeplex.com/releases/view/66587) to register global hotkeys. The module is not included and needs to be downloaded separately. Registered hotkeys are by default saved in `%appdata%` and used by a background process which listens to the activated hotkeys.

## Todo
 - Option to autostart the background listener process when a users signs in.
 - Provide help on how to use the GUI + link to github.
 
## Know issues and dependencies 
 - Virmire requires [PsEventingPlus module](http://pseventing.codeplex.com/releases/view/66587) which is released under the following [license](http://pseventing.codeplex.com/license).
 - Global hotkeys will not work if they were registered in an elevated instance of PowerShell and you have UAC enabled. Use an unelevated instance of PowerShell 

## License

The MIT License (MIT)

Copyright (c) 2016 Teddy Wong, Fred Uggla

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
