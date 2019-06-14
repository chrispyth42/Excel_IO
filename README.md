# Excel_IO
An repo containing a powershell module that I've written for myself, to make interacting with the Microsoft Excel COM object much easier and manageable.

### module.psm1
contains the module itself, with an array of methods that interact and manage the MS Excel object, and makes scripting it far easier

### demoscript.ps1
Is a powershell script that utilizes 'module.psm1', and demonstrates the functionality of the methods that it contains. It generates the file 'demofile.xlsx' in the current directory, complete with a populated table on Sheet 1, and a checkerboard pattern on Sheet 2

### notes.txt
Contains 5 lines of powershell that allow you to easily fire up a new Excel object in powershell, to demo things before placing them in the actual script; as well as any other discoveries I've made

(Fun fact, Excel indexes start at 1 instead of 0)
