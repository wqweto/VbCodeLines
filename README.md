## VbCodeLines
(Legacy project w/ first commit 2004-11-12)

Allows to add line numbers to VB6 sources before compiling a .vbp project. Each line number put in the source files matches the `Ln` info as shown in VB IDE for this source line, so `Erl` function at run-time returns correct source line ready to be logged.

### Usage

c:> VbCodeLines MyProject.vbp

Note: This modifies source files **in-place** so make sure to perform only on a working copy.