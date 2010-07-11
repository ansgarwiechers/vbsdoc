Motivation
----------
API documentation is nice, and being able to generate it from the code
is even nicer. However, unlike Perl, Python, Java, or several other
languages, VBScript doesn't have a feature or tool that supports this.
Which kinda sucks.

I tried VBDOX [1], but didn't find usability or results too convincing.
I also tried doxygen [2] by adapting Basti Grembowietz' Visual Basic
doxygen filter. However, doxygen does a lot of things I don't actually
need, and I didn't manage to make it do some of the things I did want.
Thus I ended up writing my own VBScript documentation generator.


Copyright
---------
See COPYING.txt.


Requirements
------------
VBSdoc uses my Logger class [3] for displaying messages.


Doc-comments
------------
VBSdoc comments begin with the string '! (apostrophe followed by an
exclamation mark) and must be placed either before the element they
refer to (without blank lines between doc-comment and code) or at the
end of the code line. Examples:

- Valid:
    '! Some procedure.
    '! @param bar Input value
    Sub Foo(bar)

- Valid:
    Const Bar = 42  '! Some constant.
                    '! @see <http://www.example.org/>

- Not valid (blank line between doc-comment and code):
    '! Some procedure.
    '! @param bar Input value

    Function Foo

- Not valid (regular comment between doc-comment and code):
    '! Some procedure.
    '! @param bar Input value
    ' other comment
    Function Foo


Tags
----
Supported tags are:

@author   Name and/or mail address of the author. Optional, multiple
          tags per documented element are allowed.
@brief    Summary information. If this tag is omitted, but @details is
          defined, summary information is auto-generated from the first
          sentence or line of the detail information. Should appear at
          most once per documented element.
@date     The release date. Valid for files and classes, otherwise
          ignored. Optional.
@details
@param
@raise
@return
@see
@todo
@version


References
----------
[1] http://vbdox.sourceforge.net/
[2] http://www.doxygen.org/
[3] http://www.planetcobalt.net/download/LoggerClass-1.0.zip
