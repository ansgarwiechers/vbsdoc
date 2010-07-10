Motivation
----------
API documentation is nice, and being able to generate it from the code
is even nicer. However, unlike Perl (perldoc) or Python (docstrings),
VBScript doesn't have a feature like that. Which kinda sucks.

I was unable to find any commercial or free tool that officially
supports VBScript, so I tried to adapt Basti Grembowietz' doxygen
filter for Visual Basic. However, doxygen does a lot of things I don't
actually need. Also I didn't manage to make it do some of the things I
did want. So I ended up writing my own documentation generator. Which
shamelessly rips off doxygen as well as javadoc.


Doc-comments and Tags
---------------------
VBSdoc comments begin with the string '! (apostrophe followed by an
exclamation mark). Several tags are supported:

  @author
  @brief
  @date
  @details
  @file
  @param
  @raise
  @return
  @see
  @todo
  @version
