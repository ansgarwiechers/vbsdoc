'! Script for automatic generation of API documentation from special
'! comments in VBScripts.
'!
'! @author  Ansgar Wiechers <ansgar.wiechers@planetcobalt.net>
'! @date    2010-07-19
'! @version 0.9b

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

'! @todo Add tag @mainpage for global documentation on the global index page?
'! @todo Add HTMLHelp generation?
'! @todo Add grouping option for doc comments? (like doxygen's @{ ... @})

Option Explicit

Import "LoggerClass.vbs"

' Some symbolic constants for internal use.
Private Const ForReading = 1
Private Const ForWriting = 2

Private Const vbReplaceAll = -1

Private Const Extension = "vbs"

Private Const StylesheetName = "vbsdoc.css"
Private Const TextFont       = "Verdana, Arial, helvetica, sans-serif"
Private Const CodeFont       = "Lucida Console, Courier New, Courier, monospace"
Private Const BaseFontSize   = "14px"

' Initialize global objects.
Private fso : Set fso = CreateObject("Scripting.FileSystemObject")
Private sh  : Set sh = CreateObject("WScript.Shell")
Private log : Set log = New Logger

'! Match line-continuations.
Private reLineCont : Set reLineCont = CompileRegExp("[ \t]+_\n[ \t]*", True, True)
'! Match End-of-Line doc comments.
Private reEOLComment : Set reEOLComment = CompileRegExp("(^|\n)([ \t]*[^' \t\n].*)('![ \t]*.*(\n[ \t]*'!.*)*)", True, True)
'! Match @todo-tagged doc comments.
Private reTodo     : Set reTodo = CompileRegExp("'![ \t]*@todo[ \t]*(.*\n([ \t]*'!([ \t]*[^@\s].*|\s*)\n)*)", True, True)
'! Match class implementations and prepended doc comments.
Private reClass    : Set reClass = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*Class[ \t]+(\w+)([\s\S]*?)End[ \t]+Class", True, True)
'! Match constructor implementations and prepended doc comments.
Private reCtor     : Set reCtor = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*((Public|Private)[ \t]+)?Sub[ \t]+(Class_Initialize)[ \t]*(\(\))?[\s\S]*?End[ \t]+Sub", True, True)
'! Match destructor implementations and prepended doc comments.
Private reDtor     : Set reDtor = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*((Public|Private)[ \t]+)?Sub[ \t]+(Class_Terminate)[ \t]*(\(\))?[\s\S]*?End[ \t]+Sub", True, True)
'! Match implementations of methods/procedures as well as prepended
'! doc comments.
Private reMethod   : Set reMethod = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*((Public|Private)[ \t]+)?(Function|Sub)[ \t]+(\w+)[ \t]*(\([\w\t ,]*\))?[\s\S]*?End[ \t]+\6", True, True)
'! Match property implementations and prepended doc comments.
Private reProperty : Set reProperty = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*((Public|Private)[ \t]+)?Property[ \t]+(Get|Let|Set)[ \t]+(\w+)[ \t]*(\([\w\t ]*\))?[\s\S]*?End[ \t]+Property", True, True)
'! Match definitions of constants and prepended doc comments.
Private reConst    : Set reConst = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*((Public|Private)[ \t]+)?Const[ \t]+(\w+)[ \t]*=[ \t]*(.*)", True, True)
'! Match variable declarations and prepended doc comments. Allow for combined
'! declaration:definition as well as multiple declarations of variables on
'! one line, e.g.:
'!   - Dim foo : foo = 42
'!   - Dim foo, bar, baz
Private reVar      : Set reVar = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*(Public|Private|Dim|ReDim)[ \t]+(((\w+)([ \t]*\(\))?)[ \t]*(:[ \t]*(Set[ \t]+)?\7[ \t]*=.*|(,[ \t]*\w+[ \t]*(\(\))?)*))", True, True)

'! Dictionary listing the tags that VBSdoc accepts.
Private isValidTag : Set isValidTag = CreateObject("Scripting.Dictionary")
	isValidTag.Add "@author", True
	isValidTag.Add "@brief", True
	isValidTag.Add "@date", True
	isValidTag.Add "@details", True
	isValidTag.Add "@param", True
	isValidTag.Add "@raise", True
	isValidTag.Add "@return", True
	isValidTag.Add "@see", True
	isValidTag.Add "@todo", True
	isValidTag.Add "@version", True

Private beQuiet, projectName, docRoot, indexFile

Main WScript.Arguments


'! The starting point. Evaluates commandline arguments, initializes
'! global variables and starts the documentation generation.
'!
'! @param  args   The list of arguments passed to the script.
Public Sub Main(args)
	Dim includePrivate, srcRoot

	' initialize global variables with default values
	beQuiet = False
	includePrivate = False
	projectName = ""
	log.Debug = True

	' evaluate commandline arguments
	With args
		If args.Named.Exists("?") Then PrintUsage
		If args.Named.Exists("a") Then includePrivate = True
		If args.Named.Exists("q") Then beQuiet = True
		If args.Named.Exists("p") Then projectName = args.Named("p")

		If args.Named.Exists("s") Then
			srcRoot = args.Named("s")
		Else
			PrintUsage
		End If

		If args.Named.Exists("d") Then
			docRoot = args.Named("d")
		Else
			PrintUsage
		End If
	End With

	' start documentation generation
	CreateDirectory docRoot
	If fso.FileExists(srcRoot) Then
		GenFileDoc srcRoot, docRoot, StylesheetName, includePrivate
		CreateStylesheet fso.BuildPath(fso.BuildPath(docRoot, fso.GetBaseName(srcRoot)), StylesheetName)
	Else
		Set indexFile = fso.OpenTextFile(fso.BuildPath(docRoot, "index.html"), ForWriting, True)
		WriteHeader indexFile, "Main Page", StylesheetName
		If projectName <> "" Then indexFile.WriteLine "<h1>" & projectName & "</h1>"

		GenDoc fso.GetFolder(srcRoot), docRoot, "../" & StylesheetName, includePrivate

		WriteFooter indexFile
		indexFile.Close

		CreateStylesheet fso.BuildPath(docRoot, StylesheetName)
	End If

	WScript.Quit(0)
End Sub

'! Traverse all subdirecotries of the given srcDir and generate documentation
'! for all VBS files. If includePrivate is set to True, then documentation for
'! private elements is generated as well, otherwise only public elements are
'! included in the documentation. The procedure also creates an index file with
'! links to the documentation pages of each VBS file.
'!
'! @param  srcDir         Directory containing the source files for the
'!                        documentation generation.
'! @param  docDir         Absolute or relative path to the directory where the
'!                        documentation for the files in srcDir should be
'!                        generated.
'! @param  stylesheet     Relative path to the stylesheet for the HTML files.
'! @param  includePrivate Include documentation for private elements.
Public Sub GenDoc(srcDir, docDir, stylesheet, includePrivate)
	Dim f, dir, docName

	For Each f In srcDir.Files
		If fso.GetExtensionName(f.Name) = Extension Then
			docName = GenFileDoc(fso.BuildPath(f.ParentFolder, f.Name), docDir, stylesheet, includePrivate)
			If Not IsNull(docName) Then
				' Write (global) index file entry.
				docName = Replace(docName, "\", "/")
				indexFile.WriteLine "<p><a href=""" & docName & """>" _
					& Replace(docName, "/index.html", ".vbs") & "</a></p>"
			End If
		End If
	Next

	For Each dir In srcDir.SubFolders
		GenDoc dir, fso.BuildPath(docDir, dir.Name), "../" & stylesheet, includePrivate
	Next
End Sub

'! Generate the documentation for a given file. The documentation files are
'! created in the given documentation directory. If includePrivate is set to
'! True, then documentation for private elements is generated as well,
'! otherwise only public elements are included in the documentation.
'!
'! @param  filename       Name of the source file for documentation generation.
'! @param  docDir         Absolute or relative path to the directory where the
'!                        documentation for the given file should be generated.
'! @param  stylesheet     Relative path to the stylesheet for the HTML files.
'! @param  includePrivate Include documentation for private elements.
'! @return The name and path of the generated documentation file.
Public Function GenFileDoc(filename, docDir, stylesheet, includePrivate)
	Dim outDir, inFile, outFile, content, m
	Dim todoList, classInfo, classes, entry, line, tags
	Dim fileDoc, procDoc, constDoc, varDoc, reDocComment, docName

	GenFileDoc = Null
	If fso.GetFile(filename).Size = 0 Or Not fso.FileExists(filename) Then Exit Function ' nothing to do

	log.LogInfo "Generating documentation for " & filename & " ..."

	outDir = fso.BuildPath(docDir, fso.GetBaseName(filename))
	CreateDirectory(outDir)

	Set inFile = fso.OpenTextFile(filename, ForReading)
	content = inFile.ReadAll
	inFile.Close

	' Convert all linebreaks to LF, otherwise regular expressions might produce
	' strings with a trailing CR.
	content = Replace(Replace(content, vbNewLine, vbLf), vbCr, vbLf)

	' Join continued lines.
	content = reLineCont.Replace(content, " ")

	' Move End-of-Line doc comments to front.
	For Each m In reEOLComment.Execute(content)
		With m.SubMatches
			' Move doc comment to front only if the substring left of the doc comment
			' signifier ('!) contains an even number of double quotes. Otherwise the
			' signifier is inside a string, i.e. does not start an actual doc comment.
			If (Len(.Item(1)) - Len(Replace(.Item(1), """", ""))) Mod 2 = 0 Then
				content = Replace(content, m, vbLf & .Item(2) & vbLf & .Item(1), 1, 1)
			End If
		End With
	Next

	todoList = GenTodoList(content)

	classInfo = ""
	classes = GenClassDoc(content, outDir, stylesheet, includePrivate)
	For Each entry In classes
		If classInfo <> "" Then classInfo = classInfo & "<hr/>" & vbNewLine
		classInfo = classInfo & "<p><span class=""name""><a href=""" & entry(0) & ".html"">" & entry(0) & "</a></span></p>" & vbNewLine
		If entry(1) <> "" Then classInfo = classInfo & "<p class=""description"">" & entry(1) & "</p>" & vbNewLine
	Next

	Set procDoc = GenMethodDoc(content, includePrivate)
	Set constDoc = GenConstDoc(content, includePrivate)
	Set varDoc = GenVariableDoc(content, includePrivate)

	' process file-global doc comments
	fileDoc = ""
	Set reDocComment = CompileRegExp("^[ \t]*('!.*)", True, True)
	For Each line In Split(content, vbLf)
		If reDocComment.Test(line) Then fileDoc = fileDoc & reDocComment.Replace(line, "$1") & vbLf
	Next
	Set tags = ProcessComments(fileDoc)

	CheckRemainingCode content

	docName = fso.BuildPath(outDir, "index.html")
	Set outFile = fso.OpenTextFile(docName, ForWriting, True)

	WriteHeader outFile, fso.GetFileName(filename), stylesheet

	outFile.WriteLine "<h1>" & fso.GetFileName(filename) & "</h1>"
	outFile.Write GenDetailsInfo(tags)
	outFile.Write GenVersionInfo(tags)
	outFile.Write GenAuthorInfo(tags)
	outFile.Write GenReferencesInfo(tags)

	' Write ToDo list.
	If UBound(todoList) > -1 Then
		outFile.WriteLine "<h2>ToDo List</h2>" & vbNewLine & "<ul>"
		For Each entry In todoList
			outFile.WriteLine "  <li>" & EncodeHTMLEntities(entry) & "</li>"
		Next
		outFile.WriteLine "</ul>"
	End If

	WriteSection outFile, "Global Constant Summary", constDoc("summary")
	WriteSection outFile, "Global Variable Summary", varDoc("summary")
	WriteSection outFile, "Classes Summary", classInfo
	WriteSection outFile, "Global Procedure Summary", procDoc("summary")
	WriteSection outFile, "Global Constant Detail", constDoc("details")
	WriteSection outFile, "Global Variable Detail", varDoc("details")
	WriteSection outFile, "Global Procedure Detail", procDoc("details")

	WriteFooter outFile
	outFile.Close

	' Return documentation file name without the leading docRoot directory.
	GenFileDoc = Replace(docName, docRoot & "\", "", 1, 1)
End Function

'! Generate a todo list from the @todo tags in the code.
'!
'! @param  code   code fragment to check for @todo items.
'! @return A list with the todo items from the code.
Private Function GenTodoList(ByRef code)
	Dim list, m, line

	list = Array()

	For Each m in reTodo.Execute(code)
		ReDim Preserve list(UBound(list)+1)
		list(UBound(list)) = ""
		For Each line In Split(m.SubMatches.Item(0), vbLf)
			list(UBound(list)) = Trim(list(UBound(list)) & " " _
				& Trim(Replace(Replace(line, vbTab, " "), "'!", "")))
		Next
	Next
	code = reTodo.Replace(code, "")

	GenTodoList = list
End Function

'! Generate documentation for the give class code. Write it to a file named
'! after the class in the given directory.
'!
'! @param  code           Code fragment containing the class implementation.
'! @param  dir            Directory to write the class documentation to.
'! @param  stylesheet     Relative path to the stylesheet for the HTML files.
'! @param  includePrivate Include private members/methods in the documentation.
'! @return Dictionary with summary and detail documentation.
Private Function GenClassDoc(ByRef code, ByVal dir, ByVal stylesheet, ByVal includePrivate)
	Dim m, outFile, tags, classBody
	Dim ctorDtorDoc, propertyDoc, methodDoc, fieldDoc

	Dim classes : classes = Array()

	For Each m In reClass.Execute(code)
		With m.SubMatches
			ReDim Preserve classes(UBound(classes)+1)
			Set outFile = fso.OpenTextFile(fso.BuildPath(dir, .Item(3) & ".html"), ForWriting, True)
			Set tags = ProcessComments(.Item(1))
			classes(UBound(classes)) = Array(.Item(3), "")
			If tags.Exists("@brief") Then
				classes(UBound(classes))(1) = tags("@brief")
			Else
				If Not beQuiet Then log.LogWarning "No description for class " & .Item(3) & " found."
			End If

			classBody = .Item(4)
			Set ctorDtorDoc = GenCtorDtorDoc(classBody, includePrivate)
			Set propertyDoc = GenPropertyDoc(classBody)
			Set methodDoc = GenMethodDoc(classBody, includePrivate)
			Set fieldDoc = GenVariableDoc(classBody, includePrivate)

			WriteHeader outFile, .Item(3), stylesheet

			' Write global class information.
			outFile.WriteLine "<h1>Class " & .Item(3) & "</h1>"
			outFile.Write GenDetailsInfo(tags)
			outFile.Write GenVersionInfo(tags)
			outFile.Write GenAuthorInfo(tags)
			outFile.Write GenReferencesInfo(tags)

			' Write summaries of fields, properties, ctor/dtor and methods.
			WriteSection outFile, "Field Summary", fieldDoc("summary")
			WriteSection outFile, "Property Summary", propertyDoc("summary")
			WriteSection outFile, "Constructor/Destructor Summary", ctorDtorDoc("summary")
			WriteSection outFile, "Method Summary", methodDoc("summary")

			' Write details for fields, properties, ctor/dtor and methods.
			WriteSection outFile, "Field Detail", fieldDoc("details")
			WriteSection outFile, "Property Detail", propertyDoc("details")
			WriteSection outFile, "Constructor/Destructor Detail", ctorDtorDoc("details")
			WriteSection outFile, "Method Detail", methodDoc("details")

			WriteFooter outFile

			outFile.Close
		End With
	Next
	code = reClass.Replace(code, vbLf)

	GenClassDoc = classes
End Function

'! Generate Constructor and Destructor documentation.
'!
'! @param  code           Code fragment containing constructor and/or destructor.
'! @param  includePrivate Create documentation for private constructor or
'!                        destructor as well.
'! @return Dictionary with summary and detail documentation.
Private Function GenCtorDtorDoc(ByRef code, ByVal includePrivate)
	Dim m, visibility, tags

	Dim summary : summary = Array()
	Dim details : details = Array()

	For Each m In reCtor.Execute(code)
		With m.SubMatches
			visibility = GetVisibility(.Item(4))
			If visibility = "Public" Or includePrivate Then
				Set tags = ProcessComments(.Item(1))
				CheckConsistency .Item(5), Array(), tags, "sub"
				ReDim Preserve summary(UBound(summary)+1)
				summary(UBound(summary)) = GenMethodSummary(.Item(5), Array(), tags)
				ReDim Preserve details(UBound(details)+1)
				details(UBound(details)) = GenMethodDetails(.Item(5), visibility, Array(), tags)
			End If
		End With
	Next
	code = reCtor.Replace(code, vbLf)

	For Each m In reDtor.Execute(code)
		With m.SubMatches
			visibility = GetVisibility(.Item(4))
			If visibility = "Public" Or includePrivate Then
				Set tags = ProcessComments(.Item(1))
				CheckConsistency .Item(5), Array(), tags, "sub"
				ReDim Preserve summary(UBound(summary)+1)
				summary(UBound(summary)) = GenMethodSummary(.Item(5), Array(), tags)
				ReDim Preserve details(UBound(details)+1)
				details(UBound(details)) = GenMethodDetails(.Item(5), visibility, Array(), tags)
			End If
		End With
	Next
	code = reDtor.Replace(code, vbLf)

	Set GenCtorDtorDoc = CreateObject("Scripting.Dictionary")
	GenCtorDtorDoc.Add "summary", Join(summary, "<hr/>" & vbNewLine)
	GenCtorDtorDoc.Add "details", Join(details, "<hr/>" & vbNewLine)
End Function

'! Generate documentation for class methods or global procedures.
'!
'! @param  code           Code fragment containing the methods/procedures.
'! @param  includePrivate Create documentation for private methods as well.
'! @return Dictionary with summary and detail documentation.
Private Function GenMethodDoc(ByRef code, ByVal includePrivate)
	Dim m, visibility, params, tags

	Dim summary : summary = Array()
	Dim details : details = Array()

	For Each m In reMethod.Execute(code)
		With m.SubMatches
			visibility = GetVisibility(.Item(4))
			If visibility = "Public" Or includePrivate Then
				If Len(.Item(7)) > 0 Then
					params = Mid(.Item(7), 2, Len(.Item(7))-2)  ' remove enclosing parentheses
				Else
					params = ""
				End If
				params = Replace(params, vbTab, " ")
				params = Replace(params, "ByVal ", "", 1, vbReplaceAll, vbTextCompare)
				params = Replace(params, "ByRef ", "", 1, vbReplaceAll, vbTextCompare)
				params = Split(Replace(params, " ", ""), ",")

				Set tags = ProcessComments(.Item(1))
				CheckConsistency .Item(6), params, tags, .Item(5)

				ReDim Preserve summary(UBound(summary)+1)
				summary(UBound(summary)) = GenMethodSummary(.Item(6), params, tags)
				ReDim Preserve details(UBound(details)+1)
				details(UBound(details)) = GenMethodDetails(.Item(6), visibility, params, tags)
			End If
		End With
	Next
	code = reMethod.Replace(code, vbLf)

	Set GenMethodDoc = CreateObject("Scripting.Dictionary")
	GenMethodDoc.Add "summary", Join(summary, "<hr/>" & vbNewLine)
	GenMethodDoc.Add "details", Join(details, "<hr/>" & vbNewLine)
End Function

'! Generate documentation for class properties.
'!
'! @param  code   Code fragment containing class properties.
'! @return Dictionary with summary and detail documentation.
Private Function GenPropertyDoc(ByRef code)
	Dim m, prop

	Dim summary : summary = Array()
	Dim details : details = Array()

	Dim readable  : Set readable = CreateObject("Scripting.Dictionary")
	Dim writable  : Set writable = CreateObject("Scripting.Dictionary")
	Dim readwrite : Set readwrite = CreateObject("Scripting.Dictionary")

	For Each m In reProperty.Execute(code)
		With m.SubMatches
			If GetVisibility(.Item(4)) = "Public" Then
				' Private getter and setter methods are disregarded, because even with
				' the method present, the property is not readable/writable from an
				' interface point of view.
				If LCase(.Item(5)) = "get" Then
					' getter function
					readable.Add .Item(6), .Item(1)
				Else
					' setter function
					writable.Add .Item(6), .Item(1)
				End If
			End If
		End With
	Next
	code = reProperty.Replace(code, vbLf)

	' Move properties that are both read- and writable to another dictionary.
	' That way we'll have separate dictionaries for read-only properties,
	' write-only properties, and read-write properties.
	For Each prop In readable.Keys
		If writable.Exists(prop) Then
			readwrite.Add prop, readable(prop) & writable(prop)
			readable.Remove(prop)
			writable.Remove(prop)
		End If
	Next

	GenPropertyInfo readable, "read-only", summary, details
	GenPropertyInfo writable, "write-only", summary, details
	GenPropertyInfo readwrite, "read-write", summary, details

	Set GenPropertyDoc = CreateObject("Scripting.Dictionary")
	GenPropertyDoc.Add "summary", Join(summary, "<hr/>" & vbNewLine)
	GenPropertyDoc.Add "details", Join(details, "<hr/>" & vbNewLine)
End Function

'! Generate documentation for variables.
'!
'! @param  code   Code fragment containing variable declarations.
'! @param  includePrivate Create documentation for private variables as well.
'! @return Dictionary with summary and detail documentation.
Private Function GenVariableDoc(ByRef code, ByVal includePrivate)
	Dim m, visibility, tags, vars

	Dim summary : summary = Array()
	Dim details : details = Array()

	For Each m In reVar.Execute(code)
		With m.SubMatches
			visibility = GetVisibility(.Item(3))
			If visibility = "Public" Or includePrivate Then
				Set tags = ProcessComments(.Item(1))
				vars = .Item(4)
				' If the match contains a declaration/definition combination: remove the
				' definition part.
				If Left(.Item(8), 1) = ":" Then vars = Trim(Split(vars, ":")(0))
				CheckIdentifierTags vars, tags
				vars = Split(Replace(Replace(vars, vbTab, ""), " ", ""), ",")
				ReDim Preserve summary(UBound(summary)+1)
				summary(UBound(summary)) = GenVariableSummary(vars, tags)
				ReDim Preserve details(UBound(details)+1)
				details(UBound(details)) = GenVariableDetails(vars, visibility, tags)
			End If
		End With
	Next
	code = reVar.Replace(code, vbLf)

	Set GenVariableDoc = CreateObject("Scripting.Dictionary")
	GenVariableDoc.Add "summary", Join(summary, "<hr/>" & vbNewLine)
	GenVariableDoc.Add "details", Join(details, "<hr/>" & vbNewLine)
End Function

'! Generate documentation for (global) constants.
'!
'! @param  code   Code fragment containing constant definitions.
'! @param  includePrivate Create documentation for private constants as well.
'! @return Dictionary with summary and detail documentation.
Private Function GenConstDoc(ByRef code, ByVal includePrivate)
	Dim m, visibility, tags

	Dim summary : summary = Array()
	Dim details : details = Array()

	For Each m In reConst.Execute(code)
		With m.SubMatches
			visibility = GetVisibility(.Item(4))
			If visibility = "Public" Or includePrivate Then
				Set tags = ProcessComments(.Item(1))
				CheckIdentifierTags .Item(5), tags
				ReDim Preserve summary(UBound(summary)+1)
				summary(UBound(summary)) = GenConstSummary(.Item(5), .Item(6), tags)
				ReDim Preserve details(UBound(details)+1)
				details(UBound(details)) = GenConstDetails(.Item(5), .Item(6), visibility, tags)
			End If
		End With
	Next
	code = reConst.Replace(code, "")

	Set GenConstDoc = CreateObject("Scripting.Dictionary")
	GenConstDoc.Add "summary", Join(summary, "<hr/>" & vbNewLine)
	GenConstDoc.Add "details", Join(details, "<hr/>" & vbNewLine)
End Function

'! Parse the given comment and return a dictionary with all present tags and
'! their values. A line that does not begin with a tag is appended to the value
'! of the previous tag, or to "@details" if there was no previous tag. Values
'! of tags that can appear more than once (e.g. "@param", "@see", ...) are
'! stored in Arrays.
'!
'! @param  comments  The comments to parse.
'! @return Dictionary with tag/value pairs.
Private Function ProcessComments(ByVal comments)
	Dim line, re, myMatches, m, currentTag

	Dim tags    : Set tags = CreateObject("Scripting.Dictionary")
	Dim authors : authors = Array()
	Dim params  : params  = Array()
	Dim errors  : errors  = Array()
	Dim refs    : refs    = Array()

	currentTag = Null
	For Each line in Split(comments, vbLf)
		line = Trim(Replace(line, vbTab, " "))

		Set re = CompileRegExp("'![ \t]*(@\w+)[ \t]*(.*)", True, True)
		Set myMatches = re.Execute(line)
		If myMatches.Count > 0 Then
			' line starts with a tag
			For Each m in myMatches
				currentTag = LCase(m.SubMatches(0))
				If Not isValidTag(currentTag) Then
					If Not beQuiet Then log.LogWarning "Unknown tag " & currentTag & "."
				Else
					Select Case currentTag
						Case "@author" Append authors, m.SubMatches(1)
						Case "@param" Append params, m.SubMatches(1)
						Case "@raise" Append errors, m.SubMatches(1)
						Case "@see" Append refs, m.SubMatches(1)
						Case Else
							If tags.Exists(currentTag) Then
								' Re-definiton of a tag that's supposed to be unique per
								' documentation block may be undesired.
								If Not beQuiet Then log.LogWarning "Duplicate definition of tag " & currentTag _
									& ": " & m.SubMatches(1)
								tags(currentTag) = m.SubMatches(1)
							Else
								tags.Add currentTag, m.SubMatches(1)
							End If
					End Select
				End If
			Next
		Else
			' line does not begin with a tag
			' => line must be either empty, first line of detail description, or
			'    continuation of previous line.
			line = Trim(Mid(line, 3))   ' strip '! from beginning of line
			If line = "" Then
				If currentTag = "@details" Then
					tags("@details") = tags("@details") & vbNewLine
				Else
					currentTag = Null
				End If
			Else
				' Make "@details" currentTag if currentTag is not set. Then append
				' comment text to currentTag.
				If IsNull(currentTag) Then currentTag = "@details"
				Select Case currentTag
					Case "@author" Append authors(UBound(authors)), line
					Case "@param" Append params(UBound(params)), line
					Case "@raise" Append errors(UBound(errors)), line
					Case "@see" Append refs(UBound(refs)), line
					Case Else
						If tags.Exists(currentTag) Then
							If currentTag = "@details" And Left(line, 2) = "- " Then
								' line is list element => new line
								tags(currentTag) = tags(currentTag) & vbNewLine & line
							Else
								' line is not a list element (or the continuation of a list
								' element) => append text
								tags(currentTag) = tags(currentTag) & " " & line
							End If
						Else
							tags.Add currentTag, line
						End If
				End Select
			End If
		End If
	Next

	If UBound(authors) > -1 Then tags.Add "@author", authors
	If UBound(params) > -1 Then tags.Add "@param", params
	If UBound(errors) > -1 Then tags.Add "@raise", errors
	If UBound(refs) > -1 Then tags.Add "@see", refs

	' If no short description was given, set it to the first full sentence (or
	' the first line, whichever is the shorter match) of the long description.
	' If no long description was given, set it to the short description.
	' Do nothing if neither short nor long description were given.
	If Not tags.Exists("@brief") And tags.Exists("@details") Then
		Set re = CompileRegExp("([.!?\n]).*", True, True)
		tags.Add "@brief", Replace(re.Replace(tags("@details"), "$1"), vbLf, "")
		re.Pattern = "[,;:]$"
		tags("@brief") = re.Replace(tags("@brief"), ".")
	ElseIf tags.Exists("@brief") And Not tags.Exists("@details") Then
		tags.Add "@details", tags("@brief")
	End If

	Set ProcessComments = tags
End Function

' ------------------------------------------------------------------------------
' HTML code generation
' ------------------------------------------------------------------------------

'! Write HTML headers to the given file. The headers are parametrized with
'! title and stylesheet.
'!
'! @param  outFile    Handle to a file.
'! @param  title      Title of the HTML page
'! @param  stylesheet Path to the stylesheet for the HTML page.
Private Sub WriteHeader(outFile, title, stylesheet)
	If projectName <> "" Then title = projectName & ": " & title
	outFile.WriteLine "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Frameset//EN""" & vbNewLine _
		& vbTab & """http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd"">" & vbNewLine _
		& "<html>" & vbNewLine & "<head>" & vbNewLine _
		& "<title>" & title & "</title>" & vbNewLine _
		& "<meta name=""date"" content=""" & FormatDate(Now) & """ />" & vbNewLine _
		& "<meta http-equiv=""content-type"" content=""text/html; charset=iso-8859-1"" />" & vbNewLine _
		& "<meta http-equiv=""content-language"" content=""en"" />" & vbNewLine _
		& "<link rel=""stylesheet"" type=""text/css"" href=""" & stylesheet & """ />" & vbNewLine _
		& "</head>" & vbNewLine & "<body>"
End Sub

'! Write HTML closing tags to the given file.
'!
'! @param  outFile  Handle to a file.
Private Sub WriteFooter(outFile)
	outFile.WriteLine "</body>" & vbNewLine & "</html>"
End Sub

'! Generate documentation for class properties.
'!
'! @param  properties    A dictionary with name/comment pairs of properties.
'! @param  accessibility The accessibility (read-only/write-only/read-write) of
'!                       the properties.
'! @param  summary       Reference to an array that will receive the summary
'!                       documentation for the properties.
'! @param  details       Reference to an array that will receive the detail
'!                       documentation for the properties.
Private Sub GenPropertyInfo(ByVal properties, ByVal accessibility, ByRef summary, ByRef details)
	Dim name, tags

	For Each name In properties.Keys
		Set tags = ProcessComments(properties(name))
		ReDim Preserve summary(UBound(summary)+1)
		summary(UBound(summary)) = GenPropertySummary(name, accessibility, tags)
		ReDim Preserve details(UBound(details)+1)
		details(UBound(details)) = GenPropertyDetails(name, accessibility, tags)
	Next
End Sub

'! Generate summary documentation for a class property. The documentation is
'! generated in HTML format.
'!
'! @param  name          Name of the property.
'! @param  accessibility The accessibility (read-only/write-only/read-write) of
'!                       the properties.
'! @param  tags          Dictionary with the tag/value pairs from the
'!                       documentation header.
'! @return The summary documentation in HTML format.
Private Function GenPropertySummary(ByVal name, ByVal accessibility, ByVal tags)
	Dim signature

	name = EncodeHTMLEntities(name)
	signature = "<span class=""name""><a href=""#" & LCase(name) & """>" & name & "</a></span>"
	GenPropertySummary = GenSummary(signature, tags)
End Function

'! Generate summary documentation for a class methods or (global) procedure.
'! The documentation is generated in HTML format.
'!
'! @param  name   Name of the method/procedure.
'! @param  params Array with the parameter names of the method/procedure.
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return The summary documentation in HTML format.
Private Function GenMethodSummary(ByVal name, ByVal params, ByVal tags)
	Dim signature

	name = EncodeHTMLEntities(name)
	params = EncodeHTMLEntities(Join(params, ", "))
	signature = "<span class=""name""><a href=""#" & LCase(name) & "(" & LCase(Replace(params, " ", "")) _
		& ")"">" & name & "</a></span>(" & params & ")"
	GenMethodSummary = GenSummary(signature, tags)
End Function

'! Generate summary documentation for variables. All variables passed to this
'! function were declared in one list (Dim a, b, ...) and had thus the same
'! documentation header. The documentation is generated in HTML format.
'!
'! @param  vars   Array with the names of the variables.
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return The summary documentation in HTML format.
Private Function GenVariableSummary(ByVal vars, ByVal tags)
	Dim i, name, signature
	ReDim summary(UBound(vars))

	For i = LBound(vars) To UBound(vars)
		name = EncodeHTMLEntities(vars(i))
		signature = "<span class=""name""><a href=""#" & LCase(name) & """>" & name & "</a></span>"
		summary(i) = GenSummary(signature, tags)
	Next
	GenVariableSummary = Join(summary, "<hr/>" & vbNewLine)
End Function

'! Generate summary documentation for a constant. The documentation is
'! generated in HTML format.
'!
'! @param  name   Name of the constant.
'! @param  value  Value of the constant.
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return The summary documentation in HTML format.
Private Function GenConstSummary(ByVal name, ByVal value, ByVal tags)
	Dim signature

	name = EncodeHTMLEntities(name)
	signature = "<span class=""name""><a href=""#" & LCase(name) & """>" & name & "</a></span>: " & Trim(value)
	GenConstSummary = GenSummary(signature, tags)
End Function

'! Generate detail documentation for a class property. The documentation is
'! generated in HTML format.
'!
'! @param  name          Name of the property.
'! @param  accessibility The accessibility (read-only/write-only/read-write) of
'!                       the properties.
'! @param  tags          Dictionary with the tag/value pairs from the
'!                       documentation header.
'! @return The detail documentation in HTML format.
Private Function GenPropertyDetails(ByVal name, ByVal accessibility, ByVal tags)
	Dim heading, signature

	name = EncodeHTMLEntities(name)
	heading = "<a name=""" & LCase(name) & """></a>" & name
	signature = "<span class=""name"">" & name & "</span> (" & accessibility & ")"
	GenPropertyDetails = GenDetailsHeading(heading, signature) _
		& GenDetailsInfo(tags) _
		& GenExceptionInfo(tags) _
		& GenReferencesInfo(tags)
End Function

'! Generate detail documentation for a class methods or (global) procedure.
'! The documentation is generated in HTML format.
'!
'! @param  name       Name of the method/procedure.
'! @param  visibility visibility (Private/Public) of the method.
'! @param  params     Array with the parameter names of the method/procedure.
'! @param  tags       Dictionary with the tag/value pairs from the
'!                    documentation header.
'! @return The detail documentation in HTML format.
Private Function GenMethodDetails(ByVal name, ByVal visibility, ByVal params, ByVal tags)
	Dim heading, signature

	name = EncodeHTMLEntities(name)
	params = EncodeHTMLEntities(Join(params, ", "))
	heading = "<a name=""" & LCase(name) & "(" & LCase(Replace(params, " ", "")) & ")""></a>" & name
	signature = visibility & " <span class=""name"">" & name & "</span>(" & params & ")"
	GenMethodDetails = GenDetailsHeading(heading, signature) _
		& GenDetailsInfo(tags) _
		& GenParameterInfo(tags) _
		& GenReturnValueInfo(tags) _
		& GenExceptionInfo(tags) _
		& GenReferencesInfo(tags)
End Function

'! Generate detail documentation for variables. All variables passed to this
'! function were declared in one list (Dim a, b, ...) and had thus the same
'! documentation header and visibility. The documentation is generated in HTML
'! format.
'!
'! @param  vars       Array with the names of the variables.
'! @param  visibility visibility (Private/Public) of the variables.
'! @param  tags       Dictionary with the tag/value pairs from the
'!                    documentation header.
'! @return The detail documentation in HTML format.
Private Function GenVariableDetails(ByVal vars, ByVal visibility, ByVal tags)
	Dim description, references, i, name, heading, signature
	ReDim details(UBound(vars))

	description = GenDetailsInfo(tags)
	references = GenReferencesInfo(tags)

	For i = LBound(vars) To UBound(vars)
		name = EncodeHTMLEntities(vars(i))
		heading = "<a name=""" & LCase(name) & """></a>" & name
		signature = visibility & " <span class=""name"">" & name & "</span>"
		details(i) = GenDetailsHeading(heading, signature) & description & references
	Next

	GenVariableDetails = Join(details, "<hr/>" & vbNewLine)
End Function

'! Generate detail documentation for a constant. The documentation is generated
'! in HTML format.
'!
'! @param  name       Name of the constant.
'! @param  value      Value of the constant.
'! @param  visibility visibility (Private/Public) of the variables.
'! @param  tags       Dictionary with the tag/value pairs from the
'!                    documentation header.
'! @return The detail documentation in HTML format.
Private Function GenConstDetails(ByVal name, ByVal value, ByVal visibility, ByVal tags)
	Dim heading, signature

	name = EncodeHTMLEntities(name)
	heading = "<a name=""" & LCase(name) & """></a>" & name
	signature = visibility & " Const <span class=""name"">" & name & "</span> = " & Trim(value)
	GenConstDetails = GenDetailsHeading(heading, signature) & GenDetailsInfo(tags) & GenReferencesInfo(tags)
End Function

'! Generate author information from @author tags. Should the tag also contain
'! an e-mail address, that address is made into a hyperlink.
'!
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return HTML snippet with the author information.
Private Function GenAuthorInfo(tags)
	Dim re, author
	Dim info : info = ""

	If tags.Exists("@author") Then
		info = "<h4>Author:</h4>" & vbNewLine
		Set re = CompileRegExp("\S+@\S+", True, True)
		For Each author In tags("@author")
			If re.Test(author) Then
				' author data contains e-mail address => create link
				info = info & "<p class=""value"">" & Trim(EncodeHTMLEntities(Trim(re.Replace(author, ""))) _
					& " &lt;" & CreateLink(re.Execute(author)(0), True)) & "&gt;</p>" & vbNewLine
			Else
				' author data does not contain e-mail address => use as-is
				info = info & "<p class=""value"">" & EncodeHTMLEntities(Trim(author)) & "</p>" & vbNewLine
			End If
		Next
	End If

	GenAuthorInfo = info
End Function

'! Generate references ("see also") information from @see tags. All values are
'! made into hyperlinks (either internal or external).
'!
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return HTML snippet with the references information.
Private Function GenReferencesInfo(tags)
	Dim ref
	Dim info : info = ""

	If tags.Exists("@see") Then
		info = "<h4>See Also:</h4>" & vbNewLine
		For Each ref In tags("@see")
			info = info & "<p class=""value"">" & CreateLink(ref, False) & "</p>" & vbNewLine
		Next
	End If

	GenReferencesInfo = info
End Function

'! Generate version information from the @version tag. If an @date tag is
'! present as well, its value is appended to the version.
'!
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return HTML snippet with the version information.
Private Function GenVersionInfo(tags)
	Dim info : info = ""

	If tags.Exists("@version") Then
		info = "<h4>Version:</h4>" & vbNewLine & "<p class=""value"">" & EncodeHTMLEntities(tags("@version"))
		If tags.Exists("@date") Then info = info & ", " & EncodeHTMLEntities(tags("@date"))
		info = info & "</p>" & vbNewLine
	End If

	GenVersionInfo = info
End Function

'! Generate HTML code for heading and signature line in a "detail" section.
'! The heading is created as <h3>.
'!
'! @param  heading   The heading text.
'! @param  signature The signature of the procedure, variable or constant.
'! @return HTML snippet with the heading and signature.
Private Function GenDetailsHeading(heading, signature)
	GenDetailsHeading = "<h3>" & heading & "</h3>" & vbNewLine & "<p class=""function""><code>" _
		& signature & "</code></p>" & vbNewLine
End Function

'! Generate detail information from the @detail tag.
'!
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return HTML snippet with the detail information. Empty string if no detail
'!         information was present.
Private Function GenDetailsInfo(tags)
	Dim info : info = ""

	If tags.Exists("@details") Then
		info = MangleBlankLines(tags("@details"), 1)
		info = EncodeHTMLEntities(info)
		info = "<p>" & Replace(info, vbNewLine, "</p>" & vbLf & "<p>") & "</p>"

		' Remove blank lines.
		info = Replace(info, "<p></p>" & vbLf, "")

		' Mark list items as such.
		Dim re : Set re = CompileRegExp("<p>- (.*)</p>", True, True)
		info = re.Replace(info, "<li>$1</li>")
		' Enclose blocks of list items in <ul></ul> tags.
		re.Pattern = "(^|</p>\n)<li>"
		info = re.Replace(info, "$1<ul>" & vbLf & "<li>")
		re.Pattern = "</li>(\n<p|$)"
		info = re.Replace(info, "</li>" & vbLf & "</ul>$1")

		' Add classifications.
		info = Replace(info, "<p>", "<p class=""description"">")
		info = Replace(info, "<ul>", "<ul class=""description"">")
	End If

	GenDetailsInfo = Replace(info, vbLf, vbNewLine) & vbNewLine
End Function

'! Generate parameter information from @param tags.
'!
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return HTML snippet with the parameter information. Empty string if no
'!         parameter information was present.
Private Function GenParameterInfo(tags)
	Dim param
	Dim info : info = ""

	If tags.Exists("@param") Then
		info = "<h4>Parameters:</h4>" & vbNewLine
		For Each param In tags("@param")
			param = Split(param, " ", 2)
			info = info & "<p class=""value""><code>" & EncodeHTMLEntities(param(0)) & "</code>"
			If UBound(param) > 0 Then info = info & " - " & EncodeHTMLEntities(Trim(param(1)))
			info = info & "</p>" & vbNewLine
		Next
	End If

	GenParameterInfo = info
End Function

'! Generate return value information from the @return tag.
'!
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return HTML snippet with the return value information. Empty string if no
'!         return value information was present.
Private Function GenReturnValueInfo(tags)
	GenReturnValueInfo = ""
	If tags.Exists("@return") Then GenReturnValueInfo = "<h4>Returns:</h4>" & vbNewLine _
		& "<p class=""value"">" & EncodeHTMLEntities(tags("@return")) & "</p>" & vbNewLine
End Function

'! Generate information on the errors raised by a method/procedure from @raise
'! tags.
'!
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return HTML snippet with the error information. Empty string if no error
'!         information was present.
Private Function GenExceptionInfo(tags)
	Dim errType
	Dim info : info = ""

	If tags.Exists("@raise") Then
		info = "<h4>Raises:</h4>" & vbNewLine
		For Each errType In tags("@raise")
			info = info & "<p class=""value"">" & EncodeHTMLEntities(errType) & "</p>" & vbNewLine
		Next
	End If

	GenExceptionInfo = info
End Function

'! Create a hyperlink from the given reference. If isMailTo is True, the
'! reference is considered an e-mail address.
'!
'! @param  ref        The reference.
'! @param  isMailTo   Treat the reference as an e-mail address.
'! @return HTML snippet with the hyperlink to the reference.
Private Function CreateLink(ByVal ref, ByVal isMailTo)
	Dim reURL, link

	Set reURL = CompileRegExp("<(.*)>", True, True)
	ref = reURL.Replace(ref, "$1")

	If isMailTo Then
		' Link is mailto: link.
		link = ">" & ref
		ref = "mailto:" & ref
	ElseIf Left(ref, 1) = "#" Then
		' Link is internal reference.
		link = ">" & Mid(ref, 2)
		ref = LCase(ref)
	Else
		' Link is external reference.
		link = "target=""_blank"">" & ref
	End If

	CreateLink = "<a href=""" & ref & """" & link & "</a>"
End Function

'! Generate summary information from the @brief tag.
'!
'! @param  name   Name of the documented element.
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return HTML snippet with the summary information. Empty string if no
'!         summary information was present.
Private Function GenSummary(name, tags)
	Dim summary
	summary = "<p class=""function""><code>" & name & "</code></p>" & vbNewLine
	If tags.Exists("@brief") Then summary = summary & "<p class=""description"">" _
		& EncodeHTMLEntities(tags("@brief")) & "</p>" & vbNewLine
	GenSummary = summary
End Function

'! Create a stylesheet with the given filename.
'!
'! @param  filename   Name (including relative or absolute path) of the file
'!                    to create.
Private Sub CreateStylesheet(filename)
	Dim f : Set f = fso.OpenTextFile(filename, ForWriting, True)
	f.WriteLine "* { margin: 0; padding: 0; border: 0; }" & vbNewLine _
		& "body { margin: 10px; margin-bottom: 30px; font-family: " & TextFont & "; font-size: " & BaseFontSize & "; }" & vbNewLine _
		& "h1,h2,h3,h4 { font-weight: bold; }" & vbNewLine _
		& "h1 { font-size: 200%; margin-bottom: 10px; }" & vbNewLine _
		& "h2 { background-color: #ccccff; border: 1px solid black; font-size: 150%; margin: 20px 0 10px; padding: 10px 5px; }" & vbNewLine _
		& "h3,p { margin-bottom: 5px; }" & vbNewLine _
		& "h4,p.description { margin: 3px 0 0 50px; }" & vbNewLine _
		& "h4 { margin-top: 6px; margin-bottom: 4px; }" & vbNewLine _
		& "p.value { margin-left: 100px; }" & vbNewLine _
		& "code { font-family: " & CodeFont & "; }" & vbNewLine _
		& "span.name { font-weight: bold; }" & vbNewLine _
		& "hr { border: 1px solid #a0a0a0; width: 94%; margin: 10px 3%; }" & vbNewLine _
		& "ul { list-style: disc inside; margin-left: 50px; padding: 5px 0; }" & vbNewLine _
		& "ul.description { margin-left: 75px; }" & vbNewLine _
		& "li { text-indent: -1em; margin-left: 1em; }"
	f.Close
End Sub

' ------------------------------------------------------------------------------
' Consistency checks
' ------------------------------------------------------------------------------

'! Apply some consistency checks to documentation of procedures and functions.
'!
'! @param  name      Name of the procedure or function.
'! @param  params    Array with the parameter names of the procedure or function.
'! @param  tags      A dictionary with the values of the documentation tags.
'! @param  funcType  Type of the procedure or function (function/sub).
Private Sub CheckConsistency(name, params, tags, funcType)
	If Not tags.Exists("@brief") And Not beQuiet Then log.LogWarning "No description for " & name & "() found."
	If tags.Exists("@param") Then
		CheckParameterMismatch params, tags("@param"), name
	Else
		CheckParameterMismatch params, Array(), name
	End If
	CheckRetvalMismatch funcType, name, tags.Exists("@return")
End Sub

'! Check for mismatches between the documented and the actual parameters of a
'! procedure or function. Logs a warning if there is a mismatch.
'!
'! @param  codeParams  Array with the actual parameters from the code.
'! @param  docParams   Array with the documented parameters.
'! @param  name        Name of the procedure or function.
Private Sub CheckParameterMismatch(ByVal codeParams, ByVal docParams, ByVal name)
	Dim cPrms, dPrms, param

	Set cPrms = CreateObject("Scripting.Dictionary")
	For Each param In codeParams
		cPrms.Add LCase(param), True
	Next

	Set dPrms = CreateObject("Scripting.Dictionary")
	For Each param In docParams
		dPrms.Add Split(LCase(param & " "))(0), True
	Next

	For Each param In cPrms.Keys
		If dPrms.Exists(param) Then
			cPrms.Remove(param)
			dPrms.Remove(param)
		End If
	Next

	If cPrms.Count > 0 And Not beQuiet Then log.LogWarning "Undocumented parameters in " & name & "(): " & Join(cPrms.Keys, ", ")
	If dPrms.Count > 0 And Not beQuiet Then log.LogWarning "Parameters not implemented in " & name & "(): " & Join(dPrms.Keys, ", ")
End Sub

'! Check for return value mismatches. Functions must have a return value, subs
'! must not have a return value. In case of a mismatch a warning is logged.
'!
'! @param  funcType    Type of the procedure or function (function/sub).
'! @param  name        Name of the procedure or function.
'! @param  hasRetval   Flag indicating if the procedure or function has a
'!                     documented return value.
Private Sub CheckRetvalMismatch(ByVal funcType, ByVal name, ByVal hasRetval)
	Select Case LCase(funcType)
		Case "function" If Not hasRetval And Not beQuiet Then log.LogWarning "Undocumented return value for method " & name & "()."
		Case "sub" If hasRetval And Not beQuiet Then log.LogWarning "Method " & name & "() cannot have a return value."
		Case Else log.LogError "CheckRetvalMismatch(): Invalid type " & funcType & "."
	End Select
End Sub

'! Check for pointless documentation tags in the doc comments of identifiers.
'! Applies to both variables and constants.
'!
'! @param  name   Name of the identifier.
'! @param  tags   A dictionary with the values of the documentation tags.
Private Sub CheckIdentifierTags(ByVal name, ByVal tags)
	If Not tags.Exists("@brief") And Not beQuiet Then log.LogWarning "No description for identifier(s) " & Trim(name) & " found."
	If tags.Exists("@param") And Not beQuiet Then log.LogWarning "Parameter documentation found, but " & name & " is an identifier."
	If tags.Exists("@return") And Not beQuiet Then log.LogWarning "Return value documentation found, but " & name & " is an identifier."
	If tags.Exists("@raise") And Not beQuiet Then log.LogWarning "Exception documentation found, but " & name & " is an identifier."
End Sub

'! Check the given string for remaining code. Issue a warning, if there's still
'! code left (i.e. the string does not consist of whitespace only). Should the
'! string contain an "Option Explicit" statement, that statement is ignored.
'!
'! @param  str  The string to check.
Private Sub CheckRemainingCode(str)
	' "Option Explicit" statement can be ignored, so remove it.
	Dim re : Set re = CompileRegExp("Option[ \t]+Explicit", True, True)
	str = re.Replace(str, "")
	' Also remove comment lines.
	re.Pattern = "(^|\n)[ \t]*'.*"
	str = re.Replace(str, vbLf)

	If Trim(Replace(Replace(str, vbTab, ""), vbLf, "")) <> "" Then
		' there's still some (global) code left
		str = MangleBlankLines(Replace(str, vbLf, vbNewLine), 2)
		str = "> " & Replace(str, vbNewLine, vbNewLine & "> ")  ' prepend each line with "> "
		If Not beQuiet Then log.LogWarning "Unencapsulated global code:" & vbNewLine & str
	End If
End Sub

' ------------------------------------------------------------------------------
' Helper functions
' ------------------------------------------------------------------------------

'! Compile a new regular expression.
'!
'! @param  pattern      The pattern for the regular expression.
'! @param  ignoreCase   Boolean value indicating whether the regular expression
'!                      should be treated case-insensitive or not.
'! @param  searchGlobal Boolean value indicating whether all matches or just
'!                      the first one should be returned.
'! @return The prepared regular expression object.
Private Function CompileRegExp(pattern, ignoreCase, searchGlobal)
	Set CompileRegExp = New RegExp
	CompileRegExp.Pattern = pattern
	CompileRegExp.IgnoreCase = Not Not ignoreCase
	CompileRegExp.Global = Not Not searchGlobal
End Function

'! Recursively create a directory and all non-existent parent directories.
'!
'! @param  dir  The directory to create.
Private Sub CreateDirectory(ByVal dir)
	dir = fso.GetAbsolutePathName(dir)
	' The recursion terminates once an existing parent folder is found. Which in
	' the worst case is the drive's root folder.
	If Not fso.FolderExists(dir) Then
		CreateDirectory fso.GetParentFolderName(dir)
		' At this point it's certain that the parent folder does exist, so we can
		' carry on and create the subfolder.
		fso.CreateFolder dir
	End If
End Sub

'! Format the given date with ISO format "yyyy-mm-dd".
'!
'! @param  val  The date to format.
'! @return The formatted date.
Private Function FormatDate(val)
	FormatDate = Year(val) _
		& "-" & Right("0" & Month(val), 2) _
		& "-" & Right("0" & Day(val), 2)
End Function

'! Append the given value to an array or a (string) variable.
'!
'! @param  var   Reference to a variable that the given value should be appended to.
'! @param  val   The value to append.
Private Sub Append(ByRef var, ByVal val)
	Select Case TypeName(var)
		Case "Variant()"
			ReDim Preserve var(UBound(var)+1)
			var(UBound(var)) = val
		Case "Empty"
			var = val
		Case Else
			var = var & " " & val
	End Select
End Sub

'! Return the canonicalized visibility classifier. Can be either Private or
'! Public. If anything else but the string "private" is given, the function
'! defaults to Public. That allows to determine the visibility of variable
'! and constant declarations as well.
'!
'! @param  classifier   The visibility classifier.
'! @return The canonicalized visibility.
Private Function GetVisibility(ByVal classifier)
	If LCase(Trim(Replace(classifier, vbTab, " "))) = "private" Then
		GetVisibility = "Private"
	Else
		GetVisibility = "Public"
	End If
End Function

'! Replace special characters (and some particular character sequences) with
'! their respective HTML entity encoding.
'!
'! @param  text  The text to encode.
'! @return The encoded text.
Private Function EncodeHTMLEntities(ByVal text)
	' Ampersand (&) must be encoded first.
	text = Replace(text, "&", "&amp;")

	' replace/encode character sequences with "special" meanings
	text = Replace(text, "<->", "&harr;")
	text = Replace(text, "<-", "&larr;")
	text = Replace(text, "->", "&rarr;")
	text = Replace(text, "<=>", "&hArr;")
	text = Replace(text, "<=", "&lArr;")
	text = Replace(text, "=>", "&&rArr;")
	text = Replace(text, "...", "…")
	text = Replace(text, "(c)", "©", 1, vbReplaceAll, vbTextCompare)
	text = Replace(text, "(r)", "®", 1, vbReplaceAll, vbTextCompare)

	' encode all other HTML entities
	text = Replace(text, "ä", "&auml;")
	text = Replace(text, "Ä", "&Auml;")
	text = Replace(text, "ë", "&euml;")
	text = Replace(text, "Ë", "&Euml;")
	text = Replace(text, "ï", "&iuml;")
	text = Replace(text, "Ï", "&Iuml;")
	text = Replace(text, "ö", "&ouml;")
	text = Replace(text, "Ö", "&Ouml;")
	text = Replace(text, "ü", "&uuml;")
	text = Replace(text, "Ü", "&Uuml;")
	text = Replace(text, "ÿ", "&yuml;")
	text = Replace(text, "Ÿ", "&Yuml;")
	text = Replace(text, "¨", "&uml;")
	text = Replace(text, "á", "&aacute;")
	text = Replace(text, "Á", "&Aacute;")
	text = Replace(text, "é", "&eacute;")
	text = Replace(text, "É", "&Eacute;")
	text = Replace(text, "í", "&iacute;")
	text = Replace(text, "Í", "&Iacute;")
	text = Replace(text, "ó", "&oacute;")
	text = Replace(text, "Ó", "&Oacute;")
	text = Replace(text, "ú", "&uacute;")
	text = Replace(text, "Ú", "&Uacute;")
	text = Replace(text, "ý", "&yacute;")
	text = Replace(text, "Ý", "&Yacute;")
	text = Replace(text, "à", "&agrave;")
	text = Replace(text, "À", "&Agrave;")
	text = Replace(text, "è", "&egrave;")
	text = Replace(text, "È", "&Egrave;")
	text = Replace(text, "ì", "&igrave;")
	text = Replace(text, "Ì", "&Igrave;")
	text = Replace(text, "ò", "&ograve;")
	text = Replace(text, "Ò", "&Ograve;")
	text = Replace(text, "ù", "&ugrave;")
	text = Replace(text, "Ù", "&Ugrave;")
	text = Replace(text, "â", "&acirc;")
	text = Replace(text, "Â", "&Acirc;")
	text = Replace(text, "ê", "&ecirc;")
	text = Replace(text, "Ê", "&Ecirc;")
	text = Replace(text, "î", "&icirc;")
	text = Replace(text, "Î", "&Icirc;")
	text = Replace(text, "ô", "&ocirc;")
	text = Replace(text, "Ô", "&Ocirc;")
	text = Replace(text, "û", "&ucirc;")
	text = Replace(text, "Û", "&Ucirc;")
	text = Replace(text, "ˆ", "&circ;")
	text = Replace(text, "ã", "&atilde;")
	text = Replace(text, "Ã", "&Atilde;")
	text = Replace(text, "ñ", "&ntilde;")
	text = Replace(text, "Ñ", "&Ntilde;")
	text = Replace(text, "õ", "&otilde;")
	text = Replace(text, "Õ", "&Otilde;")
	text = Replace(text, "˜", "&tilde;")
	text = Replace(text, "å", "&aring;")
	text = Replace(text, "Å", "&Aring;")
	text = Replace(text, "ç", "&ccedil;")
	text = Replace(text, "Ç", "&Ccedil;")
	text = Replace(text, "¸", "&cedil;")
	text = Replace(text, "ø", "&oslash;")
	text = Replace(text, "Ø", "&Oslash;")
	text = Replace(text, "ß", "&szlig;")
	text = Replace(text, "æ", "&aelig;")
	text = Replace(text, "Æ", "&AElig;")
	text = Replace(text, "œ", "&oelig;")
	text = Replace(text, "Œ", "&OElig;")
	text = Replace(text, "š", "&scaron;")
	text = Replace(text, "Š", "&Scaron;")
	text = Replace(text, "µ", "&micro;")
	' quotation marks
	text = Replace(text, """", "&quot;")
	text = Replace(text, "'", "&apos;")
	text = Replace(text, "«", "&laquo;")
	text = Replace(text, "»", "&raquo;")
	text = Replace(text, "‹", "&lsaquo;")
	text = Replace(text, "›", "&rsaquo;")
	text = Replace(text, "‘", "&lsquo;")
	text = Replace(text, "’", "&rsquo;")
	text = Replace(text, "‚", "&sbquo;")
	text = Replace(text, "“", "&ldquo;")
	text = Replace(text, "”", "&rdquo;")
	text = Replace(text, "„", "&bdquo;")
	' currency symbols
	text = Replace(text, "¢", "&cent;")
	text = Replace(text, "€", "&euro;")
	text = Replace(text, "£", "&pound;")
	text = Replace(text, "¥", "&yen;")
	' other special character
	text = Replace(text, ">", "&gt;")
	text = Replace(text, "<", "&lt;")
	text = Replace(text, "°", "&deg;")
	text = Replace(text, "©", "&copy;")
	text = Replace(text, "®", "&reg;")
	text = Replace(text, "¡", "&iexcl;")
	text = Replace(text, "¿", "&iquest;")
	text = Replace(text, "·", "&middot;")
	text = Replace(text, "•", "&bull;")
	text = Replace(text, "§", "&sect;")
	text = Replace(text, "ª", "&ordf;")
	text = Replace(text, "º", "&ordm;")
	text = Replace(text, "¼", "&frac14;")
	text = Replace(text, "½", "&frac12;")
	text = Replace(text, "¾", "&frac34;")
	text = Replace(text, "¹", "&sup1;")
	text = Replace(text, "²", "&sup2;")
	text = Replace(text, "³", "&sup3;")
	text = Replace(text, "¯", "&macr;")
	text = Replace(text, "±", "&plusmn;")
	text = Replace(text, "–", "&ndash;")
	text = Replace(text, "—", "&mdash;")
	text = Replace(text, "…", "&hellip;")

	EncodeHTMLEntities = text
End Function

'! Write the given heading and (HTML) text to the given file, unless the text
'! is the empty string. The heading is written as <h2>.
'!
'! @param  file     File to write to.
'! @param  heading  The heading for the given text.
'! @param  text     Text to write to the file.
Private Sub WriteSection(file, heading, text)
	If text <> "" Then file.WriteLine "<h2>" & heading & "</h2>" & vbNewLine & text
End Sub

'! Mangle multiple blank lines in the given string into the given number of
'! newlines. Also remove leading and trailing newlines.
'!
'! @param  str    The string to process.
'! @param  number Number of newlines to use as replacement.
'! @return The string without the unwanted newlines.
Private Function MangleBlankLines(ByVal str, ByVal number)
	Dim re, i, replacement

	Set re = New RegExp
	re.Global = True

	' Remove spaces and tabs from lines consisting only of spaces and/or tabs.
	re.Pattern = "^[ \t]*$"
	str = Split(str, vbNewLine)
	For i = LBound(str) To UBound(str)
		str(i) = re.Replace(str(i), "")
	Next
	str = Join(str, vbNewLine)

	' Mangle multiple newlines. It would've been nice to be able to create the
	' replacement (a string of newlines) like this: String(n, vbNewLine).
	' Unfortunately, String() only works with single characters, while vbNewLine
	' might consist of two characters (CR + LF), depending on the platform.
	' Therefore the workaround to create a string with the desired number of
	' spaces, and then replace the spaces with newlines.
	re.Pattern = "(" & vbNewLine & "){3,}"
	replacement = Replace(Space(number), " ", vbNewLine)
	str = re.Replace(str, replacement)

	' Remove leading/trailing newlines.
	re.Pattern = "(^" & replacement & "|" & replacement & "$)"
	str = re.Replace(str, "")

	MangleBlankLines = str
End Function

'! Display usage information and exit.
Private Sub PrintUsage()
	log.LogInfo "Usage:" & vbTab & WScript.ScriptName & " [/a] [/q] [/p:NAME] /s:SOURCE /d:DOC_DIR" & vbNewLine _
		& vbTab & WScript.ScriptName & " /?" & vbNewLine & vbNewLine _
		& vbTab & "/?   Print this help." & vbNewLine _
		& vbTab & "/a   Generate documentation for all elements (public and private)." & vbNewLine _
		& vbTab & "     Without this option, documentation is generated for public" & vbNewLine _
		& vbTab & "     elements only." & vbNewLine _
		& vbTab & "/d   Generate documentation in DOC_DIR." & vbNewLine _
		& vbTab & "/p   Use NAME as the project name." & vbNewLine _
		& vbTab & "/q   Don't print warnings." & vbNewLine _
		& vbTab & "/s   Source to generate documentation from. Can be a file or a" & vbNewLine _
		& vbTab & "     directory."
	WScript.Quit 0
End Sub

' ==============================================================================

'! Import the first occurrence of the given filename from the working directory
'! or any directory in the %PATH%.
'!
'! @param  filename  Name of the file to import (can be either absolute or relative)
'!
'! @raise  Path not found (0x4c)
'!
'! @see http://gazeek.com/coding/importing-vbs-files-in-your-vbscript-project/
Private Sub Import(ByVal filename)
	Dim fso, sh, file, code, dir

	' Create my own objects, so the function is self-contained and can be called
	' before anything else in the script.
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set sh = CreateObject("WScript.Shell")

	filename = Trim(sh.ExpandEnvironmentStrings(filename))
	If Not (Left(filename, 2) = "\\" Or Mid(filename, 2, 2) = ":\") Then
		' filename is not absolute
		If Not fso.FileExists(fso.GetAbsolutePathName(filename)) Then
			' file doesn't exist in the working directory => iterate over the
			' directories in the %PATH% and take the first occurrence
			' if no occurrence is found => use filename as-is, which will result
			' in an error when trying to open the file
			For Each dir In Split(sh.ExpandEnvironmentStrings("%PATH%"), ";")
				If fso.FileExists(fso.BuildPath(dir, filename)) Then
					filename = fso.BuildPath(dir, filename)
					Exit For
				End If
			Next
		End If
		filename = fso.GetAbsolutePathName(filename)
	End If

	Set file = fso.OpenTextFile(filename, 1, False)
	code = file.ReadAll
	file.Close

	ExecuteGlobal(code)
End Sub
