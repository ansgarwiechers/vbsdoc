'! Script for automatic generation of API documentation from special
'! comments in VBScripts.
'!
'! @author  Ansgar Wiechers <ansgar.wiechers@planetcobalt.net>
'! @date    2010-11-21
'! @version 1.0

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

'! @todo build list of internal references and warn when conflicts occur
'! @todo Add tag @mainpage for global documentation on the global index page?
'! @todo Add HTMLHelp generation?
'! @todo Add grouping option for doc comments? (like doxygen's @{ ... @})

Option Explicit

Import "LoggerClass.vbs"

' Some symbolic constants for internal use.
Private Const ForReading = 1
Private Const ForWriting = 2

Private Const vbReplaceAll = -1

Private Const Ext = "vbs"

Private Const StylesheetName = "vbsdoc.css"
Private Const TextFont       = "Verdana, Arial, helvetica, sans-serif"
Private Const CodeFont       = "Lucida Console, Courier New, Courier, monospace"
Private Const BaseFontSize   = "14px"

Private Const DefaultLanguage = "en"

' Initialize global objects.
Private fso : Set fso = CreateObject("Scripting.FileSystemObject")
Private sh  : Set sh = CreateObject("WScript.Shell")
Private log : Set log = New Logger

'! Match line-continuations.
Private reLineCont : Set reLineCont = CompileRegExp("[ \t]+_\n[ \t]*", True, True)
'! Match End-of-Line doc comments.
Private reEOLComment : Set reEOLComment = CompileRegExp("(^|\n)([ \t]*[^' \t\n].*)('![ \t]*.*(\n[ \t]*'!.*)*)", True, True)
'! Match @todo-tagged doc comments.
Private reTodo : Set reTodo = CompileRegExp("'![ \t]*@todo[ \t]*(.*\n([ \t]*'!([ \t]*[^@\s].*|\s*)\n)*)", True, True)
'! Match class implementations and prepended doc comments.
Private reClass : Set reClass = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*Class[ \t]+(\w+)([\s\S]*?)End[ \t]+Class", True, True)
'! Match constructor implementations and prepended doc comments.
Private reCtor : Set reCtor = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*((Public|Private)[ \t]+)?Sub[ \t]+(Class_Initialize)[ \t]*(\(\))?[\s\S]*?End[ \t]+Sub", True, True)
'! Match destructor implementations and prepended doc comments.
Private reDtor : Set reDtor = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*((Public|Private)[ \t]+)?Sub[ \t]+(Class_Terminate)[ \t]*(\(\))?[\s\S]*?End[ \t]+Sub", True, True)
'! Match implementations of methods/procedures as well as prepended
'! doc comments.
Private reMethod : Set reMethod = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*((Public|Private)[ \t]+)?(Function|Sub)[ \t]+(\w+)[ \t]*(\([\w\t ,]*\))?[\s\S]*?End[ \t]+\6", True, True)
'! Match property implementations and prepended doc comments.
Private reProperty : Set reProperty = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*((Public|Private)[ \t]+)?Property[ \t]+(Get|Let|Set)[ \t]+(\w+)[ \t]*(\([\w\t ]*\))?[\s\S]*?End[ \t]+Property", True, True)
'! Match definitions of constants and prepended doc comments.
Private reConst : Set reConst = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*((Public|Private)[ \t]+)?Const[ \t]+(\w+)[ \t]*=[ \t]*(.*)", True, True)
'! Match variable declarations and prepended doc comments. Allow for combined
'! declaration:definition as well as multiple declarations of variables on
'! one line, e.g.:
'!   - Dim foo : foo = 42
'!   - Dim foo, bar, baz
Private reVar : Set reVar = CompileRegExp("(^|\n)(([ \t]*'!.*\n)*)[ \t]*(Public|Private|Dim|ReDim)[ \t]+(((\w+)([ \t]*\(\))?)[ \t]*(:[ \t]*(Set[ \t]+)?\7[ \t]*=.*|(,[ \t]*\w+[ \t]*(\(\))?)*))", True, True)
'! Match doc comments. This regular expression is used to process file-global
'! doc comments after all other elements in a given file were processed.
Private reDocComment : Set reDocComment = CompileRegExp("^[ \t]*('!.*)", True, True)

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

Private localize : Set localize = CreateObject("Scripting.Dictionary")
	' English localization
	localize.Add "en", CreateObject("Scripting.Dictionary")
		localize("en").Add "AUTHOR"          , "Author"
		localize("en").Add "CLASS"           , "Class"
		localize("en").Add "CLASS_SUMMARY"   , "Classes Summary"
		localize("en").Add "CONST_DETAIL"    , "Global Constant Detail"
		localize("en").Add "CONST_SUMMARY"   , "Global Constant Summary"
		localize("en").Add "CTORDTOR_DETAIL", "Constructor/Destructor Detail"
		localize("en").Add "CTORDTOR_SUMMARY", "Constructor/Destructor Summary"
		localize("en").Add "EXCEPT"          , "Raises"
		localize("en").Add "METHOD_DETAIL"   , "Method Detail"
		localize("en").Add "METHOD_SUMMARY"  , "Method Summary"
		localize("en").Add "PARAM"           , "Parameters"
		localize("en").Add "PROC_DETAIL"     , "Global Procedure Detail"
		localize("en").Add "PROC_SUMMARY"    , "Global Procedure Summary"
		localize("en").Add "PROP_DETAIL"     , "Property Detail"
		localize("en").Add "PROP_SUMMARY"    , "Property Summary"
		localize("en").Add "RETURN"          , "Returns"
		localize("en").Add "SEE_ALSO"        , "See also"
		localize("en").Add "TODO"            , "ToDo List"
		localize("en").Add "VAR_DETAIL"      , "Global Variable Detail"
		localize("en").Add "VAR_SUMMARY"     , "Global Variable Summary"
		localize("en").Add "VERSION"         , "Version"
	' Deutsche Lokalisierung
	localize.Add "de", CreateObject("Scripting.Dictionary")
		localize("de").Add "AUTHOR"          , "Autor"
		localize("de").Add "CLASS"           , "Klasse"
		localize("de").Add "CLASS_SUMMARY"   , "Klassen - Zusammenfassung"
		localize("de").Add "CONST_DETAIL"    , "Globale Konstanten - Details"
		localize("de").Add "CONST_SUMMARY"   , "Globale Konstanten - Zusammenfassung"
		localize("de").Add "CTORDTOR_DETAIL", "Konstruktor/Destruktor - Details"
		localize("de").Add "CTORDTOR_SUMMARY", "Konstruktor/Destruktor - Zusammenfassung"
		localize("de").Add "EXCEPT"          , "Wirft"
		localize("de").Add "METHOD_DETAIL"   , "Methoden - Details"
		localize("de").Add "METHOD_SUMMARY"  , "Methoden - Zusammenfassung"
		localize("de").Add "PARAM"           , "Parameter"
		localize("de").Add "PROC_DETAIL"     , "Globale Prozeduren - Details"
		localize("de").Add "PROC_SUMMARY"    , "Global Prozeduren - Zusammenfassung"
		localize("de").Add "PROP_DETAIL"     , "Eigenschaften - Details"
		localize("de").Add "PROP_SUMMARY"    , "Eigenschaften - Zusammenfassung"
		localize("de").Add "RETURN"          , "Rückgabewert"
		localize("de").Add "SEE_ALSO"        , "Siehe auch"
		localize("de").Add "TODO"            , "Aufgabenliste"
		localize("de").Add "VAR_DETAIL"      , "Globale Variablen - Details"
		localize("de").Add "VAR_SUMMARY"     , "Globale Variablen - Zusammenfassung"
		localize("de").Add "VERSION"         , "Version"

Private beQuiet, projectName

Main WScript.Arguments


'! The starting point. Evaluates commandline arguments, initializes
'! global variables and starts the documentation generation.
'!
'! @param  args   The list of arguments passed to the script.
Public Sub Main(args)
	Dim lang, includePrivate, srcRoot, docRoot, doc
	Dim docTitle, name

	' initialize global variables/settings with default values
	beQuiet = False
	log.Debug = False
	projectName = ""

	' initialize local variables with default values
	lang = DefaultLanguage
	includePrivate = False

	' evaluate commandline arguments
	With args
		If .Named.Exists("?") Then PrintUsage(0)

		If .Named.Exists("d") Then log.Debug = True
		If .Named.Exists("a") Then includePrivate = True
		If .Named.Exists("q") And Not log.Debug Then beQuiet = True
		If .Named.Exists("p") Then projectName = .Named("p")

		' Use the default language if /l is omitted or used without specifying a
		' particular language. Use the given language if it exists in localize.Keys.
		' Otherwise print an error message and exit.
		If .Named.Exists("l") Then
			If localize.Exists(.Named("l")) Then
				lang = .Named("l")
			ElseIf .Named("l") <> "" Then
				log.LogError .Named("l") & " is not a supported Language. Valid languages are: " _
					& Join(Sort(localize.Keys), ", ")
				WScript.Quit(1)
			End If
		End If

		If .Named.Exists("i") Then
			srcRoot = .Named("i")
		Else
			PrintUsage(1)
		End If

		If .Named.Exists("o") Then
			docRoot = .Named("o")
		Else
			PrintUsage(1)
		End If
	End With

	log.LogDebug "beQuiet:        " & beQuiet
	log.LogDebug "projectName:    " & projectName
	log.LogDebug "lang:           " & lang
	log.LogDebug "includePrivate: " & includePrivate
	log.LogDebug "srcRoot:        " & srcRoot
	log.LogDebug "docRoot:        " & docRoot

	' extract the data
	Set doc = CreateObject("Scripting.Dictionary")

	If fso.FileExists(srcRoot) Then
		doc.Add "", GetFileDef(srcRoot, includePrivate)
		docTitle = fso.GetFileName(srcRoot)
	Else
		GetDef doc, fso.GetFolder(srcRoot), "", includePrivate
		docTitle = Null
	End If

	For Each name In doc.Keys
		If doc(name) Is Nothing Then doc.Remove(name)
	Next

	' generate the documentation
	GenDoc doc, docRoot, lang, docTitle

	WScript.Quit(0)
End Sub

' ------------------------------------------------------------------------------
' Data gathering
' ------------------------------------------------------------------------------

'! Traverse all subdirecotries of the given srcDir and extract documentation
'! information from all VBS files. If includePrivate is set to True, then
'! documentation for private elements is generated as well, otherwise only
'! public elements are included in the documentation.
'!
'! @param  doc            Reference to the dictionary containing the
'!                        documentation elements extracted from the source
'!                        files.
'! @param  srcDir         Directory containing the source files for the
'!                        documentation generation.
'! @param  docDir         Path relative to the documentation root directory
'!                        where the documentation for the files in srcDir
'!                        should be generated.
'! @param  includePrivate Include documentation for private elements.
Public Sub GetDef(ByRef doc, srcDir, docDir, includePrivate)
	Dim f, name, srcFile, dir

	log.LogDebug "GetDef(" & TypeName(doc) & ", " & TypeName(srcDir) & ", " & TypeName(docDir) & ", " & TypeName(includePrivate) & ")"

	For Each f In srcDir.Files
		log.LogDebug "> " & fso.BuildPath(srcDir, f.Name)
		If LCase(fso.GetExtensionName(f.Name)) = Ext Then
			name = Replace(fso.BuildPath(docDir, fso.GetBaseName(f.Name)), "\", "/")
			srcFile = fso.BuildPath(f.ParentFolder, f.Name)
			doc.Add name, GetFileDef(srcFile, includePrivate)
		End If
	Next

	For Each dir In srcDir.SubFolders
		log.LogDebug "> " & fso.BuildPath(srcDir, dir.Name)
		GetDef doc, dir, fso.BuildPath(docDir, dir.Name), includePrivate
	Next
End Sub

'! Generate the documentation for a given file. The documentation files are
'! created in the given documentation directory. If includePrivate is set to
'! True, then documentation for private elements is generated as well,
'! otherwise only public elements are included in the documentation.
'!
'! @param  filename       Name of the source file for documentation generation.
'! @param  includePrivate Include documentation for private elements.
'! @return The name and path of the generated documentation file.
'! @return Dictionary describing structural and metadata elements in the given
'!         source file.
Public Function GetFileDef(filename, includePrivate)
	Dim outDir, inFile, content, m, line, globalComments, document

	log.LogDebug "GetFileDef(" & TypeName(filename) & ", " & TypeName(includePrivate) & ")"

	Set GetFileDef = Nothing
	If fso.GetFile(filename).Size = 0 Or Not fso.FileExists(filename) Then
		If fso.FileExists(filename) Then
			log.LogDebug "File " & filename & " has size 0."
		Else
			log.LogDebug "File " & filename & " does not exist."
		End If
		Exit Function ' nothing to do
	End If

	log.LogInfo "Generating documentation for " & filename & " ..."

	log.LogDebug "Reading input file " & filename & " ..."
	Set inFile = fso.OpenTextFile(filename, ForReading)
	content = inFile.ReadAll
	inFile.Close

	' ****************************************************************************
	' preparatory steps
	' ****************************************************************************

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

	' ****************************************************************************
	' parsing the content starts here
	' ****************************************************************************

	Set document = CreateObject("Scripting.Dictionary")

	document.Add "Todo", GetTodoList(content)
	document.Add "Classes", GetClassDef(content, includePrivate)
	document.Add "Procedures", GetMethodDef(content, includePrivate)
	document.Add "Constants", GetConstDef(content, includePrivate)
	document.Add "Variables", GetVariableDef(content, includePrivate)

	' process file-global doc comments
	globalComments = ""
	For Each line In Split(content, vbLf)
		If reDocComment.Test(line) Then globalComments = globalComments & reDocComment.Replace(line, "$1") & vbLf
	Next
	document.Add "Metadata", ProcessComments(globalComments)

	CheckRemainingCode content

	Set GetFileDef = document
End Function

'! Get a list of todo items. The list is generated from the @todo tags in the
'! code.
'!
'! @param  code   code fragment to check for @todo items.
'! @return A list with the todo items from the code.
Private Function GetTodoList(ByRef code)
	Dim list, m, line

	log.LogDebug "GetTodoList(" & TypeName(code) & ")"

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

	GetTodoList = list
End Function

'! Extract definitions of classes from the given code fragment.
'!
'! @param  code           Code fragment containing the class implementation.
'! @param  includePrivate Include private members/methods in the documentation.
'! @return Dictionary of dictionaries describing the class(es). The keys of
'!         the main dictionary are the names of the classes, which point to
'!         sub-dictionaries containing the definition data.
Private Function GetClassDef(ByRef code, ByVal includePrivate)
	Dim m, classBody, d

	log.LogDebug "GetClassDef(" & TypeName(code) & ", " & TypeName(includePrivate) & ")"

	Dim classes : Set classes = CreateObject("Scripting.Dictionary")

	For Each m In reClass.Execute(code)
		With m.SubMatches
			classBody = .Item(4)

			Set d = CreateObject("Scripting.Dictionary")
			d.Add "Metadata", ProcessComments(.Item(1))
			d.Add "Constructor", GetCtorDtorDef(classBody, includePrivate, True)
			d.Add "Destructor", GetCtorDtorDef(classBody, includePrivate, False)
			d.Add "Properties", GetPropertyDef(classBody)
			d.Add "Methods", GetMethodDef(classBody, includePrivate)
			d.Add "Fields", GetVariableDef(classBody, includePrivate)

			classes.Add .Item(3), d
		End With
	Next
	code = reClass.Replace(code, vbLf)

	Set GetClassDef = classes
End Function

'! Extract definitions of constructor or destructor from the given code
'! fragment.
'!
'! @param  code           Code fragment containing constructor/destructor.
'! @param  includePrivate Include definition of private constructor or
'!                        destructor as well.
'! @param  returnCtor     If true, the function will return the constructor
'!                        definition, otherwise the destructor definition.
'! @return Dictionary describing constructor or destructor.
Private Function GetCtorDtorDef(ByRef code, ByVal includePrivate, ByVal returnCtor)
	Dim re, descr, m, isPrivate, tags, methodRedefined

	log.LogDebug "GetCtorDtorDef(" & TypeName(code) & ", " & TypeName(includePrivate) & ", " & TypeName(returnCtor) & ")"

	Dim method : Set method = CreateObject("Scripting.Dictionary")

	If returnCtor Then
		Set re = reCtor
		descr = "constructor"
	Else
		Set re = reDtor
		descr = "destructor"
	End If

	For Each m In re.Execute(code)
		With m.SubMatches
			isPrivate = CheckIfPrivate(.Item(4))
			If Not isPrivate Or includePrivate Then
				Set tags = ProcessComments(.Item(1))
				CheckConsistency .Item(5), Array(), tags, "sub"

				On Error Resume Next
				method.Add "Parameters", Array()
				method.Add "IsPrivate", isPrivate
				method.Add "Metadata", tags
				If Err.Number <> 0 Then
					If Err.Number = 457 Then
						' key is already present in dictionary
						methodRedefined = True
						' overwrite previous definition data ("last match wins")
						' no need to overwrite "Parameters", though, because those are
						' always [] for constructor as well as destructor
						method("IsPrivate") = isPrivate
						Set method("Metadata") = tags
					Else
						log.LogError "Error storing " & descr & " data: " & FormatErrorMessage(Err)
					End If
				End If
				On Error Goto 0
			End If
		End With
	Next
	code = re.Replace(code, vbLf)

	If methodRedefined And Not beQuiet Then log.LogWarning "Multiple " & descr _
		& " definitions. Using the last one."

	Set GetCtorDtorDef = method
End Function

'! Extract definitions of class methods and global procedures from the given
'! code fragment.
'!
'! @param  code           Code fragment containing the methods/procedures.
'! @param  includePrivate Include definitions of private methods as well.
'! @return Dictionary of dictionaries describing the methods. The keys of
'!         the main dictionary are the names of the methods, which point to
'!         sub-dictionaries containing the definition data.
Private Function GetMethodDef(ByRef code, ByVal includePrivate)
	Dim m, isPrivate, params, tags, d

	log.LogDebug "GetMethodDef(" & TypeName(code) & ", " & TypeName(includePrivate) & ")"

	Dim methods : Set methods = CreateObject("Scripting.Dictionary")

	For Each m In reMethod.Execute(code)
		With m.SubMatches
			isPrivate = CheckIfPrivate(.Item(4))
			If Not isPrivate Or includePrivate Then
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

				Set d = CreateObject("Scripting.Dictionary")
				d.Add "Parameters", params
				d.Add "IsPrivate", isPrivate
				d.Add "Metadata", tags

				methods.Add .Item(6), d
			End If
		End With
	Next
	code = reMethod.Replace(code, vbLf)

	Set GetMethodDef = methods
End Function

'! Extract definitons of class properties. Private getter and setter methods
'! are disregarded, because even with the method present, the property would
'! not be readable/writable from an interface point of view.
'!
'! @param  code   Code fragment containing class properties.
'! @return Dictionary with summary and detail documentation.
'! @return Dictionary of dictionaries describing the properties. The keys of
'!         the main dictionary are the names of the properties, which point to
'!         sub-dictionaries containing the definition data.
Private Function GetPropertyDef(ByRef code)
	Dim m, name, d

	log.LogDebug "GetPropertyDef(" & TypeName(code) & ")"

	Dim properties : Set properties = CreateObject("Scripting.Dictionary")
	Dim readable   : Set readable = CreateObject("Scripting.Dictionary")
	Dim writable   : Set writable = CreateObject("Scripting.Dictionary")

	For Each m In reProperty.Execute(code)
		With m.SubMatches
			If Not CheckIfPrivate(.Item(4)) Then
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

	' Add readable properties to the result dictionary. Set writable status
	' according to the property name's presence in the "writable" dictionary
	' and remove matching properties from the latter dictionary.
	For Each name In readable.Keys
		Set d = CreateObject("Scripting.Dictionary")
		d.Add "Readable", True
		If writable.Exists(name) Then
			d.Add "Writable", True
			d.Add "Metadata", ProcessComments(readable(name) & writable(name))
			writable.Remove(name)
		Else
			d.Add "Writable", False
			d.Add "Metadata", ProcessComments(readable(name))
		End If
		properties.Add name, d
	Next

	' At this point the "writable" dictionary contains only properties that are
	' not present in the "readable" dictionary, so we can process those as
	' write-only properties.
	For Each name In writable.Keys
		Set d = CreateObject("Scripting.Dictionary")
		d.Add "Readable", False
		d.Add "Writable", True
		d.Add "Metadata", ProcessComments(writable(name))
		properties.Add name, d
	Next

	Set GetPropertyDef = properties
End Function

'! Extract definitions of variables from the given code fragment.
'!
'! @param  code           Code fragment containing variable declarations.
'! @param  includePrivate Include definitions of private variables as well.
'! @return Dictionary of dictionaries describing the variables. The keys of
'!         the main dictionary are the names of the variables, which point to
'!         sub-dictionaries containing the definition data.
Private Function GetVariableDef(ByRef code, ByVal includePrivate)
	Dim m, isPrivate, tags, vars, name, d

	log.LogDebug "GetVariableDef(" & TypeName(code) & ", " & TypeName(includePrivate) & ")"

	Dim variables : Set variables = CreateObject("Scripting.Dictionary")

	For Each m In reVar.Execute(code)
		With m.SubMatches
			isPrivate = CheckIfPrivate(.Item(3))
			If Not isPrivate Or includePrivate Then
				Set tags = ProcessComments(.Item(1))
				vars = .Item(4)
				' If the match contains a declaration/definition combination: remove the
				' definition part.
				If Left(.Item(8), 1) = ":" Then vars = Trim(Split(vars, ":")(0))
				CheckIdentifierTags vars, tags
				vars = Split(Replace(Replace(vars, vbTab, ""), " ", ""), ",")

				Set d = CreateObject("Scripting.Dictionary")
				d.Add "IsPrivate", isPrivate
				d.Add "Metadata", tags

				For Each name In vars
					variables.Add name, d
				Next
			End If
		End With
	Next
	code = reVar.Replace(code, vbLf)

	Set GetVariableDef = variables
End Function

'! Extract definitions of (global) constants from the given code fragment.
'!
'! @param  code           Code fragment containing constant definitions.
'! @param  includePrivate Include definitions of private constants as well.
'! @return Dictionariy of dictionaries describing the constants. The keys of
'!         the main dictionary are the names of the constants, which point to
'!         sub-dictionaries containing the definition data.
Private Function GetConstDef(ByRef code, ByVal includePrivate)
	Dim m, isPrivate, tags, d

	log.LogDebug "GetConstDef(" & TypeName(code) & ", " & TypeName(includePrivate) & ")"

	Dim constants : Set constants = CreateObject("Scripting.Dictionary")

	For Each m In reConst.Execute(code)
		With m.SubMatches
			isPrivate = CheckIfPrivate(.Item(4))
			If Not isPrivate Or includePrivate Then
				Set tags = ProcessComments(.Item(1))
				CheckIdentifierTags .Item(5), tags

				Set d = CreateObject("Scripting.Dictionary")
				d.Add "Value", .Item(6)
				d.Add "IsPrivate", isPrivate
				d.Add "Metadata", tags

				constants.Add .Item(5), d
			End If
		End With
	Next
	code = reConst.Replace(code, "")

	Set GetConstDef = constants
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

	log.LogDebug "ProcessComments(" & TypeName(comments) & ")"

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
' Document generation
' ------------------------------------------------------------------------------

'! Generate the documentation from the extracted data.
'!
'! @param  doc      Structure containing the documentation elements extracted
'!                  from the source file(s).
'! @param  docRoot  Root directory for the documentation files.
'! @param  lang     Documentation language. All generated text that is not read
'!                  from the source document(s) is created in this language.
'! @param  title    Title of the documentation page when documentation for a
'!                  single source file is generated. Must be Null otherwise.
Sub GenDoc(doc, docRoot, lang, title)
	Dim indexFile, relPath, re, css, filename, f, dir, entry, name, section, isFirst

	log.LogDebug "GenDoc(" & TypeName(doc) & ", " & TypeName(docRoot) & ", " & TypeName(lang) & ", " & TypeName(title) & ")"

	CreateDirectory docRoot
	CreateStylesheet fso.BuildPath(docRoot, StylesheetName)

	If IsNull(title) Then
		log.LogDebug "Writing index file " & fsoBuildPath(docRoot, "index.html") & " ..."
		Set indexFile = fso.OpenTextFile(fso.BuildPath(docRoot, "index.html"), ForWriting, True)
		WriteHeader indexFile, "Main Page", StylesheetName
		If projectName <> "" Then indexFile.WriteLine "<h1>" & projectName & "</h1>"

		For Each relPath In Sort(doc.Keys)
			indexFile.WriteLine "<p><a href=""" & relPath & "/index.html"">" & relPath & ".vbs</a></p>"
		Next

		WriteFooter indexFile
		indexFile.Close
	End If

	Set re = CompileRegExp("[^\\]+\\", True, True)

	For Each relPath In doc.Keys
		css = re.Replace(fso.BuildPath(relPath, StylesheetName), "../")
		dir = fso.BuildPath(docRoot, relPath)
		CreateDirectory dir

		log.LogDebug "Writing script documentation file " & fso.BuildPath(dir, "index.html") & " ..."
		Set f = fso.OpenTextFile(fso.BuildPath(dir, "index.html"), ForWriting, True)
		WriteHeader f, fso.GetFileName(relPath), css

		If IsNull(title) Then
			f.WriteLine "<h1>" & fso.GetFileName(relPath) & "</h1>"
		Else
			f.WriteLine "<h1>" & title & "</h1>"
		End If

		With doc(relPath)
			f.Write GenDetailsInfo(.Item("Metadata"))
			f.Write GenVersionInfo(.Item("Metadata"), lang)
			f.Write GenAuthorInfo(.Item("Metadata"), lang)
			f.Write GenReferencesInfo(.Item("Metadata"), lang)

			' Write ToDo list.
			If UBound(.Item("Todo")) > -1 Then
				f.WriteLine "<h2>" & EncodeHTMLEntities(localize(lang)("TODO")) & "</h2>" & vbNewLine & "<ul>"
				For Each entry In .Item("Todo")
					f.WriteLine "  <li>" & EncodeHTMLEntities(entry) & "</li>"
				Next
				f.WriteLine "</ul>"
			End If

			WriteSection f, localize(lang)("CONST_SUMMARY"), .Item("Constants"), lang, "Constant", True
			WriteSection f, localize(lang)("VAR_SUMMARY"), .Item("Variables"), lang, "Variable", True

			' Write class summary information.
			If .Item("Classes").Count > 0 Then
				f.WriteLine "<h2>" & EncodeHTMLEntities(localize(lang)("CLASS_SUMMARY")) & "</h2>"
				isFirst = True
				For Each entry In Sort(.Item("Classes").Keys)
					If isFirst Then
						isFirst = False
					Else
						f.WriteLine "<hr/>"
					End If
					f.WriteLine "<p><span class=""name""><a href=""" & EncodeHTMLEntities(entry) & ".html"">" _
						& EncodeHTMLEntities(entry) & "</a></span></p>"
					If .Item("Classes")(entry)("Metadata").Exists("@brief") Then f.WriteLine "<p class=""description"">" _
						& EncodeHTMLEntities(.Item("Classes")(entry)("Metadata")("@brief")) & "</p>"
				Next
			End If

			WriteSection f, localize(lang)("PROC_SUMMARY"), .Item("Procedures"), lang, "Procedure", True
			WriteSection f, localize(lang)("CONST_DETAIL"), .Item("Constants"), lang, "Constant", False
			WriteSection f, localize(lang)("VAR_DETAIL"), .Item("Variables"), lang, "Variable", False
			WriteSection f, localize(lang)("PROC_DETAIL"), .Item("Procedures"), lang, "Procedure", False
		End With

		WriteFooter f
		f.Close

		For Each name In doc(relPath)("Classes").Keys
			log.LogDebug "Writing class documentation file " & fso.BuildPath(dir, name & ".html") & " ..."
			Set f = fso.OpenTextFile(fso.BuildPath(dir, name & ".html"), ForWriting, True)
			WriteHeader f, EncodeHTMLEntities(name), css

			With doc(relPath)("Classes")(name)
				f.WriteLine "<h1>" & EncodeHTMLEntities(localize(lang)("CLASS") & " " & name) & "</h1>"
				f.Write GenDetailsInfo(.Item("Metadata"))
				f.Write GenVersionInfo(.Item("Metadata"), lang)
				f.Write GenAuthorInfo(.Item("Metadata"), lang)
				f.Write GenReferencesInfo(.Item("Metadata"), lang)

				WriteSection f, localize(lang)("PROP_SUMMARY"), .Item("Properties"), lang, "Property", True
				section = ""
				If .Item("Constructor").Count > 0 Then section = GenSummary("Class_Initialize", .Item("Constructor"), "Procedure")
				If .Item("Destructor").Count > 0 Then
					If section <> "" Then section = section & vbNewLine & "<hr/>" & vbNewLine
					section = section & GenSummary("Class_Terminate", .Item("Destructor"), "Procedure")
				End If
				If section <> "" Then f.WriteLine "<h2>" & EncodeHTMLEntities(localize(lang)("CTORDTOR_SUMMARY")) _
					& "</h2>" & vbNewLine & section
				WriteSection f, localize(lang)("METHOD_SUMMARY"), .Item("Methods"), lang, "Procedure", True

				WriteSection f, localize(lang)("PROP_DETAIL"), .Item("Properties"), lang, "Property", False
				section = ""
				If .Item("Constructor").Count > 0 Then section = GenDetails("Class_Initialize", .Item("Constructor"), lang, "Procedure")
				If .Item("Destructor").Count > 0 Then
					If section <> "" Then section = section & vbNewLine & "<hr/>" & vbNewLine
					section = section & GenDetails("Class_Terminate", .Item("Destructor"), lang, "Procedure")
				End If
				If section <> "" Then f.WriteLine "<h2>" & EncodeHTMLEntities(localize(lang)("CTORDTOR_DETAIL")) _
					& "</h2>" & vbNewLine & section
				WriteSection f, localize(lang)("METHOD_DETAIL"), .Item("Methods"), lang, "Procedure", False
			End With

			WriteFooter f
			f.Close
		Next
	Next
End Sub

'! Write the given heading and data to the given file. The heading omitted in
'! case it's Null, otherwise it is written as <h2>. The data is processed into
'! summary or detail information, depending on presence (or absence) of the
'! word "summary" in the heading.
'!
'! @param  file         Filehandle to write to.
'! @param  heading      The heading for the given section.
'! @param  data         Data (sub-)structure containing the elements of the
'!                      section.
'! @param  lang         Documentation language. All generated text that is not
'!                      read from the source document(s) is created in this
'!                      language.
'! @param  sectionType  The type of the section to be written (constant,
'!                      method, property, or variable)
'! @param  isSummary    If True, a summary section is generated, otherwise a
'!                      detail section.
Private Sub WriteSection(file, heading, data, lang, sectionType, isSummary)
	Dim name, isFirst

	log.LogDebug "WriteSection(" & TypeName(file) & ", " & TypeName(heading) & ", " & TypeName(data) & ", " & TypeName(lang) & ", " & TypeName(sectionType) & ", " & TypeName(isSummary) & ")"

	If data.Count > 0 Then
		If Not IsNull(heading) Then file.WriteLine "<h2>" & EncodeHTMLEntities(heading) & "</h2>" & vbNewLine
		isFirst = True
		For Each name In Sort(data.Keys)
			If isFirst Then
				isFirst = False
			Else
				file.WriteLine "<hr/>"
			End If
			If isSummary Then
				file.WriteLine GenSummary(name, data(name), sectionType)
			Else
				file.WriteLine GenDetails(name, data(name), lang, sectionType)
			End If
		Next
	End If
End Sub

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
	log.LogDebug "WriteHeader(" & TypeName(outFile) & ", " & TypeName(title) & ", " & TypeName(stylesheet) & ")"
	log.LogDebug "title:      " & title
	log.LogDebug "stylesheet: " & stylesheet

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

'! Generate summary documentation. The documentation is generated in HTML
'! format.
'!
'! @param  name         Name of the processed element.
'! @param  properties   Dictionary with the properties of the processed
'!                      element.
'! @param elementType   The type of the processed element (constant, method,
'!                      property, or variable)
'! @return The summary documentation in HTML format.
Private Function GenSummary(ByVal name, ByVal properties, ByVal elementType)
	Dim signature, params

	log.LogDebug "GenSummary(" & TypeName(name) & ", " & TypeName(properties) & ", " & TypeName(elementType) & ")"
	log.LogDebug "name:        " & name
	log.LogDebug "properties:  " & Join(properties.Keys, ", ")
	log.LogDebug "elementType: " & elementType

	name = EncodeHTMLEntities(name)

	Select Case LCase(elementType)
	Case "constant"
		signature = "<span class=""name""><a href=""#" & LCase(name) & """>" & name & "</a></span>: " _
			& Trim(properties("Value"))
		GenSummary = GenSummaryInfo(signature, properties("Metadata"))
	Case "procedure"
		params = EncodeHTMLEntities(Join(properties("Parameters"), ", "))
		signature = "<span class=""name""><a href=""#" & LCase(name) & "(" & LCase(Replace(params, " ", "")) _
			& ")"">" & name & "</a></span>(" & params & ")"
		GenSummary = GenSummaryInfo(signature, properties("Metadata"))
	Case "property"
		signature = "<span class=""name""><a href=""#" & LCase(name) & """>" & name & "</a></span>"
		GenSummary = GenSummaryInfo(signature, properties("Metadata"))
	Case "variable"
		signature = "<span class=""name""><a href=""#" & LCase(name) & """>" & name & "</a></span>"
		GenSummary = GenSummaryInfo(signature, properties("Metadata"))
	Case Else
		log.LogError "Cannot generate summary information for unknown element type " & elementType & "."
	End Select
End Function

'! Generate detail documentation. The documentation is generated in HTML
'! format.
'!
'! @param  name         Name of the processed element.
'! @param  properties   Dictionary with the properties of the processed
'!                      element.
'! @param  lang         Documentation language. All generated text that is not
'!                      read from the source document(s) is created in this
'!                      language.
'! @param elementType   The type of the processed element (constant, method,
'!                      property, or variable)
'! @return The detail documentation in HTML format.
Private Function GenDetails(ByVal name, ByVal properties, ByVal lang, ByVal elementType)
	Dim heading, signature, params, visibility, accessibility

	log.LogDebug "GenDetails(" & TypeName(name) & ", " & TypeName(properties) & ", " & TypeName(lang) & ", " & TypeName(elementType) & ")"
	log.LogDebug "name:        " & name
	log.LogDebug "properties:  " & Join(properties.Keys, ", ")
	log.LogDebug "lang:        " & lang
	log.LogDebug "elementType: " & elementType

	GenDetails = ""

	name = EncodeHTMLEntities(name)

	If LCase(elementType) = "procedure" Then
		params = EncodeHTMLEntities(Join(properties("Parameters"), ", "))
		heading = "<a name=""" & LCase(name) & "(" & LCase(Replace(params, " ", "")) & ")""></a>" & name
	Else
		heading = "<a name=""" & LCase(name) & """></a>" & name
	End If

	If properties("IsPrivate") Then
		visibility = "Private"
	Else
		visibility = "Public"
	End If

	Select Case LCase(elementType)
	Case "constant"
		signature = visibility & " Const <span class=""name"">" & name & "</span> = " & Trim(properties("Value"))
		GenDetails = GenDetailsHeading(heading, signature) _
			& GenDetailsInfo(properties("Metadata")) _
			& GenReferencesInfo(properties("Metadata"), lang)
	Case "procedure"
		signature = visibility & " <span class=""name"">" & name & "</span>(" & params & ")"
		GenDetails = GenDetailsHeading(heading, signature) _
			& GenDetailsInfo(properties("Metadata")) _
			& GenParameterInfo(properties("Metadata"), lang) _
			& GenReturnValueInfo(properties("Metadata"), lang) _
			& GenExceptionInfo(properties("Metadata"), lang) _
			& GenReferencesInfo(properties("Metadata"), lang)
	Case "property"
		If properties("Readable") Then
			If properties("Writable") Then
				accessibility = "read-write"
			Else
				accessibility = "read-only"
			End If
		Else
			If properties("Writable") Then
				accessibility = "write-only"
			Else
				log.LogError "Property " & name & " is neither readable nor writable. This should never happen, since this kind of property is ignored by the document parser."
			End If
		End If
		signature = "<span class=""name"">" & name & "</span> (" & accessibility & ")"
		GenDetails = GenDetailsHeading(heading, signature) _
			& GenDetailsInfo(properties("Metadata")) _
			& GenExceptionInfo(properties("Metadata"), lang) _
			& GenReferencesInfo(properties("Metadata"), lang)
	Case "variable"
		signature = visibility & " <span class=""name"">" & name & "</span>"
		GenDetails = GenDetailsHeading(heading, signature) _
			& GenDetailsInfo(properties("Metadata")) _
			& GenReferencesInfo(properties("Metadata"), lang)
	Case Else
		log.LogError "Cannot generate detail information for unknown element type " & elementType & "."
	End Select
End Function

'! Generate author information from @author tags. Should the tag also contain
'! an e-mail address, that address is made into a hyperlink.
'!
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @param  lang     Documentation language. All generated text that is not read
'!                  from the source document(s) is created in this language.
'! @return HTML snippet with the author information.
Private Function GenAuthorInfo(tags, lang)
	Dim re, author
	Dim info : info = ""

	log.LogDebug "GenAuthorInfo(" & TypeName(tags) & ", " & TypeName(lang) & ")"

	If tags.Exists("@author") Then
		info = "<h4>" & EncodeHTMLEntities(localize(lang)("AUTHOR")) & ":</h4>" & vbNewLine
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
'! @param  lang   Documentation language. All generated text that is not read
'!                from the source document(s) is created in this language.
'! @return HTML snippet with the references information.
Private Function GenReferencesInfo(tags, lang)
	Dim ref
	Dim info : info = ""

	log.LogDebug "GenReferencesInfo(" & TypeName(tags) & ", " & TypeName(lang) & ")"

	If tags.Exists("@see") Then
		info = "<h4>" & EncodeHTMLEntities(localize(lang)("SEE_ALSO")) & ":</h4>" & vbNewLine
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
'! @param  lang   Documentation language. All generated text that is not read
'!                from the source document(s) is created in this language.
'! @return HTML snippet with the version information.
Private Function GenVersionInfo(tags, lang)
	Dim info : info = ""

	log.LogDebug "GenVersionInfo(" & TypeName(tags) & ", " & TypeName(lang) & ")"

	If tags.Exists("@version") Then
		info = "<h4>" & EncodeHTMLEntities(localize(lang)("VERSION")) & ":</h4>" & vbNewLine _
			& "<p class=""value"">" & EncodeHTMLEntities(tags("@version"))
		If tags.Exists("@date") Then info = info & ", " & EncodeHTMLEntities(tags("@date"))
		info = info & "</p>" & vbNewLine
	End If

	GenVersionInfo = info
End Function

'! Generate summary information from the @brief tag.
'!
'! @param  name   Name of the documented element.
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @return HTML snippet with the summary information. Empty string if no
'!         summary information was present.
Private Function GenSummaryInfo(name, tags)
	Dim summary

	log.LogDebug "GenSummaryInfo(" & TypeName(name) & ", " & TypeName(tags) & ")"

	summary = "<p class=""function""><code>" & name & "</code></p>" & vbNewLine
	If tags.Exists("@brief") Then summary = summary & "<p class=""description"">" _
		& EncodeHTMLEntities(tags("@brief")) & "</p>" & vbNewLine
	GenSummaryInfo = summary
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

	log.LogDebug "GenDetailsInfo(" & TypeName(tags) & ")"

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
'! @param  lang   Documentation language. All generated text that is not read
'!                from the source document(s) is created in this language.
'! @return HTML snippet with the parameter information. Empty string if no
'!         parameter information was present.
Private Function GenParameterInfo(tags, lang)
	Dim param
	Dim info : info = ""

	log.LogDebug "GenParameterInfo(" & TypeName(tags) & ", " & TypeName(lang) & ")"

	If tags.Exists("@param") Then
		info = "<h4>" & EncodeHTMLEntities(localize(lang)("PARAM")) & ":</h4>" & vbNewLine
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
'! @param  lang   Documentation language. All generated text that is not read
'!                from the source document(s) is created in this language.
'! @return HTML snippet with the return value information. Empty string if no
'!         return value information was present.
Private Function GenReturnValueInfo(tags, lang)
	GenReturnValueInfo = ""
	log.LogDebug "GenReturnValueInfo(" & TypeName(tags) & ", " & TypeName(lang) & ")"
	If tags.Exists("@return") Then GenReturnValueInfo = "<h4>" & EncodeHTMLEntities(localize(lang)("RETURN")) _
		& ":</h4>" & vbNewLine & "<p class=""value"">" & EncodeHTMLEntities(tags("@return")) & "</p>" & vbNewLine
End Function

'! Generate information on the errors raised by a method/procedure from @raise
'! tags.
'!
'! @param  tags   Dictionary with the tag/value pairs from the documentation
'!                header.
'! @param  lang   Documentation language. All generated text that is not read
'!                from the source document(s) is created in this language.
'! @return HTML snippet with the error information. Empty string if no error
'!         information was present.
Private Function GenExceptionInfo(tags, lang)
	Dim errType
	Dim info : info = ""

	log.LogDebug "GenExceptionInfo(" & TypeName(tags) & ", " & TypeName(lang) & ")"

	If tags.Exists("@raise") Then
		info = "<h4>" & EncodeHTMLEntities(localize(lang)("EXCEPT")) & ":</h4>" & vbNewLine
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

	log.LogDebug "CreateLink(" & TypeName(ref) & ", " & TypeName(isMailTo) & ")"

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

'! Create a stylesheet with the given filename.
'!
'! @param  filename   Name (including relative or absolute path) of the file
'!                    to create.
Private Sub CreateStylesheet(filename)
	log.LogDebug "Creating stylesheet " & filename & " ..."
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
	log.LogDebug "CreateDirectory(" & dir & ")"
	dir = fso.GetAbsolutePathName(dir)
	' The recursion terminates once an existing parent folder is found. Which in
	' the worst case is the drive's root folder.
	If Not fso.FolderExists(dir) Then
		CreateDirectory fso.GetParentFolderName(dir)
		' At this point it's certain that the parent folder does exist, so we can
		' carry on and create the subfolder.
		fso.CreateFolder dir
		log.LogDebug "Directory " & dir & " created."
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

'! Return a boolean value indicating whether the classifier is "private". If
'! anything else but the string "private" is given, the function defaults to
'! False. That allows to determine the visibility of variable and constant
'! declarations as well.
'!
'! @param  classifier   The visibility classifier.
'! @return True if visibility is "private", otherwise False.
Private Function CheckIfPrivate(ByVal classifier)
	If LCase(Trim(Replace(classifier, vbTab, " "))) = "private" Then
		CheckIfPrivate = True
	Else
		CheckIfPrivate = False
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

'! Mangle multiple blank lines in the given string into the given number of
'! newlines. Also remove leading and trailing newlines.
'!
'! @param  str    The string to process.
'! @param  number Number of newlines to use as replacement.
'! @return The string without the unwanted newlines.
Private Function MangleBlankLines(ByVal str, ByVal number)
	Dim re, i, replacement

	log.LogDebug "MangleBlankLines(" & TypeName(str) & ", " & TypeName(number) & ")"

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

'! Sort a given array in ascending order. This is merely a wrapper for
'! QuickSort(), so that I can simply call Sort(array) without having to
'! specify the boundaries in the inital function call. This is also to
'! avoid changing the original array.
'!
'! @param  arr  The array to sort.
'! @return The array sorted in ascending order.
Function Sort(ByVal arr)
	QuickSort arr, 0, UBound(arr)
	Sort = arr
End Function

'! Sort a given array in ascending order, using the quicksort algorithm.
'!
'! @param  arr    The array to sort.
'! @param  left   Left (lower) boundary of the array slice the current
'!                recursion step will operate on.
'! @param  right  Right (upper) boundary of the array slice the current
'!                recursion step will operate on.
'!
'! @see http://en.wikipedia.org/wiki/Quicksort
Sub QuickSort(arr, left, right)
	Dim pivot, leftIndex, rightIndex, buffer

	log.LogDebug "QuickSort(" & TypeName(arr) & ", " & TypeName(left) & ", " & TypeName(right) & ")"

	leftIndex = left
	rightIndex = right

	If right - left > 0 Then
		pivot = Int((left + right) / 2)

		While leftIndex <= pivot And rightIndex >= pivot
			While arr(leftIndex) < arr(pivot) And leftIndex <= pivot
				leftIndex = leftIndex + 1
			Wend
			While arr(rightIndex) > arr(pivot) And rightIndex >= pivot
				rightIndex = rightIndex - 1
			Wend

			buffer = arr(leftIndex)
			arr(leftIndex) = arr(rightIndex)
			arr(rightIndex) = buffer

			leftIndex = leftIndex + 1
			rightIndex = rightIndex - 1
			If leftIndex - 1 = pivot Then
				rightIndex = rightIndex + 1
				pivot = rightIndex
			ElseIf rightIndex + 1 = pivot Then
				leftIndex = leftIndex - 1
				pivot = leftIndex
			End If
		Wend

		QuickSort arr, left, pivot-1
		QuickSort arr, pivot+1, right
	End If
End Sub

'! Display usage information and exit with the given exit code.
'!
'! @param  exitCode   The exit code.
Private Sub PrintUsage(exitCode)
	log.LogInfo "Usage:" & vbTab & WScript.ScriptName & " [/d] [/a] [/q] [/l:LANG] [/p:NAME] /i:SOURCE /o:DOC_DIR" & vbNewLine _
		& vbTab & WScript.ScriptName & " /?" & vbNewLine & vbNewLine _
		& vbTab & "/?" & vbTab & "Print this help." & vbNewLine _
		& vbTab & "/a" & vbTab & "Generate documentation for all elements (public and private)." & vbNewLine _
		& vbTab & vbTab & "Without this option, documentation is generated for public" & vbNewLine _
		& vbTab & vbTab & "elements only." & vbNewLine _
		& vbTab & "/d" & vbTab & "Enable debug messages. (you really don't want this)" & vbNewLine _
		& vbTab & "/i" & vbTab & "Read input files from SOURCE. Can be either a file or a" & vbNewLine _
		& vbTab & vbTab & "directory. (required)" & vbNewLine _
		& vbTab & "/l" & vbTab & "Create localized output [" & Join(Sort(localize.Keys), ",") & "]. Default language is " & DefaultLanguage & "." & vbNewLine _
		& vbTab & "/o" & vbTab & "Generate output files in DOC_DIR. (required)" & vbNewLine _
		& vbTab & "/p" & vbTab & "Use NAME as the project name." & vbNewLine _
		& vbTab & "/q" & vbTab & "Don't print warnings. Ignored if debug messages are enabled."
	WScript.Quit exitCode
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
