﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "17.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("Emulator8086Program.Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to [Assembler]
        '''-Only 8086 instructions are supported.
        '''-A flat address target must be prefixed with an at sign (&quot;@&quot;) and is only supported for control flow instructions.
        '''-Single characters enclosed in quotes (&quot;&quot;&quot;) may be used in place of a hexadecimal byte.
        '''-Double characters enclosed in quotes (&quot;&quot;&quot;) may be used in place of a hexadecimal word.
        '''-Comments prefixed with a semi colon (&quot;;&quot;) are allowed.
        '''-Enter an asterisk (&quot;*&quot;) to overwrite the most recent instruction.
        '''.
        '''</summary>
        Friend ReadOnly Property Assembler() As String
            Get
                Return ResourceManager.GetString("Assembler", resourceCulture)
            End Get
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to [Help]
        '''?                             Displays this summary.
        '''$ Path                        Executes the commands in the specified file
        '''                              These files are required to have a &quot;[script]&quot; header as their first line.
        '''[Segment:Offset]              Displays the specified memory location contents as a byte, word, and first two bytes as characters.
        '''[Segment:Offset] = Value      Writes the specified byte/word/character to the specified memory address.
        '''[Segment:Offset] = {values}   Writ [rest of string was truncated]&quot;;.
        '''</summary>
        Friend ReadOnly Property Help() As String
            Get
                Return ResourceManager.GetString("Help", resourceCulture)
            End Get
        End Property
    End Module
End Namespace
