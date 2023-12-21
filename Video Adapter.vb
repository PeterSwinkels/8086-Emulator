'This interface's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Drawing

'This interface defines a video adapter.
Public Interface VideoAdapterClass

   'This procedure draws the specified video buffer's context on the specified image.
   Sub Display(Screen As Image, Memory() As Byte)

   'This procedure returns the screen size used by a video adapter.
   Function ScreenSize() As Size
End Interface
