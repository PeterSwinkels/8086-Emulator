'This interface's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

Imports System.Drawing

'This interface defines a video adapter.
Public Interface VideoAdapterClass
   'This structure defines an area on the screen in text modes.
   Structure ScreenAreaStr
      Public ULCRow As Integer      'Defines the upper left corner's row.
      Public ULCColumn As Integer   'Defines the upper left corner's column.
      Public LRCRow As Integer      'Defines the lower right corner's row.
      Public LRCColumn As Integer   'Defines the lower right corner's column.
   End Structure

   'This procedure clears video adapter's buffer.
   Sub ClearBuffer()

   'This procedure draws the specified CPU memory's content onto the specified image.
   Sub Display(Screen As Image, Memory() As Byte, CodePage() As Integer)

   'This procedure initializes the video adapter.
   Sub Initialize()

   'This procedure returns the adapter's image resolution.
   Function Resolution() As Size

   'This procedure scrolls the video adapter's buffer.
   Sub ScrollBuffer(Up As Boolean, ScrollArea As ScreenAreaStr, Count As Integer)
End Interface
