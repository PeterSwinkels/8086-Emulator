'This module's imports and settings.
Option Compare Binary
Option Explicit On
Option Infer Off
Option Strict On

'This module contains the mouse related procedures.
Public Module MouseModule
   Public Const MOUSE_BUTTON_COUNT As Integer = &H3%           'Defines the number of buttons the emulated mouse has.
   Public Const MOUSE_DRIVER_INSTALLED As Integer = &HFFFF%    'Indicates a mouse driver has been installed.
   Public Const MOUSE_DRIVER_NOT_INSTALLED As Integer = &H0%   'Indicates a mouse driver has not been installed.

End Module
