;===============================================
/* JLDev v0.1
Written using AutoHotkey_L v1.1.09.03+ (http://l.autohotkey.net/)
By jlalonde on AHK forum
2013-06-19+
*/



###(str)
{
   static blnSkip### := false
   if (blnSkip###)
      return
   intOption := 6 + 512 ; 6 Cancel/Try Again/Continue + 512 Makes the 2nd button the default
   MsgBox, % intOption, Débug, %str%
   IfMsgBox TryAgain
      ExitApp
   else IfMsgBox Cancel
      blnSkip### := true
   else
      return
}



GUID()         ; 32 hex digits = 128-bit Globally Unique ID
; Source: Laszlo in http://www.autohotkey.com/board/topic/5362-more-secure-random-numbers/
{
   format = %A_FormatInteger%       ; save original integer format
   SetFormat Integer, Hex           ; for converting bytes to hex
   VarSetCapacity(A,16)
   DllCall("rpcrt4\UuidCreate","Str",A)
   Address := &A
   Loop 16
   {
      x := 256 + *Address           ; get byte in hex, set 17th bit
      StringTrimLeft x, x, 3        ; remove 0x1
      h = %x%%h%                    ; in memory: LS byte first
      Address++
   }
   SetFormat Integer, %format%      ; restore original format
   Return h
}




ASCII()
{
   s := "ASCII TABLE`n`n"
   i := 33
   Loop
   {
       s := s . Chr(i) . " (" . i . ")"
       if Mod(i, 8)
           s := s . " "
       else
           s := s . "`n"
       i := i + 1
       if (i = 128)
       {
           i := 161
           s := s . "`n`n"
       }
       if i > 255
           break
   }
   return s
}
