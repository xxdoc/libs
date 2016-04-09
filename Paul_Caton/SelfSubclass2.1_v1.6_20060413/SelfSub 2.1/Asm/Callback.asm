;*****************************************************************************************
;** Callback.asm - Generic Class/Form/UserControl callback thunk. Assemble with nasm.
;**
;** Paul_Caton@hotmail.com
;** Copyright free, use and abuse as you see fit.
;**
;** v1.0 The original............................................................ 20060408
;** v1.1 IDE safe................................................................ 20060409
;** v1.1 Validate that the callback address is live code......................... 20060413
;*****************************************************************************************

;***********************
;Definitions
%define objOwner        [ebx]       ;Owner object address
%define addrCallback    [ebx +  4]  ;Callback address
%define fnEbMode        [ebx +  8]  ;EbMode address
%define fnIsBadCodePtr  [ebx + 12]  ;EbMode address
%define RetValue        [ebp -  4]  ;Callback/thunk return value

[bits 32]
;************
;Data storage
    dd_objOwner         dd 0        ;Owner object address
    dd_addrCallback     dd 0        ;Callback address
    dd_fnEbmode         dd 0        ;EbMode address
    dd_fnIsBadCodePtr   dd 0        ;IsBadCodePtr address
    
;***********
;Thunk start    
    mov     eax, esp                ;Get a copy of esp as it is now
    pushad                          ;Push all the cpu registers on to the stack
    mov     ebx, 012345678h         ;Address of the data, patched from VB
    mov     ebp, eax                ;Set ebp to point to the return address
    
    push    dword addrCallback      ;Callback address
    call    fnIsBadCodePtr          ;Call IsBadCodePtr
    jnz     _return                 ;If the callback code isn't live, return
    
    cmp     fnEbMode, dword 0       ;Are we running in the IDE?
    jnz     _ide_state              ;Check the IDE state

_ide_running:    
    mov     eax, ebp                ;Copy the stack frame pointer
    sub     eax, 4                  ;Address of the callback/thunk return value
    push    eax                     ;Push the return value address
    nop
    
    mov     ecx, 012345678h         ;Parameter count into ecx, patched from VB
    jecxz   _callback               ;If parameter count = 0, skip _parameter_loop
    
_parameter_loop:
    push    dword [ebp + ecx * 4]   ;Push parameter
    loop    _parameter_loop         ;Decrement ecx, if <> 0 jump to _parameter_loop

_callback:    
    push    dword objOwner          ;Owning object
    call    addrCallback            ;Make the callback
    
_return:        
    popad                           ;Restore registers
    nop
    ret     01234h                  ;Return, the number of esp stack bytes to release is patched from VB
    dw      0
_ide_state:                         ;Running under the VB IDE
    call    near fnEbMode           ;Determine the IDE state

    cmp     eax, dword 1            ;If EbMode = 1
    je      _ide_running            ;Running normally
    
    xor     eax, eax                ;Zero eax
    mov     RetValue, eax           ;Set the return value
    jmp     _return                 ;Outta here