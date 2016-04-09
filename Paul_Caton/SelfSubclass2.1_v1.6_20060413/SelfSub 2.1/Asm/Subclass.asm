;*****************************************************************************************
;** Subclass.asm - subclassing thunk. Assemble with nasm.
;**
;** Paul_Caton@hotmail.com
;** Copyright free, use and abuse as you see fit.
;**
;** v1.0 Re-write of the SelfSub/WinSubHook-2 submission to Planet Source Code... 20060322
;** v1.1 Thunk redesigned to handle memory release............................... 20060325
;** v1.2 Optimized layout so that all jumps are short............................ 20060327
;** v1.3 Optimized layout for better branch prediction & out-of-order execution.. 20060328
;** v1.4 Added the lParamUser user-defined callback parameter.................... 20060411
;*****************************************************************************************

;***************
;API definitions
%define GWL_WNDPROC     -4          ;SetWindowsLong WndProc parameter
%define M_RELEASE       8000h       ;VirtualFree memory release flag

;******************************
;Stack frame access definitions
%define lParam          [ebp + 48]  ;WndProc lParam
%define wParam          [ebp + 44]  ;WndProc wParam
%define uMsg            [ebp + 40]  ;WndProc uMsg
%define hWnd            [ebp + 36]  ;WndProc hWnd
%define lRetAddr        [ebp + 32]  ;Return address of the code that called us
%define lReturn         [ebp + 28]  ;lReturn local, restored to eax after popad
%define bHandled        [ebp + 20]  ;bHandled local, restored to edx after popad

;***********************
;Data access definitions
%define nCallStack      [ebx]       ;WndProc call stack counter
%define bShutdown       [ebx +  4]  ;Shutdown flag
%define lWnd            [ebx +  8]  ;Window handle
%define fnEbMode        [ebx + 12]  ;EbMode function address
%define fnCallWinProc   [ebx + 16]  ;CallWindowProc function address
%define fnSetWinLong    [ebx + 20]  ;SetWindowsLong function address
%define fnVirtualFree   [edx + 24]  ;VirtualFree function address (deliberately edx)
%define fnIsBadCodePtr  [ebx + 28]  ;IsBadCodePtr function address
%define objOwner        [ebx + 32]  ;Owner object address
%define addrWndProc     [ebx + 36]  ;Original WndProc address
%define addrCallback    [ebx + 40]  ;Callback address
%define addrTableB      [ebx + 44]  ;Address of before original WndProc message table
%define addrTableA      [ebx + 48]  ;Address of after original WndProc message table
%define lParamUser      [ebx + 52]  ;User defined callback parameter

[bits 32]
;************
;Data storage
    dd_nCallStack       dd 0        ;WndProc call stack counter
    dd_bShutdown        dd 0        ;Shutdown flag
    dd_lWnd             dd 0        ;Window handle
    dd_fnEbMode         dd 0        ;EbMode function address
    dd_fnCallWinProc    dd 0        ;CallWindowProc function address
    dd_fnSetWinLong     dd 0        ;SetWindowsLong function address
    dd_fnVirtualFree    dd 0        ;VirtualFree function address
    dd_fnIsBadCodePtr   dd 0        ;ISBadCodePtr function address
    dd_objOwner         dd 0        ;Owner object address
    dd_addrWndProc      dd 0        ;Original WndProc address
    dd_addrCallback     dd 0        ;Callback address
    dd_addrTableB       dd 0        ;Address of before original WndProc message table
    dd_addrTableA       dd 0        ;Address of after original WndProc message table
    dd_lParamUser       dd 0        ;User defined callback parameter
    
;***********
;Thunk start    
    xor     eax, eax                ;Zero eax, lReturn in the ebp stack frame
    xor     edx, edx                ;Zero edx, bHandled in the ebp stack frame
    pushad                          ;Push all the cpu registers on to the stack
    mov     ebp, esp                ;Setup the ebp stack frame
    mov     ebx, 012345678h         ;Address of the data, patched from VB
    xor     esi, esi                ;Zero esi

    cmp     fnEbMode, eax           ;Check if the EbMode address is set
    jnz     _ide_state              ;Running in the VB IDE

_before:                            ;Before the original WndProc
    dec     edx                     ;edx <> 0, bBefore calkback parameter = True
    mov     edi, addrTableB         ;Get the before message table
    call    _callback               ;Attempt the VB callback
    
    cmp     bHandled, esi           ;If bHandled <> False
    jne     _return                 ;The callback entirely handled the message

_original_wndproc:
    call    _wndproc                ;Call the original WndProc
    
_after:                             ;After the original WndProc
    xor     edx, edx                ;Zero edx, bBefore calkback parameter = False
    mov     edi, addrTableA         ;Get the after message table
    call    _callback               ;Attempt the VB callback

_return:                            ;Clean up and return to caller
    popad                           ;Pop all registers. lReturn is popped into eax
    ret     16                      ;Return with a 16 byte stack release

_ide_state:                         ;Running under the VB IDE
    call    near fnEbMode           ;Determine the IDE state

    cmp     eax, dword 1            ;If EbMode = 1
    je      _before                 ;Running normally
    
    test    eax, eax                ;If EbMode = 0
    jz      _shutdown               ;Ended, shutdown

    call    _wndproc                ;EbMode = 2, breakpoint... call original WndProc
    jmp     _return                 ;Return
    
_wndproc:                           ;Call the original WndProc
    push    dword lParam            ;ByVal lParam
    push    dword wParam            ;ByVal wParam
    push    dword uMsg              ;ByVal uMsg
    push    dword hWnd              ;ByVal hWnd
    push    dword addrWndProc       ;ByVal Address of the original WndProc
    inc     dword nCallStack        ;Increment the WndProc call counter
    call    near fnCallWinProc      ;Call CallWindowProc
    mov     lReturn, eax            ;Save the return value
    dec     dword nCallStack        ;Decrement the WndProc call counter
    jnz     _generic_ret            ;Original WndProc call stack is recursed
    
_check_shutdown 
    cmp     bShutdown, esi          ;If bShutdown flag = 0
    jz      _generic_ret            ;Return

    pop     eax                     ;Eat the call return address
    
_shutdown:                          ;Restore the original WndProc
    push    dword addrWndProc       ;Address of the original WndProc
    push    dword GWL_WNDPROC       ;WndProc index
    push    dword lWnd              ;Push the window handle
    call    fnSetWinLong            ;Call SetWindowsLong

_free_memory:                       ;Free the memory this code is running in.... tricky
    mov     uMsg, ebx               ;VirtualFree param #1, start address of this memory
    mov     wParam, esi             ;VirtualFree param #2, 0
    mov     lParam, dword M_RELEASE ;VirtualFree param #3, memory release flag
    mov     eax, lRetAddr           ;Return address of the code that called this thunk
    mov     bHandled, ebx           ;ebx popped to edx after the popad instruction
    mov     hWnd, eax               ;Return address to the code that called this thunk
    popad                           ;Restore the registers
    add     esp, 4                  ;Adjust the stack to point to the new return address
    jmp     fnVirtualFree           ;Jump to VirtualFree, ret to the caller of this thunk
    
_callback:                          ;Validate the callback
    mov     ecx, [edi]              ;ecx = table entry count
    jecxz   _generic_ret            ;ecx = 0, table is empty

    test    ecx, ecx                ;Set the flags as per ecx
    js      _call                   ;Table entry count is negative, all messages callback

    add     edi, 4                  ;Inc edi to point to the start of the callback table
    mov     eax, uMsg               ;eax = the value to scan for
    repne   scasd                   ;Scan the callback table for uMsg
    jne     _generic_ret            ;uMsg not in the callback table

_call:                              ;Callback required, do it...
    push    dword addrCallback      ;Push the callback address
    call    fnIsBadCodePtr          ;Check the code is live
    jnz     _generic_ret            ;If not, skip callback
    
    lea     eax, lParamUser         ;Address of lParamUser
    lea     ecx, bHandled           ;Address of bHandled
    push    eax                     ;ByRef lParamUser
    lea     eax, lReturn            ;Address of lReturn
    push    dword lParam            ;ByVal lParam
    push    dword wParam            ;ByVal wParam
    push    dword uMsg              ;ByVal uMsg
    push    dword lWnd              ;ByVal hWnd
    push    eax                     ;ByRef lReturn
    push    ecx                     ;ByRef bHandled
    push    edx                     ;ByVal bBefore
    push    dword objOwner          ;ByVal the owner object
    call    near addrCallback       ;Call the zWndProc callback procedure
    
_generic_ret:                       ;Shared return
    ret