
$asm
Mov Sp , #10h
Mov P3 , #0ffh
Mov P1 , #0ffh
Mov Scon , #52h
Mov Tmod , #25h
Mov Th1 , #0f8h
Setb TCON.6
Lcall 0150h
Lcall 0250h
Lcall 01e0h
Cjne A , #0a0h , 001ah
Mov R0 , #30h
Lcall 01e0h
Cjne A , #20h , 002ah
Mov A , #3ah
Mov @r0 , A
Inc R0
Cjne A , #0dh , 0022h
Lcall 015ch
Lcall 02f0h
Ljmp 0100h
Add A , #01h
MOVC A , @A + Pc
Ret

Addc A , R7
Inc @r0
Anl A , R3
Orl A , R7
Xrl A , @r0
Xrl A , R5
Mov R5 , #07h
Mov R7 , #6fh
Nop
Nop
Nop
Mov R5 , #08h
Mov R6 , #0ffh
Mov R7 , #0ffh
Jb Ri , 0114h
Nop
Nop
Djnz R7 , 0106h
Djnz R6 , 0104h
Djnz R5 , 0102h
Ljmp 002fh
Clr Ri
Mov A , Sbuf
Cjne A , #0a0h , 010bh
Clr Ri
Ljmp 0020h
Nop
Nop
Nop
Nop
Nop
Nop
Nop
Nop
Nop
Nop
Nop
Nop
Clr P1.7
Mov R2 , #0ffh
Clr P1.6
Setb P1.6
Djnz R2 , 0160h


Ret


Jnb Ri , 01e0h
Clr Ri
Mov A , Sbuf

Ret

Nop
Mov R0 , #30h
Mov A , #0ffh
Mov @r0 , A
Inc R0
Mov A , R0
Cjne A , #7fh , 0253h


Ret

$asm
Push 00h
Nop
Nop
Dec R0
Nop
Dec R0
Nop
Clr C
Mov A , @r0
Subb A , #30h
Nop
Lcall 00e0h
Lcall 0350h
Dec R0
Mov A , R0
Nop
Nop
Nop
Cjne A , #2fh , 02f8h
Pop 00h

$end Asm

Ret
$asm
Mov R2 , #08h
Nop
Nop
Rlc A
Nop
Mov P1.7 , C
Setb P1.6
Clr P1.6
Djnz R2 , 0354h

$end Asm
Ret
