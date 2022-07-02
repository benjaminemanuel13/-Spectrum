
Start:
  org #8000
  
ProgramLoop:
  di
  ld a, 1
  ld (0xe002), a
  ld (0xe003), a

  ld a, 0
  ld (0xe004), a
  
  //call FillBuffer

AfterInit:
  ld a, 1
  call ShowPacman

  call FillScreen

  ld a, 0
  call ShowPacman
  
  call CheckKeys

  jp AfterInit