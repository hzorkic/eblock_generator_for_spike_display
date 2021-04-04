# eblock generator for spike display
### A tool to generate custom SARS-CoV-2 spike mutation eblocks as described in Javanmardi et al. (2020)

### Input:
You absolutely **MUST** upload an .xlsx fiel with the following format:

![input example](https://user-images.githubusercontent.com/56274447/113521071-c89a0900-955c-11eb-8672-267904fe9ad9.png)

Keep the following part cut offs in mind when selecting your multiplexed mutation strings:

![cutoffs](https://user-images.githubusercontent.com/56274447/113521302-385cc380-955e-11eb-91ff-e1d95269a05f.png)

### Examples:
1. you have mutations you want to alter at positions 5, 120, 135, and 600 in the same sequence. 
- you cannot change all of these within the same sequences as they fall in diffrenet parts. you would include two lines in your .xlsx: 
  - "L5F, A120H, F135K" *in order* which are in part 1a
  - "G600W" which is in part 3
2. you have mutations you want to alter at position 135, 136 in the same sequence.
- you would include one line in your .xlsx:
  - "F135K, F136G" *in order* which is apart of part 1 (not 1a or 1b)
3. you have mutations you want to alter at positions 5, 120, 135, 600 each in different sequences. 
- you would include four lines in your .xlsx
  - "L5F"
  - "A120H"
  - "F135K"
  - "G600W"

### Output:
A .xlsx file containing the original list of mutations, the mutated sequences with their respective part overhangs, a column of notes, and a column with the length of the sequence. 
For our project, we took the mutated sequences and ordered them on IDT as eblocks. 


### How to run:
If using WINDOWS:
- press "Windows" + "R" keys together
- search "cmd" and click enter
- type "cd C:\Users\[user name used on your computer]\Downloads\eblock_generator_for_spike_display" or "cd [insert filepath to repo]"
- type "Rscript eblock.R [input file name.xlsx]

If all else fails:
- download RStudio
- change the file name in line 23 to your .xlsx input file name
- click run in thr top right corner of the page
