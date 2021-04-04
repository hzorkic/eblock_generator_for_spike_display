# eblock generator for spike display
### A tool to generate custom SARS-CoV-2 spike mutation eblocks as described in Javanmardi et al. (2020) Rapid characterization of spike variants via mammalian cell surface display. 

### Input:
You absolutely **MUST** upload an .xlsx fiel with the following format:
![input example](https://user-images.githubusercontent.com/56274447/113521071-c89a0900-955c-11eb-8672-267904fe9ad9.png)

Keep the following part cut offs in mind when selecting your multiplexed mutation strings:


If using WINDOWS:
- press "Windows" + "R" keys together
- search "cmd" and click enter
- type "cd C:\Users\[user name used on your computer]\Downloads\eblock_generator_for_spike_display" or "cd [insert filepath to repo]"
- type "Rscript eblock.R [input file name.xlsx]

If all else fails:
- download RStudio
- change the file name in line 23 to your .xlsx input file name
- click run in thr top right corner of the page
