import os
import sys
import pandas as pd

print()
print("Create an xlsx file with the following format for your spike variants:")
print("   1. Do NOT add a header/column names.")
print("   2. List each mutation as Orignial Amino Acid + Positoin + Desired Mutaiton")
print("          example: D614G would mutate a sequecne at position 614 with a G AA.")
print("          example: D614- would cause a DELETION at position 614.")
print("   3. For combinatorial mutations, create a list of each mutation followed by a ', '")
print("          example: D614G, A666V, A667V.")
print("   4. Each row in the xlsx represents one seqeunce. ")
print("   5. Your file needs to have a .xlsx extension. avoid spaces or '\\' '/' in your file name")
print()


input_file = str(input("What is the filepath to your excel spreadsheet?\nexample C:/Research/eblock_generator_for_spike_display/04032021.xlsx: "))
df = pd.read_excel(input_file, header= None, index_col= False) # can also index sheet by name or fetch all sheets
df = df.drop(index = 0).values.tolist()
mutations = df


# function to mutate position
def mutate_pos(sequence, positionNT, mutation):
    """Return Gene with a mutation to base at position pos."""
    if (mutation == "A") and (alanine_check(sequence, positionNT) == True):
        sequence1 = "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
        error = "You are trying to mutate an position that is already an Alanine into an Alanine"
        return sequence1, error
    elif mutation == "-":
        sequence1 = sequence[:positionNT-1] + "---" + sequence[positionNT:]
        error = "None"
        return sequence1, error
    else:
        base = optimize_odon(mutation)
        sequence1 = sequence[:positionNT-1] + base + sequence[positionNT:]
        error = "None"
        return sequence1, error


# function to change target AA mutation to nucleotides 
def optimize_odon(AA):
  if (AA == "A"):
    codon = "gcc"

  elif (AA == "R"):
    codon = "cgg"

  elif (AA == "N"):
    codon = "aac"

  elif (AA == "D"):
    codon = "gac"

  elif (AA == "C"):
    codon = "tgc"

  elif (AA == "E"):
    codon = "gag"

  elif (AA == "Q"):
    codon = "cag"

  elif (AA == "G"):
    codon = "ggc"

  elif (AA == "H"):
    codon = "cac"

  elif (AA == "I"):
    codon = "atc"

  elif (AA == "L"):
    codon = "ctg"

  elif (AA == "K"):
    codon = "aag"

  elif (AA == "M"):
    codon = "atg"

  elif (AA == "F"):
    codon = "ttc"

  elif (AA == "P"):
    codon = "ccc"

  elif (AA == "S"):
    codon = "tcc"

  elif (AA == "T"):
    codon = "acc"

  elif (AA == "W"):
    codon = "tgg"

  elif (AA == "Y"):
    codon = "tac"

  elif (AA == "V"):
    codon = "gtg"

  elif (AA == "-"):
    codon = "---"

  else:
    codon = "ERROR"

  return (codon)


# function to see if alanine already exsits at the target position
def alanine_check(refseq, positionAA):
    if refseq[positionAA:(positionAA+2)] == "gct" or\
       refseq[positionAA:(positionAA+2)] == "gcc" or\
       refseq[positionAA:(positionAA+2)] == "gca" or\
       refseq[positionAA:(positionAA+2)] == "gcg" :
        return True
    else:
        return False
    
    
# function to check for BsmB1 and Aar1 cut sites
def cut_site_check(refseq):
    if ("cacctgc" in refseq) or ("cgtccac" in refseq):
        error = "There is a BsmBI or AarI site in this sequence"
    else:
        error = "None"
    return error


# fucntion to add 5' 3' overhangs
def add_overhangs(refseq, partprime3, partprime5):  
    return partprime5 + refseq + partprime3


# fucntion to find part
def part_check(positionAA):
    if 1 <= positionAA & positionAA <= 135:
        # part 1a
        partstart = 1
        partend = 135
        partrefseq = "atgttcgtgttcctggtgctcctgcctctggtgagcagccagtgcgtgaacctgaccacccgaacccagctcccaccagcctacaccaacagctttacacggggcgtgtactaccctgacaaggtgttcagatctagcgtcctgcacagcactcaggacctcttcctgccgttcttcagcaacgtgacatggttccacgccatccacgtgagcggcacaaacggaaccaagcggtttgataaccccgtcctgccattcaatgatggagtttacttcgccagtaccgagaagagtaacatcatccggggctggatcttcggcaccaccctggatagcaaaacacagagcctcctgatcgtgaacaatgccacgaacgtcgtgatcaaggtgtgcgagttccagtttt"
        partprime5 = "gcatcgtctcatcggcacctgccacctgac"
        partprime3 = "gcaaggtggcaggtggacctgagacggcat"
        positionNT = 3*(positionAA)-2
        error = "None"

    elif 138 <= positionAA & positionAA <= 265:
        #part 1b
        partstart = 138
        partend = 265
        partrefseq = "tgatcctttcctgggtgtgtactaccacaagaacaacaagagctggatggaaagcgagttcagagtctacagcagcgccaacaactgcacattcgagtacgtctctcagccttttctgatggaccttgaggggaaacaaggcaacttcaagaacctgagagaattcgtgttcaagaacatcgacggctacttcaaaatctactccaagcacacacccatcaacctggtccgggacctccctcagggcttcagcgccctggaacccctggtcgacctgcccataggcatcaacataacgcggttccaaaccctgctggccctgcatagatcctacctgactcctggcgacagcagcagcggatggaccgccggagctgcagcctacta"
        partprime5 = "gcatcgtctcatcggcacctgccaccgcaa"
        partprime3 = "tgtgggtggcaggtggacctgagacggcat"
        positionNT = (3*(positionAA)-2)-410
        error = "None"
        
    elif 136 <= positionAA & positionAA <= 137:    
        #Part 1
        partstart = 136
        partend = 137
        partrefseq = "ggtggcaggtggaaagtgaaacgtgatttcatgcgtcattttgaacattttgtaaatcttatttaataatgtgtgcggcaattcacatttaatttatgaatgttttcttaacatcgcggcaactcaagaaacggcaggttcggatcttagctactagagaaagaggagaaatactagatgcgtaaaggcgaagagctgttcactggtgtcgtccctattctggtggaactggatggtgatgtcaacggtcataagttttccgtgcgtggcgagggtgaaggtgacgcaactaatggtaaactgacgctgaagttcatctgtactactggtaaactgccggttccttggccgactctggtaacgacgctgacttatggtgttcagtgctttgctcgttatccggaccatatgaagcagcatgacttcttcaagtccgccatgccggaaggctatgtgcaggaacgcacgatttcctttaaggatgacggcacgtacaaaacgcgtgcggaagtgaaatttgaaggcgataccctggtaaaccgcattgagctgaaaggcattgactttaaagaggacggcaatatcctgggccataagctggaatacaattttaacagccacaatgtttacatcaccgccgataaacaaaaaaatggcattaaagcgaattttaaaattcgccacaacgtggaggatggcagcgtgcagctggctgatcactaccagcaaaacactccaatcggtgatggtcctgttctgctgccagacaatcactatctgagcacgcaaagcgttctgtctaaagatccgaacgagaaacgcgatcatatggttctgctggagttcgtaaccgcagcgggcatcacgcatggtatggatgaactgtacaaatgaccaggcatcaaataaaacgaaaggctcagtcgaaagactgggcctttcgttttatctgttgtttgtcggtgaacgctctctactagagtcacactggctcaccttcgggtgggcctttctgcgtttatacacctgccacc"
        partprime5 = "gcatcgtctcatcggcacctgccacctgac"
        partprime3 = "tgtgggtggcaggtggacctgagacggcat"
        positionNT = 3*(positionAA)-2
        error = "None"
    
    elif 268 <= positionAA & positionAA <= 521:    
        #Part 2
        partstart = 268
        partend = 521
        partrefseq = "ggctacctgcaacctagaaccttcctgctgaagtacaacgagaacggcacaatcacagacgccgtcgactgcgccctggaccctctctctgagacaaagtgcaccctgaagtccttcaccgtggaaaagggcatctaccagaccagcaacttccgggtgcagcctacagagagcatcgtgcgatttccaaacattaccaacctctgccccttcggcgaggtgtttaacgccacaagatttgcctccgtttacgcctggaatagaaagagaatcagcaattgtgtggccgactactccgtgctgtataacagcgcctctttcagcaccttcaagtgctacggcgtttccccaacaaagctgaatgacctgtgcttcaccaacgtgtacgccgactccttcgtaattagaggcgatgaggtgcggcagatcgcaccaggccagaccggtaagatcgctgactacaactataagctgcctgatgattttacaggctgcgtgatcgcctggaactctaacaacctggatagcaaggtgggcggcaactacaactacctgtaccggctgtttcgcaagtctaacctgaaacctttcgagagagacatctccacagagatctaccaggccggttctacaccttgtaacggggtggaaggcttcaactgttacttccctctgcaaagctacggcttccagcctaccaatggagtcggctaccagccataccgggtggtcgtgctgtccttcgagttactccacgccccc"
        partprime5 = "gcatcgtctcatcggcacctgccacctgtg"
        partprime3 = "gccaggtggcaggtggacctgagacggcat"
        positionNT = (3*(positionAA)-2)-801
        error = "None"
        
    elif 524 <= positionAA & positionAA <= 786:
        #Part 3
        partstart = 524
        partend = 786
        partrefseq = "ccgtctgcggtcctaagaagtccaccaatctggttaagaacaaatgcgtgaacttcaacttcaacggcctgaccgggaccggcgtgctgaccgaaagcaacaaaaagttcctccccttccagcagttcggccgtgatatcgctgacaccacagatgccgtcagagatccacagaccctggaaatcctggatattacaccctgctccttcggaggagtttctgtgatcacccccgggaccaataccagcaaccaggtggctgtgctgtaccaaggtgttaactgcaccgaggttcctgtggccatccacgccgatcagctgacacctacttggagagtgtactccactggctccaatgtgttccagaccagggccggatgtctgatcggcgccgagcacgtgaataacagttacgagtgcgacatccctatcggcgccggcatctgtgccagctaccagacccagacaaacagccctgggtctgcttcctctgtagctagccagagcatcatcgcctacaccatgagcctgggcgcagagaacagcgtggcctattccaacaactctatcgccattcccaccaactttacaattagcgtcacaacagagatcctgcccgtgagcatgaccaagaccagcgtggactgtacaatgtacatctgtggcgacagcactgaatgcagcaacctgctgctgcaatacggctccttttgcacccaactgaaccgggcgctgaccggaatcgccgtggaacaggacaaaaatacccaggaggtgttcgcccaagtgaagca"
        partprime5 = "gcatcgtctcatcggcacctgccaccgcca"
        partprime3 = "gatcggtggcaggtggacctgagacggcat"
        positionNT = (3*(positionAA)-2)-1567
        error = "None"
        
    elif 789 <= positionAA & positionAA <= 1053:
        #Part 4
        partstart = 789
        partend = 1053
        partrefseq = "tacaagaccccacctatcaaggacttcggcggctttaactttagccagattctccctgatccttctaagcctagcaagcggagccctatcgaggatctgctgttcaacaaggtcaccctggccgatgccggctttatcaaacagtatggcgattgcctgggcgacatagccgccagagatctgatctgcgcccagaaattcaacggcctgacagttctcccacctctgctgaccgacgagatgatcgctcagtacacctctgccctgctggctggcaccatcacatctgggtggacatttggcgccggccccgccctgcagatcccctttcccatgcagatggcctatagattcaacggaatcggcgtgacccagaacgtgctgtatgaaaaccagaagctgatcgctaaccagttcaattctgccatcggcaagatccaggactccctctcctctacccccagcgccctgggcaaactgcaggacgtggtgaatcagaacgcccaagccctgaacaccctggtgaagcagctcagcagcaattttggcgccatcagctctgtgctgaacgatatcctgtctagactggaccctccagaagccgaagtccagatcgatagactgatcacaggcagactgcagtccctgcaaacctacgtgacccaacagctgatcagggccgctgaaataagagccagcgccaatctcgccgctaccaagatgtccgagtgtgtgctgggacagtctaaacgcgttgacttctgcggcaaaggctatcacctgatgagcttcccc"
        partprime5 = "gcatcgtctcatcggcacctgccaccgatc"
        partprime3 = "cagaggtggcaggtggacctgagacggcat"
        positionNT = (3*(positionAA)-2)-2364
        error = "None"

    elif 1056 <= positionAA & positionAA <= 1206:
        #Part 5
        partstart = 1056
        partend = 1206
        partrefseq = "gcgcgccgcacggcgtggtgttcctgcatgtgacatacgtgcctgcccaagagaagaatttcacaaccgcccctgccatctgccacgacggcaaggcccacttcccaagagagggcgttttcgtttccaatggcacacactggttcgtgacacaaagaaacttctacgaaccccagattatcaccaccgacaacaccttcgtgagtggcaattgtgacgtggtcatcggaatcgtgaacaacacagtgtacgaccctctgcaacctgagctggactcttttaaggaagagctggacaagtactttaaaaaccacaccagccccgatgtggacctgggcgacatcagtggcattaacgccagcgtggtgaacatccaaaaggaaatcgacagactgaacgaggtggccaagaacctgaacgagtccctgatcgacctgcaggagctcggcaaatacga"
        partprime5 = "gcatcgtctcatcggcacctgccacccaga"
        partprime3 = "gcagggtggcaggtggacctgagacggcat"
        positionNT = (3*(positionAA)-2)-3163
        error = "None"
        
    else:
        partstart = 0
        partend = 0
        partrefseq = ""
        partprime5 = ""
        partprime3 = ""
        positionNT
        error = "this mutation lies on a part overhang"
    
    return partstart, partend, partrefseq, partprime5, partprime3, positionNT, error
    
    
# mutation factory   
def main(mutations):
    output = pd.DataFrame(columns= ["Mutations", "Sequence", "Error"])
    df = []
    for row in range(len(mutations)):
        list_of_mutations = mutations[row] 
        row_muts = mutations[row][0].split(", ")
        count = 0 
        errors = []
        for mutation in row_muts:
            original, positionNT, mutation = mutation[0], int(mutation[1:-1]), mutation[-1]
            # if its out first mutation, grab the part and associated infromation for the part its in
            if count == 0:
                partstart = part_check(positionNT)[0]
                partend = part_check(positionNT)[1]
                partrefseq = part_check(positionNT)[2]
                partprime5 = part_check(positionNT)[3]
                partprime3 = part_check(positionNT)[4]
                errors.append(part_check(positionNT)[6])
                refseq = partrefseq
            positionAA = part_check(positionNT)[5]
            # mutate our refseq!
            refseq = mutate_pos(refseq, positionAA, mutation)[0]
            errors.append(mutate_pos(refseq, positionAA, mutation)[1])
            count += 1 
            
        # add overhangs
        refseq = add_overhangs(refseq, partprime3, partprime5)
        # remove deletions
        refseq.replace("---", "")
        # check for cutsites
        errors.append(cut_site_check(refseq))
        
        df.append([", ".join(list_of_mutations), refseq, ", ".join(errors)])
    #print(*df, sep="\n")
    # make
    output = pd.DataFrame(df, columns = ["Mutations", "Sequence", "Errors"])
    #name = "SpikeVariantMutaions" + today.strftime("mdY")
    #print(name)
    output.to_excel("spike_variants.xlsx", index = False)
    print("Your file has been saved in your current dirctory as 'spike_variants.xlsx' !")
            
        
            
            
            
            

main(mutations)