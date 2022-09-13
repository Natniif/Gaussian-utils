# A script to collect all the cool quotes at the end of gaussian files 
# Takes an input from the terminal as follows: python3 quote_collector.py GAUSSIAN_OUTPUT_FILE_HERE.out 
# Will either create a new txt file called quassian_quotes.txt
import sys

gauss_file = open(sys.argv[1], 'r')
quotes_file = open('gaussian_quotes.txt', 'a')

lines = gauss_file.readlines()

quote = []
count = 1
for line in lines:
    line1 = ' '.join(line.rstrip().split())
    line = line1.strip(' ')
    #for i in range(len(line)):
    if len(line) > 0 and line[-1] =='@':
        for j in range(count+1,len(lines)):
            line1 = lines[j].split() 
            if "Job" in line1:
                quotes_file.write('\n')
                quotes_file.write('\n') 
                break 
            elif len(line1) ==0:
                pass
            else:
                sentence = ' '.join(line1)
                quote.append(sentence)
                quotes_file.write(sentence)
                quotes_file.write('\n')
        break
    else:
        count += 1
output = ' '.join(quote)
#print(output)

gauss_file.close()        


quotes_file.close()
