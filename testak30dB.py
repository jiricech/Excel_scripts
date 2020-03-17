import sys
import math

parse = False
sum = 0.0

if len(sys.argv) != 2:
    print("Usage: testak.py <FILE.trc>")
    sys.exit(1)

with open(sys.argv[1]) as f_data:
    line_id = 1
    data_line_cnt = 0
    fil_line_cnt = 0

    for line in f_data:
        if not parse:
            if line.startswith("A-X/512"):
                    parse = True
        else:
            cols = line.split("\t")
            if len(cols) < 11:
                print("Wrong cols number on line" + str(line_id) + ";expected: 11, got: " + str(len(cols)))
            else:
                data_line_cnt += 1
                freq = float(cols[0])
                vol = float(cols[1])
                if (freq > 20000.0 and freq < 80000.0):
                    fil_line_cnt += 1
                    sum += vol * vol                    

        line_id += 1
    print("Data lines: " + str(data_line_cnt))
    print("Filtered lines:" + str(fil_line_cnt))

res = 20*math.log10(math.sqrt(sum)/(0.0995*0.00002))
print("Result: " + str(res))

