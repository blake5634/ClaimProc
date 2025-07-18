import os
import sys
import csv
import xlsx2csv as x2c

def error(msg):
    print('Error:',msg)
    print('Usage: ')
    print(' > proClaims  file.xls')
    quit()

def main():
    if len(sys.argv) != 2:
        print("Usage: python procClaims.py input_file.xlsx")
        print("Creates a csv compatible with HealthBills pasting")
        return

    input_file = sys.argv[1]

    if x2c.xlsx_to_csv(input_file):
        print(f"{input_file} converted to .csv")
    else:
        print("File conversion failed!")

    csv_file = os.path.splitext(input_file)[0] + '.csv'

    # select and convert cols to HealthBills cols Format

    try:
        fp = open(csv_file,'r')
    except:
        error(f"Couldn't open file for input:  {csv_file}")
    rows = csv.reader(fp, delimiter=',')
    claims = []
    for r in rows:
        claims.append(r)

    print(f'Successfully read in {len(claims)-3} claims.')
    oldheaders = claims[1]
    for r in claims[2:]:
        for c in r:  # get rid of Grand total lines
            if c.startswith('Grand'):
                claims.remove(r)

    outheaders = ['DOS', 'Bill Date', 'Provider', 'Description', 'Organiz.',
                  'EOB Resp.', 'Bill Amt.', 'Paymt', 'E/P', 'Pymt Date',
                  'Deductable', 'Balance', 'Claim #', 'CoPay', 'Coinsur.']

    known_patients = ['Blake','Cynthia']
    pdict = {}
    for p in known_patients:
        pdict[p] = []
    for r in claims[2:]:
        cnum = r[0]
        DOS  = r[1]
        stat = r[4]
        prov = r[5]
        patient = r[6]
        p1 = patient.replace('Ruggeiro','')
        p2 = p1.replace('Hannaford','')
        patient = p2.strip()
        amt = r[7]      # '$123.45'
        presp = r[8]
        ded = r[9]
        copay = r[10]
        coins = r[11]

        if stat.startswith('Pending'):  # only get completed claims
            continue

        newrow = []
        newrow.append(DOS)
        newrow.append('')    # no bill date for claims
        newrow.append(prov)
        newrow.append('')    # no descrip for claims(!>W*#@() *)
        newrow.append('')    # no org for claims
        newrow.append(presp)
        newrow.append('')    # no bill amt for claims
        newrow.append('')    # no payment info for claims
        newrow.append('')    # no payment info for claims
        newrow.append('')    # no payment info for claims
        newrow.append(ded)
        newrow.append('')    # deductable balance
        newrow.append(cnum)
        newrow.append(copay) # co-pay
        newrow.append(coins) #co-insurance
        pdict[patient].append(newrow)  # separate claims by patient


    outfile = 'NewSheetforHealthBills.csv'
    try:
        ofp = open(outfile,'w')
    except:
        error(f'Couldnt open {outfile} for output')

    writer = csv.writer(ofp,delimiter=',',quotechar='"',quoting=csv.QUOTE_MINIMAL)
    writer.writerow(outheaders)
    for p in known_patients:
        namerow = [p]
        writer.writerow(namerow)
        writer.writerows(pdict[p])
    ofp.close()
    print(f'\n\n Job completed. Open new data in {outfile}\n')

if __name__ == "__main__":
    main()
