import pdfquery
import sqlite3
import glob
import Tkinter, tkFileDialog


def scrape_OPR800 (filename):
    pdf = pdfquery.PDFQuery(filename)
    pages = pdf.doc.catalog['Pages'].resolve()['Count']
    print ("pdf has %d pages" % pages)
    totals = []
    date = None
    for i in range((pages/10)+1):
        try:
            # load the next 10 pages
            pdf = pdfquery.PDFQuery(filename)
            pdf.load(range(i*10, min((i+1)*10, pages)))

            # if we don't already have the date, search for it
            if date is None:
                date_field = "EI Date To: "
                date = pdf.pq("LTTextBoxHorizontal:contains('%s')" % date_field)[0].text
                date = date[date.index(date_field)+len(date_field):]
                if " " in date: date = date[:date.index(" ")]
                print ("got date (%s)" % date)

            # search for the totals and append them    
            total = pdf.pq('LTTextLineHorizontal:contains("Total for")')
            for t in total:
                totals.append(t.text)
                
            print ("finished page %d" % (min((i+1)*10, pages)))
            del pdf
        except Exception as exc:
            print ("error in %s on page %s (%s)" % (filename, ((i+1)*10), exc))
            
    return (totals, date)

def str_to_tuple (s):
    i=0
    # remove spaces between the - and the numbers 
    while "-" in s[i:]:
        i = i+ s[i:].index("-")
        while s[i+1] == " ":
            s = s[:i+1]+s[i+2:]
        i=i+1
        
    # remove the commas
    s = s.replace(",","")
        
    p = s[:s.find(":")]
    p = p[p.rfind(" ")+1:]
    
    vals = [float(f) for f in s[s.find(":")+1:].strip().split(" ")]
    return (p,vals)
        
if __name__ == "__main__":
    fields = [ "Date", "Project", "Total Hours", "Bill Hours", "Raw Cost", "Burden Cost", "Revenue", "Bill Amount", "Unbilled"]
    outfilename = tkFileDialog.asksaveasfilename()
    oprfiles = tkFileDialog.askopenfilenames()
    with open (outfilename, "w") as outfile:
        outfile.write("\t".join(fields)+"\n")
        for oprfile in oprfiles:
            print ("opening %s"  % oprfile)
            f = scrape_OPR800(oprfile)
            for s in f[0]:
                (p, vals) = str_to_tuple (s)
                outfile.write("%s\t%s\t%s\n" % (f[1], p, "\t".join([str(v) for v in vals])))
            
