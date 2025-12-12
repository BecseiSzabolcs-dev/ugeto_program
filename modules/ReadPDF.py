import PyPDF2
import os 
try:
    from .GetData import GetData, Futam, Horses
except:
    from GetData import GetData, Futam, Horses

def removeTXT(search, txt):
    if " "+search in txt: return txt[0:txt.index(" "+search)]
    elif search+" " in txt:   return txt[0:txt.index(search)]
    else:                 return txt

def remove_dupl(li:list):
    ndli = []
    for ln in li:
        if ln not in ndli:
            ndli.append(ln)
    
    return ndli

                       


class ReadPDF:

    def __init__(self, file_name):
        self.rome_num = ["I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII","XIII","XIV"]
        self.horses = []
        self.futams = []
        self.opinions = []
        self.pdf = []
        self.read(file_name)


    def read(self,file_name):
        if not os.path.exists(file_name): 
            print("File not exsits")
            return False
        
        reader = PyPDF2.PdfReader(file_name)
        num_pages = len(reader.pages)


        for page_num in range(num_pages):
            page = reader.pages[page_num]
            text = page.extract_text()
            self.pdf.append(text.split("\n"))

        #print(file_name.split("/")[-1].split("_")[1:-1])
        y,m,d = file_name.split("/")[-1].split("_")[1:-1]

        data = GetData(f"https://mla.kincsempark.hu/racecards/trotting/{y}-{m}-{d}")
        for ln in data.futam_data:
            title = Futam()
            title.load_json(ln)  # or however you load futam
            self.futams.append(title)

            for horse in ln["participants"]:
                driver = Horses()
                driver.load_json(horse, title.id)
                self.horses.append(driver)


        opinions = []
        for page in self.pdf:
            for ln in page:
                for horse in self.horses:
                    if "Véleményünk:" in ln and horse.Hname in ln:
                        opinions.append(ln)
                        break

        for i, op in enumerate(opinions):
            for num in range(9, 14):
                opinions[i] = removeTXT(str(num), opinions[i])

            opinions[i] = removeTXT("Elérhetőségek", opinions[i])
            opinions[i] = removeTXT("100.000 Ft", opinions[i])
            opinions[i] = removeTXT("200.000 Ft", opinions[i])
            opinions[i] = removeTXT("300.000 Ft", opinions[i])
            opinions[i] = removeTXT("101.190 Ft", opinions[i])
            opinions[i] = removeTXT("Esélyelemzés", opinions[i])
            #Véleményünk: 
            opinions[i] = opinions[i].replace("Véleményünk: ",'')
            opinions[i] = opinions[i].replace("Véleményünk:",'')
            opinions[i] = opinions[i].strip()

        
        self.opinions = remove_dupl(opinions)

        """
        print("\nopinions:")
        for i,op in enumerate(self.opinions):
            print(f"{i}: {op}")
        """
        
        cnt = 0
        for futam in self.futams:
            if not futam.daily in self.rome_num:
                cnt = futam.id+1
        #print("\nfutam opinions:")
        for i,futam in enumerate(self.futams):
            if futam.daily in self.rome_num:
                futam.opinion = self.opinions[futam.id-cnt]
                #print(f"{i},{futam.id-cnt}: {futam.opinion}")


if __name__ == "__main__":
    PDF_data = ReadPDF(r"C:\Users\Becsei Szabolcs\Downloads\versenyprogram_2025_12_13_ugeto.pdf")

    print("titles:")
    for i in PDF_data.futams:
        print(i)



    



    
