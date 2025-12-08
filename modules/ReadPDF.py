import PyPDF2
import os 
from GetData import GetData, Futam, Horses


class ReadPDF:
    def __init__(self, file_name):
        data = self.read(file_name)

    def read(self,file_name):
        if not os.path.exists(file_name): 
            print("File not exsits")
            return False
        
        reader = PyPDF2.PdfReader(file_name)
        num_pages = len(reader.pages)
        pages = []

        for page_num in range(num_pages):
            page = reader.pages[page_num]
            text = page.extract_text()
            pages.append(text.split("\n"))

        y,m,d = file_name.split("\\")[-1].split("_")[1:-1]

        data = GetData(f"https://mla.kincsempark.hu/racecards/trotting/{y}-{m}-{d}")
        for ln in data.futam_data:
            title = Futam(ln)
            print(f"{title}")
            for horse in ln["participants"]:
                print(f"        {Horses(horse,title.id)}")



if __name__ == "__main__":
    PDF_data = ReadPDF(r"C:\Users\Becsei Szabolcs\Apps\projects\python\ugeto_program\versenyprogram_2025_06_01_ugeto.pdf")



    
