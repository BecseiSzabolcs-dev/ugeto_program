import re
import requests
import json
from bs4 import BeautifulSoup

class Futam:
    #Id;Daily;Title;Distance;Time;Start type;Opinion
    def __init__(self,data):
        self.id = data["id"]
        self.daily = data["daily"]
        self.title = data["title"]
        self.dist = data["distance"][:-1] if data["distance"][-1] == 'A' else data["distance"]
        self.time = data["time"]
        self.track = data["track"]
        self.start = "Autóstart!" if data["distance"][-1] == 'A' else "Fordulóstart!"
        self.opnion = ""
    
    def __str__(self):
        return f"{self.id};{self.daily};{self.title};{self.dist};{self.time};{self.start};{self.opnion}"
  
class Horses:
    #Horse Number;Horse Name;Horse Distance;Driver Name;Futam Number;is run
    def __init__(self,data,Fnum):
        self.Hnum   = data['number']
        self.Hname  = data['name']
        self.dist   = data['distance']
        self.DJname = data['driver_jockey']
        self.Fnum   = Fnum
        self.isRun  = 1

    def __str__(self):
        return f"{self.Hnum};{self.Hname};{self.dist};{self.DJname};{self.Fnum};{self.isRun}"


class GetData:
    def __init__(self,url):
        self.rome_num = ["I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII","XIII","XIV"]
        self.futam_data = self.get_race_data(url)

    def part_join(self,line, fch, lch):
        forlen = len(line)
        for i in range(0, len(line)):
            if fch in line[i] and not (lch in line[i]):
                for d in range(i, len(line)):
                    if lch in line[d]:
                        text = " ".join(line[i:d + 1])
                        del line[i:d + 1]
                        line.insert(i, text)
                        break
            if forlen != len(line):
                break
        return line

        
    def get_race_data(self,url):
        all_races = []
        clean_futam = []
        try:
            response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'})

            soup = BeautifulSoup(response.text, 'html.parser')

            scripts = soup.find_all('script')

            pattern = re.compile(r'races_table_divs\[".*?"\] = (\{.*?\});', re.DOTALL)

            for script in scripts:
                if script.string:
                    
                    matches = pattern.findall(script.string)
                    for match in matches:
                        try:
                           
                            race_dict = json.loads(match)
                            all_races.append(race_dict)
                        except json.JSONDecodeError as e:
                            print(f"Figyelmeztetés: Nem sikerült dekódolni egy futam JSON adatát. Hiba: {e}")
                            
        except requests.exceptions.RequestException as e:
            print(f"Hiba az URL lekérése közben: {e}")
            return []
        
        for id, race in enumerate(all_races, 0):
            surface = ""
            if race.get('surface', 'N/A') == 'GYEP':                surface = "Gyep pálya"
            elif race.get('surface', 'N/A') == 'HOMOK/SZINTETIKUS': surface = "Szintetikus pálya"
            else:                                                   surface = race.get('surface', 'N/A')

            race_name = race.get('race_name', 'N/A')
            if race_name != 'N/A':
                line = race_name.split(" ")
                line = self.part_join(line,"(",")")


                for i in range(0,10):
                    line = [data for data in line if data != f"(HUN Gd-{i})"                                 ]
                    line = [data for data in line if data != f"({self.rome_num[i]}.o.)"                      ]
                    line = [data for data in line if data != f"({self.rome_num[i]}.kat.)"                    ]
                    line = [data for data in line if data != f"({self.rome_num[i]}. kat.)"                   ]
                    line = [data for data in line if data != f"(Elit kat.)"                                  ]
                    line = [data for data in line if data != f"({self.rome_num[i]}/b.)"                      ]
                    line = [data for data in line if data != f"({self.rome_num[i]}.kat.)(szintetikus pálya)" ]
                    line = [data for data in line if data != f"(szintetikus pálya)"                          ]
                #print(line)
                race_name = " ".join(line)
            
            clean_futam.append({
                "id":           id, # self.rome_num.index(race.get('daily', '')), 
                "daily":        race.get('daily', 'N/A'),
                "title":        race_name,
                "distance":     race.get('distance', 'N/A'),
                "participants": race.get('participants', 'N/A'),
                "time":         race.get('start', 'N/A'),
                "track":        surface,
                #"opinion": self.opinion[id]
            })
        
        
        return clean_futam
    
if __name__ == "__main__":
    data = GetData("https://mla.kincsempark.hu/racecards/trotting/2025-12-06")
    for ln in data.futam_data:
        title = Futam(ln)
        print(f"{title}")
        for horse in ln["participants"]:
            print(f"        {Horses(horse,title.id)}")
            
        


