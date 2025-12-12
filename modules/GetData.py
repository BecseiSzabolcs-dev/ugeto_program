import re
import requests
import json
from bs4 import BeautifulSoup


class Futam:
    #Id;Daily;Title;Distance;Time;Start type;Opinion
    def __init__(self,line=""):
        if line!="":
            try:
                id, daily, title, dist, time, start, opinion = line.strip().split(";")
            except:
                id, daily, title, dist, time, start = line.strip().split(";")
            self.id      = id
            self.daily   = daily
            self.title   = title
            self.dist    = dist
            self.time    = time
            #self.track   = track
            self.start   = start
            try:
                self.opinion = opinion
            except:
                self.opinion = ""
        else:
            self.id      = ""
            self.daily   = ""
            self.title   = ""
            self.dist    = ""
            self.time    = ""
            self.track   = ""
            self.start   = ""
            self.opinion = ""


    def load_json(self,data):
        self.id = data["id"]
        self.daily = data["daily"].replace('.',"")
        self.title = data["title"]
        self.dist = data["distance"][:-1] if data["distance"][-1] == 'A' else data["distance"]
        self.time = data["time"]
        self.track = data["track"]
        self.start = "Autóstart!" if data["distance"][-1] == 'A' else "Fordulóstart!"
        self.opinion = ""
        return self

    #{'Id': '0', 'Daily': 'Q', 'Title': 'Q11 KVALIFIKÁCIÓ', 'Distance': '1800', 'Start time': '13:15', 'Start type': 'Autóstart!', 'Opinion': ''}
    def load_dict(self,data):
        self.id = data["Id"]
        self.daily = data["Daily"].replace('.',"")
        self.title = data["Title"]
        self.dist = data["Distance"]
        self.time = data["Start time"]
        #self.track = data["track"]
        self.start =  data["Start type"]
        self.opinion =  data["Opinion"]
        return self

    
    def __str__(self):
        return f"{self.id};{self.daily};{self.title};{self.dist};{self.time};{self.start};{self.opinion}"
    
    def to_dict(self):
        return {'Id': self.id, 'Daily': self.daily, 'Title': self.title, 'Distance': self.dist, 'Start time': self.time, 'Start type': self.start, 'Opinion': self.opinion}
  
class Horses: 
    #Horse Number;Horse Name;Horse Distance;Driver Name;Futam Number;is run
    def __init__(self, line=""):
        if(line!=""):
            Hnum, Hname, dist, DJname, Fnum, isRun  = line.strip().split(";")
            self.Hnum   = Hnum
            self.Hname  = Hname
            self.dist   = dist
            self.DJname = DJname
            self.Fnum   = Fnum
            self.isRun  = isRun
        else:
            self.Hnum   = ""
            self.Hname  = ""
            self.dist   = ""
            self.DJname = ""
            self.Fnum   = ""
            self.isRun  = "0"

    def load_json(self,data,Fnum):
        self.Hnum   = data['number']
        self.Hname  = data['name']
        self.dist   = data['distance']
        self.DJname = data['driver_jockey']
        self.Fnum   = Fnum
        self.isRun  = "1"
        return self
    
    #{'Start number': '7', 'Horse name': 'Zippy Boy', 'Distance': '1900', 'Driver name': 'Fazekas Andrea', 'Futam id': '10', 'Run': '1'}
    def load_dict(self,data):
        self.Hnum   = data['Start number']
        self.Hname  = data['Horse name']
        self.dist   = data['Distance']
        self.DJname = data['Driver name']
        self.Fnum   = data['Futam id']
        self.isRun  = data['Run']
        return self

    def __str__(self):
        return f"{self.Hnum};{self.Hname};{self.dist};{self.DJname};{self.Fnum};{self.isRun}"
    
    def to_dict(self):
        return {'Start number': self.Hnum, 'Horse name': self.Hname, 'Distance': self.dist, 'Driver name': self.DJname, 'Futam id': self.Fnum, 'Run': self.isRun}


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
        title = Futam()
        title.load_json(ln)
        print(f"{title}")
        for horse in ln["participants"]:
            Horse = Horses()
            Horse.load_json(horse,title.id)
            print(f"        {Horse}")
            
        


