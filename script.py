import pandas as pd
import json

main_json = []

def convert_slot_to_time_range(start_slot, end_slot=None): # converts time slots to range
      
    start_time = 8 + start_slot - 1
    if end_slot is None:
        end_time = start_time + 1  
    else:
        end_time = 8 + end_slot 

    return [start_time, end_time]

def parse(sheet_no):              # main code for parsing the data
    
    time_table = pd.read_excel("timetable.xlsx", sheet_name=sheet_no, header=1)
    

    
    course_code = time_table.iloc[1, 1]  
    course_title = time_table.iloc[1, 2]

    lecture_units = int(time_table.iloc[1, 3]) if time_table.iloc[1, 3] != "-" else 0
    practical_units = int(time_table.iloc[1, 4]) if time_table.iloc[1, 4] != "-" else 0
    total_units = int(time_table.iloc[1, 5]) if time_table.iloc[1, 5] != "-" else 0




    course_json = {
        "course_code": course_code,
        "course_title": course_title,
        "credits": {
            "lecture": lecture_units,
            "practical": practical_units,
            "units": total_units
        },
        "sections": []
    }

    
    secs = time_table["SEC"]
    for i in range(1, len(secs)):
        if pd.notna(secs[i]):  
            
            if 'L' in str(time_table.iloc[i, 6]):
                section_type = "Lecture"
            elif 'P' in str(time_table.iloc[i, 6]):
                section_type = "Practical"
            else:
                section_type = "Tutorial"

            
            section = {
                "section_type": section_type,
                "section_number": str(time_table.iloc[i, 6]),
                "instructors": [str(time_table.iloc[i, 7])],
                "room": str(time_table.iloc[i, 8]),
                "timings": give_timings(str(time_table.iloc[i, 9]))  
            }

            
            course_json["sections"].append(section)

    
    main_json.append(course_json)

def give_timings(cell):    # gives the timing slots
    time = []
    parts = cell.split() 
    j = 0
    while j < len(parts):
        day = parts[j]  
        slots = []


        j += 1
        while j < len(parts) and parts[j].isdigit():
            slots.append(int(parts[j]))
            j += 1
        

        if slots:
            if len(slots) == 1:
                time_range = convert_slot_to_time_range(slots[0])
            else:
                time_range = convert_slot_to_time_range(slots[0], slots[-1])

            time.append({"day": day, "slots": time_range})

    return time

sheets = ["S1", "S2", "S3", "S4", "S5", "S6"]  # looping code for all the sheets
for sheet in sheets:
    parse(sheet)


with open('timetable.json', 'w') as json_file: 
    json.dump(main_json, json_file, indent=4)