from moviepy.editor import *
import pandas
import datetime
from tkinter import filedialog
import os
def make_video(name):
    vid = VideoFileClip("HBD.mp4")
    
    text_clip = TextClip(txt=name,
                         fontsize=120,
                         font='LibreBaskerville-Bold.ttf',
                         color='White').set_position(("center", 675)).set_duration(4)
    
    text_clip = text_clip.crossfadein(2.0)
    
    finalvid = CompositeVideoClip([vid, text_clip])
    datenow = datetime.datetime.now()
    pdate = datenow.strftime("%d-%m-%y")
    try:
        os.makedirs(pdate)
    except:
        pass
    outname = "{pdate}\\{name}.mp4".format(name=name, pdate=pdate)
    finalvid.write_videofile(outname)



excel = filedialog.askopenfilename(initialdir="/", title="Select A File", filetypes=(("Excel File", "*.xlsx"), ("all files", "*.*")))
df = pandas.read_excel(excel)


for i in range(len(df)):
    name = (df.loc[i, "NAME"])
    name = name.title()
    make_video(name)
