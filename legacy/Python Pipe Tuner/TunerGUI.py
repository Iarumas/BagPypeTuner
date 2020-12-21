#--------------------------------------------------------------
# GUI
#--------------------------------------------------------------

from  PIL import Image
import tkinter as tk
import numpy as np 
import matplotlib.pyplot as plt
import matplotlib.cm as cm

def btn_L3_Click():
    global Reference_Frequency
    Reference_Frequency -= 10.0
    Update_Ref_Freq(Reference_Frequency)

def btn_L2_Click():
    global Reference_Frequency
    Reference_Frequency -= 1.0
    Update_Ref_Freq(Reference_Frequency)

def btn_L1_Click():
    global Reference_Frequency
    Reference_Frequency -= 0.1
    Update_Ref_Freq(Reference_Frequency)

def btn_R1_Click():
    global Reference_Frequency
    Reference_Frequency += 0.1
    Update_Ref_Freq(Reference_Frequency)

def btn_R2_Click():
    global Reference_Frequency
    Reference_Frequency += 1.0
    Update_Ref_Freq(Reference_Frequency)

def btn_R3_Click():
    global Reference_Frequency
    Reference_Frequency += 10.0
    Update_Ref_Freq(Reference_Frequency)

def Update_Ref_Freq(Reference_Frequency):
    txt_ref_freqA = "A  :  {0:4.1f}".format(Reference_Frequency)
    txt_ref_freqB = "Bb :  {0:4.1f}".format(Reference_Frequency/Bb_factor)
    lbl_FrequencyA.config(text=txt_ref_freqA)
    lbl_FrequencyB.config(text=txt_ref_freqB)


Reference_Frequency = 480.0
Bb_factor = 2**(1/12)
win_xsize = 640
win_ysize = 480
win_size=str(win_xsize)+'x'+str(win_ysize)

frm_left = 5
frm_top = 5
frm_xsize = win_xsize - 2*frm_left
frm_ysize = win_ysize - 2*frm_top

lbl_left = 5
lbl_top = 5

btn_xsize = 30
btn_ysize = 15
btn_xpos = lbl_left
btn_ypos = 40
btn_x_sep = 2

lbl_Output_left = 200
lbl_Output_top = frm_top
lbl_Output_xsize = 425
lbl_Output_ysize = frm_ysize-2*frm_top

lbl_Freq_ysize = 30
lbl_Freq_xsize = 3*btn_xsize+2*btn_x_sep-1

# Erzeugung des Fensters
win = tk.Tk()
win.title("Tuner")
win.geometry(win_size)
# frame 
frm_Tuner = tk.Frame(background="gray")
frm_Tuner.place(x=frm_left, y=frm_top, width=frm_xsize, height=frm_ysize)
# Label outpu
lbl_Output = tk.Label(master=frm_Tuner, background="white", text=str('Output'))
lbl_Output.place(x=lbl_Output_left, y=lbl_Output_top, width=lbl_Output_xsize, height=lbl_Output_ysize)      
#label frequencies
lbl_FrequencyA = tk.Label(master=frm_Tuner, background="white")
lbl_FrequencyA.place(x=lbl_left, y=lbl_top, width=lbl_Freq_xsize, height=lbl_Freq_ysize)
lbl_FrequencyB = tk.Label(master=frm_Tuner, background="white")
lbl_FrequencyB.place(x=btn_xpos+btn_x_sep+1+lbl_Freq_xsize, y=lbl_top, width=lbl_Freq_xsize, height=lbl_Freq_ysize)
Update_Ref_Freq(Reference_Frequency)

canvas = tk.Canvas(master=lbl_Output, background="gray")
canvas.place(relx=0, rely=0, relwidth=0.15, relheight=1, bordermode='outside',anchor='nw')
#canvas.pack()

delta = 0.025
x = y = np.arange(-3.0, 3.0, delta)
X, Y = np.meshgrid(x, y)
Z1 = np.exp(-X**2 - Y**2)
Z2 = np.exp(-(X - 1)**2 - (Y - 1)**2)
Z = (Z1 - Z2) * 2
Z3 = 10*np.sin(x)
for el,val in enumerate(x):
    Z[el,120]= Z3[el]
fig, ax = plt.subplots()
im = ax.imshow(Z, interpolation='bilinear', cmap=cm.RdYlGn,
               origin='lower', extent=[-3, 3, -3, 3],
               vmax=abs(Z).max(), vmin=-abs(Z).max())

            


plt.show()


bitmaps = ["error", "gray75", "gray50", "gray25", "gray12", "hourglass", "info", "questhead", "question", "warning"]
nsteps = len(bitmaps)
step_y = 30

for i in range(0, nsteps):
   canvas.create_bitmap(30,(i+1)*step_y, bitmap=bitmaps[i])



# Buttons
btn_L3 = tk.Button(master=frm_Tuner, text="<<<", command=btn_L3_Click)
btn_L2 = tk.Button(master=frm_Tuner, text="<<", command=btn_L2_Click)
btn_L1 = tk.Button(master=frm_Tuner, text="<", command=btn_L1_Click)
btn_R1 = tk.Button(master=frm_Tuner, text=">", command=btn_R1_Click)
btn_R2 = tk.Button(master=frm_Tuner, text=">>", command=btn_R2_Click)
btn_R3 = tk.Button(master=frm_Tuner, text=">>>", command=btn_R3_Click)
btn_L3.place(x=btn_xpos, y=btn_ypos, width=btn_xsize, height=btn_ysize)
btn_L2.place(x=btn_xpos+1*(btn_xsize+btn_x_sep), y=btn_ypos, width=btn_xsize, height=btn_ysize)
btn_L1.place(x=btn_xpos+2*(btn_xsize+btn_x_sep), y=btn_ypos, width=btn_xsize, height=btn_ysize)
btn_R1.place(x=btn_xpos+3*(btn_xsize+btn_x_sep), y=btn_ypos, width=btn_xsize, height=btn_ysize)
btn_R2.place(x=btn_xpos+4*(btn_xsize+btn_x_sep), y=btn_ypos, width=btn_xsize, height=btn_ysize)
btn_R3.place(x=btn_xpos+5*(btn_xsize+btn_x_sep), y=btn_ypos, width=btn_xsize, height=btn_ysize)

# Aktivierung der Ereignisschleife
win.mainloop()
