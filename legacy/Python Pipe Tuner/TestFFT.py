import numpy as np
from scipy.io import wavfile
from scipy.signal import savgol_filter
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.animation import FuncAnimation
import sounddevice as sd
import soundfile as sf
import time
#import sys
from classFIFO import ArrayFIFO
import FFT_Functions as MyFFT
#import soundfile as sf


filename = "boum mono.wav"
fs, all_data = wavfile.read(filename)

Z=2048
N = 4 * Z
Cep = 2048
stop = 0

fifo = ArrayFIFO(N)
WF=MyFFT.WindowFunction(4*Z,8)
#CWF = np.ones(2*Z)
CWF = MyFFT.WindowFunction(2*Z,8)

t1 =time.time()
print(len(all_data))

position = 0
last_position = 44100*10  #(x seconds)
#last_position = len(all_data)-Z
#while position < len(all_data)-Z:

while position < last_position:

    new_data=all_data[position:position+Z]
    fifo.InOut(new_data) 

    Spectrum=MyFFT.FFTW(fifo.Content,WF)
    AmplitudeSpectrum=MyFFT.AmplitudeSpectrum(Spectrum)[0:2*Z]
    PhaseSpectrum=MyFFT.PhaseSpectrum(Spectrum)[0:2*Z]

    #Cepstrum=MyFFT.FFT(np.log(AmplitudeSpectrum))
    Cepstrum=MyFFT.FFTW(np.log(AmplitudeSpectrum),CWF)
    AmplitudeCepstrum = MyFFT.AmplitudeSpectrum(Cepstrum)
    #SAC = MyFFT.smooth(AmplitudeCepstrum,51)
    #AmplitudeCepstrum -= SAC
    #SAC = savgol_filter(AmplitudeCepstrum, 51, 3)
    #AmplitudeCepstrum -= SAC

    if position == 0:
        FullSpec = AmplitudeSpectrum
        FullCeps = AmplitudeCepstrum
    else:
        FullSpec = np.vstack((FullSpec,AmplitudeSpectrum))
        FullCeps = np.vstack((FullCeps,AmplitudeCepstrum))

    if position == 454656:
        stop = 1
    else:
        stop = 0

    positions, peaks = MyFFT.Peak_Finder(AmplitudeSpectrum,0.2,5,1,stop)
    if np.size(positions) < 2 : #or position == 454656:
        print(positions,peaks)
        print(position)
        break
    print(position)
    MyFFT.basic_index(positions,peaks)

    position += Z

print (time.time()-t1)


#cpositions, cpeaks = MyFFT.Peak_Finder(AmplitudeCepstrum,0.8,10)
#basic_cposition = cpositions[0]

#print(positions,peaks)
#print(cpositions,cpeaks)

#MyFFT.basic_index(positions,peaks)

#plt.plot(PhaseSpectrum)
#plt.plot(np.log(AmplitudeSpectrum))
#plt.plot(AmplitudeSpectrum)

#plt.plot(AmplitudeCepstrum)
#plt.plot(SAC)
#plt.axis([0,200,0,2000])
#plt.show()

print(np.shape(FullSpec))

Fs = 44100
s = fifo.Content*WF
t = np.arange(0,N)/Fs
freqs = np.arange(0,N/2)*Fs/N
ceps = np.arange(0,N/4)*2/Fs
maxframe = int(np.size(FullSpec)/(N/2)-1)
maxframe = 10 #select frame

fig, axes = plt.subplots(nrows=3, ncols=1, figsize=(10,10))

# plot time signal:
axes[0].set_title("Signal")
axes[0].plot(t*1000, s, color='C0')
axes[0].set_xlabel("Time in ms")
axes[0].set_ylabel("Amplitude")

# plot spectrum:
axes[1].set_xlim(0,5000)
axes[1].set_ylim(0,max(AmplitudeSpectrum))
axes[1].set_title("Magnitude Spectrum")
axes[1].set_xlabel("Frequency")
#axes[1].plot(freqs, AmplitudeSpectrum, color='C1')
axes[1].plot(freqs, FullSpec[maxframe,:], color='C1')

# plot cepstrum:
axes[2].set_xlim(0,10)
axes[2].set_ylim(0,2000)
axes[2].set_title("Cepstrum")
#axes[2].plot(ceps, AmplitudeCepstrum, color="C2")
axes[2].plot(ceps*1000, FullCeps[maxframe,:], color='C2')
axes[2].set_xlabel("Time in ms")

fig.tight_layout()
#plt.show()



data = FullSpec[:,0:200]

length, width = data.shape
x, y = np.meshgrid(np.arange(length), np.arange(width) , indexing='ij')
fig = plt.figure()
ax = fig.add_subplot(1,1,1, projection='3d')
ax.plot_surface(x, y, data)
ax.set_xlabel('x')
ax.set_ylabel('y')
ax.set_zlabel('z')

plt.show()