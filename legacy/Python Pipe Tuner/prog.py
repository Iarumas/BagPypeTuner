import numpy as np
from scipy.io import wavfile
from scipy.signal import savgol_filter
import matplotlib.pyplot as plt
import matplotlib.ticker as plticker
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.animation import FuncAnimation
import sounddevice as sd
import soundfile as sf
import time
#import sys
from classFIFO import ArrayFIFO
import FFT_Functions as MyFFT
#import soundfile as sf


def audio_callback(indata, frames, time, status):
    """This is called (from a separate thread) for each audio block."""
    #fifo.InOut(indata)

#def init():
    #ax.set_xlim(0, 2000)
    #ax.set_ylim(0, 10)
    #return line,

def update_plot(frame):
    """This is called by matplotlib for each plot update.    """
    global position
    fifo.InOut(all_data[position:position+SampleWidth])
    position += SampleWidth
    
    #t= time.time()
    Spectrum=MyFFT.FFTW(fifo.Content,WF)
    AmplitudeSpectrum=MyFFT.AmplitudeSpectrum(Spectrum)
    PhaseSpectrum=MyFFT.PhaseSpectrum(Spectrum)

    line.set_data(freq,np.log10(AmplitudeSpectrum))
    line2.set_data(freq,PhaseSpectrum)
    #print(time.time()-t)
    return line, line2

#Definitions
filename = "boum mono.wav"
fs, all_data = wavfile.read(filename)
global position
position = 0

SampleWidth = 8192
SampleProgress = 4410
SampleRate = 44100.0
SampleChannels = 1
SampleDataType = 'int16'

fifo = ArrayFIFO(SampleWidth)
WF=MyFFT.WindowFunction(SampleWidth,8)
AmplitudeSpectrum = np.zeros(int(SampleWidth/2))
PhaseSpectrum = np.zeros(int(SampleWidth/2))
freq = np.linspace(0,SampleRate/2,int(SampleWidth/2),dtype='float32')

device_info = sd.query_devices(kind='input')
print(device_info)
sd.default.blocksize = SampleProgress
sd.default.channels = SampleChannels
sd.default.dtype = SampleDataType
sd.default.samplerate = SampleRate


fig, ax = plt.subplots()
intervals = 1
loc = plticker.MultipleLocator(base=intervals)
ax.yaxis.set_major_locator(loc)
plt.grid(axis='y')
ax.set_xlim(0, 2000)
ax.set_ylim(0, 13)
ax.set_xticklabels([])
ax.set_yticklabels([])
line, = ax.plot(AmplitudeSpectrum)
line2, = ax.plot(PhaseSpectrum)

fig.tight_layout(pad=0)
#plt.show()

stream = sd.InputStream(callback=audio_callback)
ani = FuncAnimation(fig, update_plot, blit=True)

with stream:
    #print('1')
    plt.show()