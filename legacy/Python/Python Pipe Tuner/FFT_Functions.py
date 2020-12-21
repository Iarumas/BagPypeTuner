import numpy as np
import math
from detect_peak import detect_peaks
import matplotlib.pyplot as plt


def WindowFunction(DataLength=16,GaussOrder=8):
    """returns window function for FFT"""
    # GaussOrder = DataLength / Sigma  -> Sigma = Datalength / GaussOrder     
    return np.exp(-1/2*(np.arange(-DataLength/2,DataLength/2))**2/(DataLength/GaussOrder)**2)

def FFTW(Raw_Data, Window_Function):
    " FFT with function multiplied with window function "
    DataLength = int(len(Raw_Data)/2)
    return np.fft.rfft(Raw_Data*Window_Function)[0:DataLength]

def FFT_Gauss_Sigma (GaussOrder = 8):
     " expected sigma for frequency in units of data index"
     return GaussOrder / (2*math.pi)

def FFT(Raw_Data):
    " FFT without window function "
    length = int(len(Raw_Data)/2)
    return np.fft.rfft(Raw_Data)[0:length]

def AmplitudeSpectrum(Spectrum):
    return abs(Spectrum)

def PhaseSpectrum(Spectrum):
    return np.arctan2(Spectrum.imag,Spectrum.real)

def Gauss_Fit(Array, Index, GaussOrder = 8):
    "return center, sigma & amplitude of 3 point gauss fit around index position"
    #return none if index at the end of array or beyond
    if Index <= 0 or Index >= len(Array):
        return None
    #take logarithm of values   
    LogValues = list(np.log(Array[Index-1:Index+2]))
    #fit of parabolic curve: y = ax^2 + bx +c / data points (x0,y0), (x1,y1), (x2,y2) and x(i+1)-x(i) = 1 (equidistant) 
    # => (c = y1, b = (y2-y0)/2, a = (y2+y0)/2 - y1
    Parabolic_Fit = [LogValues[1], (LogValues[2]-LogValues[0])/2,(LogValues[0]+LogValues[2])/2-LogValues[1]]
    #fit of gauss curve
    GaussCenter = Index - Parabolic_Fit[1]/(2*Parabolic_Fit[2])             # = Index - b/(2*a)
    GaussSigma= np.sqrt(-1/(2*Parabolic_Fit[2]))                           # = sqrt(-1/(2*a))
    GaussAmplitude = np.exp( Parabolic_Fit[0]-Parabolic_Fit[1]**2/(4 * Parabolic_Fit[2])   )       # = ln(Amp) = c - b^2/(4*a)

    expectedSigma = GaussOrder/(2*math.pi)   # you might ewant to compared this to the measured sigma

    return GaussCenter, GaussSigma, GaussAmplitude

def Peak_Finder(array,fraction=0.2,separation=10,SigmaFilter = False ,stop = 0):
    # finds peaks in array
    #fraction = 0.2
    #separation = 10
    threshold = fraction * max(array)
    #I found "detec_peaks" function in the web: similar to findpeaks in MatLab
    Peaks = detect_peaks(array,threshold,separation) 

    if len(Peaks) < 2:
        return np.array(0), np.array(0)
    
    Idx_f = []      #frquency index (not just index but exact position floating point)
    Sigma = []      #sigma of peak
    Amplitude = []  #amplitude of peak

    for position in Peaks:
        fit = Gauss_Fit(array,position)
        Idx_f = np.append(Idx_f,fit[0])          #list of measured idx_f
        Sigma = np.append(Sigma,fit[1])          #list of measued sigmas
        Amplitude = np.append(Amplitude,fit[2])  # list of measured amplitudes

    #sorted by amplitude (decreasing)
    SortedIdx = np.argsort(Amplitude)[::-1]
    SortedIdx_f = Idx_f[SortedIdx]
    SortedSigma = Sigma[SortedIdx]
    SortedAmp = Amplitude[SortedIdx]

    if stop:
        print()
        print(SortedIdx)
        print(SortedIdx_f)
        print(SortedSigma)
        print(SortedAmp)
        print(Peaks)

    if SigmaFilter == True: 
        # sorted by sigma (decreasing)
        SortedIdx = np.argsort(Sigma)[::-1]
        SortedIdx_f = Idx_f[SortedIdx]
        SortedSigma = Sigma[SortedIdx]
        SortedAmp = Amplitude[SortedIdx]
        # exclude peaks that are to broad: sigma larger by "factor" than min. sigma  
        factor = 2. 
        minSigma = min(Sigma)  
        #minSigma = 8/(2*math.pi)  
        while Sigma[0] > factor * minSigma:
            Idx_f = np.delete(Idx_f, 0)
            Sigma = np.delete(Sigma,0)
            Amplitude = np.delete(Amplitude,0)
    
    if stop:
        print()
        print(SortedIdx)
        print(SortedIdx_f)
        print(SortedSigma)
        print(SortedAmp)
        print(Peaks)

    if len(SortedIdx_f) < 2:
        return np.array(0), np.array(0)

    #sorted by index (increasing)
    SortedIdx = np.argsort(Idx_f)
    SortedIdx_f = Idx_f[SortedIdx]
    SortedSigma = Sigma[SortedIdx]
    SortedAmp = Amplitude[SortedIdx]
  
    return SortedIdx_f, SortedAmp


def smooth(y, box_pts):
    box = np.ones(box_pts)/box_pts
    y_smooth = np.convolve(y, box, mode='same')
    return y_smooth

def basic_index(Idx_f, Amplitude):
    #sorted by amplitude (decreasing)
    SortedIdx = np.argsort(Amplitude)[::-1]
    SortedIdx_f = Idx_f[SortedIdx]
    SortedAmp = Amplitude[SortedIdx]

    max1 = SortedIdx_f[0]
    max2 = SortedIdx_f[1]

    if max1 > max2:
        max1,max2 = max2,max1

    l = 1
    k = 1
    f1 = 1
    f2 = 1

    while max1 * l < 10000:
        while k < l:
            ratio = l/k*max1/max2
            if abs(ratio-1) < 0.02:
                f1 = k
                f2 = l
                print(max1/k, max2/l, max2-max1)
            k += 1
        l += 1

    return np.mean([max1/f1, max2/f2, max2-max1])
    








