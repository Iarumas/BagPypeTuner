import numpy as np
import math

def windowFunction(DataLength=16,GaussOrder=8):
    """returns window function for FFT"""
    # GaussOrder = DataLength / Sigma  -> Sigma = Datalength / GaussOrder
    return np.exp(-1/2*(np.arange(-DataLength/2,DataLength/2))**2/(DataLength/GaussOrder)**2)

def fftw(Raw_Data, Window_Function):
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