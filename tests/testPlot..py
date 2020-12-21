import struct
import wave

import numpy as np
import matplotlib.pyplot as plt  # TODO: remove matplotlib, just for initIal testing..

def draw_graph():
    # Extract Raw Audio from Wav File
    sound_file = wave.open("../resources/boum mono.wav", 'r')
    framerate = sound_file.getframerate()
    file_length = sound_file.getnframes()
    data = sound_file.readframes(file_length)
    data = struct.unpack('{n}h'.format(n=file_length), data)
    data = np.array(data)
    sound_file.close()

    # x = np.linspace(0.0, 1, 600)
    # y = np.sin(50.0 * 2.0*np.pi*x)
    # yf = fft(y)
    yf = np.fft.fft(data)
    freq = np.fft.fftfreq(data.shape[0],framerate)

    """
    plt.xlim(0, 4000)
    plt.ylim(0,3e6)
    plt.plot(np.abs(yf))
    """
    plt.plot(freq, yf.real, freq, yf.imag)
    plt.grid()
    plt.show()
    print("Plot should be shown")

def draw_graph2():
    import pyaudio
    import matplotlib.pyplot as plt
    import numpy as np
    import time

    form_1 = pyaudio.paInt16  # 16-bit resolution
    chans = 1  # 1 channel
    samp_rate = 44100  # 44.1kHz sampling rate
    chunk = 8192  # 2^12 samples for buffer
    dev_index = 2  # device index found by p.get_device_info_by_index(ii)

    # Open sound file  in read binary form.
    file = wave.open("../resources/boum mono.wav", 'rb')

    audio = pyaudio.PyAudio()  # create pyaudio instantiation

    # create pyaudio stream
    """ From microphone
    stream = audio.open(format=form_1, \
                        channels=chans, \
                        rate=samp_rate, \
                        input_device_index=dev_index, \
                        input=True, \
                        frames_per_buffer=chunk)
    """

    stream = audio.open(format=audio.get_format_from_width(file.getsampwidth()),
                    channels=file.getnchannels(),
                    rate=file.getframerate(),
                    input=True)

    # record data chunk
    stream.start_stream()
    data = np.fromstring(stream.read(chunk), dtype=np.int16)
    stream.stop_stream()

    # mic sensitivity correction and bit conversion
    mic_sens_dBV = -47.0  # mic sensitivity in dBV + any gain
    mic_sens_corr = np.power(10.0, mic_sens_dBV / 20.0)  # calculate mic sensitivity conversion factor

    # (USB=5V, so 15 bits are used (the 16th for negatives)) and the manufacturer microphone sensitivity corrections
    data = ((data / np.power(2.0, 15)) * 5.25) * (mic_sens_corr)

    # compute FFT parameters
    f_vec = samp_rate * np.arange(chunk / 2) / chunk  # frequency vector based on window size and sample rate
    mic_low_freq = 100  # low frequency response of the mic (mine in this case is 100 Hz)
    low_freq_loc = np.argmin(np.abs(f_vec - mic_low_freq))
    fft_data = (np.abs(np.fft.fft(data))[0:int(np.floor(chunk / 2))]) / chunk
    fft_data[1:] = 2 * fft_data[1:]

    max_loc = np.argmax(fft_data[low_freq_loc:]) + low_freq_loc

    # plot
    plt.style.use('ggplot')
    plt.rcParams['font.size'] = 18
    fig = plt.figure(figsize=(13, 8))
    ax = fig.add_subplot(111)
    plt.plot(f_vec, fft_data)
    ax.set_ylim([0, 2 * np.max(fft_data)])
    plt.xlabel('Frequency [Hz]')
    plt.ylabel('Amplitude [Pa]')
    ax.set_xscale('log')
    plt.grid(True)

    # max frequency resolution
    plt.annotate(r'$\Delta f_{max}$: %2.1f Hz' % (samp_rate / (2 * chunk)), xy=(0.7, 0.92), \
                 xycoords='figure fraction')

    # annotate peak frequency
    annot = ax.annotate('Freq: %2.1f' % (f_vec[max_loc]), xy=(f_vec[max_loc], fft_data[max_loc]), \
                        xycoords='data', xytext=(0, 30), textcoords='offset points', \
                        arrowprops=dict(arrowstyle="->"), ha='center', va='bottom')

    plt.show()

if __name__ == '__main__':
    draw_graph2()