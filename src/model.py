import wave


class Model():
    def __init__(self, filename=None):
        self.wave = wave.open(filename)

