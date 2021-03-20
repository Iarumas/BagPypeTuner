from kivy.app import App
from kivy.lang import Builder

from model import Model


class BagPypeTuner(App):
    def build(self):
        pass


if __name__ == '__main__':
    model = Model("../resources/boum mono.wav")
    BagPypeTuner().run()
