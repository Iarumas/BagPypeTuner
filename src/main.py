from kivy.app import App
from kivy.uix.screenmanager import ScreenManager

from model import Model

from src.gui.startscreen import StartScreen
from src.gui.tunescreen import TuneScreen
from src.gui.spectrumscreen import SpectrumScreen


class BagPypeTuner(App):
    def build(self):
        sm = ScreenManager()
        sm.add_widget(StartScreen(name='start'))
        sm.add_widget(TuneScreen(name='tune'))
        sm.add_widget(SpectrumScreen(name='spectrum'))
        sm.current = 'spectrum'

        return sm


if __name__ == '__main__':
    model = Model("../resources/boum mono.wav")
    model.draw_graph()
    BagPypeTuner().run()
