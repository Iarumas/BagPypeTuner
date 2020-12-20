from kivy.app import App

from src.gui.startscreen import StartScreen


class BagPypeTuner(App):
    def build(self):
        return StartScreen()


if __name__ == '__main__':
    BagPypeTuner().run()