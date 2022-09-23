from updater import Updater
from updater_gui import UpdaterGUI


# The main script instantiates the updater object, then pass that object as an argument in the GUI constructor
if __name__ == '__main__':

    updater = Updater()
    UpdaterGUI(updater)
