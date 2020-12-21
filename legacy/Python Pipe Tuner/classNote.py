


import numpy as np

class Note(object):

    def __init__(self,name,ratio):
        self.__name = name 
        self.__ratio = ratio 
        self.__cent = 1200*np.log(self.__ratio)/np.log(2)

    def setName(self,name):
        self.__name = name 

    def getName(self):
        return self.__name

    def setRatio(self,ratio):
        self.ratio = ratio
        self.__cent = 1200*np.log(self.__ratio)/np.log(2)

    def getRatio(self):
        return self.__ratio

    def getCent(self):
        return self.__cent

    name = property (getName,setName)
    ratio = property (getRatio,setRatio)
    cent = property (getCent)

