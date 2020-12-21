import numpy as np

#FIFO class only for arrays as input and output
class ArrayFIFO(object):

    def __init__(self,fifo_size=1):
        self.__fifo_size = fifo_size
        self.ClearFIFO()

    def ClearFIFO(self):
        self.__FIFO = np.zeros(self.__fifo_size)
        self.__filled_elements = 0      # number of elements that have been filled
        self.__full = False
        self.__empty = True
        
    def InOut(self, fifo_array_in):
        # push and array and pop an array of the same size

        # check if input is an array
        if type(fifo_array_in) != 'numpy.ndarray':
            self.__in = np.array(fifo_array_in)
        
        self.__inout_size =  np.size(self.__in)

        #check dimensions and change to 1 dimension
        if np.ndim(self.__in) > 1:
            self.__in = self.__in.reshape(self.__inout_size,)

        # check if input array is not larger ther FIFO size
        if self.__inout_size > self.__fifo_size:
            return None
        
        # shift FIFO
        self.__out = self.__FIFO[:self.__inout_size]
        self.__FIFO = np.append(self.__FIFO[self.__inout_size:],self.__in)

        # check for number of filled elements
        self.__empty = False
        self.__filled_elements += self.__inout_size 
        if self.__filled_elements >= self.__fifo_size :
            self.__filled_elements = self.__fifo_size
            self.__full = True

        #returns FIFOarray of the same size as the input array
        return self.__out

    def getContent(self):
        return self.__FIFO

    def getAverage(self):
        return np.mean(self.__FIFO)
   
    def getStatusEmpty(self):
        return self.__empty

    def getStatusFull(self):
        return self.__full

    Content = property(getContent)
    Average = property(getAverage)
    Empty = property(getStatusEmpty)
    Full = property(getStatusFull)




