# Deklaration der Klasse Roboter

class Roboter(object):
    def __init__(self):
        self.x = 0
        self.y = 0
        self.r = 'S'
        self.w = None      

    def getX(self):
        return self.x

    def getY(self):
        return self.y

    def getR(self):
        return self.r

    def getZustand(self):
        return (self.x, self.y, self.r)

    def setZustand(self, x, y, r):
        self.x = x
        self.y = y
        self.r = r

    def setWelt(self, w):
        self.w = w

    def getWelt(self):
        return self.w


    def schritt(self):
        if self.r == 'O' and self.x < self.w.getFelderX()-1:
            self.x = self.x + 1
        elif self.r == 'S' and self.y < self.w.getFelderY()-1:
            self.y = self.y + 1
        elif self.r == 'W' and self.x > 0:
            self.x = self.x - 1
        elif self.r == 'N' and self.y > 0:
            self.y = self.y - 1

    def rechts(self):
        if self.r == 'O':
            self.r = 'S'
        elif self.r == 'S':
            self.r = 'W'
        elif self.r == 'W':
            self.r = 'N'
        elif self.r == 'N':
            self.r = 'O'

    def links(self):
        if self.r == 'O':
            self.r = 'N'
        elif self.r == 'N':
            self.r = 'W'
        elif self.r == 'W':
            self.r = 'S'
        elif self.r == 'S':
            self.r = 'O'

    def ziegelHinlegen(self):
        if self.r == 'O' and self.x < self.w.getFelderX()-1:
            self.w.incZiegel(self.x+1, self.y)
        elif self.r == 'S' and self.y < self.w.getFelderY()-1:
            self.w.incZiegel(self.x, self.y+1)
        elif self.r == 'W' and self.x > 0:
            self.w.incZiegel(self.x-1, self.y)
        elif self.r == 'N' and self.y > 0:
            self.w.incZiegel(self.x, self.y-1)

    def ziegelAufheben(self):
        if self.r == 'O' and self.x < self.w.getFelderX()-1:
            self.w.decZiegel(self.x+1, self.y)
        elif self.r == 'S' and self.y < self.w.getFelderY()-1:
            self.w.decZiegel(self.x, self.y+1)
        elif self.r == 'W' and self.x > 0:
            self.w.decZiegel(self.x-1, self.y)
        elif self.r == 'N' and self.y > 0:
            self.w.decZiegel(self.x, self.y-1)

    def vorWand(self):
        if self.r == 'O' and self.x < self.w.getFelderX()-1:
            return False
        elif self.r == 'S' and self.y < self.w.getFelderY()-1:
            return False
        elif self.r == 'W' and self.x > 0:
            return False
        elif self.r == 'N' and self.y > 0:
            return False
        else:
            return True

    def nichtVorWand(self):
        return not self.vorWand()

    def vorZiegel(self):
        if self.r == 'O' and self.x < self.w.getFelderX()-1:
            return (self.w.getZiegel(self.x+1, self.y) > 0)
        elif self.r == 'S' and self.y < self.w.getFelderY()-1:
            return (self.w.getZiegel(self.x, self.y+1) > 0)
        elif self.r == 'W' and self.x > 0:
            return (self.w.getZiegel(self.x-1, self.y) > 0)
        elif self.r == 'N' and self.y > 0:
            return (self.w.getZiegel(self.x, self.y-1) > 0)

    def nichtVorZiegel(self):
        return not self.vorZiegel()

    def markeSetzen(self):
        self.w.setMarke(self.x, self.y)

    def markeLoeschen(self):
        self.w.delMarke(self.x, self.y)

    def aufMarke(self):
        if self.w.getMarke(self.x, self.y) == 0:
            return False
        else:
            return True

    def nichtAufMarke(self):
        return not self.aufMarke()


# Deklaration der Klasse Welt

class Welt(object):
    def __init__(self, x, y):
        self.felderX = x
        self.felderY = y
        l = []
        for i in range(self.felderY):
            m = []
            for j in range(self.felderX):
                m = m + [0]
            l = l + [m]
        self.ziegel = l
        l = []
        for i in range(self.felderY):
            m = []
            for j in range(self.felderX):
                m = m + [False]
            l = l + [m]
        self.marken = l
        
    def getFelderX(self):
        return self.felderX

    def getFelderY(self):
        return self.felderY

    def incZiegel(self, x, y):
        self.ziegel[y][x] = self.ziegel[y][x] + 1

    def decZiegel(self, x, y):
        if self.ziegel[y][x] > 0:
            self.ziegel[y][x] = self.ziegel[y][x] - 1

    def getZiegel(self, x, y):
        return self.ziegel[y][x]

    def getAlleZiegel(self):
        return self.ziegel    
    
    def setMarke(self, x, y):
        self.marken[y][x] = True

    def delMarke(self, x, y):
        self.marken[y][x] = False

    def getMarke(self, x, y):
        return self.marken[y][x]

    def getAlleMarken(self):
        return self.marken
