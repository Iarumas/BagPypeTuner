import numpy as np
from classNote import Note


Noten=[]

Noten.append(Note("LG",7./8.))
Noten.append(Note("LA",1.))
Noten.append(Note("B",9./8.))
Noten.append(Note("C",5./4.))
Noten.append(Note("D",4./3.))
Noten.append(Note("E",3./2.))
Noten.append(Note("F",5./3.))
Noten.append(Note("HG",7./4.))
Noten.append(Note("HA",2.))

print(len(Noten))
i = 0 
while i < len(Noten):
    print (Noten[i].name, Noten[i].cent)
    i+=1

print("")

for el in Noten:
    print (el.name, el.cent)
    
  
