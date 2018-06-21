import pandas as pd

a = pd.DataFrame({"Mat": [1,4,6,7,3], 2: [1,4,5,5,5], 6: [2,4,5,5,0]})
b = pd.DataFrame({"Mat": [1,4,10,5,7], 2: [11,4,5,5,65], 4: [12,41,51,51,11]})

c = a.merge(b,on=["Mat",2],how='outer')
print(a)
print(b)
print(c)